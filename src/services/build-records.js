'use strict';

const fs = require('fs');
const path = require('path');
const { processAgreements } = require('../pdf/agreement');
const { closeOcr } = require('../pdf/ocr');
const {
  parseTally, mergeTallies, applyAllocations, applyRenumber, assignBuildings, totalReceipts,
} = require('../tally/parse');
const { buildRecord } = require('../engine/collection');
const { buildTallyOnlyRecords } = require('../engine/tallyOnly');
const { buildUnallocated } = require('../engine/unallocated');
const { buildNoFlatRecords, partnerGstSetting } = require('../engine/noFlat');
const {
  compareFlatKeys, normalizeFlatKey, renumberMap, groupKey,
} = require('../util/flatKey');
const { canonicalBuilding } = require('../util/building');

/** config/overrides.json (or `file`); a missing file is {}, a broken one is an error. */
function loadOverrides(file) {
  const target = file || path.join(__dirname, '..', '..', 'config', 'overrides.json');
  if (!fs.existsSync(target)) return {};
  try {
    return JSON.parse(fs.readFileSync(target, 'utf8'));
  } catch (e) {
    throw new Error(`Could not parse overrides ${target}: ${e.message}`);
  }
}

/** Read every Tally file (all sheets), merge them and apply "renumber" (old -> new flat). */
function readTally(tallyFiles, overrides = {}) {
  const tallies = tallyFiles.map((tf) => parseTally(tf.path, undefined, tf.originalname));
  return applyRenumber(mergeTallies(tallies), overrides.renumber);
}

/** "allocate" targets may be written with the old flat number: follow the renumber. */
function allocationsOf(overrides) {
  const map = renumberMap(overrides.renumber);
  const allocate = {};
  for (const [vch, target] of Object.entries(overrides.allocate || {})) {
    const key = normalizeFlatKey(target);
    allocate[vch] = (key && map.get(key)) || target;
  }
  return allocate;
}

/**
 * Tally books -> grouped receipts, with renumber + allocate applied. No building
 * handling (flats are keyed by number alone): used by the "update existing sheet"
 * flows. The full pipeline below also assigns buildings.
 */
function prepareTally(tallyFiles, overrides = {}) {
  return applyAllocations(readTally(tallyFiles, overrides), allocationsOf(overrides));
}

/** Per-flat override: "Building|flat" first, then the flat (new number, then the old one). */
function overrideFor(overrides, flatNo, oldFlatNo, building) {
  return (building && overrides[groupKey(building, flatNo)]) || overrides[flatNo] || overrides[oldFlatNo] || {};
}

/**
 * The whole pipeline, shared by the web server and the CLI: Tally books + OCR'd
 * agreements (PDF or photo) -> sorted flat records + one block per party without a flat no., and the
 * receipt totals. Agreements may be empty (tally-only run). Progress goes to
 * `onProgress(message)`. Uploaded files are NOT deleted here.
 *
 * Agreements are read BEFORE receipts are assigned to flats, because the building
 * an agreement names decides whether two "403"s are the same flat.
 *
 * @param {object} opts
 * @param {{path:string, originalname?:string}[]} opts.tallyFiles
 * @param {{path:string, originalname?:string}[]} [opts.agreementFiles]  .pdf / .jpg / .jpeg / .png
 * @param {object[]} [opts.agreementObjects]  agreements that are already parsed (skip OCR)
 * @param {object} [opts.overrides]
 * @param {string} opts.workDir          scratch folder for extracted page images
 * @param {(msg:string)=>void} [opts.onProgress]
 */
async function buildRecords(opts) {
  const {
    tallyFiles, agreementFiles = [], agreementObjects = [], overrides = {}, workDir, onProgress = () => {},
  } = opts;
  const renumber = renumberMap(overrides.renumber);
  const aliases = overrides.buildings || {};

  onProgress(`Reading ${tallyFiles.length} Tally file(s)`);
  const tally = readTally(tallyFiles, overrides);

  const agreements = [];
  const addAgreement = (agreement) => {
    const oldHint = agreement.flatNoHint;
    const hint = oldHint && renumber.get(normalizeFlatKey(oldHint));
    agreement.oldFlatNoHint = oldHint;
    if (hint) agreement.flatNoHint = hint;
    if (agreement.buildingRaw && agreement.building == null) agreement.building = canonicalBuilding(agreement.buildingRaw, aliases);
    agreements.push(agreement);
  };
  try {
    let n = 0;
    for (const af of agreementFiles) {
      n += 1;
      const label = af.originalname || path.basename(af.path);
      onProgress(`Reading agreement ${n} of ${agreementFiles.length}: ${label}`);
      // One PDF may hold several Index-II pages for different flats.
      for (const agreement of await processAgreements(af.path, { workDir })) addAgreement(agreement);
    }
    // already-parsed agreements (tests, or a CA typing the values in)
    for (const agreement of agreementObjects) addAgreement(agreement);
  } finally {
    await closeOcr();
  }

  onProgress('Matching receipts to flats');
  const named = agreements.map((a) => a.building).filter(Boolean);
  assignBuildings(tally, {
    names: named,
    agreementBuildings: named,
    buildings: aliases,
    books: overrides.books,
    agreements: agreements.map((a) => ({
      flatNo: a.flatNoHint, building: a.building, agreementDate: a.agreementDate, names: a.names, name: a.primaryName,
    })),
  });
  applyAllocations(tally, allocationsOf(overrides));

  const records = [];
  for (const agreement of agreements) {
    const building = agreement.building || '';
    const rec = buildRecord(agreement, tally, overrideFor(overrides, agreement.flatNoHint, agreement.oldFlatNoHint, building));
    rec.pdf = agreement.pdf;
    if (agreement.flatNoHint !== agreement.oldFlatNoHint && agreement.oldFlatNoHint) {
      rec.flags.push(`Renumbered: agreement names flat ${agreement.oldFlatNoHint}, now ${agreement.flatNoHint}`);
    }
    records.push(rec);
  }

  onProgress('Adding flats that exist only in Tally');
  const matched = new Set(records.map((r) => r.tallyKey).filter(Boolean));
  records.push(...buildTallyOnlyRecords(tally, matched, overrides, {}));
  records.sort((a, b) => (a.building || '').localeCompare(b.building || '') || compareFlatKeys(a.flatNo, b.flatNo));

  // Receipts without a flat number: one block per party, printed below the flats on the main sheet.
  const partners = [...new Set([
    ...(Array.isArray(overrides.partners) ? overrides.partners : []),
    ...agreements.flatMap((a) => a.partners || []),
  ])];
  const noFlat = buildNoFlatRecords(tally, { partners, partnerGst: partnerGstSetting(overrides), records });
  records.push(...noFlat);

  return {
    records,
    noFlat,
    tally,
    unallocated: buildUnallocated(tally, records),
    tallyTotal: totalReceipts(tally),
    tallyFiles: tally.fileStats,
  };
}

module.exports = {
  buildRecords, prepareTally, readTally, loadOverrides, overrideFor,
};
