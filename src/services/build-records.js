'use strict';

const fs = require('fs');
const path = require('path');
const { processAgreements } = require('../pdf/agreement');
const { closeOcr } = require('../pdf/ocr');
const {
  parseTally, mergeTallies, applyAllocations, applyRenumber, totalReceipts,
} = require('../tally/parse');
const { buildRecord } = require('../engine/collection');
const { buildTallyOnlyRecords } = require('../engine/tallyOnly');
const { buildUnallocated } = require('../engine/unallocated');
const { compareFlatKeys, normalizeFlatKey, renumberMap } = require('../util/flatKey');

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

/**
 * Read every Tally file (all sheets), merge them, then apply the overrides
 * "renumber" (old -> new flat) and "allocate" (vch no -> flat).
 *
 * @param {{path:string, originalname?:string}[]} tallyFiles
 */
function prepareTally(tallyFiles, overrides = {}) {
  const tallies = tallyFiles.map((tf) => parseTally(tf.path, undefined, tf.originalname));
  const tally = applyRenumber(mergeTallies(tallies), overrides.renumber);

  // an allocation target may be written with the old flat number
  const map = renumberMap(overrides.renumber);
  const allocate = {};
  for (const [vch, target] of Object.entries(overrides.allocate || {})) {
    const key = normalizeFlatKey(target);
    allocate[vch] = (key && map.get(key)) || target;
  }
  return applyAllocations(tally, allocate);
}

/** Per-flat override: keyed by the new flat number, falling back to the old one. */
function overrideFor(overrides, flatNo, oldFlatNo) {
  return overrides[flatNo] || overrides[oldFlatNo] || {};
}

/**
 * The whole pipeline, shared by the web server and the CLI: Tally books + OCR'd
 * agreements -> sorted flat records, the Unallocated list and the receipt totals.
 * Agreements may be empty (tally-only run). Progress is reported through
 * `onProgress(message)`. Uploaded files are NOT deleted here.
 *
 * @param {object} opts
 * @param {{path:string, originalname?:string}[]} opts.tallyFiles
 * @param {{path:string, originalname?:string}[]} [opts.agreementFiles]
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

  onProgress(`Reading ${tallyFiles.length} Tally file(s)`);
  const tally = prepareTally(tallyFiles, overrides);
  const records = [];

  const addAgreement = (agreement) => {
    const oldHint = agreement.flatNoHint;
    const hint = oldHint && renumber.get(normalizeFlatKey(oldHint));
    if (hint) agreement.flatNoHint = hint;
    const rec = buildRecord(agreement, tally, overrideFor(overrides, agreement.flatNoHint, oldHint));
    rec.pdf = agreement.pdf;
    if (hint) rec.flags.push(`Renumbered: agreement names flat ${oldHint}, now ${hint}`);
    records.push(rec);
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

  onProgress('Adding flats that exist only in Tally');
  const matched = new Set(records.flatMap((r) => [r.flatNo, r.tallyFlat].filter(Boolean)));
  records.push(...buildTallyOnlyRecords(tally, matched, overrides, {}));
  records.sort((a, b) => compareFlatKeys(a.flatNo, b.flatNo));

  return {
    records,
    tally,
    unallocated: buildUnallocated(tally, records),
    tallyTotal: totalReceipts(tally),
    tallyFiles: tally.fileStats,
  };
}

module.exports = {
  buildRecords, prepareTally, loadOverrides, overrideFor,
};
