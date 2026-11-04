'use strict';

/**
 * Client fixtures: every folder in test/fixtures/ holding an expected.json is
 * replayed through the real pipeline (Tally books -> OCR'd agreements -> records
 * -> styled workbook) and compared with what the CA verified.
 *
 *   test/fixtures/<client>/
 *     <input files>      Tally .xlsx books, agreement .pdf files
 *     expected.json      { client, inputs, expected }
 *
 * expected.json
 *   inputs.tally[]            Tally files (all sheets of each are read)
 *   inputs.agreements[]       agreement PDFs (OCR)
 *   inputs.agreementStubs[]   already-parsed agreements (skips OCR)
 *   inputs.overrides{}        overrides.json content for this client
 *   expected.files[]          per Tally file: { file, receipts, total }
 *   expected.tallyTotal       sum of every receipt read
 *   expected.unallocated      { count, total }
 *   expected.reconciliation   "MATCHED" (flat + unallocated = Tally total)
 *   expected.sheets[]         worksheet names of the downloaded workbook
 *   expected.warning          true when the workbook starts with CHECK BEFORE USE
 *   expected.flats{}          per flat: name, totals, payments [date, amount, gst, vch]
 */

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const ExcelJS = require('exceljs');
const { buildRecords } = require('../src/services/build-records');
const { writeStyledCollectionWorkbook, reconcile } = require('../src/output/workbook-styled');

const FIXTURES = path.join(__dirname, 'fixtures');
const round2 = (n) => Math.round(n * 100) / 100;
const ddmmyyyy = (d) => (d instanceof Date && !Number.isNaN(d.getTime())
  ? `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}` : '');

/** Run one client folder through the pipeline and describe the result. */
async function runClient(dir) {
  const spec = JSON.parse(fs.readFileSync(path.join(dir, 'expected.json'), 'utf8'));
  const file = (name) => ({ path: path.join(dir, name), originalname: name });
  const inputs = spec.inputs;
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'client-'));
  const out = path.join(workDir, 'collection.xlsx');
  try {
    const result = await buildRecords({
      tallyFiles: inputs.tally.map(file),
      agreementFiles: (inputs.agreements || []).map(file),
      agreementObjects: JSON.parse(JSON.stringify(inputs.agreementStubs || [])),
      overrides: inputs.overrides || {},
      workDir: path.join(workDir, 'pages'),
    });
    const opts = {
      projectName: result.tally.projectName || undefined,
      unallocated: result.unallocated,
      reconciliation: { tallyTotal: result.tallyTotal, files: result.tallyFiles },
    };
    await writeStyledCollectionWorkbook(result.records, out, opts);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);

    const rec = reconcile(result.records, opts);
    const flats = {};
    for (const r of result.records) {
      flats[r.flatNo] = {
        name: r.name,
        totalReceived: r.totalReceived,
        totalGst: r.totalGst,
        balance: r.balance,
        agreementValue: r.agreementValue,
        agreementDate: ddmmyyyy(r.agreementDateObj),
        carpetArea: r.carpetArea,
        payments: r.payments.map((p) => [p.dateText, p.amount, p.gst, p.vchNo]),
      };
    }
    return {
      spec,
      actual: {
        files: result.tallyFiles.map((f) => ({ file: f.file, receipts: f.receipts, total: f.total })),
        tallyTotal: result.tallyTotal,
        unallocated: {
          count: result.unallocated.length,
          total: round2(result.unallocated.reduce((s, u) => s + u.amount, 0)),
        },
        reconciliation: rec.ok ? 'MATCHED' : `DIFFERENCE ${rec.diff}`,
        sheets: wb.worksheets.map((w) => w.name),
        warning: String(wb.getWorksheet('Summary').getRow(1).getCell(1).value).startsWith('CHECK BEFORE USE'),
        flats,
      },
    };
  } finally {
    fs.rmSync(workDir, { recursive: true, force: true });
  }
}

function clientDirs() {
  return fs.readdirSync(FIXTURES)
    .map((n) => path.join(FIXTURES, n))
    .filter((d) => fs.statSync(d).isDirectory() && fs.existsSync(path.join(d, 'expected.json')))
    .sort();
}

async function main() {
  let passed = 0;
  let failed = 0;
  const only = process.argv[2];
  for (const dir of clientDirs()) {
    const name = path.basename(dir);
    if (only && only !== name) continue;
    console.log(`\n${name}`);
    let run;
    try {
      run = await runClient(dir);
    } catch (e) {
      failed += 1;
      console.log(`  FAIL- could not run: ${e.stack || e.message}`);
      continue;
    }
    const { spec, actual } = run;
    const exp = spec.expected;
    const check = (label, fn) => {
      try { fn(); passed += 1; console.log(`  ok  - ${label}`); } catch (e) {
        failed += 1; console.log(`  FAIL- ${label}\n       ${e.message.split('\n').slice(0, 14).join('\n       ')}`);
      }
    };
    check(`each Tally file: receipts and total (${exp.files.length} file(s))`, () => assert.deepStrictEqual(actual.files, exp.files));
    check(`Tally total ${exp.tallyTotal}`, () => assert.strictEqual(actual.tallyTotal, exp.tallyTotal));
    check(`Unallocated: ${exp.unallocated.count} receipt(s), total ${exp.unallocated.total}`, () => assert.deepStrictEqual(actual.unallocated, exp.unallocated));
    check('flat receipts + Unallocated = Tally total (MATCHED)', () => assert.strictEqual(actual.reconciliation, exp.reconciliation));
    check(`workbook sheets: ${exp.sheets.join(', ')}`, () => assert.deepStrictEqual(actual.sheets, exp.sheets));
    check(`workbook warning row ${exp.warning ? 'shown' : 'absent'}`, () => assert.strictEqual(actual.warning, exp.warning));
    check(`${Object.keys(exp.flats).length} flat(s) identical (names, totals, GST, every payment)`, () => {
      assert.deepStrictEqual(Object.keys(actual.flats), Object.keys(exp.flats));
      // only the fields written in expected.json are compared (the CA's own figures
      // for a client list just what was checked; Space Home lists every field)
      for (const k of Object.keys(exp.flats)) {
        for (const field of Object.keys(exp.flats[k])) {
          assert.deepStrictEqual(actual.flats[k][field], exp.flats[k][field], `flat ${k}, ${field}`);
        }
      }
    });
  }
  console.log(`\n${passed} client checks passed${failed ? `, ${failed} FAILED` : ''}`);
  process.exit(failed ? 1 : 0);
}

if (require.main === module) main();
module.exports = { runClient, clientDirs, FIXTURES };
