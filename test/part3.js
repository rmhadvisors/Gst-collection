'use strict';

/**
 * Matching, overrides (allocate / renumber), GST receipt + pre-agreement GST and
 * the workbook warning row / manual-receipts line. No OCR - runs in seconds.
 */

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const ExcelJS = require('exceljs');
const { mergeTallies, applyRenumber, applyAllocations, totalReceipts } = require('../src/tally/parse');
const { prepareTally } = require('../src/services/build-records');
const { buildRecord } = require('../src/engine/collection');
const { buildTallyOnlyRecords } = require('../src/engine/tallyOnly');
const { buildUnallocated } = require('../src/engine/unallocated');
const { writeStyledCollectionWorkbook } = require('../src/output/workbook-styled');

let passed = 0;
let failed = 0;
async function check(name, fn) {
  try { await fn(); passed += 1; console.log(`  ok  - ${name}`); } catch (e) {
    failed += 1; console.log(`  FAIL- ${name}\n       ${e.message}`);
  }
}

const S = (f) => path.join(__dirname, '..', 'config', 'samples', f);
const row = (o) => ({
  date: new Date(2026, 0, 5), vchNo: '1', amount: 100000, name: '', particulars: '', file: 'a.xlsx', ...o,
});
const agreement = (o) => ({
  flatNoHint: null, flatSource: 'filename', primaryName: null, agreementValue: 4490000,
  agreementDateText: '', valueCandidates: [], names: [], ...o,
});

(async () => {
  console.log('1. Matching: flat number first, then transliterated name');
  const people = mergeTallies([{
    projectName: 'P',
    rows: [row({ flatNo: '403', name: 'PRADNYA AJIT PRABHUKAR', particulars: '403 PRADNYA AJIT PRABHUKAR' })],
  }]);
  const marathi = 'प्रज्ञा अजित प्रभुकरवतकर';
  await check('Marathi buyer name matches the English Tally name when the flat is unknown', () => {
    const rec = buildRecord(agreement({ flatNoHint: '999', primaryName: marathi, names: [marathi] }), people, {});
    assert.strictEqual(rec.matchReason, 'name');
    assert.strictEqual(rec.totalReceived, 100000);
    assert.strictEqual(rec.name, 'PRADNYA AJIT PRABHUKAR'); // English on the sheet
    assert.strictEqual(rec.purchaserMarathi, marathi); // original kept for Audit
    assert.strictEqual(rec.flatNo, '999'); // the agreement's own flat is not relabelled
  });
  await check('flat number wins over a different name', () => {
    const rec = buildRecord(agreement({ flatNoHint: '403', primaryName: 'SOMEONE ELSE ENTIRELY' }), people, {});
    assert.strictEqual(rec.matchReason, 'filename');
    assert.strictEqual(rec.totalReceived, 100000);
  });
  await check('an unrelated Marathi name does not match', () => {
    const rec = buildRecord(agreement({ flatNoHint: '999', primaryName: 'सुरेश कुमार पाटील' }), people, {});
    assert.strictEqual(rec.matchReason, 'unmatched');
    assert.strictEqual(rec.totalReceived, 0);
  });
  await check('Capital Ac ledger: flagged (review), receipts never auto-allocated', () => {
    const t = mergeTallies([{
      projectName: 'P',
      rows: [row({ flatNo: null, isCapital: true, name: 'Pradnya Ajit Prabhukar Capital Ac', particulars: 'Pradnya Ajit Prabhukar Capital Ac' })],
    }]);
    const rec = buildRecord(agreement({ flatNoHint: '777', primaryName: marathi }), t, {});
    assert.ok(/\(review\)/.test(rec.matchReason), rec.matchReason);
    assert.ok(rec.needsReview);
    assert.strictEqual(rec.totalReceived, 0);
    assert.ok(rec.flags.some((f) => /\(review\)/.test(f)));
    assert.strictEqual(t.unallocated.length, 1); // still in Unallocated
  });

  console.log('\n2. overrides.json: renumber and allocate');
  await check('renumber moves Tally receipts old -> new flat and flags them', () => {
    const t = mergeTallies([{
      projectName: 'P',
      rows: [row({ flatNo: '403', vchNo: '1' }), row({ flatNo: '403', vchNo: '2', amount: 50000 }), row({ flatNo: '404', vchNo: '3' })],
    }]);
    applyRenumber(t, { _comment: 'x', 403: '503' });
    assert.ok(!t.byFlat.has('403') && t.byFlat.has('503') && t.byFlat.has('404'));
    assert.strictEqual(t.byFlat.get('503').payments.length, 2);
    assert.deepStrictEqual(t.renumbered, [{ from: '403', to: '503', receipts: 2 }]);
    const rec = buildTallyOnlyRecords(t, [], {}, {}).find((r) => r.flatNo === '503');
    assert.ok(rec.flags.some((f) => /Renumbered.*403/.test(f)), rec.flags.join('|'));
    assert.strictEqual(totalReceipts(t), 250000); // nothing dropped
  });
  await check('renumber can swap two flats in one pass', () => {
    const t = mergeTallies([{
      projectName: 'P', rows: [row({ flatNo: '403', amount: 1 }), row({ flatNo: '404', vchNo: '2', amount: 20 })],
    }]);
    applyRenumber(t, { 403: '404', 404: '403' });
    assert.strictEqual(t.byFlat.get('404').payments[0].amount, 1);
    assert.strictEqual(t.byFlat.get('403').payments[0].amount, 20);
  });
  await check('allocate moves an unallocated receipt onto a flat; an unknown voucher is reported', () => {
    const t = mergeTallies([{
      projectName: 'P',
      rows: [row({ flatNo: '403', vchNo: '1' }), row({ flatNo: null, vchNo: '81', amount: 6200, name: 'X Capital Ac', particulars: 'X Capital Ac' })],
    }]);
    applyAllocations(t, { 81: '403', 99: '403' });
    assert.strictEqual(t.byFlat.get('403').payments.length, 2);
    assert.strictEqual(t.unallocated.length, 0);
    assert.deepStrictEqual(t.manualAllocations.map((m) => [m.vchNo, m.found]), [['81', true], ['99', false]]);
  });
  await check('prepareTally: allocate target written with the OLD flat number follows the renumber', () => {
    const base = prepareTally([{ path: S('mahira-federal-bank.xlsx') }], {});
    const flat = [...base.byFlat.keys()][0];
    const unalloc = base.unallocated[0];
    assert.ok(unalloc, 'sample has unallocated receipts');
    const t = prepareTally([{ path: S('mahira-federal-bank.xlsx') }], {
      renumber: { [flat]: 'A-4242' }, allocate: { [unalloc.vchNo]: flat },
    });
    assert.ok(t.byFlat.has('A-4242') && !t.byFlat.has(flat));
    assert.ok(t.byFlat.get('A-4242').payments.some((p) => p.manualAllocation && p.vchNo === unalloc.vchNo));
    assert.strictEqual(totalReceipts(t), totalReceipts(base));
  });

  console.log('\n3 + 4. GST receipt and advances before the agreement date');
  const gstTally = mergeTallies([{
    projectName: 'P',
    rows: [
      row({ flatNo: '1804', vchNo: '1', date: new Date(2025, 11, 1), amount: 100000 }), // advance before agreement
      row({ flatNo: '1804', vchNo: '2', date: new Date(2026, 0, 5), amount: 44900 }), // = GST on 44,90,000
      row({ flatNo: '1804', vchNo: '3', date: new Date(2026, 3, 1), amount: 200000 }), // after agreement
    ],
  }]);
  const gstRec = buildRecord(agreement({ flatNoHint: '1804', agreementDateText: '3/10/26' }), gstTally, {});
  await check('receipt equal to the GST on the agreement value is a "GST receipt" with no GST', () => {
    const p = gstRec.payments.find((x) => x.amount === 44900);
    assert.strictEqual(p.gst, 0);
    assert.ok(p.gstReceipt);
    assert.ok(gstRec.flags.some((f) => /GST receipt/.test(f)));
    assert.strictEqual(gstRec.payments.find((x) => x.amount === 200000).gst, 2000);
  });
  await check('pre-agreement advances keep the Case 2 logic (GST receipt is not an advance)', () => {
    assert.strictEqual(gstRec.gstCase, 2);
    assert.strictEqual(gstRec.advancesBeforeAgreement, 100000);
    assert.strictEqual(gstRec.reverseGstOnAdvances, 1000);
    assert.strictEqual(gstRec.outstandingGst, 43900);
    assert.strictEqual(gstRec.preAgreementPayments.length, 2);
    assert.strictEqual(gstRec.postAgreementPayments.length, 1);
  });

  console.log('\n5. Warning row at the top of the workbook');
  const mahira = prepareTally([{ path: S('mahira-bank-of-maharashtra.xlsx'), originalname: 'bom.xlsx' }, { path: S('mahira-federal-bank.xlsx'), originalname: 'fed.xlsx' }], {});
  const readSheet = async (records, opts) => {
    const out = path.join(os.tmpdir(), `part3-${Date.now()}-${Math.round(Math.random() * 1e4)}.xlsx`);
    await writeStyledCollectionWorkbook(records, out, opts);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    fs.unlinkSync(out);
    return wb;
  };
  const cell = (wb, sheet) => String(wb.getWorksheet(sheet).getRow(1).getCell(1).value);
  await check('no warning row when everything is MATCHED and nothing needs a check', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const wb = await readSheet(records, {
      projectName: 'P', unallocated: buildUnallocated(mahira, records),
      reconciliation: { tallyTotal: totalReceipts(mahira), files: mahira.fileStats },
    });
    assert.ok(cell(wb, 'Summary').startsWith('PROJECT'), cell(wb, 'Summary'));
    assert.ok(cell(wb, 'Audit').startsWith('Flat No'), cell(wb, 'Audit'));
  });
  await check('warning row for OCR FAILED, (review), check manually and a reconciliation difference', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const extra = (flatNo, o) => ({ ...records[0], flatNo, hasAgreement: true, payments: [], flags: [], ...o });
    records.push(extra('Z-1', { ocrFailed: true }));
    records.push(extra('Z-2', { needsReview: true, matchReason: 'name-capital-ac (review)', agreementValue: 1 }));
    records.push(extra('Z-3', { agreementValue: null }));
    const wb = await readSheet(records, {
      projectName: 'P', unallocated: [], reconciliation: { tallyTotal: totalReceipts(mahira) + 5, files: mahira.fileStats },
    });
    for (const sheet of ['Summary', 'Audit']) {
      const t = cell(wb, sheet);
      assert.ok(t.startsWith('CHECK BEFORE USE'), t);
      assert.ok(/OCR FAILED on flat\(s\) Z-1/.test(t), t);
      assert.ok(/\(review\) on flat\(s\) Z-2/.test(t), t);
      assert.ok(/check manually: flat\(s\) [^|]*Z-3/.test(t), t);
      assert.ok(/NOT MATCHED/.test(t), t);
      assert.strictEqual(wb.getWorksheet(sheet).getRow(1).getCell(1).fill.fgColor.argb, 'FFC00000');
    }
  });
  await check('manual receipts from overrides.json get their own Audit line; Tally total stays the Tally books', async () => {
    const overrides = { 'B-607': { extraPayments: [{ date: '3/4/26', amount: 44900 }, { date: '4/25/26', amount: 140699 }] } };
    const records = buildTallyOnlyRecords(mahira, [], overrides, {});
    const wb = await readSheet(records, {
      projectName: 'P', unallocated: buildUnallocated(mahira, records),
      reconciliation: { tallyTotal: totalReceipts(mahira), files: mahira.fileStats },
    });
    const rows = [];
    wb.getWorksheet('Audit').eachRow((r) => rows.push([r.getCell(1).value, r.getCell(4).value]));
    const line = (re) => rows.find(([label]) => re.test(String(label)));
    assert.strictEqual(line(/^Total Tally receipts/)[1], totalReceipts(mahira)); // the real Tally books
    assert.strictEqual(line(/^Manual receipts \(not in Tally\)/)[1], 185599);
    assert.strictEqual(line(/Flat B-607 {2}4\/25\/26/)[1], 140699);
    assert.strictEqual(rows[rows.length - 1][1], 'MATCHED'); // manual receipts do not break the check
    assert.ok(!cell(wb, 'Summary').startsWith('CHECK BEFORE USE'));
  });
  await check('no manual-receipts line when there are none', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const wb = await readSheet(records, {
      projectName: 'P', unallocated: buildUnallocated(mahira, records),
      reconciliation: { tallyTotal: totalReceipts(mahira), files: mahira.fileStats },
    });
    let found = false;
    wb.getWorksheet('Audit').eachRow((r) => { if (/^Manual receipts/.test(String(r.getCell(1).value))) found = true; });
    assert.ok(!found);
  });
  await check('a Tally file that gave nothing is called out in the warning', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const wb = await readSheet(records, {
      projectName: 'P', unallocated: buildUnallocated(mahira, records),
      reconciliation: { tallyTotal: totalReceipts(mahira), files: [...mahira.fileStats, { file: 'empty.xlsx', receipts: 0, total: 0 }] },
    });
    assert.ok(/NOT READ: empty\.xlsx/.test(cell(wb, 'Summary')), cell(wb, 'Summary'));
  });

  console.log(`\n${passed} part-3 checks passed${failed ? `, ${failed} FAILED` : ''}`);
  process.exit(failed ? 1 : 0);
})();
