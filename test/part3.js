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
const { buildNoFlatRecords } = require('../src/engine/noFlat');
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
    const wb = await readSheet([...records, ...buildNoFlatRecords(mahira, { records })], {
      projectName: 'P',
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
      projectName: 'P', reconciliation: { tallyTotal: totalReceipts(mahira) + 5, files: mahira.fileStats },
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
    const wb = await readSheet([...records, ...buildNoFlatRecords(mahira, { records })], {
      projectName: 'P',
      reconciliation: { tallyTotal: totalReceipts(mahira), files: mahira.fileStats },
    });
    const rows = [];
    wb.getWorksheet('Audit').eachRow((r) => rows.push([r.getCell(1).value, r.getCell(4).value]));
    const line = (re) => rows.find(([label]) => re.test(String(label)));
    assert.strictEqual(line(/^Total Tally receipts/)[1], totalReceipts(mahira)); // the real Tally books
    assert.strictEqual(line(/^Manual receipts \(not in Tally\)/)[1], 185599);
    assert.strictEqual(line(/Flat B-607 {2}25\/04\/2026/)[1], 140699); // dd/mm/yyyy
    assert.strictEqual(rows[rows.length - 1][1], 'MATCHED'); // manual receipts do not break the check
    assert.ok(!cell(wb, 'Summary').startsWith('CHECK BEFORE USE'));
  });
  await check('no manual-receipts line when there are none', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const wb = await readSheet([...records, ...buildNoFlatRecords(mahira, { records })], {
      projectName: 'P',
      reconciliation: { tallyTotal: totalReceipts(mahira), files: mahira.fileStats },
    });
    let found = false;
    wb.getWorksheet('Audit').eachRow((r) => { if (/^Manual receipts/.test(String(r.getCell(1).value))) found = true; });
    assert.ok(!found);
  });
  await check('a Tally file that gave nothing is called out in the warning', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const wb = await readSheet([...records, ...buildNoFlatRecords(mahira, { records })], {
      projectName: 'P',
      reconciliation: { tallyTotal: totalReceipts(mahira), files: [...mahira.fileStats, { file: 'empty.xlsx', receipts: 0, total: 0 }] },
    });
    assert.ok(/NOT READ: empty\.xlsx/.test(cell(wb, 'Summary')), cell(wb, 'Summary'));
  });

  await check('Unallocated: "Sashikant Thakur Capital Ac" gets the best flat guess B-603, flagged (review), and is not allocated', () => {
    const t = mergeTallies([{
      projectName: 'P',
      rows: [
        row({ flatNo: 'B-607', name: 'RAJESH SINGH', vchNo: '5', amount: 10 }),
        row({ flatNo: null, isCapital: true, name: 'Sashikant Thakur Capital Ac', particulars: 'Sashikant Thakur Capital Ac', vchNo: '81', amount: 6200 }),
      ],
    }]);
    const records = [{ flatNo: 'B-603', name: 'Shashikant B. Thakur (HUF)', allottees: ['Shashikant B. Thakur (HUF)'] }];
    const u = buildUnallocated(t, records);
    assert.strictEqual(u.length, 1);
    assert.strictEqual(u[0].guess, 'B-603');
    assert.ok(/\(review\)/.test(u[0].flag), u[0].flag);
    assert.strictEqual(t.byFlat.has('B-603'), false); // nothing was allocated
    assert.strictEqual(totalReceipts(t), 6210);
  });

  console.log('\n7. Alternate accommodation, market value, shops');
  const noTally = mergeTallies([{ projectName: 'P', rows: [] }]);
  const alt = (flat, market) => buildRecord(agreement({
    flatNoHint: flat, docType: 'alternate', agreementValue: null, marketValue: market, building: 'Maitri Heights',
  }), noTally, {});
  await check('पर्यायी जागेचा करार: value 0, GST on the market value (602 = 5,625; 603 = 10,540; 604 = 8,190)', () => {
    for (const [flat, market, gst] of [['602', 562500, 5625], ['603', 1054000, 10540], ['604', 819000, 8190]]) {
      const rec = alt(flat, market);
      assert.strictEqual(rec.agreementValue, 0, flat);
      assert.strictEqual(rec.gstOnAgreement, gst, flat);
      assert.strictEqual(rec.gstRate, 0.01, flat);
      assert.ok(rec.flags.includes('alternate accommodation – GST on market value'), rec.flags.join('|'));
      assert.ok(!rec.needsManual, flat);
    }
  });
  await check('market value 1 (601) or unreadable: no GST is calculated and the flat says "check manually"', () => {
    for (const market of [1, 0, null]) {
      const rec = alt('601', market);
      assert.strictEqual(rec.gstOnAgreement, null);
      assert.ok(rec.needsManual);
      assert.ok(rec.flags.some((f) => /^check manually/.test(f)), rec.flags.join('|'));
      assert.strictEqual(rec.valueSource, 'manual needed');
    }
  });
  await check('an alternate-accommodation receipt equal to the GST is a GST receipt, balance 0', () => {
    const t = mergeTallies([{ projectName: 'P', rows: [row({ flatNo: '602', amount: 5625, vchNo: '9' })] }]);
    const rec = buildRecord(agreement({
      flatNoHint: '602', docType: 'alternate', agreementValue: null, marketValue: 562500, building: 'Maitri Heights',
    }), t, {});
    assert.strictEqual(rec.gstReceived, 5625);
    assert.strictEqual(rec.gstOutstanding, 0);
    assert.strictEqual(rec.balance, 0);
  });
  await check('a shop is taxed at 5% with "commercial – confirm rate" (from the agreement and from Tally alone)', () => {
    const rec = buildRecord(agreement({ flatNoHint: 'Shop-1', agreementValue: 6000000, commercial: true }), noTally, {});
    assert.strictEqual(rec.gstRate, 0.05);
    assert.strictEqual(rec.gstOnAgreement, 300000);
    assert.ok(rec.flags.includes('commercial – confirm rate'));
    const t = mergeTallies([{ projectName: 'P', rows: [row({ flatNo: 'Shop-1', amount: 100000 })] }]);
    const only = buildTallyOnlyRecords(t, [], {}, {})[0];
    assert.strictEqual(only.gstRate, 0.05);
    assert.strictEqual(only.payments[0].gst, 5000);
    assert.ok(only.flags.includes('commercial – confirm rate'));
  });
  await check('a flat is not charged the shop rate: 44,90,000 agreement stays 1%', () => {
    assert.strictEqual(buildRecord(agreement({ flatNoHint: '403', agreementValue: 4490000 }), noTally, {}).gstRate, 0.01);
  });

  console.log('\n8. Buildings: Maitri Heights 403 and Siddhivinayak Palace 403 never merge');
  const { assignBuildings } = require('../src/tally/parse');
  const cfg = { 'Maitri Heights': ['मैत्री हाईट्स', 'maitri'], 'Siddhivinayak Palace': ['सिद्धिविनायक पॅलेस', 'siddhivinayak'] };
  const book = (file, title, rows) => ({
    projectName: 'P',
    rows: rows.map((r) => row({ file, ...r })),
    fileStats: [{ file, title, receipts: rows.length, total: rows.reduce((s, r) => s + r.amount, 0) }],
  });
  const buildingTally = (assign) => {
    const t = mergeTallies([
      book('hdfc 444.xlsx', 'H D F C BANK 44444 Siddhivinayak Palace  Book', [{ flatNo: '403', amount: 100, name: 'SIDDHI BUYER', vchNo: '1' }]),
      book('idfc 5112.xlsx', 'IDFC FIRST BANK- RERA COLLECTION A/C 100% - 5112  Book', [{ flatNo: '403', amount: 200, name: 'MAITRI BUYER', vchNo: '1' }]),
      book('other.xlsx', 'Some bank', [{ flatNo: '403', amount: 50, name: 'UNKNOWN BUYER', vchNo: '1' }]),
    ]);
    return assignBuildings(t, { buildings: cfg, books: assign });
  };
  const tb = buildingTally({ 'IDFC 5112.XLSX': 'Maitri Heights' }); // matched on the file name, any case
  await check('building from the book title (hdfc 444) or from overrides "books"; two flats 403, not one', () => {
    assert.ok(tb.buildingsInUse);
    assert.deepStrictEqual([...tb.byFlat.keys()].sort(), ['Maitri Heights|403', 'Siddhivinayak Palace|403']);
    assert.strictEqual(tb.byFlat.get('Siddhivinayak Palace|403').payments[0].amount, 100);
    assert.strictEqual(tb.byFlat.get('Maitri Heights|403').payments[0].amount, 200);
  });
  await check('each agreement gets only its own building\'s receipts', () => {
    const m = buildRecord(agreement({ flatNoHint: '403', building: 'Maitri Heights' }), tb, {});
    const s = buildRecord(agreement({ flatNoHint: '403', building: 'Siddhivinayak Palace' }), tb, {});
    assert.deepStrictEqual([m.totalReceived, s.totalReceived], [200, 100]);
    assert.deepStrictEqual([m.building, s.building], ['Maitri Heights', 'Siddhivinayak Palace']);
  });
  await check('an agreement whose building is unknown does not guess between two 403s: check manually', () => {
    const rec = buildRecord(agreement({ flatNoHint: '403' }), tb, {});
    assert.strictEqual(rec.totalReceived, 0);
    assert.ok(rec.needsManual);
    assert.ok(rec.flags.some((f) => /building not read/.test(f)));
  });
  await check('a receipt whose book has no building is Unallocated with a (review) guess, never merged', () => {
    assert.strictEqual(tb.unallocated.length, 1);
    assert.strictEqual(tb.unallocated[0].amount, 50);
    assert.strictEqual(tb.unallocated[0].guessFlat, '403');
    const u = buildUnallocated(tb, [])[0];
    assert.strictEqual(u.guess, '403');
    assert.ok(/\(review\)/.test(u.flag), u.flag);
    assert.strictEqual(totalReceipts(tb), 350); // nothing lost
  });
  await check('allocating to a flat that exists in two buildings needs "Building|flat"', () => {
    const t = buildingTally({ 'idfc 5112.xlsx': 'Maitri Heights' });
    applyAllocations(t, { 1: '403' });
    assert.strictEqual(t.manualAllocations[0].found, false);
    assert.ok(/several buildings/.test(t.manualAllocations[0].reason));
    applyAllocations(t, { 1: 'Maitri Heights|403' });
    assert.strictEqual(t.byFlat.get('Maitri Heights|403').payments.length, 2);
  });
  await check('buildings are not used when no book has one and the agreements name just one', () => {
    const t = mergeTallies([book('a.xlsx', 'Some bank', [{ flatNo: '403', amount: 1 }])]);
    assignBuildings(t, { buildings: cfg, names: ['Maitri Heights'], agreementBuildings: ['Maitri Heights'] });
    assert.ok(!t.buildingsInUse);
    assert.ok(t.byFlat.has('403'));
  });
  await check('...but two buildings in the agreements switch it on; a flat found in no known building stays its own building-less flat (not merged)', () => {
    const t = mergeTallies([book('a.xlsx', 'Some bank', [{ flatNo: '403', amount: 1 }])]);
    assignBuildings(t, { buildings: cfg, agreementBuildings: ['Maitri Heights', 'Siddhivinayak Palace'] });
    assert.ok(t.buildingsInUse);
    assert.strictEqual(t.unallocated.length, 0);
    assert.strictEqual(t.byFlat.get('403').building, '');
    assert.strictEqual(t.byFlat.get('403').payments.length, 1);
  });
  await check('workbook: one sheet per building', async () => {
    const records = [
      buildRecord(agreement({ flatNoHint: '403', building: 'Maitri Heights' }), tb, {}),
      buildRecord(agreement({ flatNoHint: '403', building: 'Siddhivinayak Palace' }), tb, {}),
    ];
    const wb = await readSheet(records, { projectName: 'P' });
    assert.deepStrictEqual(wb.worksheets.map((w) => w.name), ['Summary', 'Maitri Heights', 'Siddhivinayak Palace', 'Audit']);
  });
  await check('Marathi building names resolve to the configured English names', () => {
    const { canonicalBuilding } = require('../src/util/building');
    assert.strictEqual(canonicalBuilding('मैत्री हाईट्स', cfg), 'Maitri Heights');
    assert.strictEqual(canonicalBuilding('त्री सिद्धिविनायक पॅलेस को- ऑप.हौ.सो.लि', cfg), 'Siddhivinayak Palace');
  });

  console.log('\n9. Dates are Indian (d/m/y) and real Excel dates; the GST receipt is reported apart');
  await check('a Tally date written 04/01/2025 is 4 January, never 1 April', () => {
    const XLSX = require('xlsx');
    const ws = XLSX.utils.aoa_to_sheet([
      ['Date', 'Particulars', '', '', 'Vch Type', 'Vch No.', 'Debit', 'Credit'],
      ['04/01/2025', 'To', '403 SOMEONE', '', 'Receipt', '1', '1000', ''],
      ['16/08/2024', 'To', '403 SOMEONE', '', 'Receipt', '2', '2000', ''],
    ]);
    const wbk = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbk, ws, 'Anything');
    const file = path.join(os.tmpdir(), `dates-${Date.now()}.xlsx`);
    XLSX.writeFile(wbk, file);
    const t = require('../src/tally/parse').parseTally(file);
    fs.unlinkSync(file);
    const d = t.rows.map((r) => [r.date.getFullYear(), r.date.getMonth() + 1, r.date.getDate()]);
    assert.deepStrictEqual(d, [[2025, 1, 4], [2024, 8, 16]]);
  });
  const base403 = [
    ['16/08/2024', 100000], ['21/08/2024', 439000], ['26/08/2024', 1000000], ['27/12/2024', 500000], ['04/01/2025', 250900], ['23/01/2025', 44900],
  ].map(([dmy, amount], i) => {
    const [d, m, y] = dmy.split('/').map(Number);
    return row({ flatNo: '403', vchNo: String(i + 1), date: new Date(y, m - 1, d), amount });
  });
  const t403 = mergeTallies([{ projectName: 'P', rows: base403 }]);
  const rec403 = buildRecord(agreement({
    flatNoHint: '403', agreementValue: 4490000, agreementDate: new Date(2024, 11, 11), regNo: '4611/2024', marketValue: 3649500,
  }), t403, {});
  await check('GST receipt split: base advances, GST received, GST outstanding, O/S amount', () => {
    assert.strictEqual(rec403.totalReceived, 2334800);
    assert.strictEqual(rec403.baseReceived, 2289900);
    assert.strictEqual(rec403.gstReceived, 44900);
    assert.strictEqual(rec403.gstOutstanding, 0);
    assert.strictEqual(rec403.osAmount, 2200100);
    assert.strictEqual(rec403.balance, 2200100);
    assert.strictEqual(rec403.regNo, '4611/2024');
    assert.strictEqual(rec403.marketValue, 3649500);
  });
  await check('workbook: dates are real dates (dd/mm/yyyy); Summary O/S, GST PAID, Adv Rec, REG. NO., MARKET VALUE filled', async () => {
    const wb = await readSheet([rec403], { projectName: 'P' });
    const flats = wb.getWorksheet('Flats');
    const dates = [];
    flats.eachRow((r) => {
      const c = r.getCell(11);
      if (c.value instanceof Date) dates.push([c.value.getUTCDate(), c.value.getUTCMonth() + 1, c.value.getUTCFullYear(), c.numFmt]);
    });
    assert.deepStrictEqual(dates.map((d) => d.slice(0, 3)), [[16, 8, 2024], [21, 8, 2024], [26, 8, 2024], [27, 12, 2024], [4, 1, 2025], [23, 1, 2025]]);
    assert.ok(dates.every((d) => d[3] === 'dd/mm/yyyy'));
    let agr = null;
    flats.eachRow((r) => { if (r.getCell(3).value === '403') agr = r.getCell(7); });
    assert.ok(agr.value instanceof Date && agr.numFmt === 'dd/mm/yyyy');
    assert.deepStrictEqual([agr.value.getUTCDate(), agr.value.getUTCMonth() + 1, agr.value.getUTCFullYear()], [11, 12, 2024]);
    const s = wb.getWorksheet('Summary');
    let line = null;
    s.eachRow((r) => { if (typeof r.getCell(1).value === 'number' && r.getCell(3).value === '403') line = r; });
    assert.ok(line.getCell(6).value instanceof Date);
    assert.strictEqual(line.getCell(10).value, 2289900); // ADVANCES = base
    assert.strictEqual(line.getCell(12).value, 2200100); // O/S AMOUNT
    assert.strictEqual(line.getCell(13).value, 44900); // GST PAID = the GST receipt, not 1% of the advances (22,899)
    assert.strictEqual(line.getCell(14).value, 2289900); // Adv Rec for Form-3
    assert.strictEqual(line.getCell(15).value, 2200100); // Outstanding
    assert.strictEqual(line.getCell(16).value, '4611/2024');
    assert.strictEqual(line.getCell(17).value, 3649500);
  });

  console.log(`\n${passed} part-3 checks passed${failed ? `, ${failed} FAILED` : ''}`);
  process.exit(failed ? 1 : 0);
})();
