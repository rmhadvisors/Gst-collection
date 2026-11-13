'use strict';

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const ExcelJS = require('exceljs');
const { snapshot } = require('./snapshot-spacehome');
const {
  parseTally, mergeTallies, applyAllocations, totalReceipts,
} = require('../src/tally/parse');
const { buildRecord } = require('../src/engine/collection');
const { transliterateName } = require('../src/util/translit');
const { findDates } = require('../src/pdf/dateRetry');
const { processAgreements } = require('../src/pdf/agreement');
const { closeOcr } = require('../src/pdf/ocr');
const { buildTallyOnlyRecords } = require('../src/engine/tallyOnly');
const { buildUnallocated } = require('../src/engine/unallocated');
const { buildNoFlatRecords } = require('../src/engine/noFlat');
const { writeStyledCollectionWorkbook } = require('../src/output/workbook-styled');
const { parseFlatFromParticulars } = require('../src/util/flatKey');
const { extractAgreementValue } = require('../src/pdf/agreement');

let passed = 0;
let failed = 0;
function ok(name, fn) {
  try { fn(); passed += 1; console.log(`  ok  - ${name}`); } catch (e) {
    failed += 1; console.log(`  FAIL- ${name}\n       ${e.message}`);
  }
}
async function okAsync(name, fn) {
  try { await fn(); passed += 1; console.log(`  ok  - ${name}`); } catch (e) {
    failed += 1; console.log(`  FAIL- ${name}\n       ${e.message}`);
  }
}

const F = (f) => path.join(__dirname, 'fixtures', 'mahira-tower', f);
const S = (f) => path.join(__dirname, '..', 'config', 'samples', f);

console.log('\nNarration parsing');
ok('all narration formats give wing-aware keys', () => {
  const cases = {
    'Flat No.1804 Mrs. X': ['1804', 'Mrs. X'],
    'flat no 1804 Mrs. X': ['1804', 'Mrs. X'],
    'B 607 -RAJESH SINGH': ['B-607', 'RAJESH SINGH'],
    'B-1002 Foo Bar': ['B-1002', 'Foo Bar'],
    'A- 503 Foo Bar': ['A-503', 'Foo Bar'],
    'A 904 SANJAY RAM': ['A-904', 'SANJAY RAM'],
    'B- 1401 Mithali G': ['B-1401', 'Mithali G'],
    'Flat No. B-603 Z': ['B-603', 'Z'],
  };
  for (const [text, [key, name]] of Object.entries(cases)) {
    const r = parseFlatFromParticulars(text);
    assert.ok(r, text);
    assert.strictEqual(r.key, key, text);
    assert.strictEqual(r.rest, name, text);
  }
  assert.strictEqual(parseFlatFromParticulars('Riyaz AHmed Khan Capital Ac'), null);
});

console.log('\nAgreement value: consideration beats market value');
ok('1804 receipt text -> 4490000 without override', () => {
  const text = 'बाजार मुल्य: रु.4296000 /-\nमोबदला रु.4490000/-\nभरलेले मुद्रांक शुल्क : रु. 314300/-\n1) रक्कम: रु.5400/-';
  assert.strictEqual(extractAgreementValue(text).value, 4490000);
  // order reversed: consideration first, market value second must still win
  const rev = 'मोबदला रु.4490000/-\nबाजार मुल्य: रु.4296000 /-\nभरलेले मुद्रांक शुल्क : रु. 314300/-';
  assert.strictEqual(extractAgreementValue(rev).value, 4490000);
  assert.strictEqual(extractAgreementValue('Market value Rs.5000000/-\nConsideration Rs.4800000/-').value, 4800000);
});

console.log('\nSpace Home output unchanged (vs pre-change baseline; only flat 109 +Rs 2 for two Re.1 rows, approved)');
ok('every flat identical to baseline', () => {
  const base = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'space-home', 'expected.json'), 'utf8')).expected.flats;
  const now = snapshot(S('tally receipts.xlsx'));
  assert.deepStrictEqual(Object.keys(now), Object.keys(base));
  for (const k of Object.keys(base)) assert.deepStrictEqual(now[k], base[k], `flat ${k}`);
});
ok('flags the genuine 381 voucher (different dates) but keeps both payments', () => {
  const t = parseTally(S('tally receipts.xlsx'));
  assert.strictEqual(t.duplicateVouchers.length, 1);
  assert.strictEqual(t.byFlat.get('1804').payments.filter((p) => p.vchNo === '381').length, 2);
});

console.log('\nMahira Tower (two books, two wings)');
const mahira = mergeTallies([parseTally(S('mahira-bank-of-maharashtra.xlsx')), parseTally(S('mahira-federal-bank.xlsx'))]);
ok('18 flats, wings never merged', () => {
  assert.strictEqual(mahira.byFlat.size, 18);
  const keys = [...mahira.byFlat.keys()];
  assert.ok(keys.includes('A-503') && keys.includes('B-607') && keys.includes('B-1002'));
  assert.ok(keys.every((k) => /^[A-Z]-\d+$/.test(k)));
  const nums = keys.map((k) => k.split('-')[1]);
  assert.ok(keys.includes('A-904') && !keys.includes('B-904'));
  assert.ok(nums.length >= new Set(nums).size);
});
ok('B-607 payments total 3,090,100', () => {
  const total = mahira.byFlat.get('B-607').payments.reduce((s, p) => s + p.amount, 0);
  assert.strictEqual(total, 3090100);
});
ok('11 unallocated parties, none dropped, none auto-added to a flat', () => {
  assert.strictEqual(new Set(mahira.unallocated.map((r) => r.name)).size, 11);
  const flatRows = [...mahira.byFlat.values()].reduce((s, g) => s + g.payments.length, 0);
  assert.strictEqual(flatRows + mahira.unallocated.length, mahira.rows.length);
});
ok('Capital Ac guess points at a flat but never allocates', () => {
  const u = buildUnallocated(mahira, []);
  assert.strictEqual(u.length, mahira.unallocated.length);
  assert.ok(u.every((r) => r.amount > 0 && r.party));
});

// ── BOM.xlsx + FEDRAL BANK RECEIPT 26-27.xlsx + B-603_1.pdf ─────────────────
const bom = () => mergeTallies([parseTally(F('BOM.xlsx')), parseTally(F('FEDRAL BANK RECEIPT 26-27.xlsx'))]);
const sumPay = (g) => g.payments.reduce((s, p) => s + p.amount, 0);

console.log('\n1. Nothing from Tally is dropped');
ok('flat totals equal Tally: A-1504 = 3,31,002; B-1801 = 24,25,556', () => {
  const t = bom();
  assert.strictEqual(sumPay(t.byFlat.get('A-1504')), 331002);
  assert.strictEqual(sumPay(t.byFlat.get('B-1801')), 2425556);
});
ok('Re.1 vouchers 42, 65, 68, 70 each appear under a flat or in Unallocated', () => {
  const t = bom();
  const where = (v) => [
    ...[...t.byFlat.values()].flatMap((g) => g.payments.filter((p) => p.vchNo === v).map(() => 'flat')),
    ...t.unallocated.filter((r) => r.vchNo === v).map(() => 'unallocated'),
  ];
  for (const v of ['42', '65', '68', '70']) assert.strictEqual(where(v).length, 1, `Vch ${v}`);
  assert.ok(t.unallocated.some((r) => r.vchNo === '42' && r.amount === 1));
  assert.ok(t.byFlat.get('B-1801').payments.some((p) => p.vchNo === '65' && p.amount === 1));
  assert.ok(t.byFlat.get('A-1504').payments.some((p) => p.vchNo === '68'));
  assert.ok(t.byFlat.get('A-1504').payments.some((p) => p.vchNo === '70'));
});
ok('every Tally row is under a flat or Unallocated; Opening/Closing Balance excluded', () => {
  const t = bom();
  const flats = [...t.byFlat.values()].reduce((s, g) => s + sumPay(g), 0);
  const unalloc = t.unallocated.reduce((s, r) => s + r.amount, 0);
  assert.strictEqual(Math.round((flats + unalloc) * 100) / 100, totalReceipts(t));
  assert.ok(!t.rows.some((r) => /balance/i.test(r.particulars)));
});

console.log('\n3. Devanagari purchaser names are transliterated');
ok('names transliterate to English', () => {
  assert.strictEqual(transliterateName('शशीकांत बी ठाकूर (एच यू एफ)'), 'Shashikant B. Thakur (HUF)');
  assert.strictEqual(transliterateName('राजेश लालबचन सिंह'), 'Rajesh Lalbachan Singh');
  assert.strictEqual(transliterateName('Mrs. Chhaya Kadam'), 'Mrs. Chhaya Kadam');
});

console.log('\n4. Agreement date: dd/mm/yyyy and dd-mm-yyyy accepted');
ok('date reader accepts both separators and rejects impossible dates', () => {
  assert.deepStrictEqual(findDates('11/08/2026 x 27-08-2026'), ['11/08/2026', '27/08/2026']);
  assert.deepStrictEqual(findDates('31/02/2026 11/08/2076'), []);
});

ok('5. allocating a voucher that is not in Tally is reported and nothing moves', () => {
  const t = applyAllocations(bom(), { 999: 'B-603' });
  assert.strictEqual(t.manualAllocations[0].found, false);
  assert.ok(!t.byFlat.has('B-603'));
});

(async () => {
  let agreements = null;
  if (!process.env.SKIP_OCR) {
    console.log('\n   (OCR of B-603_1.pdf, ~1-2 min)');
    agreements = await processAgreements(F('B-603_1.pdf'), { workDir: path.join(os.tmpdir(), `b603-${Date.now()}`) });
    await closeOcr();
  }
  const ocrTest = async (name, fn) => (agreements ? okAsync(name, fn) : console.log(`  skip- ${name}`));
  const agr = (flat) => agreements.find((a) => a.flatNoHint === flat);

  await ocrTest('4. B-603 agreement date read from Index-II = 11/08/2026 (retry OCR, deskew)', () => {
    const a = agr('B-603');
    assert.ok(a, 'B-603 found');
    assert.ok(a.agreementDate, 'date read');
    assert.deepStrictEqual([a.agreementDate.getFullYear(), a.agreementDate.getMonth() + 1, a.agreementDate.getDate()], [2026, 8, 11]);
    assert.strictEqual(a.dateSource, 'ocr-retry (deskew)');
    assert.strictEqual(a.agreementValue, 2610000);
  });

  await ocrTest('5. without override Capital Ac receipts stay unallocated and B-603 has none', () => {
    const t = bom();
    const rec = buildRecord(agr('B-603'), t, {});
    assert.strictEqual(rec.payments.length, 0);
    assert.deepStrictEqual(t.unallocated.filter((r) => ['81', '82'].includes(r.vchNo)).map((r) => r.amount).sort((a, b) => a - b), [6200, 78200]);
  });
  await ocrTest('5. allocate {"81":"B-603","82":"B-603"} moves them under B-603, flagged "(manual allocation)"', () => {
    const t = applyAllocations(bom(), { 81: 'B-603', 82: 'B-603' });
    const rec = buildRecord(agr('B-603'), t, {});
    assert.deepStrictEqual(rec.payments.map((p) => [p.vchNo, p.amount]), [['81', 6200], ['82', 78200]]);
    assert.strictEqual(rec.totalReceived, 84400);
    assert.strictEqual(rec.matchReason, 'ocr-flat');
    assert.strictEqual(rec.flags.filter((f) => /\(manual allocation\)/.test(f)).length, 2);
    assert.ok(!t.unallocated.some((r) => ['81', '82'].includes(r.vchNo)));
    assert.strictEqual(rec.name, 'Shashikant B. Thakur (HUF)');
    assert.strictEqual(rec.purchaserMarathi, 'शशीकांत बी ठाकूर (एच यू एफ)');
    assert.ok(t.manualAllocations.every((m) => m.found));
  });

  await ocrTest('6. B-607 Vch 63 = 30,300 (1% of 30,30,000): shown under the flat, no GST charged', () => {
    const rec = buildRecord(agr('B-607'), bom(), {});
    assert.strictEqual(rec.agreementValue, 3030000);
    assert.strictEqual(rec.gstOnAgreement, 30300);
    const v63 = rec.payments.find((p) => p.vchNo === '63');
    assert.ok(v63 && v63.amount === 30300 && v63.gst === 0 && v63.gstReceipt);
    assert.strictEqual(rec.totalReceived, 3090100);
    assert.strictEqual(rec.totalGst, Math.round((3090100 - 30300)) / 100);
    assert.strictEqual(rec.balance, 3030000 + 30300 - 3090100);
    assert.ok(rec.flags.some((f) => /Vch 63.*GST receipt/.test(f)));
    assert.strictEqual(rec.payments.filter((p) => p.gstReceipt).length, 1);
  });

  await ocrTest('2. Audit ends with reconciliation MATCHED; Marathi original in own column; DIFFERENCE shows red', async () => {
    const t = applyAllocations(bom(), { 81: 'B-603', 82: 'B-603' });
    const records = [buildRecord(agr('B-603'), t, {}), buildRecord(agr('B-607'), t, {})];
    const matched = new Set(records.flatMap((r) => [r.flatNo, r.tallyFlat].filter(Boolean)));
    records.push(...buildTallyOnlyRecords(t, matched, {}, {}));
    records.push(...buildNoFlatRecords(t, { records }));
    const out = path.join(os.tmpdir(), `bom-${Date.now()}.xlsx`);
    const opts = { projectName: 'M', reconciliation: { tallyTotal: totalReceipts(t) } };
    await writeStyledCollectionWorkbook(records, out, opts);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    const rows = [];
    wb.getWorksheet('Audit').eachRow((r) => rows.push(r.values.slice(1)));
    const header = rows[0];
    const mrCol = header.indexOf('Purchaser (Marathi original)');
    assert.ok(mrCol > 0);
    const b603 = rows.find((r) => r[0] === 'B-603');
    assert.strictEqual(b603[1], 'Shashikant B. Thakur (HUF)');
    assert.strictEqual(b603[mrCol], 'शशीकांत बी ठाकूर (एच यू एफ)');
    assert.ok(/\(manual allocation\)/.test(b603[header.indexOf('Flags')]));
    assert.ok(/GST receipt/.test(rows.find((r) => r[0] === 'B-607')[header.indexOf('Flags')]));
    assert.strictEqual(rows[rows.length - 1][3], 'MATCHED');
    assert.strictEqual(rows.find((r) => /Total Tally receipts/.test(String(r[0])))[3], totalReceipts(t));
    assert.ok(!wb.getWorksheet('Unallocated Receipts'), 'no separate Unallocated sheet');
    assert.strictEqual(records.filter((r) => r.noFlat).flatMap((r) => r.payments).length, t.unallocated.length); // all on the main sheet

    records[0].payments.pop(); // a receipt goes missing -> DIFFERENCE in red
    await writeStyledCollectionWorkbook(records, out, opts);
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.readFile(out);
    fs.unlinkSync(out);
    let diffCell = null;
    wb2.getWorksheet('Audit').eachRow((r) => { const c = r.getCell(4); if (/^DIFFERENCE/.test(String(c.value))) diffCell = c; });
    assert.ok(diffCell, 'DIFFERENCE shown');
    assert.strictEqual(diffCell.value, 'DIFFERENCE: ₹78,200.00');
    assert.strictEqual(diffCell.font.color.argb, 'FFC00000');
  });

  await okAsync('multi-wing workbook: Wing A + Wing B + Audit; no-flat parties on the first sheet', async () => {
    const records = buildTallyOnlyRecords(mahira, [], {}, {});
    const out = path.join(os.tmpdir(), `mahira-${Date.now()}.xlsx`);
    records.push(...buildNoFlatRecords(mahira, { records }));
    await writeStyledCollectionWorkbook(records, out, { projectName: 'T' });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    fs.unlinkSync(out);
    const names = wb.worksheets.map((w) => w.name);
    for (const n of ['Summary', 'Wing A', 'Wing B', 'Audit']) assert.ok(names.includes(n), n);
    assert.ok(!names.includes('=='));
    assert.ok(!names.includes('Unallocated Receipts'));
    const partyRows = [];
    wb.getWorksheet('Wing A').eachRow((r) => { if (r.getCell(11).value === 'TOTAL ::' && !r.getCell(3).value) partyRows.push(r.getCell(8).value); });
    assert.strictEqual(partyRows.length, new Set(mahira.unallocated.map((r) => r.name)).size, 'one block per party, below the flats');
  });
  await okAsync('single-wing/wing-less project gets one "Flats" sheet (no default wing)', async () => {
    const t = parseTally(S('tally receipts.xlsx'));
    const out = path.join(os.tmpdir(), `space-${Date.now()}.xlsx`);
    await writeStyledCollectionWorkbook(buildTallyOnlyRecords(t, [], {}, {}), out, { projectName: 'S' });
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(out);
    fs.unlinkSync(out);
    assert.deepStrictEqual(wb.worksheets.map((w) => w.name), ['Summary', 'Flats', 'Audit']);
  });

  console.log(`\n${passed} regression checks passed${failed ? `, ${failed} FAILED` : ''}`);
  process.exit(failed ? 1 : 0);
})();
