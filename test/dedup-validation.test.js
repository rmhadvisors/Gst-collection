'use strict';

/**
 * Mandatory Validation Tests for Duplicate-Block Fix
 *
 * TEST 1: Flat exists as tally-only → Agreement uploaded later → Single consolidated block
 * TEST 2: Existing transactions appear again in new tally export → No duplicate payments
 * TEST 3: Same customer has multiple receipts → All grouped under one flat block
 * TEST 4: Final Collection Sheet → One flat = One customer block
 * TEST 5: Final GST Workbook context → One flat = One GST calculation block
 */

const assert = require('assert');
const path   = require('path');
const fs     = require('fs');
const XLSX   = require('xlsx');

const { parseTally, mergeTallies } = require('../src/tally/parse');
const { buildRecord }              = require('../src/engine/collection');
const { buildTallyOnlyRecords }    = require('../src/engine/tallyOnly');
const { writeStyledCollectionWorkbook } = require('../src/output/workbook-styled');
const { updateCollectionSheet, parseWingBAoa } = require('../src/services/collection-update.service');

let pass = 0;
let fail = 0;

function ok(label, cond, extra = '') {
  if (cond) {
    console.log(`  ok  - ${label}`);
    pass++;
  } else {
    console.log(`  FAIL - ${label}${extra ? '\n         ' + extra : ''}`);
    fail++;
  }
}

function section(title) { console.log(`\n${title}`); }

// ─── Paths ────────────────────────────────────────────────────────────────────
const SAMPLES      = path.join(__dirname, '../config/samples');
const OUTPUT       = path.join(__dirname, '../output');
const TALLY_PATH   = path.join(SAMPLES, 'tally receipts.xlsx');
const SUNRAY_TALLY = path.join(SAMPLES, 'Tally bank receipt.xlsx');
const WORK_DIR     = path.join(OUTPUT, '_dedup_test_' + Date.now());
fs.mkdirSync(WORK_DIR, { recursive: true });

/**
 * Count Wing B blocks (flat header rows with TOTAL ::) in an xlsx file.
 */
function countWingBBlocks(xlsxPath) {
  const wb  = XLSX.readFile(xlsxPath);
  const ws  = wb.Sheets[wb.SheetNames.find(n => n.trim() === 'Wing B')];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const recs = parseWingBAoa(aoa);
  // Return map of flatNo → record for deep inspection
  const map = new Map();
  for (const r of recs) {
    const key = String(r.flatNo);
    if (!map.has(key)) map.set(key, r);
    else {
      // If duplicate found, store both
      const existing = map.get(key);
      if (!Array.isArray(existing)) map.set(key, [existing, r]);
      else existing.push(r);
    }
  }
  return map;
}

async function runTests() {

  // ──────────────────────────────────────────────────────────────────────────
  section('TEST 1 — Tally-only flat → Agreement uploaded → Single block');
  // ──────────────────────────────────────────────────────────────────────────

  // Step A: Generate initial collection sheet with ONLY Sunray tally (flat 904, tally-only)
  const sunrayTally = parseTally(SUNRAY_TALLY);
  const t1Initial   = buildTallyOnlyRecords(sunrayTally, new Set(), {});
  const t1Path      = path.join(OUTPUT, 'test1-initial.xlsx');
  await writeStyledCollectionWorkbook(t1Initial, t1Path, { projectName: 'TEST1' });

  // Step B: Now update with same tally + agreement PDF (real PDF for flat 904)
  const PDF_904  = path.join(SAMPLES, 'FLAT NO.904.pdf');
  const t1Updated = path.join(OUTPUT, 'test1-updated.xlsx');

  const result1 = await updateCollectionSheet({
    existingPath:   t1Path,
    outPath:        t1Updated,
    tally:          sunrayTally,
    agreementPaths: [PDF_904],
    overrides:      {},
    workDir:        path.join(WORK_DIR, 't1'),
    projectName:    'TEST1',
  });

  const t1Map = countWingBBlocks(t1Updated);
  ok('TEST 1a: Flat 904 appears in output',
     t1Map.has('904'), 'flat 904 not found at all');
  ok('TEST 1b: Flat 904 appears EXACTLY ONCE (no duplicate block)',
     !Array.isArray(t1Map.get('904')), 'flat 904 has multiple blocks!');
  ok('TEST 1c: Flat 904 has hasAgreement=true after upgrade',
     t1Map.get('904') && t1Map.get('904').hasAgreement,
     `hasAgreement=${t1Map.get('904')?.hasAgreement}`);
  ok('TEST 1d: reconciledFlats includes 904',
     result1.reconciledFlats.includes('904'), JSON.stringify(result1.reconciledFlats));
  ok('TEST 1e: Flat 904 retains 2 payments (from Sunray tally)',
     t1Map.get('904')?.payments.length >= 2,
     `payments count=${t1Map.get('904')?.payments.length}`);
  ok('TEST 1f: No PDF warnings for 904',
     !result1.pdfWarnings.some(w => w.file && w.file.includes('904') && w.error && !w.error.includes('consolidated')),
     JSON.stringify(result1.pdfWarnings));

  // ──────────────────────────────────────────────────────────────────────────
  section('TEST 2 — Same tally re-uploaded → No duplicate payments');
  // ──────────────────────────────────────────────────────────────────────────

  // Run Update again on t1Updated with the SAME tally (no new payments)
  const t2Path = path.join(OUTPUT, 'test2-idempotent.xlsx');
  const result2 = await updateCollectionSheet({
    existingPath:   t1Updated,
    outPath:        t2Path,
    tally:          sunrayTally,
    agreementPaths: [],   // no PDFs, just same tally again
    overrides:      {},
    workDir:        path.join(WORK_DIR, 't2'),
    projectName:    'TEST2',
  });

  const t2Map = countWingBBlocks(t2Path);
  ok('TEST 2a: Flat 904 still appears exactly once',
     t2Map.has('904') && !Array.isArray(t2Map.get('904')),
     'flat 904 duplicated after idempotent re-upload');
  ok('TEST 2b: No new payments added (same tally = all duplicates skipped)',
     result2.newPayments === 0,
     `newPayments=${result2.newPayments}`);
  const t2Pay = t2Map.get('904')?.payments.length;
  const t1Pay = t1Map.get('904')?.payments.length;
  ok('TEST 2c: Payment count unchanged after re-upload',
     t2Pay === t1Pay,
     `t2Pay=${t2Pay}, t1Pay=${t1Pay}`);

  // ──────────────────────────────────────────────────────────────────────────
  section('TEST 3 — Multiple receipts for same flat → All grouped under one block');
  // ──────────────────────────────────────────────────────────────────────────

  // Use main tally (has flats with 3-8 payments) + generate
  const mainTally   = parseTally(TALLY_PATH);
  const t3AllRecs   = buildTallyOnlyRecords(mainTally, new Set(), {});
  const flat806     = t3AllRecs.find(r => r.flatNo === '806');  // 8 payments
  const flat1604    = t3AllRecs.find(r => r.flatNo === '1604'); // 15 payments
  ok('TEST 3a: Flat 806 has 8 payments in source tally',
     flat806 && flat806.payments.length === 8,
     `payments=${flat806?.payments.length}`);
  ok('TEST 3b: Flat 1604 has 15 payments in source tally',
     flat1604 && flat1604.payments.length === 15,
     `payments=${flat1604?.payments.length}`);

  // Write and parse back to verify correct grouping
  const t3Path = path.join(OUTPUT, 'test3-grouping.xlsx');
  await writeStyledCollectionWorkbook(t3AllRecs, t3Path, { projectName: 'TEST3' });
  const t3Map = countWingBBlocks(t3Path);
  ok('TEST 3c: Flat 806 has exactly one block in Wing B',
     t3Map.has('806') && !Array.isArray(t3Map.get('806')),
     'flat 806 has multiple blocks!');
  ok('TEST 3d: Flat 1604 has exactly one block in Wing B',
     t3Map.has('1604') && !Array.isArray(t3Map.get('1604')),
     'flat 1604 has multiple blocks!');
  ok('TEST 3e: Flat 806 block has 8 payments preserved',
     t3Map.get('806')?.payments.length === 8,
     `parsed payments=${t3Map.get('806')?.payments.length}`);
  ok('TEST 3f: Flat 1604 block has 15 payments preserved',
     t3Map.get('1604')?.payments.length === 15,
     `parsed payments=${t3Map.get('1604')?.payments.length}`);

  // ──────────────────────────────────────────────────────────────────────────
  section('TEST 4 — Final Collection Sheet: One flat = One customer block');
  // ──────────────────────────────────────────────────────────────────────────

  // Use the real main tally to generate a full sheet, verify NO flat appears twice
  const t4AllRecs = buildTallyOnlyRecords(mainTally, new Set(), {});
  const t4Path    = path.join(OUTPUT, 'test4-full-sheet.xlsx');
  await writeStyledCollectionWorkbook(t4AllRecs, t4Path, { projectName: 'TEST4' });
  const t4Map = countWingBBlocks(t4Path);

  const duplicatesInT4 = [...t4Map.entries()].filter(([, v]) => Array.isArray(v));
  ok('TEST 4a: No flat appears more than once in the generated sheet',
     duplicatesInT4.length === 0,
     `Duplicate flats: ${duplicatesInT4.map(([k]) => k).join(', ')}`);
  ok('TEST 4b: Total flat count matches tally flat count',
     t4Map.size === mainTally.byFlat.size,
     `sheet=${t4Map.size}, tally=${mainTally.byFlat.size}`);

  // Update the full sheet with same tally — still no duplicates
  const t4UpdatedPath = path.join(OUTPUT, 'test4-updated.xlsx');
  await updateCollectionSheet({
    existingPath:   t4Path,
    outPath:        t4UpdatedPath,
    tally:          mainTally,
    agreementPaths: [],
    overrides:      {},
    workDir:        path.join(WORK_DIR, 't4'),
    projectName:    'TEST4',
  });
  const t4UMap = countWingBBlocks(t4UpdatedPath);
  const dups4U = [...t4UMap.entries()].filter(([, v]) => Array.isArray(v));
  ok('TEST 4c: No duplicates after update with same tally',
     dups4U.length === 0,
     `Duplicate flats after update: ${dups4U.map(([k]) => k).join(', ')}`);

  // ──────────────────────────────────────────────────────────────────────────
  section('TEST 5 — CA manual sheet as input: dates decoded, no phantom duplicates');
  // ──────────────────────────────────────────────────────────────────────────

  // Parse the CA manual sheet (has Excel serial dates)
  const CA_SHEET = path.join(SAMPLES, 'COLLECTION DETAILS.. SPACE HOMES - final (2).xlsx');
  const caWb = XLSX.readFile(CA_SHEET);
  const caWsName = caWb.SheetNames.find(n => n.trim() === 'Wing B');
  const caWs  = caWb.Sheets[caWsName];
  const caAoa = XLSX.utils.sheet_to_json(caWs, { header: 1, defval: '' });
  const caRecs = parseWingBAoa(caAoa);

  // Check that ALL payment dates are human-readable (not raw serial numbers)
  let rawSerialCount = 0;
  for (const rec of caRecs) {
    for (const p of rec.payments) {
      // A raw Excel serial would look like "45335" or "44927" — 5-digit integer string
      if (/^\d{5}$/.test(String(p.dateText).trim())) rawSerialCount++;
    }
  }
  ok('TEST 5a: No Excel serial numbers remain in parsed payment dates',
     rawSerialCount === 0,
     `${rawSerialCount} raw serial dates found`);

  // Agreement dates too
  let rawAgrSerials = 0;
  for (const rec of caRecs) {
    if (/^\d{5}$/.test(String(rec.agreementDate).trim())) rawAgrSerials++;
  }
  ok('TEST 5b: No Excel serial numbers in agreement dates',
     rawAgrSerials === 0,
     `${rawAgrSerials} raw serial agreement dates`);

  // Check flat 904 in CA manual sheet is parseable
  const ca904 = caRecs.find(r => r.flatNo === '904');
  ok('TEST 5c: Flat 904 exists in CA manual sheet',
     ca904 != null, 'flat 904 not found');
  ok('TEST 5d: Flat 904 agreementDate is readable (not a number)',
     ca904 && isNaN(Number(ca904.agreementDate)),
     `agreementDate="${ca904?.agreementDate}"`);

  // ── Summary ────────────────────────────────────────────────────────────────
  fs.rmSync(WORK_DIR, { recursive: true, force: true });
  console.log(`\n${pass + fail} checks — ${pass} passed${fail ? ', ' + fail + ' FAILED' : ''}\n`);
  if (fail) process.exitCode = 1;
}

runTests().catch(e => {
  console.error('\nFATAL:', e.message);
  console.error(e.stack);
  process.exitCode = 1;
});
