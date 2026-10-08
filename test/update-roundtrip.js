
'use strict';
const { parseTally }     = require('./src/tally/parse');
const { buildRecord }    = require('./src/engine/collection');
const { buildTallyOnlyRecords } = require('./src/engine/tallyOnly');
const { writeCollectionWorkbook } = require('./src/output/workbook');
const { updateCollectionSheet, parseWingBAoa } = require('./src/services/collection-update.service');
const XLSX = require('xlsx');

async function main() {
  let pass = 0, fail = 0;
  function ok(label, cond) {
    if (cond) { console.log('  ok  -', label); pass++; }
    else       { console.log('  FAIL-', label); fail++; process.exitCode = 1; }
  }

  const tally     = parseTally('config/samples/tally receipts.xlsx');
  const overrides = require('./config/overrides.json');

  const rec1804 = buildRecord(
    { flatNoHint:'1804', agreementValue:4490000, names:['Mrs. Chhaya Tukaram Kadam'], valueCandidates:[] },
    tally
  );
  const tallyOnly = buildTallyOnlyRecords(tally, ['1804'], {});
  const allRecs   = [rec1804, ...tallyOnly];

  // Step 1: generate fresh
  const gen = 'output/rt-fresh.xlsx';
  writeCollectionWorkbook(allRecs, gen, { projectName: 'TEST' });

  console.log('Update round-trip tests');

  // Step 2: same tally → 0 new payments
  const r1 = await updateCollectionSheet({
    existingPath: gen, outPath: 'output/rt-same.xlsx',
    tally, overrides, projectName: 'TEST',
  });
  ok('same-tally update: 0 new payments (no duplicates)', r1.newPayments === 0);

  // Step 3: add 1 truly new payment
  tally.byFlat.get('1804').payments.push({
    date: new Date('2026-05-01'), amount: 250000, vchNo: '999', gstRate: null,
  });
  const r2 = await updateCollectionSheet({
    existingPath: gen, outPath: 'output/rt-one-new.xlsx',
    tally, overrides, projectName: 'TEST',
  });
  ok('one-new update: exactly 1 new payment for 1804', r2.newPayments === 1);

  // Step 4: verify content of updated file
  const wb  = XLSX.readFile('output/rt-one-new.xlsx');
  const ws  = wb.Sheets[wb.SheetNames.find(n => n.trim() === 'Wing B')];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const recs = parseWingBAoa(aoa);
  const r1804 = recs.find(r => r.flatNo === '1804');
  ok('1804 has 6 payments in updated file', r1804 && r1804.payments.length === 6);

  // Step 5: update again with same tally → still 0 new (no re-adding)
  const r3 = await updateCollectionSheet({
    existingPath: 'output/rt-one-new.xlsx', outPath: 'output/rt-idempotent.xlsx',
    tally, overrides, projectName: 'TEST',
  });
  ok('idempotent: second update adds 0 new payments', r3.newPayments === 0);

  console.log('');
  console.log((pass + fail) + ' checks: ' + pass + ' passed, ' + fail + ' failed');
}
main().catch(e => { console.error('ERROR:', e.message); process.exitCode = 1; });
