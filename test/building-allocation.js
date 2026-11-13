'use strict';

/**
 * Receipts in a Tally book with no building (Mrunmayee: idfc 5112 / 47288) are placed by rule:
 *   1. the flat number exists in only one building -> that building;
 *   2. several buildings -> flat + name (>= 0.85) + AGMT DT must equal the agreement;
 *   3. the Tally buyer name is shown when it matches; the transliteration stays in Audit.
 * Run with and without 403.pdf (3 books: hdfc 444 names Siddhivinayak, the IDFC books have no building).
 */

const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { buildRecords } = require('../src/services/build-records');

const DIR = path.join(__dirname, 'fixtures', 'mrunmayee-developers');
const file = (n) => ({ path: path.join(DIR, n), originalname: n });
const BOOKS = ['hdfc 444.xlsx', 'idfc 47288.xlsx', 'idfc 5112.xlsx'];
const overrides = () => ({
  buildings: {
    'Maitri Heights': ['मैत्री हाईट्स', 'maitri'],
    'Siddhivinayak Palace': ['सिद्धिविनायक पॅलेस', 'siddhivinayak'],
  },
});

async function run(agreements, extra = {}) {
  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'bld-'));
  try {
    return await buildRecords({
      tallyFiles: BOOKS.map(file), agreementFiles: agreements.map(file), overrides: { ...overrides(), ...extra }, workDir,
    });
  } finally { fs.rmSync(workDir, { recursive: true, force: true }); }
}

let passed = 0;
let failed = 0;
const check = (label, fn) => {
  try { fn(); passed += 1; console.log(`  ok  - ${label}`); } catch (e) { failed += 1; console.log(`  FAIL- ${label}\n       ${e.message}`); }
};

(async () => {
  console.log('\nMrunmayee 3 books + 403.pdf (IDFC books have no building)');
  const a = await run(['403.pdf']);
  const rec = a.records.find((r) => r.flatNo === '403');
  check('403 is one record with its agreement', () => { assert.ok(rec); assert.strictEqual(rec.hasAgreement, true); assert.strictEqual(rec.building, 'Maitri Heights'); });
  check('received 23,34,800; Vch 72 44,900 is the GST receipt; GST outstanding 0', () => {
    assert.strictEqual(rec.totalReceived, 2334800);
    assert.strictEqual(rec.gstReceived, 44900);
    assert.strictEqual(rec.gstOutstanding, 0);
    assert.strictEqual(rec.payments.find((p) => String(p.vchNo) === '72').gstReceipt, true);
  });
  check('base received 22,89,900; balance outstanding 22,00,100', () => {
    assert.strictEqual(rec.baseReceived, 2289900);
    assert.strictEqual(rec.osAmount, 2200100);
    assert.strictEqual(rec.balance, 2200100);
  });
  check('6 receipts, none left in Unallocated', () => {
    assert.strictEqual(rec.payments.length, 6);
    assert.ok(!a.unallocated.some((u) => /403/.test(u.guess)));
    assert.strictEqual(a.unallocated.length, 63);
  });
  check('buyer name is the Tally name; allocation is recorded in Audit flags', () => {
    assert.strictEqual(rec.name, 'PRADNYA AJIT PRABHUKAR');
    assert.ok(rec.flags.some((f) => /allocated/.test(f) && /only one building|flat\+name\+date/.test(f)), rec.flags.join(' | '));
  });

  console.log('\nMrunmayee 3 books WITHOUT 403.pdf');
  const b = await run([]);
  const t = b.records.find((r) => r.flatNo === '403');
  check('403 appears tally-only with 23,34,800', () => {
    assert.ok(t);
    assert.strictEqual(t.hasAgreement, false);
    assert.strictEqual(t.totalReceived, 2334800);
    assert.strictEqual(t.payments.length, 6);
  });
  check('only the 63 flat-less receipts remain Unallocated', () => assert.strictEqual(b.unallocated.length, 63));

  console.log('\nReceipts without a flat no.: partner setting, allocate, partner names');
  const party = (r, n) => r.noFlat.find((x) => x.name.toLowerCase() === n.toLowerCase());
  const blank = await run([], { partners: ['Bhumik Hemant Patil'] });
  const one = await run([], { partners: ['Bhumik Hemant Patil'], partnerGst: '1%' });
  check('partner: GST blank + flag by default; CHQ RETURN no GST; other parties 1%', () => {
    const p = party(blank, 'Bhumik Hemant Patil');
    assert.strictEqual(p.totalGst, 0);
    assert.ok(p.payments.every((x) => x.gst == null));
    assert.ok(p.flags[0].startsWith('partner – check if flat booking or capital'));
    assert.strictEqual(party(blank, 'CHQ RETURN').totalGst, 0);
    assert.strictEqual(party(blank, 'CHQ RETURN').payments.length, 9);
    assert.strictEqual(party(blank, 'Jai Hemant Patil').totalGst, 23585);
    assert.ok(party(blank, 'Jai Hemant Patil').flags[0].startsWith('no flat no. – assign flat to confirm'));
  });
  check('partnerGst "1%" switches the partner to 1% GST', () => {
    assert.strictEqual(party(one, 'Bhumik Hemant Patil').totalGst, 22350);
    assert.strictEqual(party(one, 'CHQ RETURN').totalGst, 0); // never taxed
  });
  check('parties sorted by total, largest first; every receipt on a block or a flat', () => {
    const t = blank.noFlat.map((x) => x.totalReceived);
    assert.deepStrictEqual(t, [...t].sort((x, y) => y - x));
    const sum = (l) => l.reduce((s, x) => s + x.totalReceived, 0);
    assert.strictEqual(Math.round((sum(blank.noFlat) + sum(blank.records.filter((r) => !r.noFlat))) * 100) / 100, blank.tallyTotal);
  });
  const moved = await run([], { allocate: { 16: '403' } });
  check('"allocate" moves a voucher from its party block into the flat', () => {
    assert.strictEqual(party(moved, 'Bhumik Hemant Patil').totalReceived, party(blank, 'Bhumik Hemant Patil').totalReceived - 106000);
    assert.strictEqual(moved.records.find((r) => r.flatNo === '403').totalReceived, 2334800 + 106000);
  });
  check('partner names are read from Index-II field (7) after भागीदार', () => {
    const { parsePartners } = require('../src/pdf/indexII');
    assert.deepStrictEqual(parsePartners('नाव:-मेः मूण्मयी डेव्हलपर्स तर्फे भागीदार दार भुमिक हेमंत पाटील तर्फे कु.मु. रमेश - वय:-46'), ['भुमिक हेमंत पाटील']);
    assert.deepStrictEqual(parsePartners('x तर्फे भागीदार नारायण लखू मांकर तर्फे y'), ['नारायण लखू मांकर']);
  });

  console.log(`\n${passed} building-allocation checks passed${failed ? `, ${failed} FAILED` : ''}`);
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
