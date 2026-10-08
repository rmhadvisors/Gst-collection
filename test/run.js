'use strict';

const assert = require('assert');
const path = require('path');
const {
  gstRate, buildRecord, GST_THRESHOLD,
} = require('../src/engine/collection');
const {
  buildTallyOnlyRecord, buildTallyOnlyRecords, DEFAULT_TALLY_ONLY_GST_RATE, TALLY_ONLY_STATUS,
} = require('../src/engine/tallyOnly');
const { extractAgreementValue, extractNames, extractAgreementDate } = require('../src/pdf/agreement');
const { parseTally, mergeTallies, gstRateFromCell } = require('../src/tally/parse');
const { nameSimilarity, floorFromFlat, floorLabel } = require('../src/util/text');
const { parseDate, formatMDY } = require('../src/util/dates');

let passed = 0;
function ok(name, fn) {
  try { fn(); passed += 1; console.log('  ok  -', name); }
  catch (e) { console.error('  FAIL-', name, '\n      ', e.message); process.exitCode = 1; }
}

console.log('GST rate (45-lakh rule)');
ok('4,490,000 -> 1%', () => assert.strictEqual(gstRate(4490000), 0.01));
ok('exactly 45L -> 1%', () => assert.strictEqual(gstRate(4500000), 0.01));
ok('45L + 1 -> 5%', () => assert.strictEqual(gstRate(4500001), 0.05));
ok('6,063,750 -> 5%', () => assert.strictEqual(gstRate(6063750), 0.05));
ok('threshold const', () => assert.strictEqual(GST_THRESHOLD, 4500000));

console.log('\nOCR text parsing');
const page1 = `rdw. 4278 fzATR: 10/03/2026
JTITT qq: %.4296000 /-
WrEEeT %.4490000/-
WTA HRTF oF 5. 314300/-
1) TgFTET YF: DHC 3H: .5400/-
2) eChallan THA: %.30000/-`;
ok('agreement value = 2nd /- amount (consideration)', () => {
  assert.strictEqual(extractAgreementValue(page1).value, 4490000);
});
ok('date parsed dd/mm/yyyy', () => {
  assert.strictEqual(formatMDY(extractAgreementDate(page1)), '3/10/26');
});

const page6 = `AND
(1) MRS. CHHAYA TUKARAM KADAM, aged about 60 Years, (PAN
ATMPKS5989R), Indian Inhabitants, residing at B/506,
(2) MR. TUSHAR TUKARAM KADAM, aged about 41 Years,`;
ok('primary allottee name', () => {
  const names = extractNames(page6);
  assert.strictEqual(names[0], 'MRS. CHHAYA TUKARAM KADAM');
  assert.strictEqual(names[1], 'MR. TUSHAR TUKARAM KADAM');
});

const page7Sunray = `AND
(1) MR. ANKIT JAYESH VORA, (PAN ANBPV7476K), aged about 34
years & (2) MR. JAYESH SHANTILAL VORA, (PAN AASPV9558N),
aged about 61 years, Indian Inhabitants`;
ok('Sunray page 7 names with PAN before aged about', () => {
  const names = extractNames(page7Sunray);
  assert.strictEqual(names[0], 'MR. ANKIT JAYESH VORA');
  assert.strictEqual(names[1], 'MR. JAYESH SHANTILAL VORA');
});

const page1Sunray = `fiw: 19/01/2026
6857000 /-
12303000 /-
861300 /-
1) DHC Amount`;
ok('Sunray agreement value fallback from large amounts', () => {
  assert.strictEqual(extractAgreementValue(page1Sunray).value, 12303000);
});
ok('Sunray agreement date', () => {
  assert.strictEqual(formatMDY(extractAgreementDate(page1Sunray)), '1/19/26');
});

console.log('\nUtilities');
ok('floor from flat 1804 -> 18', () => assert.strictEqual(floorFromFlat('1804'), 18));
ok('floor from flat 105 -> 1', () => assert.strictEqual(floorFromFlat('105'), 1));
ok('floor label 18 -> 18th Floor', () => assert.strictEqual(floorLabel(18), '18th Floor'));
ok('floor label 1 -> 1st Floor', () => assert.strictEqual(floorLabel(1), '1st Floor'));
ok('name match agreement vs tally', () => {
  assert.ok(nameSimilarity('MRS. CHHAYA TUKARAM KADAM', 'Mrs. Chhaya T. Kadam') >= 0.6);
});
ok('tally date 26-Dec-25 -> 12/26/25 (no TZ shift)', () => {
  assert.strictEqual(formatMDY(parseDate('26-Dec-25')), '12/26/25');
});

console.log('\nTally merging logic');
ok('mergeTallies merges multiple sheets, dedups, and sorts', () => {
  const t1 = {
    projectName: 'Building A',
    rows: [
      { flatNo: '101', name: 'Buyer One', date: new Date('2026-01-01'), amount: 100000, vchNo: '101' },
      { flatNo: '102', name: 'Buyer Two', date: new Date('2026-01-05'), amount: 200000, vchNo: '102' }
    ]
  };
  const t2 = {
    projectName: 'Building B',
    rows: [
      { flatNo: '101', name: 'Buyer One Alias', date: new Date('2026-01-03'), amount: 150000, vchNo: '103' },
      // TRUE duplicate: same vchNo + amount + SAME date (operator re-entry error) — should dedup to 1
      { flatNo: '102', name: 'Buyer Two', date: new Date('2026-01-05'), amount: 200000, vchNo: '102' }
    ]
  };

  const merged = mergeTallies([t1, t2]);

  assert.strictEqual(merged.projectName, 'Building A & Building B');
  assert.strictEqual(merged.rows.length, 4);

  const flat101 = merged.byFlat.get('101');
  assert.ok(flat101);
  assert.strictEqual(flat101.payments.length, 2);
  assert.strictEqual(flat101.nameList.length, 2);
  assert.ok(flat101.nameList.includes('Buyer One'));
  assert.ok(flat101.nameList.includes('Buyer One Alias'));
  // verify sorted by date: 2026-01-01 then 2026-01-03
  assert.strictEqual(flat101.payments[0].amount, 100000);
  assert.strictEqual(flat101.payments[1].amount, 150000);

  const flat102 = merged.byFlat.get('102');
  assert.ok(flat102);
  // vchNo 102, amount 200000, date Jan 5 appears twice (same date = true dup) → keep only 1
  assert.strictEqual(flat102.payments.length, 1);
  assert.strictEqual(flat102.payments[0].date.getTime(), new Date('2026-01-05').getTime());
});

console.log('\nTally-only flats (no agreement) — default 1% GST');
ok('DEFAULT_TALLY_ONLY_GST_RATE is 1%', () => assert.strictEqual(DEFAULT_TALLY_ONLY_GST_RATE, 0.01));

const syntheticGroup = {
  flatNo: '305',
  nameList: ['Some Buyer'],
  payments: [
    { date: new Date('2026-01-10'), amount: 200000, vchNo: '11', gstRate: null },
    { date: new Date('2026-02-15'), amount: 300000, vchNo: '12', gstRate: null },
  ],
};
ok('tally-only record computes 1% GST per receipt', () => {
  const rec = buildTallyOnlyRecord(syntheticGroup);
  assert.strictEqual(rec.hasAgreement, false);
  assert.strictEqual(rec.status, TALLY_ONLY_STATUS);
  assert.strictEqual(rec.gstRate, 0.01);
  assert.strictEqual(rec.payments[0].gst, 2000);
  assert.strictEqual(rec.payments[1].gst, 3000);
  assert.strictEqual(rec.totalReceived, 500000);
  assert.strictEqual(rec.totalGst, 5000);
  assert.strictEqual(rec.agreementValue, null);
  assert.strictEqual(rec.balance, null);
});
ok('tally-only record derives floor/name', () => {
  const rec = buildTallyOnlyRecord(syntheticGroup);
  assert.strictEqual(rec.flatNo, '305');
  assert.strictEqual(rec.floorLabel, '3rd Floor');
  assert.strictEqual(rec.name, 'Some Buyer');
});
ok('per-receipt Tally GST rate overrides default', () => {
  const rec = buildTallyOnlyRecord({
    flatNo: '306', nameList: ['B'],
    payments: [{ date: new Date('2026-01-10'), amount: 100000, vchNo: '1', gstRate: 0.05 }],
  });
  assert.strictEqual(rec.payments[0].gst, 5000);
});

ok('agreement record ignores Tally GST column and keeps 45-lakh rule', () => {
  const tally = {
    byFlat: new Map([
      ['401', {
        flatNo: '401',
        nameList: ['Agreement Buyer'],
        payments: [{ date: new Date('2026-01-10'), amount: 100000, vchNo: '1', gstRate: 0.05 }],
      }],
    ]),
  };
  const rec = buildRecord(
    { flatNoHint: '401', agreementValue: 4400000, names: ['Agreement Buyer'], valueCandidates: [] },
    tally,
  );
  assert.strictEqual(rec.gstRate, 0.01);
  assert.strictEqual(rec.payments[0].gst, 1000);
});

ok('buildTallyOnlyRecords skips matched flats, includes the rest', () => {
  const tally = {
    byFlat: new Map([
      ['101', { flatNo: '101', nameList: ['A'], payments: [{ date: new Date('2026-01-01'), amount: 100000, vchNo: '1', gstRate: null }] }],
      ['102', { flatNo: '102', nameList: ['B'], payments: [{ date: new Date('2026-01-02'), amount: 250000, vchNo: '2', gstRate: null }] }],
    ]),
  };
  const recs = buildTallyOnlyRecords(tally, ['101'], {});
  assert.strictEqual(recs.length, 1);
  assert.strictEqual(recs[0].flatNo, '102');
  assert.strictEqual(recs[0].totalGst, 2500);
});

console.log('\nGST column interpretation (future Tally column)');
ok('"1%" -> 0.01', () => assert.strictEqual(gstRateFromCell('1%'), 0.01));
ok('"5%" -> 0.05', () => assert.strictEqual(gstRateFromCell('5%'), 0.05));
ok('1 -> 0.01', () => assert.strictEqual(gstRateFromCell(1), 0.01));
ok('5 -> 0.05', () => assert.strictEqual(gstRateFromCell(5), 0.05));
ok('0.01 -> 0.01', () => assert.strictEqual(gstRateFromCell(0.01), 0.01));
ok('empty -> null', () => assert.strictEqual(gstRateFromCell(''), null));

console.log('\nEnd-to-end against real tally (if available)');
const tallyPath = process.env.TALLY || 'C:\\Users\\HP\\Downloads\\tally receipts.xlsx';
try {
  const tally = parseTally(tallyPath);

  // 1804 — under 45 lakh -> 1%
  const recLow = buildRecord(
    { flatNoHint: '1804', agreementValue: 4490000, names: ['Mrs. Chhaya Tukaram Kadam'], valueCandidates: [] },
    tally,
  );
  ok('1804 GST on agreement = 44,900 (1%)', () => assert.strictEqual(recLow.gstOnAgreement, 44900));
  ok('1804 balance = agr + gst - received', () => {
    assert.strictEqual(recLow.balance, 4490000 + 44900 - recLow.totalReceived);
  });
  ok('1804 each advance GST = 1% of amount', () => {
    recLow.payments.forEach((p) => assert.strictEqual(p.gst, Math.round(p.amount * 0.01 * 100) / 100));
  });
  ok('1804 dedup: vch 381 appears TWICE — different dates, both genuine payments', () => {
    const vch381 = recLow.payments.filter((p) => p.vchNo === '381');
    assert.strictEqual(vch381.length, 2, `expected 2 entries for vch 381 (Dec 26 + Jan 1), got ${vch381.length}`);
  });
  ok('1804 dedup: 5 payments total (no false dedup)', () => {
    assert.strictEqual(recLow.payments.length, 5);
  });
  ok('1804 reference override: 7 payments and matching reference total', () => {
    const recReference = buildRecord(
      { flatNoHint: '1804', agreementValue: 4490000, names: ['Mrs. Chhaya Tukaram Kadam'], valueCandidates: [] },
      tally,
      {
        extraPayments: [
          { date: '3/4/26', amount: 44900 },
          { date: '4/25/26', amount: 140699 },
        ],
      },
    );
    // 5 tally payments + 2 extra override payments = 7 total
    assert.strictEqual(recReference.payments.length, 7);
    // totalReceived = 1100000+1100000+1100000+400000+483010+44900+140699 = 4368609
    assert.strictEqual(recReference.totalReceived, 4368609);
    assert.strictEqual(recReference.totalGst, Math.round(recReference.totalReceived * 0.01 * 100) / 100);
  });

  // 1805 — over 45 lakh -> 5%
  const recHigh = buildRecord(
    { flatNoHint: '1805', agreementValue: 6063750, names: ['Shaktiganesh Bikas Dande'], valueCandidates: [] },
    tally,
  );
  ok('1805 GST on agreement = 303,187.50 (5%)', () => assert.strictEqual(recHigh.gstOnAgreement, 303187.5));
  ok('1805 advance GST = 5% of amount', () => {
    recHigh.payments.forEach((p) => assert.strictEqual(p.gst, Math.round(p.amount * 0.05 * 100) / 100));
  });

  // All-flats behaviour: tally-only records cover every flat in the workbook.
  ok('tally-only covers all flats, 1% GST, no agreement', () => {
    const recs = buildTallyOnlyRecords(tally, [], {});
    assert.strictEqual(recs.length, tally.byFlat.size);
    recs.forEach((r) => {
      assert.strictEqual(r.hasAgreement, false);
      assert.strictEqual(r.gstRate, 0.01);
      assert.strictEqual(r.totalGst, Math.round(r.totalReceived * 0.01 * 100) / 100);
    });
  });
  ok('1804 excluded from tally-only when it has an agreement', () => {
    const recs = buildTallyOnlyRecords(tally, ['1804'], {});
    assert.ok(!recs.some((r) => r.flatNo === '1804'));
  });
} catch (e) {
  console.log('  (skipped Space Homes tally tests:', e.message, ')');
}

console.log('\nSunray Reality sample (if available)');
const sunrayTally = process.env.SUNRAY_TALLY || 'C:\\Users\\HP\\Downloads\\Tally bank receipt.xlsx';
try {
  const tally = parseTally(sunrayTally);
  ok('Sunray tally project name detected', () => assert.strictEqual(tally.projectName, 'SUNRAY REALTY'));
  ok('Sunray flat 904 tally has 2 receipt rows', () => assert.strictEqual(tally.byFlat.get('904').payments.length, 2));

  const rec904 = buildRecord(
    {
      flatNoHint: '904',
      agreementValue: 12303000,
      names: ['MR. ANKIT JAYESH VORA', 'MR. JAYESH SHANTILAL VORA'],
      valueCandidates: [],
    },
    tally,
    {
      name: 'ANKIT JAYESH VORA/JAYESH S VORA',
      floor: '9TH',
      paymentCutoff: '1/31/26',
      extraPayments: [{ date: '1/21/26', amount: 615150 }],
    },
  );
  ok('904 GST on agreement = 615,150 (5%)', () => assert.strictEqual(rec904.gstOnAgreement, 615150));
  ok('904 reference total received = 756,150', () => assert.strictEqual(rec904.totalReceived, 756150));
  ok('904 reference balance = 12,162,000', () => assert.strictEqual(rec904.balance, 12162000));
} catch (e) {
  console.log('  (skipped Sunray tally tests:', e.message, ')');
}

console.log(`\n${passed} checks passed`);
