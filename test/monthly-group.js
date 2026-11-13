'use strict';

/**
 * groupByMonth: every receipt lands in the month it was paid, for ANY data.
 * Uses made-up records only (no client fixture), so it guards the rule itself.
 */
const assert = require('assert');
const { groupByMonth } = require('../src/engine/monthlyGroup');

const d = (y, m, day) => new Date(y, m - 1, day);
const pay = (date, amount, extra = {}) => ({ date, amount, gst: Math.round(amount) / 100, gstRate: 0.01, ...extra });
let n = 0;
function check(name, fn) { fn(); n++; console.log('  ok  - ' + name); }

const records = [
  // agreement flat: two advances before the agreement (Jan, Feb), one after (Apr)
  {
    hasAgreement: true, flatNo: 'A1', name: 'Buyer One', agreementValue: 4000000, gstRate: 0.01,
    gstOnAgreement: 40000, agreementDateObj: d(2025, 3, 10), advancesBeforeAgreement: 300000,
    reverseGstOnAdvances: 3000, outstandingGst: 37000,
    payments: [pay(d(2025, 1, 5), 100000), pay(d(2025, 2, 5), 200000), pay(d(2025, 4, 5), 50000)],
  },
  // agreement flat whose date could not be read
  {
    hasAgreement: true, flatNo: 'A2', name: 'Buyer Two', agreementValue: 3000000, gstRate: 0.01,
    gstOnAgreement: 30000, agreementDateObj: null, payments: [pay(d(2025, 2, 9), 70000)],
  },
  // tally-only flats (no agreement yet)
  { hasAgreement: false, flatNo: 'T1', name: 'Tally One', payments: [pay(d(2025, 1, 7), 10000), pay(d(2025, 3, 7), 20000)] },
  { hasAgreement: false, flatNo: 'T2', name: 'Tally Two', payments: [pay(d(2025, 3, 8), 30000, { gstRate: 0.05, gst: 1500 })] },
  // receipt with no flat no. (parties share the empty flat no.)
  { hasAgreement: false, noFlat: true, flatNo: '', name: 'Party X', payments: [pay(d(2025, 2, 11), 5000)] },
  { hasAgreement: false, noFlat: true, flatNo: '', name: 'Party Y', payments: [pay(d(2025, 2, 12), 6000, { gst: 0, gstRate: null })] },
];

console.log('Monthly grouping (generic data)');
const months = groupByMonth(records);
const adv = (label) => (months.get(label) ? months.get(label).advances : []);
const sum = (label) => adv(label).reduce((s, a) => s + a.amountThisMonth1pct + a.amountThisMonth5pct, 0);

check('months are created and sorted for every payment month', () => {
  assert.deepStrictEqual([...months.keys()], ['JAN-25', 'FEB-25', 'MAR-25']);
});
check('tally-only receipts sit in the month they were paid', () => {
  assert.ok(adv('JAN-25').some((a) => a.flatNo === 'T1' && a.amountThisMonth1pct === 10000));
  assert.ok(adv('MAR-25').some((a) => a.flatNo === 'T1' && a.amountThisMonth1pct === 20000));
  assert.ok(adv('MAR-25').some((a) => a.flatNo === 'T2' && a.amountThisMonth5pct === 30000 && a.gst5pct === 1500));
});
check('receipts before an agreement are advances in their own month', () => {
  assert.ok(adv('JAN-25').some((a) => a.flatNo === 'A1' && a.amountThisMonth1pct === 100000));
  assert.ok(adv('FEB-25').some((a) => a.flatNo === 'A1' && a.amountThisMonth1pct === 200000));
});
check('receipts after the agreement are not taxed twice', () => {
  assert.ok(!months.has('APR-25'));
  assert.strictEqual(months.get('MAR-25').agreements.length, 1);
});
check('agreement with no readable date still reports its receipts', () => {
  assert.ok(adv('FEB-25').some((a) => a.flatNo === 'A2' && a.amountThisMonth1pct === 70000));
});
check('no-flat parties stay separate and honour zero GST', () => {
  const x = adv('FEB-25').filter((a) => a.flatNo === '');
  assert.strictEqual(x.length, 2);
  assert.strictEqual(x.find((a) => a.name === 'Party Y').gst1pct, 0);
});
check('nothing is lost: advance amounts add up to the receipts expected', () => {
  assert.strictEqual(sum('JAN-25'), 110000);
  assert.strictEqual(sum('FEB-25'), 200000 + 70000 + 5000 + 6000);
  assert.strictEqual(sum('MAR-25'), 50000);
});

console.log(`\n${n} monthly-group checks passed`);
