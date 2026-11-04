'use strict';

/**
 * Snapshot of the Space Home sample (every flat in config/samples/tally receipts.xlsx)
 * as per-flat totals and payment lists. Used to prove output is unchanged:
 *   the baseline lives in test/fixtures/space-home/expected.json (taken from the pre-change code)
 */
const path = require('path');
const root = process.env.SNAPSHOT_ROOT || path.join(__dirname, '..');
const { parseTally } = require(path.join(root, 'src/tally/parse'));
const { buildTallyOnlyRecords } = require(path.join(root, 'src/engine/tallyOnly'));
const { buildRecord } = require(path.join(root, 'src/engine/collection'));

function snapshot(tallyPath) {
  const tally = parseTally(tallyPath);
  const recs = buildTallyOnlyRecords(tally, [], {}, {});
  // 1804 as an agreement flat (value from the CA reference), the rest tally-only
  const agr = buildRecord({ flatNoHint: '1804', agreementValue: 4490000, names: [], valueCandidates: [] }, tally, {});
  const all = recs.filter((r) => r.flatNo !== '1804').concat(agr);
  const out = {};
  for (const r of all.sort((a, b) => Number(a.flatNo) - Number(b.flatNo))) {
    out[r.flatNo] = {
      name: r.name,
      totalReceived: r.totalReceived,
      totalGst: r.totalGst,
      balance: r.balance,
      agreementValue: r.agreementValue,
      payments: r.payments.map((p) => [p.dateText, p.amount, p.gst, p.vchNo]),
    };
  }
  return out;
}

if (require.main === module) {
  process.stdout.write(JSON.stringify(snapshot(path.join(root, 'config/samples/tally receipts.xlsx')), null, 1));
}
module.exports = { snapshot };
