'use strict';

const { fuzzyNameSimilarity } = require('../util/text');
const { formatMDY } = require('../util/dates');
const { compareFlatKeys } = require('../util/flatKey');

const GUESS_MIN = 0.75;

const inr = (n) => `₹${Number(n).toLocaleString('en-IN', { maximumFractionDigits: 2 })}`;

/** Audit flags for a flat's payments: duplicate vouchers, manual allocations, GST receipts. */
function voucherFlags(payments) {
  const list = payments || [];
  const flags = list
    .filter((p) => p.dupVoucher)
    .map((p) => `Possible duplicate: Vch ${p.vchNo} (${p.amount}) appears more than once`);
  for (const p of list.filter((x) => x.noVoucher)) {
    flags.push(`No voucher number (${inr(p.amount)}${p.date ? ` on ${formatMDY(p.date)}` : ''})`);
  }
  const oldFlats = [...new Set(list.map((x) => x.renumberedFrom).filter(Boolean))];
  for (const from of oldFlats) flags.push(`Renumbered: receipts booked in Tally under old flat ${from}`);
  for (const p of list.filter((x) => x.chqReturn)) {
    flags.push(`Vch ${p.vchNo} ${inr(p.amount)} is a CHQ RETURN booked as a receipt`);
  }
  for (const p of list.filter((x) => x.allocatedBy)) {
    flags.push(`Vch ${p.vchNo} ${inr(p.amount)}: ${/^allocated/.test(p.allocatedBy) ? '' : 'allocated - '}${p.allocatedBy} (Tally book has no building)`);
  }
  for (const p of list.filter((x) => x.manualAllocation)) {
    flags.push(`Vch ${p.vchNo} ${inr(p.amount)} from "${p.allocatedFrom}" (manual allocation)`);
  }
  for (const p of list.filter((x) => x.gstReceipt)) {
    flags.push(`Vch ${p.vchNo} ${inr(p.amount)} GST receipt (equals GST on agreement value; no GST charged)`);
  }
  return flags;
}

/**
 * Receipts whose narration carries no flat number. They are listed, never
 * dropped and never auto-added to a flat; `guess` is only a hint (best fuzzy
 * purchaser-name match against Tally flat names and agreement purchasers).
 */
function buildUnallocated(tally, records = []) {
  const candidates = [];
  const label = (building, flat) => (building ? `${building} ${flat}` : flat);
  for (const g of tally.byFlat.values()) {
    for (const n of [...g.nameList, ...(g.manualNames || [])]) candidates.push({ flatNo: g.flatNo, name: n, label: label(g.building, g.flatNo) });
  }
  for (const r of records) {
    for (const n of (r.allottees && r.allottees.length ? r.allottees : [r.name])) {
      if (n && r.flatNo) candidates.push({ flatNo: r.flatNo, name: n, label: label(r.building, r.flatNo) });
    }
  }

  return (tally.unallocated || []).map((row) => {
    const party = row.name.replace(/capital\s*a\/?c/i, '').trim();
    let best = null;
    for (const c of candidates) {
      const score = fuzzyNameSimilarity(party, c.name);
      if (!best || score > best.score || (score === best.score && compareFlatKeys(c.flatNo, best.flatNo) < 0)) {
        best = { ...c, score };
      }
    }
    const hit = best && best.score >= GUESS_MIN;
    return {
      date: row.date,
      dateText: row.date ? formatMDY(row.date) : '',
      party: row.name,
      vchNo: row.vchNo,
      amount: row.amount,
      // best name match; failing that the flat number the narration gave (its building is unknown)
      guess: hit ? best.label : (row.guessFlat || ''),
      guessScore: hit ? Math.round(best.score * 100) / 100 : '',
      flag: [
        hit ? '(review) best flat guess by name - NOT allocated' : '',
        row.buildingReview ? '(review) building of this Tally book is not known - flat number kept as the guess' : '',
        row.dupVoucher ? 'Possible duplicate' : '',
        !row.vchNo ? 'No voucher number' : '',
        row.chqReturn ? 'CHQ RETURN booked as receipt' : '',
      ].filter(Boolean).join('; '),
    };
  });
}

module.exports = { buildUnallocated, voucherFlags };
