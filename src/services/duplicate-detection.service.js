'use strict';

/**
 * Duplicate Detection Service
 *
 * Identifies whether a given payment already exists in the known dataset,
 * using a composite fingerprint key:
 *
 *   Primary key (when voucher number is available):
 *     flatNo + "|" + vchNo + "|" + amount
 *
 *   Fallback key (no voucher number):
 *     flatNo + "|" + normalizedDate + "|" + amount
 *
 * This matches the CA's deduplication logic where:
 *   - Same flat + same voucher + same amount → definite duplicate
 *   - Same flat + same date + same amount → probable duplicate (safe to skip)
 */

/**
 * Normalise a date string to YYYY-MM-DD for consistent comparison.
 * Handles formats: "26-Dec-25", "12/26/25", "2025-12-26", Date objects.
 *
 * IMPORTANT: uses local date parts (not toISOString) to avoid UTC shift.
 *
 * @param {string|Date} d
 * @returns {string}  YYYY-MM-DD or empty string if unparseable
 */
function normaliseDate(d) {
  if (!d) return '';

  function pad(n) { return String(n).padStart(2, '0'); }
  function localISO(dt) {
    if (!dt || isNaN(dt.getTime())) return '';
    return `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}`;
  }

  if (d instanceof Date) return localISO(d);

  const s = String(d).trim();
  if (!s) return '';

  // dd-Mon-yy  e.g. "26-Dec-25" — handle FIRST before native parse
  const m1 = s.match(/^(\d{1,2})-([A-Za-z]{3,9})-(\d{2,4})$/);
  if (m1) {
    const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const month = months[m1[2].toLowerCase().slice(0,3)];
    if (month != null) {
      const yr = Number(m1[3]);
      const year = yr < 100 ? 2000 + yr : yr;
      return `${year}-${pad(month + 1)}-${pad(Number(m1[1]))}`;
    }
  }

  // dd/mm/yyyy or mm/dd/yyyy — treat as local date by splitting manually
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m2) {
    const yr = Number(m2[3]) < 100 ? 2000 + Number(m2[3]) : Number(m2[3]);
    // Assume MM/DD/YY (US format used by existing codebase)
    return localISO(new Date(yr, Number(m2[1]) - 1, Number(m2[2])));
  }

  // YYYY-MM-DD — already normalised
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // Fallback: native Date (may have UTC issues for ambiguous formats)
  const dt = new Date(s);
  return localISO(dt);
}

/**
 * Build a composite fingerprint key for one payment.
 * Primary: flatNo|vch:vchNo|amount  (when vchNo available)
 * Fallback: flatNo|YYYY-MM-DD|amount (when no vchNo)
 *
 * @param {string|number} flatNo
 * @param {object}        payment  { dateText, date, amount, vchNo }
 * @returns {string}
 */
function paymentKey(flatNo, payment) {
  const flat = String(flatNo).trim();
  const amt  = Number(payment.amount || 0).toFixed(2);
  if (payment.vchNo) {
    return `${flat}|vch:${payment.vchNo}|${amt}`;
  }
  const date = normaliseDate(payment.dateText || payment.date);
  return `${flat}|${date}|${amt}`;
}

/**
 * Build a date-only key (no vchNo) — used as fallback when the Wing B
 * file doesn't carry vchNo data but the tally does.
 *
 * @param {string|number} flatNo
 * @param {object}        payment  { dateText, date, amount }
 * @returns {string}
 */
function dateKey(flatNo, payment) {
  const flat = String(flatNo).trim();
  const amt  = Number(payment.amount || 0).toFixed(2);
  const date = normaliseDate(payment.dateText || payment.date);
  return `${flat}|${date}|${amt}`;
}

/**
 * Build a Set of existing payment fingerprints from an array of
 * { flatNo, payments: [{dateText, amount, vchNo}] } blocks.
 *
 * For each payment we store BOTH the primary key AND the date-based fallback.
 * This ensures a tally payment (which has vchNo) matches an existing Wing B
 * payment (which has no vchNo) as long as the date + amount align.
 *
 * @param {object[]} existingBlocks  array of { flatNo, payments: [] }
 * @returns {Set<string>}
 */
function buildKeySet(existingBlocks) {
  const keys = new Set();
  for (const block of existingBlocks) {
    for (const p of (block.payments || [])) {
      keys.add(paymentKey(block.flatNo, p));  // primary key
      keys.add(dateKey(block.flatNo, p));     // date fallback (always add)
    }
  }
  return keys;
}

/**
 * Filter a flat's payment list to only NEW payments not present in keySet.
 *
 * Checks BOTH the primary key (vch-based if vchNo present) AND the date-
 * based fallback. A payment is a duplicate if EITHER key is in the set.
 * This prevents re-adding tally payments that are already in Wing B even
 * when the Wing B file didn't store vchNo.
 *
 * @param {string|number} flatNo
 * @param {object[]}      payments  array of payment objects
 * @param {Set<string>}   keySet    existing fingerprints (from buildKeySet)
 * @returns {object[]}  only genuinely new payments
 */
function filterNew(flatNo, payments, keySet) {
  return payments.filter(p => {
    if (keySet.has(paymentKey(flatNo, p))) return false;  // vch or date key matches
    if (keySet.has(dateKey(flatNo, p)))    return false;  // date fallback matches
    return true;  // genuinely new
  });
}

/**
 * Check whether a single payment is a duplicate.
 *
 * @param {string|number} flatNo
 * @param {object}        payment
 * @param {Set<string>}   keySet
 * @returns {boolean}
 */
function isDuplicate(flatNo, payment, keySet) {
  return keySet.has(paymentKey(flatNo, payment)) || keySet.has(dateKey(flatNo, payment));
}

module.exports = { paymentKey, dateKey, buildKeySet, filterNew, isDuplicate, normaliseDate };
