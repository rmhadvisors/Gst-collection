'use strict';

/**
 * Monthly grouper — assigns payments and agreement executions to MMM-YY buckets.
 *
 * Input: records[] — full output of the collection engine (buildRecord +
 *        buildTallyOnlyRecords combined and sorted).
 *
 * Output: Map<'APR-25', MonthData>  (insertion-ordered by first occurrence)
 *
 * MonthData = {
 *   label:       'APR-25'       — sheet-name-safe label
 *   displayName: 'APRIL 2025'   — used in the sheet title row
 *   agreements:  AgreementEntry[]   — flats with an agreement executed this month
 *   advances:    AdvanceEntry[]     — flats with only tally advances this month
 *   totals: {
 *     amt1pct, gst1pct,   — 1% bracket totals (agreement + advance)
 *     amt5pct, gst5pct,   — 5% bracket totals
 *     totalGst,           — sum of all GST this month
 *   }
 * }
 *
 * AgreementEntry = {
 *   flatNo, name, agreementValue, rate, gstOnAgreement,
 *   advancesBeforeAgreement, reverseGstOnAdvances, outstandingGst,
 *   amountPaidThisMonth,   — sum of payments whose date falls in this month
 *   gstCase,              — 1 or 2
 *   agreementMonth,       — which month the agreement was executed in
 * }
 *
 * AdvanceEntry = {
 *   flatNo, name,
 *   amountThisMonth1pct, gst1pct,   — 1% advances received this month
 *   amountThisMonth5pct, gst5pct,   — 5% advances received this month
 * }
 */

const MONTH_NAMES = [
  'JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
  'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER',
];
const MONTH_ABBR = [
  'JAN','FEB','MAR','APR','MAY','JUN',
  'JUL','AUG','SEPT','OCT','NOV','DEC',
];

function monthKey(date) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return null;
  const m = date.getMonth(); // 0-based
  const y = date.getFullYear() % 100; // 2-digit year
  return `${MONTH_ABBR[m]}-${String(y).padStart(2, '0')}`;
}

function monthDisplayName(key) {
  // key like 'APR-25' → 'APRIL 2025'
  const [abbr, yr] = key.split('-');
  const idx = MONTH_ABBR.indexOf(abbr.toUpperCase());
  const name = idx >= 0 ? MONTH_NAMES[idx] : abbr;
  const fullYear = 2000 + Number(yr);
  return `${name} ${fullYear}`;
}

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

/**
 * Group a combined records array (from buildRecord + buildTallyOnlyRecords)
 * into monthly buckets suitable for generating the GST Working sheets.
 *
 * @param {object[]} records
 * @returns {Map<string, object>} ordered Map keyed by month label e.g. 'APR-25'
 */
function groupByMonth(records) {
  const months = new Map(); // insertion-ordered

  function ensureMonth(key) {
    if (!key) return null;
    if (!months.has(key)) {
      months.set(key, {
        label: key,
        displayName: monthDisplayName(key),
        agreements: [],
        advances: [],
      });
    }
    return months.get(key);
  }

  for (const rec of records) {
    if (rec.hasAgreement) {
      // ── Agreement flat (Case 1 or 2) ───────────────────────────────────────
      // The entry goes into the month when the agreement was executed.
      const agrMonth = monthKey(rec.agreementDateObj);
      if (!agrMonth) continue; // no date — skip from working sheet

      const monthData = ensureMonth(agrMonth);

      // Amount paid this month = payments whose date falls in this month
      const amountPaidThisMonth = round2(
        rec.payments
          .filter((p) => monthKey(p.date) === agrMonth)
          .reduce((s, p) => s + p.amount, 0),
      );

      monthData.agreements.push({
        flatNo: rec.flatNo,
        name: rec.name,
        agreementValue: rec.agreementValue,
        rate: rec.gstRate,
        gst1pct: rec.gstRate === 0.01 ? rec.gstOnAgreement : 0,
        gst5pct: rec.gstRate === 0.05 ? rec.gstOnAgreement : 0,
        advancesBeforeAgreement: rec.advancesBeforeAgreement || 0,
        reverseGstOnAdvances: rec.reverseGstOnAdvances || 0,
        outstandingGst: rec.outstandingGst != null ? rec.outstandingGst : (rec.gstOnAgreement || 0),
        amountPaidThisMonth,
        gstCase: rec.gstCase || 1,
        agreementMonth: agrMonth,
      });
    } else {
      // ── Tally-only flat (Case 3) ────────────────────────────────────────────
      // Each payment goes into its own month bucket.
      for (const p of rec.payments) {
        const pMonth = monthKey(p.date);
        if (!pMonth) continue;
        const monthData = ensureMonth(pMonth);

        // Find or create an advance entry for this flat in this month
        let entry = monthData.advances.find((a) => a.flatNo === rec.flatNo);
        if (!entry) {
          entry = {
            flatNo: rec.flatNo,
            name: rec.name,
            amountThisMonth1pct: 0,
            gst1pct: 0,
            amountThisMonth5pct: 0,
            gst5pct: 0,
          };
          monthData.advances.push(entry);
        }

        const rate = p.gstRate != null ? p.gstRate : 0.01;
        if (rate <= 0.01) {
          entry.amountThisMonth1pct = round2(entry.amountThisMonth1pct + p.amount);
          entry.gst1pct = round2(entry.gst1pct + (p.gst || round2(p.amount * rate)));
        } else {
          entry.amountThisMonth5pct = round2(entry.amountThisMonth5pct + p.amount);
          entry.gst5pct = round2(entry.gst5pct + (p.gst || round2(p.amount * rate)));
        }
      }
    }
  }

  // ── Compute per-month totals ───────────────────────────────────────────────
  for (const monthData of months.values()) {
    let amt1pct = 0, gst1pct = 0, amt5pct = 0, gst5pct = 0;

    // Agreement section
    for (const a of monthData.agreements) {
      if (a.rate === 0.01) {
        amt1pct = round2(amt1pct + (a.agreementValue - a.advancesBeforeAgreement));
        gst1pct = round2(gst1pct + a.outstandingGst);
      } else {
        amt5pct = round2(amt5pct + (a.agreementValue - a.advancesBeforeAgreement));
        gst5pct = round2(gst5pct + a.outstandingGst);
      }
    }

    // Advance section
    for (const a of monthData.advances) {
      amt1pct = round2(amt1pct + a.amountThisMonth1pct);
      gst1pct = round2(gst1pct + a.gst1pct);
      amt5pct = round2(amt5pct + a.amountThisMonth5pct);
      gst5pct = round2(gst5pct + a.gst5pct);
    }

    monthData.totals = {
      amt1pct,
      gst1pct,
      amt5pct,
      gst5pct,
      totalGst: round2(gst1pct + gst5pct),
    };
  }

  // Sort months chronologically
  const sorted = new Map(
    [...months.entries()].sort(([a], [b]) => monthSortKey(a) - monthSortKey(b)),
  );

  return sorted;
}

/** Convert a month key to a numeric sort value. */
function monthSortKey(key) {
  const [abbr, yr] = key.split('-');
  const m = MONTH_ABBR.indexOf(abbr.toUpperCase());
  return (2000 + Number(yr)) * 12 + (m >= 0 ? m : 0);
}

module.exports = { groupByMonth, monthKey, monthDisplayName, MONTH_ABBR, MONTH_NAMES };
