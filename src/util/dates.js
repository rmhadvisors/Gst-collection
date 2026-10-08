'use strict';

const MONTHS = {
  jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
  jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
};

/**
 * Parse a date from the many shapes found in the inputs:
 *  - JS Date (from xlsx cellDates)
 *  - "26-Dec-25" / "1-Apr-25" (tally)
 *  - "10/03/2026" (dd/mm/yyyy, registration receipt)
 *  - "3/10/26" (m/d/yy)
 * Returns a JS Date (local) or null.
 */
function parseDate(value, opts = {}) {
  if (value == null || value === '') return null;
  if (value instanceof Date && !isNaN(value)) return value;

  const s = String(value).trim();

  // dd-MMM-yy  e.g. 26-Dec-25
  let m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})$/);
  if (m) {
    const d = parseInt(m[1], 10);
    const mo = MONTHS[m[2].toLowerCase()];
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    if (mo != null) return new Date(y, mo, d);
  }

  // numeric separators d/m/y or m/d/y
  m = s.match(/^(\d{1,4})[\/\-.](\d{1,2})[\/\-.](\d{1,4})$/);
  if (m) {
    let a = parseInt(m[1], 10);
    let b = parseInt(m[2], 10);
    let c = parseInt(m[3], 10);
    // If first part is a 4-digit year -> y/m/d
    if (a > 31) {
      return new Date(a, b - 1, c);
    }
    let year = c;
    if (year < 100) year += 2000;
    // dayFirst: registration receipts are dd/mm/yyyy
    if (opts.dayFirst) {
      return new Date(year, b - 1, a);
    }
    // default m/d/y, but if a>12 it must be day-first
    if (a > 12) return new Date(year, b - 1, a);
    return new Date(year, a - 1, b);
  }

  const d = new Date(s);
  return isNaN(d) ? null : d;
}

/** Format as m/d/yy to match the collection sheet style (e.g. 3/10/26). */
function formatMDY(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const m = date.getMonth() + 1;
  const d = date.getDate();
  const y = String(date.getFullYear()).slice(-2);
  return `${m}/${d}/${y}`;
}

module.exports = { parseDate, formatMDY };
