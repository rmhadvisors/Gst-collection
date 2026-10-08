'use strict';

const XLSX = require('xlsx');
const { parseDate } = require('../util/dates');
const { parseAmount } = require('../util/text');

/**
 * Parse a Tally "Receipt" export sheet into individual payment rows.
 *
 * Expected columns (after the header row that contains "Date" & "Particulars"):
 *   A Date | B "To" | C Particulars (e.g. "Flat No.1804 Mrs. Chhaya T. Kadam")
 *   F Vch Type | G Vch No. | H Debit (amount) | I Credit
 *
 * @returns {{ rows: Array, byFlat: Map<string, object> }}
 */
function parseTally(filePath, sheetName) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const ws = wb.Sheets[sheetName || wb.SheetNames[0]];
  // raw:false returns the cell's *formatted* text (e.g. "26-Dec-25", "1,100,000.00").
  // We parse those strings ourselves to avoid the xlsx Date<->timezone off-by-one issue.
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });

  // locate header row
  let headerRow = -1;
  for (let i = 0; i < grid.length; i += 1) {
    const joined = grid[i].map((c) => String(c).toLowerCase()).join('|');
    if (joined.includes('particulars') && joined.includes('date')) {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) headerRow = 10; // fallback to observed layout

  // Optional GST column: future Tally exports may carry a per-receipt GST
  // rate/percentage. Detect it from the header; when absent (current files),
  // the engine falls back to the default 1% rate.
  const gstCol = headerRow >= 0 ? findGstColumn(grid[headerRow]) : -1;

  const rows = [];
  const flatRe = /flat\s*no\.?\s*(\d{3,4})/i;

  for (let i = headerRow + 1; i < grid.length; i += 1) {
    const r = grid[i];
    const dateRaw = r[0];
    const particulars = String(r[2] || '').trim();
    const vchNo = r[6] != null && r[6] !== '' ? String(r[6]).trim() : '';
    const amount = parseAmount(r[7]); // Debit column
    const gstRate = gstCol >= 0 ? gstRateFromCell(r[gstCol]) : null;

    if (!particulars) continue;
    // skip totals / closing-balance footer rows
    if (/closing balance/i.test(particulars)) continue;

    const date = parseDate(dateRaw);
    const fm = particulars.match(flatRe);
    const flatNo = fm ? fm[1] : null;

    // strip the "Flat No.xxxx" prefix to get the buyer name
    const name = particulars.replace(flatRe, '').replace(/^[\s.:-]+/, '').trim();

    if (amount == null || amount <= 1) continue; // ignore re.1 token entries

    rows.push({
      rowIndex: i + 1,
      date,
      particulars,
      flatNo,
      name,
      vchNo,
      amount,
      gstRate,
    });
  }

  // group by flat number
  const byFlat = new Map();
  for (const row of rows) {
    if (!row.flatNo) continue;
    if (!byFlat.has(row.flatNo)) {
      byFlat.set(row.flatNo, { flatNo: row.flatNo, names: new Set(), payments: [] });
    }
    const g = byFlat.get(row.flatNo);
    g.names.add(row.name);
    g.payments.push({ date: row.date, amount: row.amount, vchNo: row.vchNo, gstRate: row.gstRate });
  }

  // Deduplicate within each flat:
  //
  // A "bounced cheque" in Tally appears as the same voucher number re-used
  // on the SAME date with the same amount (the operator corrects a mis-entry).
  // Two rows with the same vchNo + amount on DIFFERENT dates are genuine
  // separate payments (the CA used the same voucher sequence across months).
  //
  // Dedup key: vchNo + amount + date (all three must match for it to be a dup).
  for (const g of byFlat.values()) {
    const vchGroups = new Map();
    for (const p of g.payments) {
      if (!p.vchNo) continue;
      const dateStr = p.date ? p.date.toISOString().slice(0, 10) : 'nodate';
      const key = `${p.vchNo}|${p.amount}|${dateStr}`;
      if (!vchGroups.has(key)) vchGroups.set(key, []);
      vchGroups.get(key).push(p);
    }
    const toRemove = new Set();
    for (const dupes of vchGroups.values()) {
      if (dupes.length <= 1) continue;
      // True duplicates (same vchNo + amount + date): keep only one
      for (let i = 1; i < dupes.length; i += 1) {
        toRemove.add(dupes[i]);
      }
    }
    if (toRemove.size) {
      g.payments = g.payments.filter((p) => !toRemove.has(p));
    }
  }


  // sort payments by date
  for (const g of byFlat.values()) {
    g.payments.sort((a, b) => {
      const da = a.date ? a.date.getTime() : 0;
      const db = b.date ? b.date.getTime() : 0;
      return da - db;
    });
    g.nameList = Array.from(g.names);
  }

  return { rows, byFlat, projectName: detectProjectName(grid) };
}

/**
 * Find the column index of a GST rate/percentage column in the header row.
 * Returns -1 when no such column exists (the common case today).
 */
function findGstColumn(headerCells) {
  if (!Array.isArray(headerCells)) return -1;
  return headerCells.findIndex((c) => /gst/i.test(String(c || '')));
}

/**
 * Interpret a GST cell into a fractional rate.
 *   "1%" / "1" / 1      -> 0.01
 *   "5%" / "5" / 5      -> 0.05
 *   0.01               -> 0.01 (already a fraction)
 * Returns null when the cell is empty/unparseable.
 */
function gstRateFromCell(raw) {
  if (raw == null || raw === '') return null;
  const s = String(raw).trim();
  const num = parseFloat(s.replace(/[^\d.]/g, ''));
  if (!Number.isFinite(num)) return null;
  if (s.includes('%')) return num / 100;
  // A value of 1 or more is a percentage (1 -> 1%, 5 -> 5%); below 1 is already
  // a fraction (0.01 -> 1%).
  return num >= 1 ? num / 100 : num;
}

/** Read project/building label from the first header rows (e.g. "SUNRAY REALTY"). */
function detectProjectName(grid) {
  for (let i = 0; i < Math.min(grid.length, 6); i += 1) {
    const cell = String(grid[i][0] || '').trim();
    if (!cell) continue;
    if (/^\d+-[A-Za-z]{3}-\d+/i.test(cell)) continue; // skip "1-Apr-25 to ..."
    if (/book$/i.test(cell)) continue;
    if (/date|particulars/i.test(cell)) continue;
    return cell;
  }
  return '';
}

/**
 * Merge multiple parsed Tally results, rebuilding the flat mapping and
 * deduplicating voucher entries across all sheets.
 */
function mergeTallies(tallies) {
  if (!tallies || !tallies.length) {
    return { rows: [], byFlat: new Map(), projectName: '' };
  }

  const combinedRows = [];
  const projectNames = [];

  for (const t of tallies) {
    if (t.rows) {
      combinedRows.push(...t.rows);
    }
    if (t.projectName) {
      projectNames.push(t.projectName);
    }
  }

  // Reconstruct byFlat map from all combined rows to apply deduplication and sorting
  const byFlat = new Map();
  for (const row of combinedRows) {
    if (!row.flatNo) continue;
    const flatStr = String(row.flatNo);
    if (!byFlat.has(flatStr)) {
      byFlat.set(flatStr, { flatNo: row.flatNo, names: new Set(), payments: [] });
    }
    const g = byFlat.get(flatStr);
    g.names.add(row.name);
    g.payments.push({ date: row.date, amount: row.amount, vchNo: row.vchNo, gstRate: row.gstRate });
  }

  // Deduplicate by voucher number within each flat across all combined rows.
  // Same rule as parseTally: only vchNo + amount + same date = true duplicate.
  for (const g of byFlat.values()) {
    const vchGroups = new Map();
    for (const p of g.payments) {
      if (!p.vchNo) continue;
      const dateStr = p.date ? p.date.toISOString().slice(0, 10) : 'nodate';
      const key = `${p.vchNo}|${p.amount}|${dateStr}`;
      if (!vchGroups.has(key)) vchGroups.set(key, []);
      vchGroups.get(key).push(p);
    }
    const toRemove = new Set();
    for (const dupes of vchGroups.values()) {
      if (dupes.length <= 1) continue;
      for (let i = 1; i < dupes.length; i += 1) {
        toRemove.add(dupes[i]);
      }
    }
    if (toRemove.size) {
      g.payments = g.payments.filter((p) => !toRemove.has(p));
    }
  }


  // Sort payments and convert names Set to Array
  for (const g of byFlat.values()) {
    g.payments.sort((a, b) => {
      const da = a.date ? a.date.getTime() : 0;
      const db = b.date ? b.date.getTime() : 0;
      return da - db;
    });
    g.nameList = Array.from(g.names);
  }

  // Clean/unique project names
  const uniqueNames = Array.from(new Set(projectNames.map(name => name.trim()).filter(Boolean)));
  const projectName = uniqueNames.join(' & ') || '';

  return {
    rows: combinedRows,
    byFlat,
    projectName,
  };
}

module.exports = { parseTally, mergeTallies, gstRateFromCell, findGstColumn };
