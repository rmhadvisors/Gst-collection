'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const XLSX = require('xlsx');
const { parseDate, formatMDY } = require('../util/dates');
const { parseAmount, fuzzyNameSimilarity } = require('../util/text');
const { canonicalBuilding, buildingInText } = require('../util/building');
const {
  parseFlatFromParticulars, normalizeFlatKey, cleanPartyName, hasChqReturn, renumberMap, groupKey,
} = require('../util/flatKey');

/** Index of the first header cell matching `re` (-1 if none). */
function colIndex(headerCells, re) {
  return headerCells.findIndex((c) => re.test(String(c || '').trim()));
}

/** Header row = the row holding both "Date" and "Vch Type" (older exports: "Date" + "Particulars"). */
function findHeaderRow(grid) {
  const has = (row, re) => row.some((c) => re.test(String(c || '').trim()));
  for (let i = 0; i < grid.length; i += 1) {
    if (has(grid[i], /^date$/i) && has(grid[i], /^vch\.?\s*type$/i)) return i;
  }
  for (let i = 0; i < grid.length; i += 1) {
    if (has(grid[i], /^date$/i) && has(grid[i], /^particulars$/i)) return i;
  }
  return -1;
}

/** Flat key + cleaned buyer name + trailing-note facts from a party cell. */
function readParty(party) {
  const chqReturn = hasChqReturn(party);
  let flatNo = null;
  let rest = party;
  const lead = parseFlatFromParticulars(party);
  if (lead) {
    flatNo = lead.key; rest = lead.rest;
  } else {
    const inner = party.match(/(?:flat|shop)\s*(?:no\.?)?\s*([A-Za-z])?\s*[-–.]?\s*(\d{1,4})\b/i);
    if (inner) {
      flatNo = parseFlatFromParticulars(inner[0]).key;
      rest = party.replace(inner[0], '').replace(/^[\s.:-]+/, '').trim();
    }
  }
  const { name, agmtDate } = cleanPartyName(rest);
  return {
    flatNo, name, chqReturn, agmtDate: agmtDate ? parseDate(agmtDate, { dayFirst: true }) : null,
  };
}

/** Rows of one worksheet; null when the sheet has no Date + Vch Type (or Particulars) header. */
function parseSheet(ws) {
  // raw:false returns the cell's *formatted* text (e.g. "26-Dec-25", "1,100,000.00").
  // We parse those strings ourselves to avoid the xlsx Date<->timezone off-by-one issue.
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
  const headerRow = findHeaderRow(grid);
  if (headerRow === -1) return null;
  const header = grid[headerRow];

  const dateCol = Math.max(colIndex(header, /^date$/i), 0);
  const vchTypeCol = colIndex(header, /^vch\.?\s*type$/i);
  const vchNoCol = colIndex(header, /^vch\.?\s*no/i);
  const debitCol = colIndex(header, /^debit/i);
  const creditCol = colIndex(header, /^credit/i);
  const partyEnd = vchTypeCol > dateCol ? vchTypeCol : dateCol + 4;
  // Optional GST column: absent in today's exports -> engine falls back to 1%.
  const gstCol = findGstColumn(header);

  const rows = [];
  for (let i = headerRow + 1; i < grid.length; i += 1) {
    const r = grid[i];
    if (vchTypeCol >= 0 && !/receipt/i.test(String(r[vchTypeCol] || ''))) continue;
    const party = r.slice(dateCol + 1, partyEnd)
      .map((c) => String(c == null ? '' : c).trim())
      .find((c) => c && !/^(to|by|cr|dr)$/i.test(c)) || '';
    if (!party) continue;
    if (/^[\d,.\s]+$/.test(party)) continue; // subtotal lines carry only a number
    if (/^(opening|closing)\s*balance|^(grand\s*)?total/i.test(party)) continue;

    const debit = debitCol >= 0 ? parseAmount(r[debitCol]) : parseAmount(r[7]);
    const credit = creditCol >= 0 ? parseAmount(r[creditCol]) : parseAmount(r[8]);
    const amount = debit != null && debit > 0 ? debit : credit;
    // Every receipt is kept (including Re.1 token entries) so flat totals equal Tally exactly.
    if (amount == null || amount <= 0) continue;

    const vchRaw = r[vchNoCol >= 0 ? vchNoCol : 6];
    const vchNo = vchRaw != null && vchRaw !== '' ? String(vchRaw).trim() : '';
    const {
      flatNo, name, chqReturn, agmtDate,
    } = readParty(party);

    rows.push({
      rowIndex: i + 1,
      date: parseDate(r[dateCol], { dayFirst: true }),
      particulars: party,
      flatNo,
      name,
      vchNo,
      amount,
      chqReturn,
      agmtDate,
      gstRate: gstCol >= 0 ? gstRateFromCell(r[gstCol]) : null,
      isCapital: !flatNo && /capital\s*a\/?c/i.test(party),
    });
  }
  // everything above the header row ("H D F C BANK 44444 Siddhivinayak Palace  Book", ...)
  const title = grid.slice(0, headerRow).map((r) => String(r[0] == null ? '' : r[0]).trim()).filter(Boolean).join(' | ');
  return { rows, projectName: detectProjectName(grid), title };
}

const round2 = (n) => Math.round(n * 100) / 100;

/**
 * Parse a Tally "Receipt" export: EVERY sheet, whatever its name (or just
 * `sheetName` when given). Header row = "Date" + "Vch Type"; only Vch Type =
 * Receipt rows count; amount = Debit, else Credit; Opening/Closing Balance and
 * total rows are skipped. `label` is the file name shown in the audit.
 *
 * @returns {{ rows, byFlat, ledgers, unallocated, duplicateVouchers, fileStats }}
 */
function parseTally(filePath, sheetName, label) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const names = sheetName && wb.SheetNames.includes(sheetName) ? [sheetName] : wb.SheetNames;
  const rows = [];
  const sheets = [];
  const projects = [];
  const titles = [];
  for (const name of names) {
    const parsed = parseSheet(wb.Sheets[name]);
    if (!parsed) { sheets.push({ name, receipts: 0, total: 0 }); continue; }
    parsed.rows.forEach((r) => { r.sheet = name; });
    rows.push(...parsed.rows);
    if (parsed.projectName) projects.push(parsed.projectName);
    if (parsed.title) titles.push(parsed.title);
    sheets.push({ name, receipts: parsed.rows.length, total: round2(parsed.rows.reduce((s, r) => s + r.amount, 0)) });
  }
  const hash = crypto.createHash('sha1').update(fs.readFileSync(filePath)).digest('hex');
  const file = label || path.basename(filePath);
  rows.forEach((r) => { r.file = file; });
  const fileStats = [{
    file, hash, sheets, title: titles.join(' | '), receipts: rows.length, total: round2(rows.reduce((s, r) => s + r.amount, 0)),
  }];
  return {
    rows, ...groupRows(rows), fileStats, projectName: [...new Set(projects)][0] || '',
  };
}

function sortByDate(payments) {
  payments.sort((a, b) => (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0));
}

/**
 * Group rows by flat key ("B-607" / "1804") and, for ledgers without a flat
 * (e.g. "Riyaz AHmed Khan Capital Ac"), by ledger name.
 */
function groupRows(rows) {
  const byFlat = new Map();
  const ledgers = new Map();
  const duplicateVouchers = findDuplicateVouchers(rows);
  const dupKeys = new Set(duplicateVouchers.map((d) => d.key));
  for (const row of rows) {
    const payment = {
      date: row.date, amount: row.amount, vchNo: row.vchNo, gstRate: row.gstRate,
      dupVoucher: dupKeys.has(voucherKey(row)),
      noVoucher: !row.vchNo,
      renumberedFrom: row.renumberedFrom || '',
      chqReturn: !!row.chqReturn,
      allocatedBy: row.allocatedBy || '',
    };
    if (row.flatNo) {
      const key = groupKey(row.building, row.flatNo);
      if (!byFlat.has(key)) {
        byFlat.set(key, {
          key, flatNo: String(row.flatNo), building: row.building || '', names: new Set(), payments: [],
        });
      }
      const g = byFlat.get(key);
      g.names.add(row.name);
      if (row.agmtDate && !g.agmtDate) g.agmtDate = row.agmtDate;
      g.payments.push(payment);
    } else if (row.isCapital) {
      if (!ledgers.has(row.name)) ledgers.set(row.name, { name: row.name, isCapital: true, payments: [] });
      ledgers.get(row.name).payments.push(payment);
    }
  }
  for (const g of byFlat.values()) {
    sortByDate(g.payments);
    g.nameList = Array.from(g.names);
  }
  for (const l of ledgers.values()) {
    sortByDate(l.payments);
  }
  // Receipts with no flat number: never dropped, never auto-allocated.
  const unallocated = rows.filter((r) => !r.flatNo);
  for (const r of unallocated) r.dupVoucher = dupKeys.has(voucherKey(r));
  return { byFlat, ledgers, unallocated, duplicateVouchers };
}

/** Vch numbers restart in every Tally book, so a repeat only counts within the same file. A blank Vch No never matches. */
function voucherKey(r) {
  return r.vchNo ? `${r.file || ''}|${r.vchNo}|${r.amount}` : '';
}

/** Same Vch No + same amount appearing more than once in a book: kept, flagged "possible duplicate". */
function findDuplicateVouchers(rows) {
  const groups = new Map();
  for (const r of rows) {
    const key = voucherKey(r);
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }
  const out = [];
  for (const [key, list] of groups) {
    if (list.length < 2) continue;
    out.push({
      key,
      vchNo: list[0].vchNo,
      amount: list[0].amount,
      dates: [...new Set(list.map((r) => (r.date ? r.date.toISOString().slice(0, 10) : 'nodate')))],
      parties: [...new Set(list.map((r) => r.particulars))],
    });
  }
  return out;
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
 * Manual voucher allocation (overrides.json -> "allocate": {"81": "B-603"}).
 * Moves receipts that Tally books without a flat number (unallocated rows /
 * Capital Ac ledgers) onto the named flat. Nothing is ever moved without an
 * explicit entry. Mutates and returns `tally`; `tally.manualAllocations`
 * lists what happened (including vouchers that were not found).
 */
/** "B-603", or "Maitri Heights|403" when the flat number exists in more than one building. */
function resolveTarget(tally, target) {
  const raw = String(target).trim();
  if (raw.includes('|')) {
    const [b, f] = raw.split('|');
    const flatNo = normalizeFlatKey(f) || f.trim();
    return { key: groupKey(b.trim(), flatNo), flatNo, building: b.trim() };
  }
  const flatNo = normalizeFlatKey(raw) || raw;
  if (tally.buildingsInUse) {
    const hit = [...tally.byFlat.values()].filter((g) => g.flatNo === flatNo);
    if (hit.length === 1) return { key: hit[0].key, flatNo, building: hit[0].building };
    if (hit.length > 1) return { key: flatNo, flatNo, building: '', ambiguous: true };
  }
  return { key: flatNo, flatNo, building: '' };
}

function applyAllocations(tally, allocate) {
  tally.manualAllocations = [];
  for (const [vch, target] of Object.entries(allocate || {})) {
    if (vch.startsWith('_')) continue;
    const dest = resolveTarget(tally, target);
    const { flatNo } = dest;
    const hits = tally.unallocated.filter((r) => String(r.vchNo) === String(vch));
    if (!hits.length || dest.ambiguous) {
      tally.manualAllocations.push({
        vchNo: String(vch), flatNo, found: false, reason: dest.ambiguous ? 'flat is in several buildings - write "Building|flat"' : 'voucher not found',
      });
      continue;
    }
    if (!tally.byFlat.has(dest.key)) {
      tally.byFlat.set(dest.key, {
        key: dest.key, flatNo, building: dest.building, names: new Set(), nameList: [], manualNames: [], payments: [],
      });
    }
    const group = tally.byFlat.get(dest.key);
    for (const row of hits) {
      const party = row.name.replace(/capital\s*a\/?c/i, '').trim();
      group.manualNames = group.manualNames || [];
      if (party && !group.manualNames.includes(party)) group.manualNames.push(party);
      group.payments.push({
        date: row.date, amount: row.amount, vchNo: row.vchNo, gstRate: row.gstRate,
        dupVoucher: false, manualAllocation: true, allocatedFrom: row.name,
      });
      row.allocatedTo = flatNo;
      for (const l of tally.ledgers.values()) {
        l.payments = l.payments.filter((p) => !(p.vchNo === row.vchNo && p.amount === row.amount
          && (p.date ? p.date.getTime() : 0) === (row.date ? row.date.getTime() : 0)));
      }
      tally.manualAllocations.push({
        vchNo: String(vch), flatNo, found: true, amount: row.amount, party: row.name,
      });
    }
    tally.unallocated = tally.unallocated.filter((r) => !hits.includes(r));
    sortByDate(group.payments);
  }
  return tally;
}

/**
 * overrides.json -> "renumber": {"<old flat>": "<new flat>"}. Receipts booked in
 * Tally under an old flat number move to the new one (flagged in Audit); nothing
 * is dropped. Mutates and returns `tally`; `tally.renumbered` lists what moved.
 */
function applyRenumber(tally, renumber) {
  const map = renumberMap(renumber);
  tally.renumbered = [];
  if (!map.size) return tally;
  const moved = new Map();
  for (const r of tally.rows) {
    if (r.flatNo && map.has(r.flatNo)) {
      r.renumberedFrom = r.flatNo;
      r.flatNo = map.get(r.flatNo);
      moved.set(r.renumberedFrom, (moved.get(r.renumberedFrom) || 0) + 1);
    }
  }
  Object.assign(tally, groupRows(tally.rows));
  tally.renumbered = [...moved].map(([from, receipts]) => ({ from, to: map.get(from), receipts }));
  return tally;
}

const sameDay = (a, b) => a instanceof Date && b instanceof Date
  && (a.getTime() === b.getTime()
    || (a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate())
    || (a.getUTCFullYear() === b.getUTCFullYear() && a.getUTCMonth() === b.getUTCMonth() && a.getUTCDate() === b.getUTCDate()));

const HOME_NAME_MIN = 0.85;

/**
 * Building for a receipt whose Tally book has none.
 *  1. the flat number exists in only ONE building (agreements or Tally books): that one;
 *  2. otherwise flat + purchaser name (>= 0.85) + AGMT DT equal to the agreement date.
 * Returns { building, by } or null (the receipt then goes to Unallocated).
 */
function homeBuilding(row, agreements, rows, perFile) {
  const same = agreements.filter((a) => normalizeFlatKey(a.flatNo) === row.flatNo);
  const buildings = new Set(same.map((a) => a.building));
  for (const o of rows) {
    const b = perFile.get(o.file);
    if (b && o.flatNo === row.flatNo) buildings.add(b);
  }
  if (buildings.size === 1) return { building: [...buildings][0], by: 'flat number exists in only one building' };
  // no known building has this flat: it is one flat of an unnamed book, not a clash
  if (!buildings.size && !same.length) return { building: '', by: 'flat number found in no known building (building unknown)' };
  for (const a of same) {
    const names = [...(a.names || []), a.name].filter(Boolean);
    const score = Math.max(0, ...names.map((n) => fuzzyNameSimilarity(row.name, n)));
    if (score >= HOME_NAME_MIN && row.agmtDate && sameDay(row.agmtDate, a.agreementDate)) {
      return { building: a.building, by: `allocated by flat+name+date (name ${Math.round(score * 100) / 100}, AGMT DT ${formatMDY(row.agmtDate)})` };
    }
  }
  return null;
}

/**
 * Give every receipt a building, then regroup so Maitri Heights 403 and
 * Siddhivinayak Palace 403 stay apart. A book's building comes from overrides.json
 * "books" ({"idfc 5112.xlsx": "Maitri Heights"}) or from its title when that names
 * a known building. Buildings are only used when some book has one, or the
 * agreements name two or more; otherwise nothing changes. In that mode a receipt
 * whose book has no building is NEVER merged into a flat: it goes to Unallocated
 * with its flat number kept as the (review) guess.
 *
 * @param {object} opts { names: [buildings read from agreements], agreementBuildings, buildings, books }
 */
function assignBuildings(tally, opts = {}) {
  const buildings = opts.buildings || {};
  const books = opts.books || {};
  const names = [...new Set([...(opts.names || []), ...Object.keys(buildings).filter((k) => !k.startsWith('_'))])];
  const titleOf = new Map((tally.fileStats || []).map((f) => [f.file, f.title || '']));

  const bookBuilding = (file) => {
    const set = Object.keys(books).find((k) => !k.startsWith('_') && k.toLowerCase() === String(file).toLowerCase());
    if (set) return canonicalBuilding(books[set], buildings);
    return buildingInText(titleOf.get(file) || '', names, buildings);
  };
  const perFile = new Map();
  for (const file of new Set(tally.rows.map((r) => r.file))) perFile.set(file, bookBuilding(file));

  const assigned = [...perFile.values()].some(Boolean);
  const distinct = new Set((opts.agreementBuildings || []).filter(Boolean)).size;
  tally.buildingsInUse = assigned || distinct >= 2;
  tally.bookBuildings = [...perFile].map(([file, building]) => ({ file, building }));
  if (!tally.buildingsInUse) return tally;

  const agreements = (opts.agreements || []).filter((a) => a.flatNo && a.building);
  for (const r of tally.rows) {
    r.building = perFile.get(r.file) || '';
    if (r.flatNo && !r.building) {
      const home = homeBuilding(r, agreements, tally.rows, perFile);
      if (home) {
        r.building = home.building;
        r.allocatedBy = home.by;
        continue;
      }
      r.guessFlat = r.flatNo;
      r.flatNo = null;
      r.buildingReview = true;
    }
  }
  Object.assign(tally, groupRows(tally.rows));
  return tally;
}

/** Sum of every receipt row read from Tally (Opening/Closing Balance rows are never rows). */
function totalReceipts(tally) {
  return Math.round(tally.rows.reduce((s, r) => s + r.amount, 0) * 100) / 100;
}

/**
 * Merge multiple parsed Tally results, rebuilding the flat mapping and
 * deduplicating voucher entries across all sheets.
 */
function mergeTallies(tallies) {
  if (!tallies || !tallies.length) {
    return {
      rows: [], byFlat: new Map(), ledgers: new Map(), unallocated: [], duplicateVouchers: [], fileStats: [], projectName: '',
    };
  }

  // The same file uploaded twice: the second copy is ignored (and shown so in the audit).
  const seen = new Set();
  const used = [];
  const fileStats = [];
  for (const t of tallies) {
    const stat = (t.fileStats || [])[0];
    if (stat && stat.hash && seen.has(stat.hash)) {
      fileStats.push({ ...stat, receipts: 0, total: 0, duplicateUpload: true });
      continue;
    }
    if (stat && stat.hash) seen.add(stat.hash);
    used.push(t);
    fileStats.push(...(t.fileStats || []));
  }
  tallies = used;
  const rows = tallies.flatMap((t) => t.rows || []);
  const projectNames = tallies.map((t) => t.projectName).filter(Boolean);
  const uniqueNames = Array.from(new Set(projectNames.map((name) => name.trim()).filter(Boolean)));

  return {
    rows, ...groupRows(rows), fileStats, projectName: uniqueNames.join(' & ') || '',
  };
}
module.exports = { parseTally, mergeTallies, applyAllocations, applyRenumber, assignBuildings, totalReceipts, gstRateFromCell, findGstColumn };
