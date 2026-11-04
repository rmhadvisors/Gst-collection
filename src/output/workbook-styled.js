'use strict';

/**
 * Collection Sheet Styled Writer (ExcelJS)
 *
 * Produces a professionally styled Collection Sheet workbook that visually
 * matches the CA template:
 *   - Dark navy headers with white bold text
 *   - Yellow TOTAL :: rows for each flat header
 *   - Thin borders on all data cells
 *   - Currency number format #,##0.00
 *   - Calibri 11pt font throughout
 *   - Correct column widths and row heights
 *   - Three sheets: == (summary), Wing B (detail), Audit
 */

const ExcelJS = require('exceljs');
const { wingOf } = require('../util/flatKey');

// ─── Style palette ──────────────────────────────────────────────────────────

const COLOR = {
  HEADER_DARK:   '1F3864',  // dark navy blue  — project title
  HEADER_MID:    '2E75B6',  // medium blue     — section / column headers
  FLAT_ROW:      'FFD966',  // golden yellow   — flat TOTAL :: rows
  TALLY_ONLY:    'DEEBF7',  // pale blue       — tally-only flat rows
  PAYMENT_ODD:   'FFFFFF',  // white           — payment rows (odd)
  PAYMENT_EVEN:  'F5F5F5',  // very light grey — payment rows (even)
  TOTALS_ROW:    'E2EFDA',  // light green     — grand-total rows
  FONT_WHITE:    'FFFFFF',
  FONT_DARK:     '000000',
  BORDER_COLOR:  '9E9E9E',  // medium grey border
};

const FONT_BOLD_WHITE   = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF' + COLOR.FONT_WHITE } };
const FONT_BOLD_DARK    = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF' + COLOR.FONT_DARK }  };
const FONT_NORMAL       = { name: 'Calibri', size: 10, color: { argb: 'FF' + COLOR.FONT_DARK }              };
const FONT_TITLE        = { name: 'Calibri', size: 13, bold: true, color: { argb: 'FF' + COLOR.FONT_WHITE } };

const NUMFMT  = '#,##0.00';
const NUMFMT0 = '#,##0';

function fill(argbHex) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + argbHex } };
}

function border(style = 'thin') {
  const s = { style, color: { argb: 'FF' + COLOR.BORDER_COLOR } };
  return { top: s, left: s, bottom: s, right: s };
}

function applyCell(cell, opts = {}) {
  if (opts.font)       cell.font       = opts.font;
  if (opts.fill)       cell.fill       = opts.fill;
  if (opts.border)     cell.border     = opts.border;
  if (opts.alignment)  cell.alignment  = opts.alignment;
  if (opts.numFmt)     cell.numFmt     = opts.numFmt;
}

// ─── Wing B builder ─────────────────────────────────────────────────────────

const WINGB_HEADERS = [
  'Sr.No.', 'FLOOR', 'SHOP NO.', 'FLAT TYPE', 'RERA CARPET AREA (Sqmt)', 'REG. NO.',
  'DATE OF AGREEMENT', "PURCHASER'S NAMES", 'AGREEMENT VALUE', 'GST ON AGREEMENT VALUE',
  'DATE OF AMOUNT RECEIVED', 'AMOUNT RECEIVED', 'GST ON ADVANCE', 'TDS', 'BALANCE OUTSTANDING',
];

const WINGB_WIDTHS = [6, 12, 10, 10, 16, 10, 18, 34, 18, 20, 18, 16, 14, 10, 20];

function addWingBSheet(wb, records, projectName, sheetName = 'Wing B') {
  const ws = wb.addWorksheet(sheetName);

  // Column widths
  ws.columns = WINGB_WIDTHS.map(w => ({ width: w }));

  function row(data, rowOpts = {}) {
    const r = ws.addRow(data);
    if (rowOpts.height) r.height = rowOpts.height;
    return r;
  }

  // ── Row 1: Project title ──
  {
    const r = row([`PROJECT :: ${projectName}`], { height: 22 });
    ws.mergeCells(r.number, 1, r.number, WINGB_HEADERS.length);
    const cell = r.getCell(1);
    applyCell(cell, {
      font:      FONT_TITLE,
      fill:      fill(COLOR.HEADER_DARK),
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
  }

  // ── Row 2: Section header ──
  {
    const r = row(['SHOP/ FLAT BOOKING DETAILS'], { height: 18 });
    ws.mergeCells(r.number, 1, r.number, WINGB_HEADERS.length);
    const cell = r.getCell(1);
    applyCell(cell, {
      font:      FONT_BOLD_WHITE,
      fill:      fill(COLOR.HEADER_MID),
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
  }

  // ── Row 3: blank spacer ──
  row([]);

  // ── Row 4: Column headers ──
  {
    const r = row(WINGB_HEADERS, { height: 30 });
    r.eachCell((cell) => {
      applyCell(cell, {
        font:      FONT_BOLD_WHITE,
        fill:      fill(COLOR.HEADER_MID),
        border:    border('thin'),
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
      });
    });
  }

  // ── Data rows ──
  records.forEach((rec, idx) => {
    const isAgreement = rec.hasAgreement;
    const rowFill = isAgreement ? fill(COLOR.FLAT_ROW) : fill(COLOR.TALLY_ONLY);

    // Flat header row (TOTAL ::)
    const headerData = [
      idx + 1,
      rec.floorLabel || '',
      rec.flatNo,
      rec.flatType   || '',
      rec.carpetArea != null && rec.carpetArea !== '' ? rec.carpetArea : '',
      rec.regNo      || '',
      rec.agreementDate || '',
      rec.name       || '',
      rec.agreementValue   != null ? rec.agreementValue   : '',
      rec.gstOnAgreement   != null ? rec.gstOnAgreement   : '',
      'TOTAL ::',
      rec.totalReceived    != null ? rec.totalReceived    : '',
      rec.totalGst         != null ? rec.totalGst         : '',
      '',
      rec.balance          != null ? rec.balance          : '',
    ];

    const hdrRow = ws.addRow(headerData);
    hdrRow.height = 16;
    hdrRow.eachCell((cell, colIdx) => {
      applyCell(cell, {
        font:   FONT_BOLD_DARK,
        fill:   rowFill,
        border: border('thin'),
        alignment: { horizontal: colIdx <= 2 ? 'center' : colIdx >= 9 ? 'right' : 'left', vertical: 'middle' },
      });
      // Number format for money columns
      if ([9, 10, 12, 13].includes(colIdx)) {
        cell.numFmt = NUMFMT;
      }
    });

    // Payment rows
    rec.payments.forEach((p, pi) => {
      const payFill = pi % 2 === 0 ? fill(COLOR.PAYMENT_ODD) : fill(COLOR.PAYMENT_EVEN);
      const payData = [
        '', '', '', '', '', '', '', '', '', '',
        p.dateText || '',
        p.amount   != null ? p.amount : '',
        p.gst      != null ? p.gst    : '',
        '',
        '',
      ];
      const payRow = ws.addRow(payData);
      payRow.height = 14;
      payRow.eachCell((cell, colIdx) => {
        applyCell(cell, {
          font:   FONT_NORMAL,
          fill:   payFill,
          border: border('thin'),
          alignment: { horizontal: colIdx >= 12 ? 'right' : 'left', vertical: 'middle' },
        });
        if ([12, 13].includes(colIdx)) cell.numFmt = NUMFMT;
      });
    });

    // Blank spacer row between flats
    ws.addRow([]).height = 4;
  });
}

// ─── Summary (==) builder ───────────────────────────────────────────────────

const SUMMARY_HEADERS = [
  'SR.NO', 'FLOOR', 'SHOP NO.', 'RERA CARPET AREA (Sqmt)', 'PARTY NAME',
  'DATE OF AGREEMENT', 'AGREEMENT', 'GST', 'AGREEMENT VALUE (WITH GST)',
  'ADVANCES', 'TDS', 'O/S AMOUNT', 'GST PAID', 'Adv Rec for Form-3', 'Outstanding',
];

const SUMMARY_WIDTHS = [6, 12, 10, 16, 32, 18, 16, 14, 22, 16, 10, 14, 14, 18, 16];

function addSummarySheet(wb, records, projectName, warnings = []) {
  const ws = wb.addWorksheet('Summary');
  ws.columns = SUMMARY_WIDTHS.map(w => ({ width: w }));
  addWarningRow(ws, warnings, SUMMARY_HEADERS.length);

  // Row 1: title
  {
    const r = ws.addRow([`PROJECT :: ${projectName}`]);
    r.height = 22;
    ws.mergeCells(r.number, 1, r.number, SUMMARY_HEADERS.length);
    applyCell(r.getCell(1), {
      font:      FONT_TITLE,
      fill:      fill(COLOR.HEADER_DARK),
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
  }

  // Row 2: section header
  {
    const r = ws.addRow(['SHOP/ FLAT BOOKING DETAILS']);
    r.height = 18;
    ws.mergeCells(r.number, 1, r.number, SUMMARY_HEADERS.length);
    applyCell(r.getCell(1), {
      font:      FONT_BOLD_WHITE,
      fill:      fill(COLOR.HEADER_MID),
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
  }

  // Row 3: column headers
  {
    const r = ws.addRow(SUMMARY_HEADERS);
    r.height = 28;
    r.eachCell(cell => applyCell(cell, {
      font:      FONT_BOLD_WHITE,
      fill:      fill(COLOR.HEADER_MID),
      border:    border('thin'),
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    }));
  }

  // Data rows
  records.forEach((rec, idx) => {
    const rowFill = rec.hasAgreement ? fill(COLOR.PAYMENT_ODD) : fill(COLOR.TALLY_ONLY);
    const r = ws.addRow([
      idx + 1,
      rec.floorLabel    || '',
      rec.flatNo,
      rec.carpetArea    != null && rec.carpetArea !== '' ? rec.carpetArea : '',
      rec.name          || '',
      rec.agreementDate || '',
      rec.agreementValue         != null ? rec.agreementValue         : '',
      rec.gstOnAgreement         != null ? rec.gstOnAgreement         : '',
      rec.agreementValueWithGst  != null ? rec.agreementValueWithGst  : '',
      rec.totalReceived          != null ? rec.totalReceived          : '',
      '',
      '',
      rec.totalGst               != null ? rec.totalGst               : '',
      '',
      rec.balance                != null ? rec.balance                : '',
    ]);
    r.height = 15;
    r.eachCell((cell, colIdx) => {
      applyCell(cell, {
        font:   FONT_NORMAL,
        fill:   rowFill,
        border: border('thin'),
        alignment: { horizontal: colIdx >= 7 ? 'right' : 'left', vertical: 'middle' },
      });
      if ([7, 8, 9, 10, 13, 15].includes(colIdx)) cell.numFmt = NUMFMT;
    });
  });
}

// ─── Reconciliation + warnings ──────────────────────────────────────────────

const round2 = (n) => Math.round(n * 100) / 100;

/** Tally receipts vs (flat receipts + Unallocated); null when no Tally total was supplied. */
function reconcile(records, opts) {
  if (!opts.reconciliation) return null;
  const sum = (arr) => round2(arr.reduce((s, n) => s + n, 0));
  const tallyTotal = opts.reconciliation.tallyTotal;
  const flatTotal = sum(records.flatMap((r) => r.payments.filter((p) => p.source === 'tally').map((p) => p.amount)));
  const unallocTotal = sum((opts.unallocated || []).map((u) => u.amount));
  const diff = round2(tallyTotal - (flatTotal + unallocTotal));
  // receipts typed in through overrides.json (extraPayments): not in the Tally books
  const manual = records.flatMap((r) => r.payments
    .filter((p) => p.source !== 'tally')
    .map((p) => ({ flatNo: r.flatNo, dateText: p.dateText || '', amount: p.amount })));
  const manualTotal = sum(manual.map((m) => m.amount));
  return {
    tallyTotal, flatTotal, unallocTotal, manual, manualTotal, diff, ok: Math.abs(diff) < 0.005,
  };
}

const flatList = (recs) => recs.map((r) => r.flatNo).join(', ');

/**
 * Reasons the CA must look at this workbook before relying on it: flats that
 * need a manual check, OCR failures, "(review)" matches, Tally files that gave
 * nothing and a reconciliation that is not MATCHED.
 */
function collectWarnings(records, opts) {
  const out = [];
  const ocr = records.filter((r) => r.ocrFailed || (r.flags || []).some((f) => /OCR FAILED/.test(f)));
  if (ocr.length) out.push(`OCR FAILED on flat(s) ${flatList(ocr)}`);
  const review = records.filter((r) => r.needsReview || /\(review\)/.test(r.matchReason || '')
    || (r.flags || []).some((f) => /\(review\)|— review/.test(f)));
  if (review.length) out.push(`Name-only match (review) on flat(s) ${flatList(review)}`);
  const manual = records.filter((r) => r.hasAgreement && r.agreementValue == null);
  if (manual.length) out.push(`Agreement value not read - check manually: flat(s) ${flatList(manual)}`);
  const unread = ((opts.reconciliation && opts.reconciliation.files) || []).filter((f) => !f.receipts && !f.duplicateUpload);
  if (unread.length) out.push(`Tally file NOT READ: ${unread.map((f) => f.file).join(', ')}`);
  const rec = reconcile(records, opts);
  if (rec && !rec.ok) {
    out.push(`Receipt reconciliation NOT MATCHED: difference Rs ${rec.diff.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
  }
  return out;
}

/** Red banner as the first row of a sheet; nothing is added when there are no warnings. */
function addWarningRow(ws, warnings, columns) {
  if (!warnings.length) return;
  const r = ws.addRow([`CHECK BEFORE USE: ${warnings.join('  |  ')}`]);
  r.height = Math.max(22, 16 * Math.ceil((warnings.join('  |  ').length + 18) / (columns * 11)));
  ws.mergeCells(r.number, 1, r.number, columns);
  applyCell(r.getCell(1), {
    font:      FONT_BOLD_WHITE,
    fill:      fill('C00000'),
    alignment: { horizontal: 'left', vertical: 'middle', wrapText: true },
  });
}

// ─── Audit sheet ────────────────────────────────────────────────────────────

const AUDIT_HEADERS = [
  'Flat No', 'Purchaser (English)', 'Purchaser (Marathi original)', 'Agreement Value', 'Value Source',
  'GST Rate', 'GST on Agreement', 'Total Received', 'Total GST on Advance',
  'Balance', 'Tally Match', 'Match Score', 'OCR Value Candidates', 'Source PDF', 'Agreement Date Source', 'Flags',
];

function addAuditSheet(wb, records, opts = {}) {
  const ws = wb.addWorksheet('Audit');
  ws.columns = [10, 30, 34, 16, 12, 10, 16, 16, 18, 16, 14, 12, 24, 28, 20, 60].map(w => ({ width: w }));
  addWarningRow(ws, opts.warnings || [], AUDIT_HEADERS.length);

  const hdr = ws.addRow(AUDIT_HEADERS);
  hdr.height = 20;
  hdr.eachCell(cell => applyCell(cell, {
    font:      FONT_BOLD_WHITE,
    fill:      fill(COLOR.HEADER_DARK),
    border:    border('thin'),
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
  }));

  records.forEach(rec => {
    const r = ws.addRow([
      rec.flatNo,
      rec.name,
      rec.purchaserMarathi || '',
      rec.agreementValue != null ? rec.agreementValue : '',
      rec.valueSource    || '',
      rec.gstRate        != null ? `${rec.gstRate * 100}%` : '',
      rec.gstOnAgreement != null ? rec.gstOnAgreement : '',
      rec.totalReceived  != null ? rec.totalReceived : '',
      rec.totalGst       != null ? rec.totalGst : '',
      rec.balance        != null ? rec.balance : '',
      rec.matchReason    || '',
      rec.matchScore     != null ? Math.round(rec.matchScore * 100) / 100 : '',
      (rec.valueCandidates || []).join(', '),
      rec.pdf            || '',
      rec.pdf ? (rec.agreementDateSource || '') : '',
      (rec.flags || []).join('; '),
    ]);
    r.eachCell((cell, colIdx) => {
      applyCell(cell, {
        font:   FONT_NORMAL,
        border: border('thin'),
        alignment: { vertical: 'middle' },
      });
      if ([4, 7, 8, 9, 10].includes(colIdx)) cell.numFmt = NUMFMT;
    });
  });

  if (opts.reconciliation) addReconciliation(ws, records, opts);
}

/**
 * Receipt reconciliation at the bottom of the Audit sheet:
 * Tally receipts (excl. Opening/Closing Balance) = flat receipts + Unallocated receipts.
 */
function addReconciliation(ws, records, opts) {
  const {
    tallyTotal, flatTotal, unallocTotal, manual, manualTotal, diff, ok,
  } = reconcile(records, opts);

  ws.addRow([]);
  const title = ws.addRow(['RECEIPT RECONCILIATION']);
  ws.mergeCells(title.number, 1, title.number, 3);
  applyCell(title.getCell(1), { font: FONT_BOLD_WHITE, fill: fill(COLOR.HEADER_DARK), border: border('thin') });

  const files = opts.reconciliation.files || [];
  if (files.length) {
    const head = ws.addRow(['Tally file', 'Receipts', 'Total', 'Status']);
    head.eachCell((cell) => applyCell(cell, { font: FONT_BOLD_DARK, border: border('thin') }));
    for (const f of files) {
      const unread = !f.receipts && !f.duplicateUpload;
      const fr = ws.addRow([f.file, f.receipts, f.total, f.duplicateUpload ? 'duplicate upload – ignored' : (unread ? 'NOT READ' : 'Read')]);
      fr.eachCell((cell, i) => applyCell(cell, {
        font: unread ? { ...FONT_BOLD_DARK, color: { argb: 'FFC00000' } } : FONT_NORMAL,
        border: border('thin'),
      }));
      fr.getCell(3).numFmt = NUMFMT;
    }
    ws.addRow([]);
  }

  const lines = [
    ['Total Tally receipts (excluding Opening/Closing Balance)', tallyTotal],
    ['Sum of all flat receipts', flatTotal],
    ['Sum of Unallocated receipts', unallocTotal],
    ['Flat + Unallocated', Math.round((flatTotal + unallocTotal) * 100) / 100],
  ];
  // Manual receipts are shown on their own line so "Tally total" stays the real Tally books.
  if (manual.length) {
    lines.push(['Manual receipts (not in Tally) - added through overrides.json, shown in the flat sheets', manualTotal]);
    for (const m of manual) lines.push([`      Flat ${m.flatNo}  ${m.dateText}`, m.amount]);
  }
  for (const [label, value] of lines) {
    const r = ws.addRow([label, '', '', value]);
    ws.mergeCells(r.number, 1, r.number, 3);
    applyCell(r.getCell(1), { font: FONT_NORMAL, border: border('thin') });
    applyCell(r.getCell(4), { font: FONT_NORMAL, border: border('thin'), numFmt: NUMFMT });
  }
  const result = ws.addRow(['Check', '', '', ok ? 'MATCHED' : `DIFFERENCE: ₹${diff.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`]);
  ws.mergeCells(result.number, 1, result.number, 3);
  applyCell(result.getCell(1), { font: FONT_BOLD_DARK, border: border('thin') });
  applyCell(result.getCell(4), {
    font: { ...FONT_BOLD_DARK, color: { argb: ok ? 'FF2E7D32' : 'FFC00000' } },
    border: border('thin'),
  });
  if (!ok) {
    const note = ws.addRow(['Receipts dropped by overrides (paymentCutoff / excludeVchNos) or exact duplicate rows appear here as a difference.']);
    applyCell(note.getCell(1), { font: { ...FONT_NORMAL, italic: true } });
  }
}

const UNALLOC_HEADERS = ['Date', 'Party', 'Vch No.', 'Amount', 'Best flat guess', 'Guess score', 'Flag'];

/** Receipts with no flat number — listed for the CA to allocate by hand. */
function addUnallocatedSheet(wb, rows) {
  const ws = wb.addWorksheet('Unallocated Receipts');
  ws.columns = [14, 40, 10, 16, 16, 12, 40].map(w => ({ width: w }));
  const hdr = ws.addRow(UNALLOC_HEADERS);
  hdr.height = 20;
  hdr.eachCell(cell => applyCell(cell, {
    font:      FONT_BOLD_WHITE,
    fill:      fill(COLOR.HEADER_DARK),
    border:    border('thin'),
    alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
  }));
  rows.forEach((u) => {
    const r = ws.addRow([u.dateText, u.party, u.vchNo, u.amount, u.guess, u.guessScore, u.flag || '']);
    r.eachCell((cell, colIdx) => {
      applyCell(cell, { font: FONT_NORMAL, border: border('thin'), alignment: { vertical: 'middle' } });
      if (colIdx === 4) cell.numFmt = NUMFMT;
    });
  });
}
// ─── Public writer ──────────────────────────────────────────────────────────

/**
 * Write a styled Collection Sheet workbook (ExcelJS).
 * Drop-in replacement for writeCollectionWorkbook() in workbook.js.
 *
 * @param {object[]} records   Array of flat record objects
 * @param {string}   outPath   Destination .xlsx file path
 * @param {object}   [opts]    { projectName }
 */
async function writeStyledCollectionWorkbook(records, outPath, opts = {}) {
  const projectName = opts.projectName || 'YASHWANTH COUNTY';
  const wb = new ExcelJS.Workbook();

  wb.creator   = 'CA Office Automation';
  wb.company   = projectName;
  wb.created   = new Date();
  wb.modified  = new Date();

  const warnings = collectWarnings(records, opts);
  addSummarySheet(wb, records, projectName, warnings);
  // Wing-less projects get one "Flats" sheet; multi-wing projects get one sheet per wing and wings never merge.
  const wings = [...new Set(records.map((r) => wingOf(r.flatNo)))].filter(Boolean).sort();
  if (wings.length > 1) {
    for (const w of wings) {
      addWingBSheet(wb, records.filter((r) => wingOf(r.flatNo) === w), projectName, `Wing ${w}`);
    }
    const noWing = records.filter((r) => !wingOf(r.flatNo));
    if (noWing.length) addWingBSheet(wb, noWing, projectName, 'Flats');
  } else {
    // No default wing: a sheet is named "Wing X" only when a wing really exists.
    addWingBSheet(wb, records, projectName, wings.length === 1 ? `Wing ${wings[0]}` : 'Flats');
  }
  addAuditSheet(wb, records, { ...opts, warnings });
  if (opts.unallocated && opts.unallocated.length) addUnallocatedSheet(wb, opts.unallocated);

  await wb.xlsx.writeFile(outPath);
  return outPath;
}

module.exports = { writeStyledCollectionWorkbook, collectWarnings, reconcile };
