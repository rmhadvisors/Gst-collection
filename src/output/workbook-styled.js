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

function addWingBSheet(wb, records, projectName) {
  const ws = wb.addWorksheet('Wing B');

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

function addSummarySheet(wb, records, projectName) {
  const ws = wb.addWorksheet('==');
  ws.columns = SUMMARY_WIDTHS.map(w => ({ width: w }));

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

// ─── Audit sheet ────────────────────────────────────────────────────────────

const AUDIT_HEADERS = [
  'Flat No', 'Purchaser (from agreement)', 'Agreement Value', 'Value Source',
  'GST Rate', 'GST on Agreement', 'Total Received', 'Total GST on Advance',
  'Balance', 'Tally Match', 'Match Score', 'OCR Value Candidates', 'Source PDF',
];

function addAuditSheet(wb, records) {
  const ws = wb.addWorksheet('Audit');
  ws.columns = [10, 30, 16, 12, 10, 16, 16, 18, 16, 14, 12, 24, 28].map(w => ({ width: w }));

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
    ]);
    r.eachCell((cell, colIdx) => {
      applyCell(cell, {
        font:   FONT_NORMAL,
        border: border('thin'),
        alignment: { vertical: 'middle' },
      });
      if ([3, 6, 7, 8, 9].includes(colIdx)) cell.numFmt = NUMFMT;
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

  addSummarySheet(wb, records, projectName);
  addWingBSheet(wb, records, projectName);
  addAuditSheet(wb, records);

  await wb.xlsx.writeFile(outPath);
  return outPath;
}

module.exports = { writeStyledCollectionWorkbook };
