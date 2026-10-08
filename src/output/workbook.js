'use strict';

const XLSX = require('xlsx');

const NUMFMT = '#,##0.00';
const DEFAULT_PROJECT = 'YASHWANTH COUNTY';

const WINGB_HEADERS = [
  'Sr.No.', 'FLOOR', 'SHOP NO.', 'FLAT TYPE', 'RERA CARPET AREA (Sqmt)', 'REG. NO.',
  'DATE OF AGREEMENT', "PURCHASER'S NAMES", 'AGREEMENT VALUE', 'GST ON AGREEMENT VALUE',
  'DATE OF AMOUNT RECEIVED', 'AMOUNT RECEIVED', 'GST ON ADVANCE', 'TDS', 'BALANCE OUTSTANDING',
];

const SUMMARY_HEADERS = [
  'SR.NO', 'FLOOR', 'SHOP NO.', 'RERA CARPET AREA (Sqmt)', 'PARTY NAME', 'DATE OF AGREEMENT',
  'AGREEMENT', 'GST', 'AGREEMENT VALUE (WITH GST)', 'ADVANCES', 'TDS', 'O\\S AMOUNT',
  'GST PAID', 'Adv Rec for Form-3', 'Outstanding',
];

function n(v) {
  return (v == null || v === '') ? '' : v;
}

/** Build the "Wing B" detailed collection sheet (block per flat). */
function buildWingB(records, projectName = DEFAULT_PROJECT) {
  const aoa = [];
  aoa.push([`PROJECT :: ${projectName}`]);
  aoa.push(['SHOP/ FLAT BOOKING DETAILS']);
  aoa.push([]);
  aoa.push(WINGB_HEADERS.slice());

  records.forEach((rec, idx) => {
    aoa.push([
      idx + 1,
      rec.floorLabel,
      rec.flatNo,
      rec.flatType,
      n(rec.carpetArea),
      rec.regNo,
      rec.agreementDate,
      rec.name,
      n(rec.agreementValue),
      n(rec.gstOnAgreement),
      'TOTAL ::',
      n(rec.totalReceived),
      n(rec.totalGst),
      '',
      n(rec.balance),
    ]);
    rec.payments.forEach((p) => {
      aoa.push([
        '', '', '', '', '', '', '', '', '', '',
        p.dateText, n(p.amount), n(p.gst), '', '',
      ]);
    });
    aoa.push([]); // spacer between flats
  });

  return aoaToFormattedSheet(aoa, [6, 12, 10, 10, 14, 10, 16, 32, 16, 18, 16, 16, 14, 10, 18]);
}

/** Build the "==" one-row-per-flat summary sheet. */
function buildSummary(records, projectName = DEFAULT_PROJECT) {
  const aoa = [];
  aoa.push([`PROJECT :: ${projectName}`]);
  aoa.push(['SHOP/ FLAT BOOKING DETAILS']);
  aoa.push(SUMMARY_HEADERS.slice());

  records.forEach((rec, idx) => {
    aoa.push([
      idx + 1,
      rec.floorLabel,
      rec.flatNo,
      n(rec.carpetArea),
      rec.name,
      rec.agreementDate,
      n(rec.agreementValue),
      n(rec.gstOnAgreement),
      n(rec.agreementValueWithGst),
      n(rec.totalReceived),
      '', // TDS (not in scope)
      '', // O\S AMOUNT (not in scope)
      n(rec.totalGst),
      '', // Adv Rec for Form-3 (not in scope)
      n(rec.balance),
    ]);
  });

  return aoaToFormattedSheet(aoa, [6, 12, 10, 14, 30, 16, 16, 14, 20, 16, 10, 14, 14, 16, 16]);
}

/** Build a transparency sheet so the user can audit OCR + matching. */
function buildAudit(records) {
  const aoa = [[
    'Flat No', 'Purchaser (from agreement)', 'Agreement Value', 'Value Source',
    'GST Rate', 'GST on Agreement', 'Total Received', 'Total GST on Advance',
    'Balance', 'Tally Match', 'Match Score', 'OCR Value Candidates', 'Source PDF',
  ]];
  records.forEach((rec) => {
    aoa.push([
      rec.flatNo,
      rec.name,
      n(rec.agreementValue),
      rec.valueSource,
      rec.gstRate != null ? `${rec.gstRate * 100}%` : '',
      n(rec.gstOnAgreement),
      n(rec.totalReceived),
      n(rec.totalGst),
      n(rec.balance),
      rec.matchReason,
      rec.matchScore != null ? Math.round(rec.matchScore * 100) / 100 : '',
      (rec.valueCandidates || []).join(', '),
      rec.pdf || '',
    ]);
  });
  return aoaToFormattedSheet(aoa, [10, 30, 16, 12, 10, 16, 16, 18, 16, 14, 12, 24, 28]);
}

function aoaToFormattedSheet(aoa, widths, intCols = [0]) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const range = XLSX.utils.decode_range(ws['!ref']);
  const intSet = new Set(intCols);
  for (let r = range.s.r; r <= range.e.r; r += 1) {
    for (let c = range.s.c; c <= range.e.c; c += 1) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && cell.t === 'n') cell.z = intSet.has(c) ? '0' : NUMFMT;
    }
  }
  if (widths) ws['!cols'] = widths.map((w) => ({ wch: w }));
  return ws;
}

/** Assemble and write the final workbook. */
function writeCollectionWorkbook(records, outPath, opts = {}) {
  const projectName = opts.projectName || DEFAULT_PROJECT;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, buildSummary(records, projectName), '==');
  XLSX.utils.book_append_sheet(wb, buildWingB(records, projectName), 'Wing B');
  XLSX.utils.book_append_sheet(wb, buildAudit(records), 'Audit');
  XLSX.writeFile(wb, outPath);
  return outPath;
}

module.exports = { writeCollectionWorkbook, buildWingB, buildSummary };
