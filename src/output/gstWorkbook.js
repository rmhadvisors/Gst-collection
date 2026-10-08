'use strict';

/**
 * GST Calculation Workbook writer.
 *
 * Produces two sheet types per active month, exactly matching the CA's format:
 *
 *  1. "MMM-YY WORKING"  (e.g. "APR-25 WORKING")
 *     Row 1:  title — "[MONTH YEAR] GST CALCULATION"
 *     Row 2:  note about 45-lakh rule
 *     Row 3:  header — DATE | FLAT NO. | AGREEMENT DONE | AGREEMENT AMOUNT | 1% GST | 5% GST | ADVANCES AMOUNT | REVERSE GST ON ADVANCE | AMOUNT PAID
 *     Row 4+: one row per agreement flat (Case 1 / Case 2)
 *     Totals row
 *     Blank rows
 *     "CALCULATION ON ADVANCES" section header
 *     Header — FLAT NO. | NAME | ADVANCE AMOUNT | 1% GST | 5% GST
 *     Advance rows (Case 3)
 *     Totals row  |  TOTAL PAYABLE ==>
 *
 *  2. "MMM-YY" (e.g. "APR-25")
 *     Full OUTWARD SUPPLY / INWARD SUPPLY / IGST / NET PAYABLE summary grid.
 */

const XLSX = require('xlsx');

const NUMFMT = '#,##0.00';
const NUMFMT_INT = '#,##0';
const DEFAULT_PROJECT = 'SPACE HOME';

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

// ─── Cell helpers ─────────────────────────────────────────────────────────────

/** Return a numeric cell object (so xlsx applies number format). */
function num(v) {
  if (v == null || v === '' || isNaN(Number(v))) return { t: 's', v: '' };
  return { t: 'n', v: Number(v) };
}

/** Return a string cell object. */
function str(v) {
  return { t: 's', v: String(v == null ? '' : v) };
}

/** Convert array-of-arrays to a sheet and apply number formatting. */
function aoaToSheet(aoa, colWidths) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (cell && cell.t === 'n') {
        cell.z = NUMFMT;
      }
    }
  }
  if (colWidths) ws['!cols'] = colWidths.map((w) => ({ wch: w }));
  return ws;
}

// ─── WORKING sheet ────────────────────────────────────────────────────────────

/**
 * Build the "MMM-YY WORKING" sheet for one month.
 * Columns (0-based):
 *   0: SR.NO.  1: DATE (or FLAT NO.)  2: AGREEMENT DONE  3: AGREEMENT AMOUNT
 *   4: 1% GST  5: 5% GST  6: ADVANCES AMOUNT  7: REVERSE GST ON ADVANCE  8: AMOUNT PAID
 */
function buildWorkingSheet(monthData, projectName) {
  const aoa = [];

  // Row 1 — title
  aoa.push(['', '', '', `${monthData.displayName} GST CALCULATION`, '', '', '', '', '']);

  // Row 2 — note
  aoa.push(['', '', '(NOTE: IF AGREEMENT VALUE CROSS MORE THAN 45,00,000 THEN GST 5% APPLY)', '', '', '', '', '', '']);

  // Row 3 — header
  aoa.push([
    'SR.NO.', 'FLAT NO.', 'AGREEMENT DONE', 'AGREEMENT AMOUNT',
    '1% GST', '5% GST', 'ADVANCES AMOUNT', 'REVERSE GST ON ADVANCE', 'AMOUNT PAID',
  ]);

  // ── Agreement section (Case 1 & 2) ──────────────────────────────────────
  const agreements = monthData.agreements || [];
  let srNo = 1;

  for (const a of agreements) {
    const gst1 = a.rate === 0.01 ? a.outstandingGst : '';
    const gst5 = a.rate === 0.05 ? a.outstandingGst : '';
    const advAmt = a.advancesBeforeAgreement > 0 ? a.advancesBeforeAgreement : '';
    const revGst = a.reverseGstOnAdvances > 0 ? a.reverseGstOnAdvances : '';

    aoa.push([
      srNo++,
      a.flatNo,
      a.name,
      a.agreementValue,
      gst1,
      gst5,
      advAmt,
      revGst,
      a.outstandingGst,
    ]);
  }

  // Filler rows to keep the same look as the CA template (up to 3 rows for agreements)
  const agrFillerCount = Math.max(0, 3 - agreements.length);
  for (let i = 0; i < agrFillerCount; i++) {
    aoa.push([srNo++, '', '', '', 0, 0, '', 0, 0]);
  }

  // Agreement totals row
  const agrTotals = agreements.reduce(
    (t, a) => {
      t.agrAmt += a.agreementValue || 0;
      t.gst1 += a.rate === 0.01 ? a.outstandingGst : 0;
      t.gst5 += a.rate === 0.05 ? a.outstandingGst : 0;
      t.advAmt += a.advancesBeforeAgreement || 0;
      t.revGst += a.reverseGstOnAdvances || 0;
      t.paid += a.outstandingGst || 0;
      return t;
    },
    { agrAmt: 0, gst1: 0, gst5: 0, advAmt: 0, revGst: 0, paid: 0 },
  );
  aoa.push([
    '', '', '',
    round2(agrTotals.agrAmt),
    round2(agrTotals.gst1),
    round2(agrTotals.gst5),
    agrTotals.advAmt > 0 ? round2(agrTotals.advAmt) : '',
    agrTotals.revGst > 0 ? round2(agrTotals.revGst) : 0,
    round2(agrTotals.paid),
  ]);

  // Blank rows separator (rows 9-10 in the template)
  aoa.push([]);
  aoa.push([]);

  // ── Advances section (Case 3) ─────────────────────────────────────────────
  aoa.push(['', '', '', 'CALCULATION ON ADVANCES', '', '', '', '', '']);
  aoa.push([]); // blank

  // Sub-header
  aoa.push(['', 'FLAT NO.', 'NAME ', 'ADVANCE AMOUNT', '1% GST', '5% GST', '', '', '']);

  const advances = monthData.advances || [];
  for (const a of advances) {
    const advAmt = (a.amountThisMonth1pct || 0) + (a.amountThisMonth5pct || 0);
    aoa.push([
      '',
      a.flatNo,
      a.name,
      advAmt > 0 ? advAmt : '',
      a.gst1pct > 0 ? a.gst1pct : '',
      a.gst5pct > 0 ? a.gst5pct : '',
      '', '', '',
    ]);
  }

  // Advance filler rows
  const advFillerCount = Math.max(0, 4 - advances.length);
  for (let i = 0; i < advFillerCount; i++) {
    aoa.push(['', '', '', '', '', 0, '', '', '']);
  }

  // Advance totals row + TOTAL PAYABLE
  const mt = monthData.totals;
  const totalPayable = round2((mt.gst1pct || 0) + (mt.gst5pct || 0));

  aoa.push([
    '', '', '',
    mt.amt1pct > 0 ? round2(mt.amt1pct) : 0,
    mt.gst1pct > 0 ? round2(mt.gst1pct) : 0,
    mt.gst5pct > 0 ? round2(mt.gst5pct) : 0,
    0,
    totalPayable,
    '',
  ]);

  const colWidths = [7, 10, 30, 16, 12, 12, 18, 22, 14];
  return aoaToSheet(aoa, colWidths);
}

// ─── SUMMARY sheet ────────────────────────────────────────────────────────────

/**
 * Build the "MMM-YY" summary sheet — OUTWARD SUPPLY / NET PAYABLE grid.
 *
 * Layout (matches CA sample exactly):
 *  Row 1:  Project title
 *  Row 2:  OUTWARD SUPPLY | 1% | 5% | 18% | 28% | CESS
 *  Row 3:  sub-header AMT / CGST / SGST per bracket
 *  Row 4:  SALES  — agreement (outstanding) amounts
 *  Row 5:  ADVANCE — advance-only amounts
 *  Row 6:  blank
 *  Row 7:  INWARD SUPPLY header
 *  Row 8:  sub-header
 *  Row 9:  PURCHASES (zeros)
 *  Row 10: REVERSE (zeros)
 *  Row 11: RCM (zeros)
 *  Row 12: blank
 *  Row 13: IGST OUTWARD SUPPLY header
 *  Row 14: IGST sub-header
 *  Row 15: SALES (zeros)
 *  Row 16: blank
 *  Row 17: IGST INWARD SUPPLY header
 *  Row 18: IGST sub-header
 *  Row 19: PURCHASE-AUG.22 (zeros)
 *  Row 20: PURCHASES AUG.22 (zeros)
 *  Row 21: blank
 *  Row 22: PARTICULARS header
 *  Row 23: GST COLLECTED  | RCM PAYABLE
 *  Row 24: RCM PAYABLE    |
 *  Row 25: ITC CREDIT ON PURCHASES | TAX FREE ...
 *  Row 26: ITC CREDIT ON EXPENSES
 *  Row 27: ITC CREDIT ON RCM
 *  Row 28: PREV CR LEDGER BAL | TOTAL SALES
 *  Row 29: PREVIOUS CASH BAL  | TOTAL PURCHASE
 *  Row 30: NET PAYABLE/REFUND
 *  Row 31: blank
 *  Row 32: PARTICULARS header 2
 *  Row 33: FINAL AMOUNT PAYABLE
 *  Row 34: RCM FINAL PAYABLE
 *  Row 35: TOTAL PAYABLE
 *  Row 36: INTEREST PAYABLE
 *  Row 37: TOTAL CHALLAN PAYABLE
 */
function buildSummarySheet(monthData, projectName) {
  const mt = monthData.totals;

  // Agreement-level amounts (outstanding = agreement value minus advances)
  const salesAmt1pct = monthData.agreements
    .filter((a) => a.rate === 0.01)
    .reduce((s, a) => s + (a.agreementValue - (a.advancesBeforeAgreement || 0)), 0);
  const salesAmt5pct = monthData.agreements
    .filter((a) => a.rate === 0.05)
    .reduce((s, a) => s + (a.agreementValue - (a.advancesBeforeAgreement || 0)), 0);

  const salesGst1pct = round2(salesAmt1pct * 0.01);
  const salesGst5pct = round2(salesAmt5pct * 0.05);
  const salesCgst1 = round2(salesGst1pct / 2);
  const salesSgst1 = salesCgst1;
  const salesCgst5 = round2(salesGst5pct / 2);
  const salesSgst5 = salesCgst5;

  // Advance-only amounts (Case 3)
  const advAmt1pct = round2(monthData.advances.reduce((s, a) => s + (a.amountThisMonth1pct || 0), 0));
  const advAmt5pct = round2(monthData.advances.reduce((s, a) => s + (a.amountThisMonth5pct || 0), 0));
  const advGst1pct = round2(monthData.advances.reduce((s, a) => s + (a.gst1pct || 0), 0));
  const advGst5pct = round2(monthData.advances.reduce((s, a) => s + (a.gst5pct || 0), 0));
  const advCgst1 = round2(advGst1pct / 2);
  const advSgst1 = advCgst1;
  const advCgst5 = round2(advGst5pct / 2);
  const advSgst5 = advCgst5;

  const totalCgst = round2(salesCgst1 + salesCgst5 + advCgst1 + advCgst5);
  const totalSgst = totalCgst;
  const totalGst = round2(totalCgst + totalSgst);
  const totalSales = round2(salesAmt1pct + salesAmt5pct);

  const title = `${projectName.toUpperCase()} GST CALCULATION FOR THE MONTH OF ${monthData.displayName}`;

  const aoa = [
    // Row 1
    [title, '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    // Row 2 — OUTWARD SUPPLY header
    ['OUTWARD SUPPLY', 0.01, '', '', 0.05, '', '', 0.18, '', '', 0.28, '', '', 'CESS', ''],
    // Row 3 — sub-header
    ['', 'AMT', 'CGST', 'SGST', 'AMT', 'CGST', 'SGST', 'AMT', 'CGST', 'SGST', 'AMT', 'CGST', 'SGST', 'AMT', 0.12],
    // Row 4 — SALES (agreement outstanding amounts)
    ['SALES ',
      salesAmt1pct || '', salesCgst1 || 0, salesSgst1 || 0,
      salesAmt5pct || '', salesCgst5 || 0, salesSgst5 || 0,
      '', '', '', '', 0, 0, '', 0],
    // Row 5 — ADVANCE
    ['ADVANCE',
      advAmt1pct || '', advCgst1 || 0, advSgst1 || 0,
      advAmt5pct || 0, advCgst5 || 0, advSgst5 || 0,
      '', '', '', '', 0, 0, '', 0],
    // Row 6 — blank
    ['', ' ', '', '', ' ', '', '', '', '', '', '', '', '', '', ''],
    // Row 7 — INWARD SUPPLY header
    ['INWARD SUPPLY', 0.05, '', '', 0.12, '', '', 0.18, '', '', 0.28, '', '', 'CESS', ''],
    // Row 8
    ['', 'AMT', 'CGST', 'SGST', 'AMT', 'CGST', 'SGST', 'AMT', 'CGST', 'SGST', 'AMT', 'CGST', 'SGST', 'AMT', 0.12],
    // Row 9
    ['PURCHASES', '', 0, 0, '', 0, 0, '', 0, 0, '', 0, 0, '', 0],
    // Row 10
    ['REVERSE', '', 0, 0, '', '', 0, '', 0, 0, '', 0, 0, '', 0],
    // Row 11
    ['RCM', '', 0, 0, '', 0, 0, '', 0, 0, '', 0, 0, '', ''],
    // Row 12 — blank
    [],
    // Row 13 — IGST OUTWARD SUPPLY
    ['IGST OUTWARD SUPPLY', 0.05, '', 0.12, '', 0.18, '', 0.28, '', 'EXEMPT SALES', '', 'NON GST SALES', '', '', ''],
    // Row 14
    ['', 'AMT', 'IGST', 'AMT', 'IGST', 'AMT', 'IGST', 'AMT', 'IGST', 'AMOUNT', '', 'AMOUNT', '', '', ''],
    // Row 15
    ['SALES', '', 0, 0, 0, '', '', '', 0, '', '', '', '', '', ''],
    // Row 16 — blank
    [],
    // Row 17 — IGST INWARD SUPPLY
    ['IGST INWARD SUPPLY', 0.05, '', 0.12, '', 0.18, '', 0.28, '', 'EXEMPT PURCHASE', '', 'NON GST PURCHASES/URD', '', '', ''],
    // Row 18
    ['', 'AMT', 'IGST', 'AMT', 'IGST', 'AMT', 'IGST', 'AMT', 'IGST', 'AMOUNT', '', 'AMOUNT', '', '', ''],
    // Row 19
    ['PURCHASE- AUG.22', 0, 0, 0, 0, '', 0, '', 0, '', '', '', '', '', ''],
    // Row 20
    ['PURCHASES AUG.22', '', 0, '', 0, 0, 0, 0, 0, '', '', '', '', '', ''],
    // Row 21 — blank
    [],
    // Row 22 — PARTICULARS header
    ['PARTICULARS', 'IGST', 'CGST', 'SGST', 'CESS', '', 'PARTICULARS', 'IGST', 'CGST', 'SGST', 'CESS', 'L Y YR GP', '', '', ''],
    // Row 23 — GST COLLECTED
    ['GST COLLECTED', 0, totalCgst, totalSgst, 0, '', 'RCM PAYABLE', 0, 0, 0, 0, '', '', '', ''],
    // Row 24
    ['RCM PAYABLE', '', 0, 0, '', '', '', '', '', '', '', '', '', '', ''],
    // Row 25
    ['ITC CREDIT ON PURCHASES', 0, 0, 0, 0, '', 'TAX FREE', 0.05, 0.12, 0.18, 0.28, 'OVERALL', '', '', ''],
    // Row 26
    ['ITC CREDIT ON EXPENSES', 0, 0, 0, 0, '', '', '', 100, '', '', 100, '', '', ''],
    // Row 27
    ['ITC CREDIT ON RCM', 0, 0, 0, 0, '', '', 'TAXABLE', 'TOTAL', '', '', '', '', '', ''],
    // Row 28
    ['PREV CR LEDGER BAL', '', '', '', '', '', 'TOTAL SALES', totalSales || '', totalSales || '', '', '', '', '', '', ''],
    // Row 29
    ['PREVIOUS CASH BAL', 0, '', '', '', '', 'TOTAL PURCHASE', 0, '', '', '', '', '', '', ''],
    // Row 30 — NET PAYABLE
    ['NET PAYABLE/REFUND', 0, totalCgst, totalSgst, 0, totalGst, '', '', '', '', '', '', '', '', ''],
    // Row 31 — blank
    [],
    // Row 32
    ['PARTICULARS', 'IGST', 'CGST', 'SGST', 'CESS', 'TOTAL', '', '', 'PARTICULARS', 'GST', '', '', '', '', ''],
    // Row 33
    ['FINAL AMOUNT PAYABLE', 0, totalCgst, totalSgst, 0, totalGst, '', '', 'Due Date', '', '', '', '', '', ''],
    // Row 34
    ['RCM FINAL PAYABLE', 0, 0, 0, 0, 0, '', '', 'Payment', '', '', '', '', '', ''],
    // Row 35
    ['TOTAL PAYABLE', 0, totalCgst, totalSgst, 0, totalGst, '', '', 'Delayed Days', 0, '', '', '', '', ''],
    // Row 36
    ['INTEREST PAYABLE', 0, 0, 0, 0, 0, '', '', 'LATE FESS', 0, '', '', '', '', ''],
    // Row 37
    ['TOTAL CHALLAN PAYABLE', 0, totalCgst, totalSgst, 0, totalGst, '', '', 'TOTAL', totalGst, '', '', '', '', ''],
  ];

  const colWidths = [24, 12, 10, 10, 12, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10];
  return aoaToSheet(aoa, colWidths);
}

// ─── Workbook assembler ───────────────────────────────────────────────────────

/**
 * Write the complete GST Calculation Workbook.
 *
 * @param {object[]} records    Combined output of buildRecord + buildTallyOnlyRecords
 * @param {string}   outPath   Destination .xlsx file path
 * @param {object}   [opts]    { projectName }
 */
function writeGstWorkbook(records, outPath, opts = {}) {
  const { groupByMonth } = require('../engine/monthlyGroup');
  const projectName = opts.projectName || DEFAULT_PROJECT;
  const monthMap = groupByMonth(records);

  const wb = XLSX.utils.book_new();

  for (const [label, monthData] of monthMap) {
    // Insert SUMMARY sheet first (matches CA ordering in sample)
    const summarySheet = buildSummarySheet(monthData, projectName);
    XLSX.utils.book_append_sheet(wb, summarySheet, label);

    // Then WORKING sheet
    const workingSheet = buildWorkingSheet(monthData, projectName);
    XLSX.utils.book_append_sheet(wb, workingSheet, `${label} WORKING`);
  }

  XLSX.writeFile(wb, outPath);
  return outPath;
}

module.exports = {
  writeGstWorkbook,
  buildWorkingSheet,
  buildSummarySheet,
};
