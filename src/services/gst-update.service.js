'use strict';

/**
 * GST Workbook Update Service
 *
 * Reads an existing GST Calculation Workbook, identifies which months are
 * already present, and either:
 *   (a) Appends brand-new month sheets for months not yet present, or
 *   (b) Updates existing WORKING + SUMMARY sheets for months with new data
 *       (late entry handling).
 *
 * All existing sheet formatting is preserved via ExcelJS.
 * New month sheets are generated using the existing buildWorkingSheet /
 * buildSummarySheet functions and then converted to ExcelJS format.
 */

const ExcelJS = require('exceljs');
const XLSX    = require('xlsx');

const { groupByMonth, monthKey } = require('../engine/monthlyGroup');
const { buildWorkingSheet, buildSummarySheet } = require('../output/gstWorkbook');

// ── WORKING sheet column indices (1-based for ExcelJS) ──
// Row 3 (header): SR.NO.|FLAT NO.|AGREEMENT DONE|AGR AMT|1% GST|5% GST|ADV AMT|REV GST|AMT PAID
// Advance section (below "CALCULATION ON ADVANCES"):
//   FLAT NO. = col 2, NAME = col 3, ADV AMT = col 4, 1% GST = col 5, 5% GST = col 6

const ADVANCE_SECTION_MARKER = 'CALCULATION ON ADVANCES';

/**
 * Parse a WORKING sheet to extract:
 *   - agreementFlatNos: Set of flat numbers already in Section A
 *   - advanceFlatNos:   Set of flat numbers already in Section B
 *   - advanceSectionRow: row number where Section B begins
 *   - totalRow: row number of the grand totals row (last non-empty row)
 *
 * @param {ExcelJS.Worksheet} ws
 * @returns {object}
 */
function parseWorkingSheet(ws) {
  const agreementFlatNos = new Set();
  const advanceFlatNos   = new Set();
  let   advanceSectionRow = null;
  let   totalRow          = null;
  let   inAdvanceSection  = false;

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const c1 = String(row.getCell(1).value || '').trim();
    const c4 = String(row.getCell(4).value || '').trim();

    // Detect "CALCULATION ON ADVANCES" marker (col 4 in our format)
    if (c4 === ADVANCE_SECTION_MARKER) {
      advanceSectionRow  = rowNumber;
      inAdvanceSection   = true;
      return;
    }

    if (!inAdvanceSection) {
      // Section A (agreements): rows where col 1 is a number (SR.NO.)
      const sr = Number(row.getCell(1).value);
      if (!isNaN(sr) && sr > 0) {
        const flatNo = String(row.getCell(2).value || '').trim();
        if (flatNo) agreementFlatNos.add(flatNo);
      }
    } else {
      // Section B (advances): rows where col 2 has a flat number
      const flatNo = String(row.getCell(2).value || '').trim();
      if (flatNo && flatNo !== 'FLAT NO.') advanceFlatNos.add(flatNo);

      // Last non-empty row in Section B = totals row
      const hasData = [2,3,4,5,6].some(c => row.getCell(c).value != null && row.getCell(c).value !== '');
      if (hasData) totalRow = rowNumber;
    }
  });

  return { agreementFlatNos, advanceFlatNos, advanceSectionRow, totalRow };
}

/**
 * Convert a SheetJS (xlsx) worksheet into an ExcelJS worksheet and append it
 * to an ExcelJS workbook under the given name.
 *
 * This is necessary because the existing buildWorkingSheet / buildSummarySheet
 * functions return SheetJS worksheet objects.
 *
 * @param {ExcelJS.Workbook} targetWb
 * @param {object}           sheetJsWs   SheetJS worksheet object
 * @param {string}           sheetName
 * @returns {ExcelJS.Worksheet}
 */
function appendSheetJsToExcelJs(targetWb, sheetJsWs, sheetName) {
  // Convert SheetJS worksheet to array-of-arrays
  const aoa = XLSX.utils.sheet_to_json(sheetJsWs, { header: 1, defval: null });

  const ws = targetWb.addWorksheet(sheetName);

  // Apply column widths from SheetJS if available
  if (sheetJsWs['!cols']) {
    ws.columns = sheetJsWs['!cols'].map((col, i) => ({
      key: String(i),
      width: col && col.wch ? col.wch : 12,
    }));
  }

  // Write data rows
  aoa.forEach((row) => {
    const exRow = ws.addRow(row);
    // Apply number format to numeric cells
    exRow.eachCell({ includeEmpty: false }, (cell) => {
      if (cell.type === ExcelJS.ValueType.Number) {
        cell.numFmt = '#,##0.00';
      }
    });
  });

  return ws;
}

/**
 * Update the existing GST Workbook with data from the latest Tally records.
 *
 * @param {object} opts
 * @param {string}   opts.existingPath   Path to existing GST Workbook .xlsx
 * @param {string}   opts.outPath        Where to write the updated file
 * @param {object[]} opts.records        Combined records (buildRecord + buildTallyOnlyRecords)
 * @param {string}   [opts.projectName]
 * @returns {Promise<{ outPath, newMonths, updatedMonths }>}
 */
async function updateGstWorkbook(opts) {
  const { existingPath, outPath, records, projectName = 'SPACE HOME' } = opts;

  // ── Load existing workbook ──
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(existingPath);

  // ── Identify which months already exist ──
  const existingSheetNames = new Set(wb.worksheets.map(s => s.name));
  const existingMonths = new Set(
    [...existingSheetNames].filter(n => !n.endsWith(' WORKING'))
  );

  // ── Group the full record set by month ──
  const monthMap = groupByMonth(records);

  const newMonths     = [];
  const updatedMonths = [];

  for (const [label, monthData] of monthMap) {
    const summaryName = label;
    const workingName = `${label} WORKING`;

    if (!existingMonths.has(summaryName)) {
      // ── NEW MONTH: generate fresh sheets and append ──
      const summarySheetJs = buildSummarySheet(monthData, projectName);
      const workingSheetJs = buildWorkingSheet(monthData, projectName);

      appendSheetJsToExcelJs(wb, summarySheetJs, summaryName);
      appendSheetJsToExcelJs(wb, workingSheetJs, workingName);

      newMonths.push(label);
      console.info(`[update-gst] Added new month: ${label}`);
    } else {
      // ── EXISTING MONTH: check for late entries ──
      const existingWs = wb.getWorksheet(workingName);
      if (!existingWs) continue;

      const { advanceFlatNos, advanceSectionRow, totalRow } = parseWorkingSheet(existingWs);

      // Find advance entries in the new month data that are NOT already in the sheet
      const newAdvances = monthData.advances.filter(
        a => !advanceFlatNos.has(String(a.flatNo))
      );

      if (!newAdvances.length) {
        console.info(`[update-gst] ${label}: no new entries, skipping`);
        continue;
      }

      // Insert new advance rows before the totals row
      const insertAt = totalRow || (advanceSectionRow ? advanceSectionRow + 4 : 13);

      for (let i = newAdvances.length - 1; i >= 0; i--) {
        const a = newAdvances[i];
        existingWs.spliceRows(insertAt, 0, []);
        const newRow = existingWs.getRow(insertAt);
        newRow.getCell(2).value = a.flatNo;
        newRow.getCell(3).value = a.name;
        const advAmt = (a.amountThisMonth1pct || 0) + (a.amountThisMonth5pct || 0);
        newRow.getCell(4).value = advAmt || null;
        newRow.getCell(5).value = a.gst1pct || null;
        newRow.getCell(6).value = a.gst5pct || null;
        [4, 5, 6].forEach(c => {
          if (newRow.getCell(c).value != null) newRow.getCell(c).numFmt = '#,##0.00';
        });
        newRow.commit();
      }

      // Recalculate the totals row after insertion
      const newTotalRow = existingWs.getRow(insertAt + newAdvances.length);
      const mt = monthData.totals;
      if (newTotalRow) {
        newTotalRow.getCell(5).value = mt.gst1pct || 0;
        newTotalRow.getCell(6).value = mt.gst5pct || 0;
        newTotalRow.getCell(8).value = (mt.gst1pct || 0) + (mt.gst5pct || 0);
        newTotalRow.commit();
      }

      // Also regenerate the SUMMARY sheet for this month
      const existingSummary = wb.getWorksheet(summaryName);
      if (existingSummary) {
        const idx = wb.worksheets.indexOf(existingSummary);
        wb.removeWorksheet(existingSummary.id);
        const freshSummary = buildSummarySheet(monthData, projectName);
        const newSummaryWs = appendSheetJsToExcelJs(wb, freshSummary, summaryName);
        // Move it back to its original position
        wb.worksheets.splice(idx, 0, ...wb.worksheets.splice(wb.worksheets.indexOf(newSummaryWs), 1));
      }

      updatedMonths.push(label);
      console.info(`[update-gst] Updated ${label}: +${newAdvances.length} advance entries`);
    }
  }

  // ── Write updated workbook ──
  await wb.xlsx.writeFile(outPath);

  return { outPath, newMonths, updatedMonths };
}

module.exports = { updateGstWorkbook };
