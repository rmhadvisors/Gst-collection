'use strict';

/**
 * Agreement Reconciliation Service
 *
 * Handles the scenario where a flat was previously "Tally Only" (no agreement
 * PDF was available) and later an Agreement PDF is uploaded.
 *
 * This service:
 *  1. Accepts a list of new Agreement PDFs
 *  2. OCRs each one via the existing processAgreement() engine
 *  3. Checks if that flat is currently tally-only in the existing record set
 *  4. If so, rebuilds a full agreement record (with 45-lakh GST rule, Case 2 etc.)
 *  5. Returns the updated record set (original records are never mutated)
 */

const { processAgreement } = require('../pdf/agreement');
const { closeOcr }         = require('../pdf/ocr');
const { buildRecord }      = require('../engine/collection');

/**
 * Reconcile tally-only records against a batch of new Agreement PDFs.
 *
 * @param {object[]} records        Current combined record set
 * @param {object}   tally          Parsed tally output (from parseTally/mergeTallies)
 * @param {string[]} agreementPaths Absolute paths to new Agreement PDF files
 * @param {object}   overrides      Per-flat overrides from config/overrides.json
 * @param {string}   workDir        Temp working directory for OCR page extraction
 * @returns {Promise<{ records: object[], reconciled: string[] }>}
 *            updated record set + list of flat numbers that were reconciled
 */
async function reconcileAgreements(records, tally, agreementPaths, overrides = {}, workDir = '') {
  if (!agreementPaths.length) return { records, reconciled: [] };

  // Index current records by flatNo for O(1) lookup
  const byFlat = new Map(records.map(r => [String(r.flatNo), r]));

  const reconciled = [];

  for (const pdfPath of agreementPaths) {
    let agreement;
    try {
      agreement = await processAgreement(pdfPath, { workDir });
    } catch (err) {
      console.warn(`[reconcile] OCR failed for ${pdfPath}: ${err.message}`);
      continue;
    }

    const flatNo = String(agreement.flatNoHint || '').trim();
    if (!flatNo) {
      console.warn(`[reconcile] Could not determine flat number from ${pdfPath}`);
      continue;
    }

    const existing = byFlat.get(flatNo);
    if (!existing) {
      console.info(`[reconcile] Flat ${flatNo} not found in existing records — skipping`);
      continue;
    }

    if (existing.hasAgreement) {
      console.info(`[reconcile] Flat ${flatNo} already has an agreement — skipping`);
      continue;
    }

    // Build full agreement record (replaces the tally-only record)
    const flatOverrides = overrides[flatNo] || {};
    const upgraded = buildRecord(agreement, tally, flatOverrides);
    upgraded.pdf = agreement.pdf;

    byFlat.set(flatNo, upgraded);
    reconciled.push(flatNo);
    console.info(`[reconcile] Flat ${flatNo} upgraded from tally-only to agreement record`);
  }

  await closeOcr();

  // Rebuild records array preserving original order
  const updatedRecords = records.map(r => byFlat.get(String(r.flatNo)) || r);

  return { records: updatedRecords, reconciled };
}

module.exports = { reconcileAgreements };
