'use strict';

/**
 * Collection Sheet Update Service — v3
 *
 * APPROACH: Scan PDFs → Parse existing sheet → Reconcile → Regenerate
 *
 * 1. All uploaded Agreement PDFs are OCR-scanned (same pipeline as generate mode).
 * 2. The existing Wing B is parsed back into record objects.
 * 3. Agreement objects are reconciled with existing records (3 cases).
 * 4. New Tally payments are merged (dedup prevents double-counting).
 * 5. A fresh, styled collection sheet is written.
 *
 * Three reconciliation cases:
 *   Case 1  — Agreement already in sheet → refresh agreement data + new payments
 *   Case 2  — Tally-Only flat, PDF now supplied → upgrade to full agreement record
 *   Case 3  — New flat not in sheet → create fresh record, append at bottom
 *
 * Wing B column layout (0-based, as written by workbook.js):
 *   0  Sr.No.
 *   1  FLOOR
 *   2  SHOP/FLAT NO.
 *   3  FLAT TYPE
 *   4  RERA CARPET AREA
 *   5  REG. NO.
 *   6  DATE OF AGREEMENT
 *   7  PURCHASER'S NAMES
 *   8  AGREEMENT VALUE
 *   9  GST ON AGREEMENT VALUE
 *  10  DATE OF AMOUNT RECEIVED  ← 'TOTAL ::' on flat header rows
 *  11  AMOUNT RECEIVED
 *  12  GST ON ADVANCE
 *  13  TDS
 *  14  BALANCE OUTSTANDING
 */

const XLSX        = require('xlsx');
const path        = require('path');
const { processAgreement }              = require('../pdf/agreement');
const { closeOcr }                      = require('../pdf/ocr');
const { buildRecord }                   = require('../engine/collection');
const { buildTallyOnlyRecords }         = require('../engine/tallyOnly');
const { writeStyledCollectionWorkbook } = require('../output/workbook-styled');
const { buildKeySet, filterNew, normaliseDate } = require('./duplicate-detection.service');
const { formatMDY }                     = require('../util/dates');

// 0-based column indices for Wing B (as written by workbook.js)
const C = {
  SR: 0, FLOOR: 1, FLAT: 2, TYPE: 3, AREA: 4, REG: 5,
  AGR_DATE: 6, NAME: 7, AGR_VAL: 8, GST_AGR: 9,
  DATE_RCV: 10, AMT_RCV: 11, GST_ADV: 12, TDS: 13, BALANCE: 14,
};

const TOTAL_MARKER = 'TOTAL ::';

// ─────────────────────────────────────────────────────────────────────────────
// Wing B parser
// ─────────────────────────────────────────────────────────────────────────────

// ─── Excel serial date converter ────────────────────────────────────────────

/**
 * Convert a raw Wing B cell value (col 10 / col 6) to a normalised date string.
 *
 * Wing B cells may contain:
 *   - A plain text string  e.g. "1/12/26" or "12/1/2026"
 *   - An Excel serial number e.g. 45335   (days since 1900-01-00)
 *   - A JS Date object (SheetJS sometimes returns these)
 *
 * We always normalise to M/D/YY format (the same formatMDY produces) so that
 * dedup keys built from Wing B payments match keys built from Tally payments.
 */
function cellToDateText(raw) {
  if (raw == null || raw === '') return '';
  // JS Date object
  if (raw instanceof Date) return formatMDY(raw);
  const n = Number(raw);
  // Excel serial: integer > 1000 and no slash characters → treat as serial
  if (!isNaN(n) && n > 1000 && !String(raw).includes('/') && !String(raw).includes('-')) {
    // Excel epoch: Jan 0 1900 = serial 0; JS epoch differs by 25569 days
    // Also Excel incorrectly counts 1900-02-29, so subtract 1 for dates >= 60
    const offset = n >= 60 ? 25568 : 25569;
    const ms     = (n - offset) * 86400 * 1000;
    const d      = new Date(ms);
    if (!isNaN(d.getTime())) return formatMDY(d);
  }
  return String(raw).trim();
}

/**
 * Parse a Wing B sheet (SheetJS aoa format) into an array of record-like objects.
 * Each record contains its existing payments list.
 *
 * IMPORTANT: Use sheet_to_json with { raw: true } so numbers stay as numbers,
 * then convert dates via cellToDateText() which handles Excel serials.
 *
 * @param {Array[]} aoa  Raw rows from sheet_to_json({ header:1, raw:true })
 * @returns {object[]}
 */
function parseWingBAoa(aoa) {
  const records = [];
  let current = null;

  for (let i = 0; i < aoa.length; i++) {
    const row = aoa[i];
    if (!row || row.length === 0) continue;

    // col10 may be 'TOTAL ::' (string) or a date value
    const raw10 = row[C.DATE_RCV];
    const col10 = raw10 != null ? String(raw10).trim() : '';
    const col2  = String(row[C.FLAT] ?? '').trim();

    if (col10 === TOTAL_MARKER && col2 !== '') {
      // ── Flat header row ──
      if (current) records.push(current);

      const agrVal = row[C.AGR_VAL];
      const gstAgr = row[C.GST_AGR];
      const hasAgreement = agrVal != null && agrVal !== '' && Number(agrVal) > 0;

      // Agreement date may be an Excel serial or text
      const rawAgrDate = row[C.AGR_DATE];
      const agreementDate = rawAgrDate != null && rawAgrDate !== '' ? cellToDateText(rawAgrDate) : '';

      current = {
        flatNo:         col2,
        floorLabel:     String(row[C.FLOOR]   ?? '').trim(),
        flatType:       String(row[C.TYPE]    ?? '').trim(),
        carpetArea:     row[C.AREA]  != null && row[C.AREA]  !== '' ? Number(row[C.AREA])  : null,
        regNo:          String(row[C.REG]     ?? '').trim(),
        agreementDate,
        name:           String(row[C.NAME]    ?? '').trim(),
        agreementValue: hasAgreement ? Number(agrVal) : null,
        gstOnAgreement: gstAgr != null && gstAgr !== '' ? Number(gstAgr) : null,
        hasAgreement,
        gstRate:        hasAgreement ? (Number(agrVal) > 4500000 ? 0.05 : 0.01) : 0.01,
        payments:       [],
      };
    } else if (current && raw10 != null && raw10 !== '' && col10 !== TOTAL_MARKER
               && String(row[C.SR] ?? '').trim() === '') {
      // ── Payment row (no Sr.No., has a value in col 10) ──
      const amt = row[C.AMT_RCV];
      if (amt != null && amt !== '' && Number(amt) > 0) {
        // Convert Excel serial date to M/D/YY text for dedup consistency
        const dateText = cellToDateText(raw10);
        current.payments.push({
          dateText,
          amount:   Number(amt),
          gst:      row[C.GST_ADV] != null && row[C.GST_ADV] !== '' ? Number(row[C.GST_ADV]) : 0,
          vchNo:    '',
        });
      }
    }
  }
  if (current) records.push(current);
  return records;
}

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function round2(n) { return Math.round((n + Number.EPSILON) * 100) / 100; }

/**
 * Convert a Wing B existing record into a payment-source object compatible
 * with the tally dedup service (so existing payments are fingerprinted).
 */
function existingBlock(rec) {
  return { flatNo: rec.flatNo, payments: rec.payments };
}

// ─────────────────────────────────────────────────────────────────────────────
// Main service
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Update the existing Collection Sheet with new Agreement PDFs + new Tally.
 *
 * @param {object}   opts
 * @param {string}   opts.existingPath    Path to existing Collection Sheet
 * @param {string}   opts.outPath         Output path for updated file
 * @param {object}   opts.tally           Parsed tally (parseTally output)
 * @param {string[]} opts.agreementPaths  Uploaded Agreement PDF file paths
 * @param {object}   opts.overrides       config/overrides.json
 * @param {string}   [opts.workDir]       Temp dir for OCR page images
 * @param {string}   [opts.projectName]
 * @returns {Promise<{ outPath, newPayments, reconciledFlats, newFlats, pdfWarnings }>}
 */
async function updateCollectionSheet(opts) {
  const {
    existingPath,
    outPath,
    tally,
    agreementPaths = [],
    overrides      = {},
    workDir        = path.join(path.dirname(outPath), '_pages_' + Date.now()),
    projectName    = 'YASHWANTH COUNTY',
  } = opts;

  const pdfWarnings   = [];  // { file, error }
  const reconciledFlats = [];
  const newFlats        = [];

  // ── STEP 1: Read existing Wing B ──────────────────────────────────────────
  const wb = XLSX.readFile(existingPath);
  const wingBName = wb.SheetNames.find(n => n.trim() === 'Wing B');
  if (!wingBName) {
    throw new Error(
      `Existing file has no "Wing B" sheet.\nFound: ${wb.SheetNames.join(', ')}\n` +
      `Upload a Collection Sheet generated by this system.`
    );
  }

  const ws  = wb.Sheets[wingBName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const existingRecords = parseWingBAoa(aoa);
  if (!existingRecords.length) {
    throw new Error('Could not parse any flat records from Wing B.');
  }

  console.info(`[update] Parsed ${existingRecords.length} flat records from existing sheet`);

  // Index existing records by flat number
  const byFlat = new Map(existingRecords.map(r => [String(r.flatNo), r]));

  // Build duplicate-fingerprint set from every existing payment in Wing B
  const keySet = buildKeySet(existingRecords.map(existingBlock));

  // ── STEP 2: OCR scan all Agreement PDFs ──────────────────────────────────
  const agreementObjects = [];   // { flatNoHint, agreementValue, names, ... }

  for (const pdfPath of agreementPaths) {
    const basename = path.basename(pdfPath);
    try {
      const agreement = await processAgreement(pdfPath, { workDir });
      agreementObjects.push(agreement);
      console.info(`[update] PDF OK: ${basename} → flat ${agreement.flatNoHint}`);
    } catch (e) {
      const msg = `Agreement could not be processed: ${basename} — ${e.message}`;
      pdfWarnings.push({ file: basename, error: e.message });
      console.warn(`[update] PDF FAILED: ${msg}`);
    }
  }

  if (agreementPaths.length) {
    try { await closeOcr(); } catch (_) {}
  }

  // ── STEP 3: Reconcile each Agreement PDF ─────────────────────────────────
  //
  //   Case 1: flat already has agreement → refresh value / name if OCR differs
  //   Case 2: flat was Tally-Only → upgrade to full agreement record
  //   Case 3: flat not in sheet → create fresh, append later

  const case3Records = [];   // brand-new flats (not in existing sheet)

  for (const agreement of agreementObjects) {
    const flatNo = String(agreement.flatNoHint || '').trim();
    if (!flatNo) { pdfWarnings.push({ file: agreement.pdf, error: 'Could not detect flat number from PDF.' }); continue; }

    const ov  = overrides[flatNo] || {};
    const rec = byFlat.get(flatNo);

    if (!rec) {
      // ── Case 3: new flat ──
      const newRec = buildRecord(agreement, tally, ov);
      newRec.pdf   = agreement.pdf;
      case3Records.push(newRec);
      newFlats.push(flatNo);
      console.info(`[update] Case 3 (new flat): ${flatNo}`);
      continue;
    }

    // ── Case 1 or 2: flat exists in sheet ──
    // Strategy:
    //   a) Build the fresh record from the PDF + tally (correct GST rate etc.)
    //   b) Find Wing B payments that are NOT covered by the tally at all
    //      (e.g. manual CA entries from months not in this tally export).
    //      Match by normalised date + amount since Wing B has no vchNo.
    //   c) Merge: fresh tally payments + Wing B-only payments, dedup, sort.
    const freshRec = buildRecord(agreement, tally, ov);

    // Build a date+amount fingerprint set from the tally for this flat
    const tallyPaymentsForFlat = (tally.byFlat.get(flatNo) || { payments: [] }).payments;
    const tallyDateAmtKeys = new Set(
      tallyPaymentsForFlat.map(p => {
        const d = p.date ? normaliseDate(p.date) : normaliseDate(formatMDY(p.date));
        return `${d}|${Number(p.amount).toFixed(2)}`;
      })
    );

    // Wing B payments whose date+amount does NOT appear in the tally at all
    // (these are CA manual entries that must be preserved)
    const existingOnlyPayments = rec.payments.filter(p => {
      const d   = normaliseDate(p.dateText);
      const key = `${d}|${Number(p.amount).toFixed(2)}`;
      return !tallyDateAmtKeys.has(key);
    });

    // Merge and dedup by date+amount
    const seen = new Set();
    const mergedPayments = [];
    for (const p of [...freshRec.payments, ...existingOnlyPayments]) {
      const d   = normaliseDate(p.dateText || (p.date ? formatMDY(p.date) : ''));
      const key = `${d}|${Number(p.amount).toFixed(2)}`;
      if (!seen.has(key)) { seen.add(key); mergedPayments.push(p); }
    }
    mergedPayments.sort((a, b) => {
      const da = a.date ? a.date.getTime() : (normaliseDate(a.dateText) || '').localeCompare('');
      const db = b.date ? b.date.getTime() : 0;
      return da - db;
    });

    // Recompute totals with merged payments at the correct GST rate
    const rate = freshRec.gstRate || 0.01;
    const totalReceived = round2(mergedPayments.reduce((s, p) => s + p.amount, 0));
    const totalGst      = round2(mergedPayments.reduce((s, p) => s + (p.gst || round2(p.amount * rate)), 0));
    const balance = freshRec.agreementValue != null
      ? round2(freshRec.agreementValue + (freshRec.gstOnAgreement || 0) - totalReceived)
      : null;

    // Update the existing record in-place (same object reference is in existingRecords[])
    const wasUpgraded = !rec.hasAgreement;
    Object.assign(rec, {
      ...freshRec,
      payments:      mergedPayments,
      totalReceived,
      totalGst,
      balance,
    });

    reconciledFlats.push(flatNo);
    console.info(`[update] Case ${wasUpgraded ? 2 : 1} (${wasUpgraded ? 'upgraded' : 'refreshed'}): flat ${flatNo}, ${mergedPayments.length} total payments`);
  }

  // ── STEP 4: Merge NEW Tally payments for non-reconciled flats ─────────────
  let totalNewPayments = 0;

  for (const [flatNo, group] of tally.byFlat) {
    const rec = byFlat.get(flatNo);
    if (!rec) continue;

    // Skip flats that were already fully rebuilt from a PDF (Step 3)
    if (reconciledFlats.includes(flatNo)) continue;

    const tallyPayments = group.payments.map(p => ({
      dateText: p.date ? formatMDY(p.date) : '',
      date:     p.date,
      amount:   p.amount,
      gst:      p.gst != null ? p.gst : round2(p.amount * (rec.gstRate || 0.01)),
      vchNo:    p.vchNo || '',
    }));

    const newPayments = filterNew(flatNo, tallyPayments, keySet);
    if (!newPayments.length) continue;

    for (const p of newPayments) {
      rec.payments.push({
        dateText: p.dateText,
        amount:   p.amount,
        gst:      p.gst,
        vchNo:    p.vchNo,
      });
    }

    totalNewPayments += newPayments.length;

    // Recompute totals
    const rate = rec.gstRate || 0.01;
    rec.totalReceived = round2(rec.payments.reduce((s, p) => s + p.amount, 0));
    rec.totalGst      = round2(rec.payments.reduce((s, p) => s + (p.gst || round2(p.amount * rate)), 0));
    if (rec.agreementValue != null && rec.gstOnAgreement != null) {
      rec.balance = round2(rec.agreementValue + rec.gstOnAgreement - rec.totalReceived);
    }

    console.info(`[update] Flat ${flatNo}: +${newPayments.length} new payment(s)`);
  }

  // ── STEP 5: Add any brand-new tally flats not in existing sheet ───────────
  const allFlatNosInSheet = new Set([...byFlat.keys(), ...newFlats]);
  const tallyOnlyNew = buildTallyOnlyRecords(tally, allFlatNosInSheet, overrides);
  for (const rec of tallyOnlyNew) {
    newFlats.push(rec.flatNo);
    existingRecords.push(rec);
    totalNewPayments += rec.payments.length;
    console.info(`[update] New tally-only flat: ${rec.flatNo}`);
  }

  // ── STEP 6: Append Case 3 records (from PDFs for brand-new flats) ─────────
  for (const rec of case3Records) {
    existingRecords.push(rec);
  }

  // ── STEP 7: Sort and deduplicate by flat number (safety guard) ──────────
  // ONE FLAT = ONE BLOCK. Merge any accidental duplicates before writing.
  existingRecords.sort((a, b) => {
    const na = isNaN(Number(a.flatNo)) ? a.flatNo : Number(a.flatNo);
    const nb = isNaN(Number(b.flatNo)) ? b.flatNo : Number(b.flatNo);
    return na < nb ? -1 : na > nb ? 1 : 0;
  });

  // Consolidate: if same flatNo appears more than once, keep only the record
  // with the most data (agreement-based wins over tally-only; more payments wins).
  const seen = new Map();
  const consolidatedWarnings = [];
  const finalRecords = [];

  for (const rec of existingRecords) {
    const key = String(rec.flatNo);
    if (!seen.has(key)) {
      seen.set(key, rec);
      finalRecords.push(rec);
    } else {
      const existing = seen.get(key);
      consolidatedWarnings.push(key);
      console.warn(`[update] ⚠ Duplicate flat ${key} detected — merging into single block`);

      // Agreement record wins over tally-only
      const master = (rec.hasAgreement && !existing.hasAgreement) ? rec : existing;
      const other  = master === rec ? existing : rec;

      // Merge payments: keep all unique date+amount combinations
      const paymentSeen = new Set();
      const allPayments = [];
      for (const p of [...master.payments, ...other.payments]) {
        const d   = normaliseDate(p.dateText || (p.date ? formatMDY(p.date) : ''));
        const pKey = `${d}|${Number(p.amount).toFixed(2)}`;
        if (!paymentSeen.has(pKey)) { paymentSeen.add(pKey); allPayments.push(p); }
      }
      allPayments.sort((a, b) => (normaliseDate(a.dateText) || '').localeCompare(normaliseDate(b.dateText) || ''));

      const rate2 = master.gstRate || 0.01;
      master.payments      = allPayments;
      master.totalReceived = round2(allPayments.reduce((s, p) => s + p.amount, 0));
      master.totalGst      = round2(allPayments.reduce((s, p) => s + (p.gst || round2(p.amount * rate2)), 0));
      if (master.agreementValue != null && master.gstOnAgreement != null) {
        master.balance = round2(master.agreementValue + master.gstOnAgreement - master.totalReceived);
      }

      // Remove the duplicate from finalRecords and replace with master
      const idx = finalRecords.indexOf(other === existing ? existing : rec);
      if (idx >= 0) finalRecords.splice(idx, 1);
      if (!finalRecords.includes(master)) finalRecords.push(master);
      seen.set(key, master);
    }
  }

  if (consolidatedWarnings.length) {
    console.warn(`[update] Consolidated ${consolidatedWarnings.length} duplicate flat(s): ${consolidatedWarnings.join(', ')}`);
    pdfWarnings.push(...consolidatedWarnings.map(f => ({
      file: `Flat ${f}`,
      error: 'Duplicate block detected and automatically consolidated into one record.',
    })));
  }

  // Final sort of the clean list
  finalRecords.sort((a, b) => {
    const na = isNaN(Number(a.flatNo)) ? a.flatNo : Number(a.flatNo);
    const nb = isNaN(Number(b.flatNo)) ? b.flatNo : Number(b.flatNo);
    return na < nb ? -1 : na > nb ? 1 : 0;
  });

  await writeStyledCollectionWorkbook(finalRecords, outPath, { projectName });
  console.info(`[update] Written ${finalRecords.length} flat records: ${outPath}`);

  return {
    outPath,
    newPayments:     totalNewPayments,
    reconciledFlats,
    newFlats,
    pdfWarnings,
  };
}

module.exports = { updateCollectionSheet, parseWingBAoa };
