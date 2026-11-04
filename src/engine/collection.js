'use strict';

const {
  floorFromFlat, floorLabel, nameSimilarity, fuzzyNameSimilarity, HAS_DEVANAGARI,
} = require('../util/text');
const { normalizeFlatKey } = require('../util/flatKey');
const { voucherFlags } = require('./unallocated');
const { parseDate, formatMDY } = require('../util/dates');
const { transliterateName } = require('../util/translit');

const GST_THRESHOLD = 4500000; // 45 lakh
const GST_RATE_HIGH = 0.05; // > 45 lakh
const GST_RATE_LOW = 0.01; // <= 45 lakh

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

/** GST rate based on the affordable-housing 45-lakh threshold. */
function gstRate(agreementValue) {
  return agreementValue > GST_THRESHOLD ? GST_RATE_HIGH : GST_RATE_LOW;
}

function byDate(a, b) {
  const da = a.date ? a.date.getTime() : 0;
  const db = b.date ? b.date.getTime() : 0;
  return da - db;
}

/**
 * Apply per-flat overrides to a payment list:
 *   - paymentCutoff: drop payments after a date (match a CA snapshot)
 *   - excludeVchNos: drop specific voucher numbers
 *   - extraPayments: append manual rows not present in Tally
 * `gstForAmount(amount)` computes GST for appended manual rows.
 */
function applyPaymentOverrides(payments, override, gstForAmount) {
  let out = payments;
  if (override.paymentCutoff) {
    const cutoff = parseDate(override.paymentCutoff);
    if (cutoff) {
      out = out.filter((p) => !p.date || p.date.getTime() <= cutoff.getTime());
    }
  }
  if (Array.isArray(override.excludeVchNos) && override.excludeVchNos.length) {
    const skip = new Set(override.excludeVchNos.map(String));
    out = out.filter((p) => !skip.has(String(p.vchNo)));
  }
  if (Array.isArray(override.extraPayments)) {
    out = out.concat(override.extraPayments.map((p) => {
      const date = p.date instanceof Date ? p.date : parseDate(p.date);
      const amount = Number(p.amount);
      return {
        date,
        dateText: p.dateText || formatMDY(date),
        amount,
        gst: p.gst != null ? Number(p.gst) : gstForAmount(amount),
        vchNo: p.vchNo || '',
        source: p.source || 'manual',
      };
    }));
    out.sort(byDate);
  }
  return out;
}

const NAME_MATCH_MIN = 0.75;

/** Tally group for a flat key; "B 607"/"B-607"/"B- 607" are the same key. */
function findTallyGroup(tally, rawKey) {
  const key = normalizeFlatKey(rawKey);
  if (!key) return null;
  if (tally.byFlat.has(key)) return { flatNo: key, group: tally.byFlat.get(key), reason: 'flat-key' };
  // wing missing on one side: match on the bare flat number when it is unambiguous
  const digits = key.replace(/D/g, '');
  const sameNumber = [...tally.byFlat.keys()].filter((k) => k.replace(/D/g, '') === digits);
  if (sameNumber.length === 1) {
    return { flatNo: sameNumber[0], group: tally.byFlat.get(sameNumber[0]), reason: 'flat-number' };
  }
  return null;
}

function nameScore(a, b) {
  if (HAS_DEVANAGARI.test(a || '') || HAS_DEVANAGARI.test(b || '')) return fuzzyNameSimilarity(a, b);
  return Math.max(nameSimilarity(a, b), fuzzyNameSimilarity(a, b) >= NAME_MATCH_MIN ? NAME_MATCH_MIN : 0);
}

/**
 * Match an agreement to Tally receipts.
 * 1. override flat  2. flat key from OCR/filename (wing-normalised)
 * 3. fuzzy purchaser-name match against flat groups
 * 4. fuzzy purchaser-name match against "Capital Ac" ledgers (flagged review)
 */
function matchTallyFlat(agreement, tally, overrideFlat) {
  if (overrideFlat && tally.byFlat.has(String(overrideFlat))) {
    return { flatNo: String(overrideFlat), group: tally.byFlat.get(String(overrideFlat)), reason: 'override', score: 1 };
  }
  const byKey = findTallyGroup(tally, agreement.flatNoHint);
  if (byKey) {
    return { ...byKey, reason: agreement.flatSource === 'ocr' ? 'ocr-flat' : 'filename', score: 1 };
  }

  const name = agreement.primaryName || '';
  if (!name) return null;

  let best = null;
  for (const g of tally.byFlat.values()) {
    for (const nm of g.nameList) {
      const score = nameScore(name, nm);
      if (!best || score > best.score) best = { flatNo: g.flatNo, group: g, reason: 'name', score };
    }
  }
  if (best && best.score >= (HAS_DEVANAGARI.test(name) ? NAME_MATCH_MIN : 0.5)) return best;

  let ledger = null;
  for (const l of (tally.ledgers ? tally.ledgers.values() : [])) {
    const score = nameScore(name, l.name.replace(/capital\s*a\/?c/i, '').trim());
    if (!ledger || score > ledger.score) ledger = { ledger: l, reason: 'name-capital-ac (review)', score };
  }
  if (ledger && ledger.score >= NAME_MATCH_MIN) return ledger;
  return null;
}

/**
 * Build a complete collection record for one flat.
 *
 * @param {object} agreement  output of processAgreement()
 * @param {object} tally      output of parseTally()
 * @param {object} [override] per-flat overrides { flatNo, agreementValue, name,
 *                            flatType, carpetArea, regNo, agreementDate,
 *                            extraPayments: [{ date, amount }] }
 */
function buildRecord(agreement, tally, override = {}) {
  const match = matchTallyFlat(agreement, tally, override.flatNo);
  const byName = !!(match && /^name/.test(match.reason));
  // A name-only match must not relabel the flat the agreement itself names.
  const flatNo = override.flatNo
    || (byName ? agreement.flatNoHint : match && match.flatNo)
    || agreement.flatNoHint || '';
  // Capital-Ac ledger matches are only a review hint: their receipts stay in
  // the Unallocated sheet and are never auto-added to the flat.
  const group = match ? (match.group || null) : null;
  const needsReview = !!(match && /review/.test(match.reason));

  const agreementValue = override.agreementValue != null
    ? override.agreementValue
    : agreement.agreementValue;

  const rate = agreementValue != null ? gstRate(agreementValue) : null;
  const gstOnAgreement = agreementValue != null ? round2(agreementValue * rate) : null;

  // Resolve the agreement date for Case 2 detection
  // "AGMT DT" in the Tally narration is used only when neither override nor agreement gives a date.
  const agreementDateStr = override.agreementDate || agreement.agreementDateText
    || (group && group.agmtDate ? formatMDY(group.agmtDate) : '');
  const agreementDateObj = agreementDateStr ? parseDate(agreementDateStr) : null;

  let payments = (group ? group.payments : []).map((p) => ({
    date: p.date,
    dateText: formatMDY(p.date),
    amount: p.amount,
    // Agreement-based rows must keep the existing 45-lakh GST logic.
    gst: rate != null ? round2(p.amount * rate) : null,
    vchNo: p.vchNo,
    dupVoucher: !!p.dupVoucher,
    noVoucher: !!p.noVoucher,
    renumberedFrom: p.renumberedFrom || '',
    chqReturn: !!p.chqReturn,
    manualAllocation: !!p.manualAllocation,
    allocatedFrom: p.allocatedFrom || '',
    source: 'tally',
  }));

  // A single Tally receipt equal to exactly the GST on the agreement value is the
  // GST payment itself: it stays under the flat but is not charged GST again.
  if (gstOnAgreement) {
    const gstReceipt = payments.find((p) => round2(p.amount) === gstOnAgreement);
    if (gstReceipt) { gstReceipt.gst = 0; gstReceipt.gstReceipt = true; }
  }

  payments = applyPaymentOverrides(
    payments,
    override,
    (amount) => (rate != null ? round2(amount * rate) : null),
  );

  const totalReceived = round2(payments.reduce((s, p) => s + p.amount, 0));
  const totalGst = round2(payments.reduce((s, p) => s + (p.gst || 0), 0));

  // balance = (agreement value + gst on agreement) - amount received
  const balance = agreementValue != null
    ? round2(agreementValue + (gstOnAgreement || 0) - totalReceived)
    : null;

  // ── Case 2: Advances received BEFORE the agreement date ──────────────────
  // Split payments into pre- and post-agreement buckets.
  // Pre-agreement payments were initially taxed at 1% (Case 3 treatment),
  // so once the agreement is executed we must reverse that GST and compute
  // the outstanding GST on the net outstanding amount.
  let preAgreementPayments = [];
  let postAgreementPayments = payments;
  let advancesBeforeAgreement = 0;
  let reverseGstOnAdvances = 0;
  let outstandingGst = gstOnAgreement || 0;
  let gstCase = 1; // default: agreement exists, no prior advances

  if (agreementDateObj && payments.length) {
    preAgreementPayments = payments.filter(
      (p) => p.date && p.date < agreementDateObj,
    );
    postAgreementPayments = payments.filter(
      (p) => !p.date || p.date >= agreementDateObj,
    );

    if (preAgreementPayments.some((p) => !p.gstReceipt)) {
      gstCase = 2;
      advancesBeforeAgreement = round2(
        preAgreementPayments.filter((p) => !p.gstReceipt).reduce((s, p) => s + p.amount, 0),
      );
      // Reverse GST = advances × rate (same rate as the agreement rate)
      reverseGstOnAdvances = round2(advancesBeforeAgreement * rate);
      // Outstanding GST = GST on full agreement − reverse GST on advances
      outstandingGst = round2((gstOnAgreement || 0) - reverseGstOnAdvances);
    }
  }

  const floorNum = override.floorNum != null ? override.floorNum : floorFromFlat(flatNo);
  // Marathi (Devanagari) purchaser names are shown transliterated; the original
  // goes to the Audit sheet. A Latin Tally name still wins over a Marathi OCR name.
  const tallyName = group && group.nameList ? group.nameList[0] : '';
  const ocrName = agreement.primaryName || '';
  const ocrIsMarathi = HAS_DEVANAGARI.test(ocrName);
  const name = override.name
    || (ocrIsMarathi ? (tallyName || transliterateName(ocrName)) : (ocrName || tallyName));

  return {
    flatNo: String(flatNo),
    floorNum,
    floorLabel: override.floor || floorLabel(floorNum),
    flatType: override.flatType || '',
    carpetArea: override.carpetArea != null ? override.carpetArea
      : (agreement.carpetAreaSqm != null ? agreement.carpetAreaSqm : ''),
    regNo: override.regNo || '',
    agreementDate: agreementDateStr,
    agreementDateObj,
    name,
    purchaserMarathi: ocrIsMarathi ? ocrName : '',
    agreementDateSource: agreement.dateSource || (agreementDateStr ? 'agreement' : ''),
    allottees: agreement.names && agreement.names.length ? agreement.names : (group && group.nameList ? group.nameList : []),
    agreementValue,
    gstRate: rate,
    gstOnAgreement,
    agreementValueWithGst: agreementValue != null ? round2(agreementValue + gstOnAgreement) : null,
    payments,
    preAgreementPayments,
    postAgreementPayments,
    advancesBeforeAgreement,
    reverseGstOnAdvances,
    outstandingGst,
    gstCase,
    totalReceived,
    totalGst,
    balance,
    matchReason: match ? match.reason : (override.flatNo ? 'override' : 'unmatched'),
    matchScore: match ? match.score : 0,
    valueSource: override.agreementValue != null ? 'override'
      : (agreementValue == null ? 'manual needed' : (agreement.valueMethod === 'position' ? 'ocr (position)' : 'ocr')),
    needsReview,
    ocrFailed: !!agreement.ocrFailed,
    flags: [...(agreement.ocrFailed ? [`OCR FAILED (page ${agreement.ocrFailedPages.join(', ')}: under 50 characters read) — check the PDF manually`] : []), ...voucherFlags(payments), ...(needsReview ? [`Capital Ac ledger "${match.ledger.name}" matches by name (${Math.round(match.score * 100) / 100}) (review): its receipts are NOT added to this flat`] : [])],
    tallyFlat: match && match.group && match.flatNo ? match.flatNo : null,
    valueCandidates: agreement.valueCandidates || [],
    hasAgreement: true,
  };
}

module.exports = {
  buildRecord,
  gstRate,
  matchTallyFlat,
  GST_THRESHOLD,
};
