'use strict';

const { floorFromFlat, floorLabel, nameSimilarity } = require('../util/text');
const { parseDate, formatMDY } = require('../util/dates');

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

/**
 * Match an agreement to a tally flat group.
 * Prefers the flat-number hint from the PDF filename, falls back to name match.
 */
function matchTallyFlat(agreement, tally, overrideFlat) {
  if (overrideFlat && tally.byFlat.has(String(overrideFlat))) {
    return { flatNo: String(overrideFlat), reason: 'override', score: 1 };
  }
  if (agreement.flatNoHint && tally.byFlat.has(agreement.flatNoHint)) {
    return { flatNo: agreement.flatNoHint, reason: 'filename', score: 1 };
  }
  // fall back to fuzzy name match against the tally buyer names
  let best = null;
  for (const g of tally.byFlat.values()) {
    for (const nm of g.nameList) {
      const score = nameSimilarity(agreement.primaryName || '', nm);
      if (!best || score > best.score) best = { flatNo: g.flatNo, reason: 'name', score };
    }
  }
  if (best && best.score >= 0.5) return best;
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
  const flatNo = override.flatNo || (match && match.flatNo) || agreement.flatNoHint || '';
  const group = flatNo ? tally.byFlat.get(String(flatNo)) : null;

  const agreementValue = override.agreementValue != null
    ? override.agreementValue
    : agreement.agreementValue;

  const rate = agreementValue != null ? gstRate(agreementValue) : null;
  const gstOnAgreement = agreementValue != null ? round2(agreementValue * rate) : null;

  // Resolve the agreement date for Case 2 detection
  const agreementDateStr = override.agreementDate || agreement.agreementDateText || '';
  const agreementDateObj = agreementDateStr ? parseDate(agreementDateStr) : null;

  let payments = (group ? group.payments : []).map((p) => ({
    date: p.date,
    dateText: formatMDY(p.date),
    amount: p.amount,
    // Agreement-based rows must keep the existing 45-lakh GST logic.
    gst: rate != null ? round2(p.amount * rate) : null,
    vchNo: p.vchNo,
    source: 'tally',
  }));

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

    if (preAgreementPayments.length > 0) {
      gstCase = 2;
      advancesBeforeAgreement = round2(
        preAgreementPayments.reduce((s, p) => s + p.amount, 0),
      );
      // Reverse GST = advances × rate (same rate as the agreement rate)
      reverseGstOnAdvances = round2(advancesBeforeAgreement * rate);
      // Outstanding GST = GST on full agreement − reverse GST on advances
      outstandingGst = round2((gstOnAgreement || 0) - reverseGstOnAdvances);
    }
  }

  const floorNum = override.floorNum != null ? override.floorNum : floorFromFlat(flatNo);
  const name = override.name
    || agreement.primaryName
    || (group && group.nameList[0])
    || '';

  return {
    flatNo: String(flatNo),
    floorNum,
    floorLabel: override.floor || floorLabel(floorNum),
    flatType: override.flatType || '',
    carpetArea: override.carpetArea != null ? override.carpetArea : '',
    regNo: override.regNo || '',
    agreementDate: agreementDateStr,
    agreementDateObj,
    name,
    allottees: agreement.names && agreement.names.length ? agreement.names : (group ? group.nameList : []),
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
    valueSource: override.agreementValue != null ? 'override' : 'ocr',
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
