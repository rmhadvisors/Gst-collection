'use strict';

const { floorFromFlat, floorLabel } = require('../util/text');
const { parseDate, formatMDY } = require('../util/dates');

const DEFAULT_TALLY_ONLY_GST_RATE = 0.01;
const TALLY_ONLY_STATUS = 'Tally Only – Agreement Not Available';

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function byDate(a, b) {
  const da = a.date ? a.date.getTime() : 0;
  const db = b.date ? b.date.getTime() : 0;
  return da - db;
}

function applyTallyOnlyOverrides(payments, override, defaultRate) {
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
      const rate = p.gstRate != null ? p.gstRate : defaultRate;
      return {
        date,
        dateText: p.dateText || formatMDY(date),
        amount,
        gst: p.gst != null ? Number(p.gst) : round2(amount * rate),
        gstRate: rate,
        vchNo: p.vchNo || '',
        source: p.source || 'manual',
      };
    }));
    out.sort(byDate);
  }

  return out;
}

/**
 * Build a collection record for a flat that exists only in Tally receipts.
 * Agreement fields are intentionally blank; GST defaults to 1% unless a future
 * GST column is present in the Tally row.
 */
function buildTallyOnlyRecord(group, override = {}, opts = {}) {
  const defaultRate = opts.defaultRate != null ? opts.defaultRate : DEFAULT_TALLY_ONLY_GST_RATE;
  const flatNo = override.flatNo || group.flatNo;

  let payments = (group.payments || []).map((p) => {
    const rate = p.gstRate != null ? p.gstRate : defaultRate;
    return {
      date: p.date,
      dateText: formatMDY(p.date),
      amount: p.amount,
      gst: round2(p.amount * rate),
      gstRate: rate,
      vchNo: p.vchNo,
      source: 'tally',
    };
  });

  payments = applyTallyOnlyOverrides(payments, override, defaultRate);

  const totalReceived = round2(payments.reduce((sum, p) => sum + p.amount, 0));
  const totalGst = round2(payments.reduce((sum, p) => sum + (p.gst || 0), 0));
  const floorNum = override.floorNum != null ? override.floorNum : floorFromFlat(flatNo);
  const name = override.name || (group.nameList && group.nameList[0]) || '';
  const rates = new Set(payments.map((p) => p.gstRate));
  const displayRate = rates.size === 1 ? [...rates][0] : defaultRate;

  return {
    flatNo: String(flatNo),
    floorNum,
    floorLabel: override.floor || floorLabel(floorNum),
    flatType: override.flatType || '',
    carpetArea: override.carpetArea != null ? override.carpetArea : '',
    regNo: override.regNo || '',
    agreementDate: '',
    name,
    allottees: group.nameList || [],
    agreementValue: null,
    gstRate: displayRate,
    gstOnAgreement: null,
    agreementValueWithGst: null,
    payments,
    totalReceived,
    totalGst,
    balance: null,
    matchReason: 'tally-only',
    matchScore: 0,
    valueSource: 'tally',
    valueCandidates: [],
    hasAgreement: false,
    status: TALLY_ONLY_STATUS,
  };
}

function buildTallyOnlyRecords(tally, agreementFlatNos = [], overrides = {}, opts = {}) {
  const matched = new Set(Array.from(agreementFlatNos, (flatNo) => String(flatNo)));
  const out = [];

  for (const [flatNo, group] of tally.byFlat) {
    if (matched.has(String(flatNo))) continue;
    const override = overrides[flatNo] || overrides[String(flatNo)] || {};
    out.push(buildTallyOnlyRecord(group, override, opts));
  }

  return out;
}

module.exports = {
  buildTallyOnlyRecord,
  buildTallyOnlyRecords,
  DEFAULT_TALLY_ONLY_GST_RATE,
  TALLY_ONLY_STATUS,
};
