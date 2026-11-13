'use strict';

const { fuzzyNameSimilarity } = require('../util/text');
const { buildUnallocated, voucherFlags } = require('./unallocated');
const { formatMDY } = require('../util/dates');

const GST_RATE_NO_FLAT = 0.01; // same 1% advance rate as a flat
const PARTNER_NAME_MIN = 0.85;
const CHQ_RETURN = 'CHQ RETURN';

const FLAG_ASSIGN = 'no flat no. – assign flat to confirm';
const FLAG_PARTNER = 'partner – check if flat booking or capital';
const FLAG_CHQ = 'CHQ RETURN booked as receipt – no GST';

const round2 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
const partyKey = (s) => String(s || '').toLowerCase().replace(/[^\p{L}\p{N}]+/gu, ' ').trim();

/**
 * overrides.json "partnerGst": how GST is shown on receipts from the developer's partners.
 *   "blank" (default) - no GST, flagged "check if flat booking or capital"
 *   "1%"              - GST at 1% like any other party (the CA confirmed they are flat bookings)
 */
function partnerGstSetting(overrides = {}) {
  const v = String(overrides.partnerGst == null ? 'blank' : overrides.partnerGst).trim().toLowerCase();
  return v === '1%' || v === '1' || v === 'on' || v === 'gst' ? '1%' : 'blank';
}

/**
 * Receipts that carry no flat number, as pseudo-records (one per party) that the workbook
 * prints like a flat: "TOTAL ::" row, then every receipt. Nothing is dropped and nothing is
 * allocated to a flat (use overrides "allocate" for that). Largest party first.
 *
 * @param {object} tally
 * @param {object} [opts] { partners: [names], partnerGst: 'blank' | '1%', records: [] (for the name-guess hint) }
 */
function buildNoFlatRecords(tally, opts = {}) {
  const rows = tally.unallocated || [];
  if (!rows.length) return [];
  const hints = buildUnallocated(tally, opts.records || []); // same order as rows
  const partners = (opts.partners || []).filter(Boolean);
  const partnerMode = opts.partnerGst === '1%' ? '1%' : 'blank';
  const isPartner = (name) => partners.some((p) => fuzzyNameSimilarity(name, p) >= PARTNER_NAME_MIN);

  const groups = new Map();
  rows.forEach((row, i) => {
    const chq = !!row.chqReturn || !String(row.name || '').trim();
    const party = chq ? CHQ_RETURN : String(row.name).replace(/capital\s*a\/?c/i, '').trim() || row.name;
    const key = chq ? CHQ_RETURN : partyKey(party);
    if (!groups.has(key)) groups.set(key, { name: party, chq, rows: [], hints: new Set() });
    const g = groups.get(key);
    g.rows.push(row);
    if (hints[i] && hints[i].guess && /best flat guess/.test(hints[i].flag)) g.hints.add(`${hints[i].guess} (${hints[i].guessScore})`);
  });

  const records = [...groups.values()].map((g) => {
    const partner = !g.chq && isPartner(g.name);
    const noGst = g.chq || (partner && partnerMode === 'blank');
    const payments = g.rows
      .map((r) => ({
        date: r.date,
        dateText: r.date ? formatMDY(r.date) : '',
        amount: r.amount,
        gst: noGst ? null : round2(r.amount * GST_RATE_NO_FLAT),
        vchNo: r.vchNo,
        dupVoucher: !!r.dupVoucher,
        noVoucher: !r.vchNo,
        chqReturn: false, // said once on the party, not on every receipt
        source: 'tally',
      }))
      .sort((a, b) => (a.date ? a.date.getTime() : 0) - (b.date ? b.date.getTime() : 0));
    const totalReceived = round2(payments.reduce((s, p) => s + p.amount, 0));
    const flags = [];
    if (g.chq) flags.push(FLAG_CHQ);
    else if (partner) flags.push(partnerMode === 'blank' ? FLAG_PARTNER : `${FLAG_PARTNER} (GST shown at 1% by setting partnerGst)`);
    else flags.push(FLAG_ASSIGN);
    if (g.rows.some((r) => r.buildingReview)) flags.push('(review) some receipts are in a Tally book whose building is not known');
    for (const h of g.hints) flags.push(`(review) best flat guess by name: ${h} - NOT allocated`);
    flags.push(...voucherFlags(payments));
    return {
      noFlat: true,
      partner,
      flatNo: '',
      floorNum: null,
      floorLabel: '',
      flatType: '',
      carpetArea: '',
      regNo: '',
      marketValue: null,
      agreementDate: '',
      agreementDateObj: null,
      name: g.name,
      purchaserMarathi: '',
      allottees: [],
      agreementValue: null,
      gstRate: noGst ? null : GST_RATE_NO_FLAT,
      gstOnAgreement: null,
      agreementValueWithGst: null,
      payments,
      totalReceived,
      baseReceived: totalReceived,
      gstReceived: 0,
      gstOutstanding: null,
      osAmount: null,
      totalGst: round2(payments.reduce((s, p) => s + (p.gst || 0), 0)),
      balance: null,
      matchReason: 'no flat no.',
      matchScore: 0,
      valueSource: '',
      building: '',
      docType: 'none',
      needsReview: false,
      needsManual: false,
      ocrFailed: false,
      hasAgreement: false,
      flags,
    };
  });
  // largest party first (ties by name)
  return records.sort((a, b) => b.totalReceived - a.totalReceived || a.name.localeCompare(b.name));
}

module.exports = { buildNoFlatRecords, partnerGstSetting, CHQ_RETURN };
