'use strict';

/**
 * Canonical flat key: "B-603" when a wing letter is known, bare "1804" otherwise.
 * "B 607", "B-607", "B- 607", "b607" all normalise to "B-607".
 */
function makeFlatKey(wing, number) {
  const num = String(number || '').replace(/\D/g, '');
  if (!num) return '';
  const w = String(wing || '').trim().toUpperCase().slice(0, 1);
  return /[A-Z]/.test(w) ? `${w}-${num}` : num;
}

/** Normalise any "B 607" / "B-607" / "B- 607" / "607" string to a flat key (or ''). */
function normalizeFlatKey(raw) {
  const m = String(raw == null ? '' : raw).trim().match(/^([A-Za-z])?\s*[-–._]?\s*(\d{2,4})$/);
  return m ? makeFlatKey(m[1], m[2]) : '';
}

/** Wing letter of a flat key ('' when none). */
function wingOf(key) {
  const m = String(key || '').match(/^([A-Z])-/);
  return m ? m[1] : '';
}

/**
 * Flat key at the START of a Tally particulars cell:
 *   "B 607 -RAJESH SINGH", "B- 1401 Mithali ...", "A-1104 Tara ...", "Flat No.1804 Mrs. X"
 * Returns { key, rest } (rest = the buyer name) or null.
 */
function parseFlatFromParticulars(text) {
  const s = String(text || '').trim();
  // "Flat No.1804", "flat no 1804", "Shop 1", "Flat B-1002" (wing optional)
  let m = s.match(/^(?:flat|shop)\s*(?:no\.?)?\s*([A-Za-z])?\s*[-–.]?\s*(\d{1,4})\b/i);
  // "B 607", "B-1002", "A- 503"
  if (!m) m = s.match(/^([A-Za-z])\s*[-–._]?\s*(\d{3,4})\b/);
  // bare "403 NAME"
  if (!m) m = s.match(/^()(\d{3,4})\b/);
  if (!m) return null;
  const rest = s.slice(m[0].length).replace(/^[\s.:,-]+/, '').trim();
  return { key: makeFlatKey(m[1], m[2]), rest };
}

const CHQ_RETURN_RE = /\b(?:chq|cheque|check)\.?\s*(?:return(?:ed)?|bounce[d]?|ret)\b/i;
const AGMT_RE = /[-–,\s]*\b(?:agmt|agreement|agrmt)\.?\s*(?:dt|date)\.?\s*[:-]?\s*(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4})/i;

/**
 * Strip trailing notes from a buyer name and report what they said:
 *   "RAJESH SINGH - AGMT DT 11/12/2024" -> { name: "RAJESH SINGH", agmtDate: "11/12/2024" }
 *   "X FC", "X (Adv Agst Flat)", "X CHQ RETURN" -> name "X" (chqReturn for the last)
 */
function cleanPartyName(raw) {
  let s = String(raw || '').trim();
  let agmtDate = '';
  const a = s.match(AGMT_RE);
  if (a) { agmtDate = a[1]; s = s.replace(AGMT_RE, ' '); }
  const chqReturn = CHQ_RETURN_RE.test(s);
  s = s.replace(CHQ_RETURN_RE, ' ');
  s = s.replace(/\s*\([^)]*\)\s*$/g, ' ');
  s = s.replace(/[\s-–]+\bFC\b\.?\s*$/i, ' ').replace(/\s+\bFC\b\.?\s*$/i, ' ');
  s = s.replace(/[\s.,:;–-]+$/, '').replace(/\s+/g, ' ').trim();
  return { name: s, agmtDate, chqReturn };
}

function hasChqReturn(text) { return CHQ_RETURN_RE.test(String(text || '')); }

/**
 * Flat key from an uploaded agreement filename.
 * Strips the multer prefix (^\d+-\d+-), the extension and copy suffixes
 * (" (1)", "_1"), then reads wing+flat from "B-603", "B 603", "B603".
 * "1804-CHHAYA KADAM.pdf" -> "1804" (1804 is the flat, not a multer prefix).
 */
function flatKeyFromFilename(filename) {
  let s = String(filename || '').replace(/^.*[\/]/, '').replace(/\.[A-Za-z0-9]+$/, '');
  s = s.replace(/^\d+-\d+-/, '');
  s = s.replace(/\s*\(\d+\)\s*$/, '').replace(/[_\s-]+\d{1,2}$/, (m, off, str) => (
    /\d{3,4}$/.test(str.slice(0, off)) ? '' : m
  ));

  let m = s.match(/(?:flat|shop)\s*no\.?\s*([A-Za-z])?\s*[-–._]?\s*(\d{3,4})/i);
  if (m) return makeFlatKey(m[1], m[2]);

  m = s.match(/(?:^|[^A-Za-z])([A-Za-z])\s*[-–._ ]?\s*(\d{3,4})(?!\d)/);
  if (m) return makeFlatKey(m[1], m[2]);

  const groups = [...s.matchAll(/(\d{3,4})/g)];
  return groups.length ? groups[groups.length - 1][1] : null;
}

/**
 * overrides.json "renumber": {"<old flat>": "<new flat>"} -> Map of normalised
 * keys. Keys starting with "_" are comments. Applied in one pass, so a swap
 * ("403": "404", "404": "403") works.
 */
function renumberMap(raw) {
  const map = new Map();
  for (const [from, to] of Object.entries(raw || {})) {
    if (from.startsWith('_')) continue;
    const a = normalizeFlatKey(from); const b = normalizeFlatKey(to);
    if (a && b && a !== b) map.set(a, b);
  }
  return map;
}

/** Sort comparator for flat keys: by wing, then flat number. */
function compareFlatKeys(a, b) {
  const wa = wingOf(String(a)); const wb = wingOf(String(b));
  if (wa !== wb) return wa < wb ? -1 : 1;
  return Number(String(a).replace(/\D/g, '')) - Number(String(b).replace(/\D/g, ''));
}

module.exports = {
  makeFlatKey, normalizeFlatKey, renumberMap, wingOf, parseFlatFromParticulars, cleanPartyName, hasChqReturn,
  flatKeyFromFilename, compareFlatKeys,
};
