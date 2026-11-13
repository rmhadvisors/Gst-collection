'use strict';

/**
 * Extra Index-II (सूची क्र.2) fields read from one OCR variant of a page:
 * document type, market value, building, shop number, registration number and
 * every buyer. Each parser returns null when it is not sure - the caller votes
 * across OCR variants and anything still null is shown as "check manually".
 */

const { normalizeDigits } = require('../util/text');

const UNIT_WORDS = 'सदनिका|दुकान|गाळा|शॉप|शाप';

/** पर्यायी जागेचा करार (alternate accommodation) vs an ordinary agreement. OCR often eats the first letters. */
function parseDocType(text) {
  return /यायी\s*जा[ऀ-ॿ]{0,3}/.test(text) ? 'alternate' : 'agreement';
}

const toNumber = (s) => Number(String(s).replace(/[^\d]/g, ''));

/**
 * The two bare numbers after the document type in the header block:
 * (2) मोबदला (consideration) then (3) बाजारभाव (market value).
 * Used when the labels themselves are unreadable.
 */
function headerNumbers(text) {
  const t = normalizeDigits(text);
  const at = t.search(/यायी\s*जा[ऀ-ॿ]{0,3}|करारनामा|दानपत्र|बक्षीस/);
  if (at < 0) return [];
  const nums = [];
  for (const line of t.slice(at, at + 500).split('\n').slice(1)) {
    const s = line.trim();
    if (!s) continue;
    if (/^[^\p{L}\d]{0,3}\d[\d,]*[^\p{L}\d]{0,3}$/u.test(s)) nums.push(toNumber(s));
    else if (/पालिके|^\d\)/.test(s) || nums.length) break;
    if (nums.length === 2) break;
  }
  return nums;
}

/** Market value (बाजारभाव). Labelled line first, else the second header number. */
function parseMarketValue(text) {
  const t = normalizeDigits(text);
  // "(3) बाजारभाव(भाडेपट्ट्याच्या ...) 562500", "बाजार मुल्य: रु.3649500"
  for (const line of t.split('\n')) {
    if (!/ाजार\s*(?:भ|मु|मू)/.test(line) || /प्रमाणे|शुल्क|मुद्रांक|नोंदणी/.test(line)) continue;
    // drop the row number ("(3) ") first, or it is read as the value; a real market value is never under 1000
    const body = line.replace(/^[^\p{L}\d]*\(?\d{1,2}\)\s*/u, '');
    const m = body.match(/(?:रु\.?\s*)?(\d[\d,]{0,11})\s*(?:\.?\s*\/-+)?[^\d\n]*$/);
    if (m && toNumber(m[1]) >= 1000) return toNumber(m[1]);
  }
  const nums = headerNumbers(t);
  return nums.length === 2 ? nums[1] : null;
}

/** Consideration printed as 0 on an alternate-accommodation agreement. */
function parseConsiderationZero(text) {
  const t = normalizeDigits(text);
  const m = t.match(/मो[ऀ-ॿA-Za-z]{0,4}द[ऀ-ॿ]{0,3}\s*\S{0,3}\s*(\d{1,9})(?!\d)/);
  if (m) return Number(m[1]) === 0;
  const nums = headerNumbers(t);
  return nums.length === 2 ? nums[0] === 0 : null;
}

/** Building name: "... इमारत मैत्री हाईट्स मधील सदनिका क्र. 602". */
function parseBuilding(text) {
  const m = text.match(new RegExp(`इमारत\\s+([\\s\\S]{3,90}?)\\s*म[ऀ-ॿ]{0,4}\\s*ल\\s*(?:${UNIT_WORDS})`));
  if (!m) return null;
  const name = m[1].replace(/\s+/g, ' ').replace(/[,.:;-]+$/, '').trim();
  return name.length >= 3 ? name : null;
}

/** Shop / gala number: "मधील शॉप क्र. 1" -> "1" (null when the page is about a flat). */
function parseShop(text) {
  if (/सदनिका\s*(?:क[्]?र?|\S{1,3})\s*[.,:]?\s*\d{3,4}/.test(text)) return null;
  // only the property description names the unit: "... इमारत X मधील शॉप क्र. 1"; a seller's address ("रोड नं: शॉप नं.3") does not
  const m = normalizeDigits(text).match(/म[ऀ-ॿ]{0,4}\s*ल\s*(?:दुकान|गाळा|शॉप|शाप)\s*(?:क्र\S{0,3}|\S{1,3})?\s*[.,:]?\s*(\d{1,3}[A-Za-z]?)(?![\dA-Za-z])/);
  return m ? m[1].toUpperCase() : null;
}

/** Commercial unit (shop / gala), not a residential flat. */
function isCommercial(text) {
  return new RegExp(`(?:दुकान|गाळा|शॉप|शाप)\\s*(?:क्र\\S{0,3}|\\S{1,3})?\\s*[.,:]?\\s*\\d`).test(normalizeDigits(text))
    && !/सदनिका\s*(?:क[्]?र?|\S{1,3})\s*[.,:]?\s*\d{3,4}/.test(text);
}

/** Registration number "7064/2026" (दस्त क्रमांक / अनुक्रमांक / वसई6-4611-2024). */
function parseRegNo(text) {
  const t = normalizeDigits(text);
  let m = t.match(/दस्त\s*क्रमांक\s*[:ः]?\s*(\d{1,6})\s*\/\s*(20\d{2})/);
  if (!m) m = t.match(/अनुक्रमांक[^\n]{0,60}?(\d{2,6})\s*[-/]\s*(20\d{2})(?!\d)/);
  // column layouts print the value far from its label: a bare "4611/2024" (not part of a date)
  if (!m) m = t.match(/(?<![\d/])(\d{3,6})\s*\/\s*(20\d{2})(?!\d)/);
  return m ? `${Number(m[1])}/${m[2]}` : null;
}

module.exports = {
  parseDocType, parseMarketValue, parseConsiderationZero, parseBuilding, parseShop, isCommercial, parseRegNo, headerNumbers,
};
