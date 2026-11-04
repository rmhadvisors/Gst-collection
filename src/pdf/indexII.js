'use strict';

/**
 * Parser for Marathi Index-II (सूची क्र.2) registration pages.
 *
 * Tesseract output for these tables is noisy and differs per page-segmentation
 * mode, so every field is parsed from each OCR variant of the page and the
 * most common answer wins (ties go to the earlier variant).
 */

const { normalizeDigits } = require('../util/text');
const { parseDate } = require('../util/dates');

const WING_MAP = { बी: 'B', वी: 'B', ए: 'A', सी: 'C', डी: 'D' };

/** Does the OCR text look like an Index-II page at all? */
function looksLikeIndexII(text) {
  // Not मोबदला alone: the registration *receipt* (pavti) also carries it and is
  // handled by the legacy receipt parser.
  return /सूची\s*क्र|Index\s*I[IlI]|सदनिका|दस्त\s*क्रमांक/.test(text);
}

function vote(values) {
  const counts = new Map();
  for (const v of values) {
    if (v == null || v === '') continue;
    counts.set(v, (counts.get(v) || 0) + 1);
  }
  let best = null;
  let bestN = 0;
  for (const [v, n] of counts) {
    if (n > bestN) { best = v; bestN = n; }
  }
  return best;
}

/** First plausible money amount (5-9 digits, >= 1 lakh) within `span` chars after `from`. */
function amountAfter(text, from, span = 70) {
  const tokens = text.slice(from, from + span).match(/\d[\d,]*\d/g) || [];
  for (const token of tokens) {
    const digits = token.replace(/\D/g, '');
    const n = Number(digits);
    if (digits.length >= 5 && digits.length <= 9 && n >= 100000) return n;
  }
  return null;
}

/** Consideration: the number following the "मोबदला" label (same line or next). */
function parseValue(text) {
  const re = /(?:मो|शि|मे)\S{0,3}दला/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const v = amountAfter(text, m.index + m[0].length);
    if (v != null) return v;
  }
  return null;
}

/**
 * Positional fallback when the "मोबदला" label is unreadable: the header lists
 * consideration then market value as bare numbers ahead of the property
 * description. Returns the first bare 6-8 digit number in that region.
 */
function parseValuePositional(text) {
  const stop = text.search(/सदनिका|भू-?\s*मापन|पालिके/);
  const head = stop > 0 ? text.slice(0, stop) : text;
  const m = head.match(/^\D{0,8}(\d{6,8})\s*$/m);
  return m ? Number(m[1]) : null;
}

/** Flat number after "सदनिका क्र."; OCR often mangles "क्र", so allow a short junk token. */
function parseFlat(text) {
  const m = text.match(/सदनिका\s*(?:क[्]?र?|\S{1,3})\s*[.,:]?\s*(\d{3,4})(?!\d)(?!\s*(?:\.\d+)?\s*चौ)/);
  return m ? m[1] : null;
}

function parseWing(text, flatIdx) {
  const scope = flatIdx != null ? text.slice(flatIdx, flatIdx + 120) : text;
  let m = scope.match(/(बी|वी|ए|सी|डी)\s*विं?ग/);
  if (m) return WING_MAP[m[1]];
  m = scope.match(/\b([A-D])\s*[- ]?\s*wing\b/i) || scope.match(/\bwing\s*[- ]?\s*([A-D])\b/i);
  return m ? m[1].toUpperCase() : null;
}

function cleanName(raw) {
  // cut at the next field ("वय:", "पॅन", "पत्ता", "तर्फे"), but keep a Marathi
  // parenthetical such as "(एच यू एफ)" (HUF) that belongs to the purchaser's name
  const head = raw.split(/\s+[-–]?\s*(?:वय|पॅन|पत्ता|तर्फे)|,/)[0];
  const paren = head.match(/^([^(]*?)\s*(\([ऀ-ॿ\s]{1,16}\))/);
  const base = paren ? `${paren[1]} ${paren[2]}` : head.split('(')[0];
  const name = base
    .replace(/[|[\]_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/[\s\-–.:]+$/, '')
    .trim();
  return name.length >= 3 ? name : null;
}

const NAME_FIELD = /नाव\s*(?:[:ः;]\s*[-–]+|[-–]+)\s*([^\n]+)/g;

// the seller / builder side of an Index-II (never the purchaser)
const SELLER = /^(?:मेसर्स|मेः|मे\.|मा\s*देणार)|डेव्हलपर्स|बिल्डर|ब्लॉक|रोड|प्लॉट|पत्ता/;

/** Purchaser: first "नाव:-" after the "दस्तऐवज करुन घेणा-या" label (field 6/8). */
function parsePurchaser(text) {
  const label = /दस्त\S*\s*क[ं]?रुन\s*घे\S{0,4}णा/g;
  let m;
  let last = null;
  while ((m = label.exec(text)) !== null) last = m;
  if (last) {
    const from = last.index + last[0].length;
    // the first party after the label that is not the seller / builder
    const window = text.slice(from, from + 1200);
    for (const nm of window.matchAll(NAME_FIELD)) {
      const name = cleanName(nm[1]);
      if (name && !SELLER.test(name)) return name;
    }
  }
  // Label unreadable: the purchaser is the last "नाव:-" that is not the
  // builder / seller ("मेसर्स ...", "मा देणार ...").
  let found = null;
  const all = /नाव\s*(?:[:ः;]\s*[-–]+|[-–]+)\s*([^\n]+)/g; // "नाव:-" / "नाव-" only (not "पालिकेचे नाव:...")
  while ((m = all.exec(text)) !== null) {
    const name = cleanName(m[1]);
    if (name && name.includes(' ') && !SELLER.test(name)) found = name;
  }
  return found;
}

/** Registration date: the date following "दस्त नोंदणी केल्याचा दिनांक" (dd/mm/yyyy text). */
function parseDateField(text) {
  const loose = /(\d{1,2})\s*[-/.]?\s*(\d{2})\s*[-/.]?\s*(20\d{2})/; // "27-082026" is OCR-merged
  const strict = /\b(\d{1,2})[-/.](\d{2})[-/.](20\d{2})\b/;
  const label = /नों?दणी\s*केल्याचा\s*दिनांक/.exec(text);
  let m = label ? text.slice(label.index, label.index + 80).match(loose) : null;
  if (!m) m = text.match(strict);
  if (!m) return null;
  const text2 = `${m[1].padStart(2, '0')}/${m[2]}/${m[3]}`;
  return parseDate(text2, { dayFirst: true }) ? text2 : null;
}

/** Carpet area in sq.m: the first number before "चौ.मी." (the flat; balconies follow). */
function parseAreaSqm(text) {
  const m = text.match(/(\d{2,3}(?:\.\d{1,2})?)\s*चौ\s*[.,]?\s*मी/);
  return m ? Number(m[1]) : null;
}

/** Area in sq.ft: the number before "चौ.फूट". */
function parseArea(text) {
  const m = text.match(/(\d{2,4}(?:\.\d+)?)\s*चौ\s*[.,]?\s*फू?ट/);
  return m ? Number(m[1]) : null;
}

/**
 * Parse one page from its OCR variants.
 * @param {string[]} variants OCR text under different PSM modes (priority order)
 */
function parseIndexIIPage(variants) {
  const texts = variants.map((t) => normalizeDigits(t));
  const wings = texts.map((t) => {
    const idx = t.search(/सदनिका\s*\S/);
    return parseWing(t, idx >= 0 ? idx : null);
  });

  let values = texts.map(parseValue).filter((v) => v != null);
  let valueMethod = 'label';
  if (!values.length) {
    // positional fallback must be corroborated by at least two OCR variants
    const pos = texts.map(parseValuePositional).filter((v) => v != null);
    const top = vote(pos);
    if (top != null && pos.filter((v) => v === top).length >= 2) {
      values = [top];
      valueMethod = 'position';
    }
  }
  const dateText = vote(texts.map(parseDateField));

  return {
    isIndexII: texts.some(looksLikeIndexII),
    flat: vote(texts.map(parseFlat)),
    wing: vote(wings),
    value: vote(values),
    valueMethod,
    valueCandidates: [...new Set(values)],
    purchaser: vote(texts.map(parsePurchaser)),
    date: dateText ? parseDate(dateText, { dayFirst: true }) : null,
    area: vote(texts.map(parseArea)),
    areaSqm: vote(texts.map(parseAreaSqm)),
  };
}

module.exports = { parseIndexIIPage, parseAreaSqm, amountAfter, looksLikeIndexII, parseValue, parseFlat, parseWing, parsePurchaser };
