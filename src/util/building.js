'use strict';

/**
 * Building names. A building is named in an Index-II page ("इमारत मैत्री हाईट्स
 * मधील सदनिका क्र. 403") and sometimes in the title of a Tally book ("H D F C
 * BANK 44444 Siddhivinayak Palace Book"). Both are reduced to one canonical English
 * name so that Maitri Heights 403 and Siddhivinayak Palace 403 never merge.
 *
 * config/overrides.json can give the canonical names and their spellings:
 *   "buildings": { "Maitri Heights": ["मैत्री हाईट्स", "maitri"],
 *                  "Siddhivinayak Palace": ["सिद्धिविनायक पॅलेस", "siddhivinayak"] }
 */

const { fuzzyNameSimilarity, HAS_DEVANAGARI, nameTokens } = require('./text');
const { transliterateName } = require('./translit');

// words that say "this is a building" but not WHICH building
const GENERIC = new Set([
  'palace', 'pales', 'heights', 'hights', 'haits', 'hits', 'tower', 'towers', 'building', 'apartment', 'apartments',
  'residency', 'society', 'chs', 'shri', 'shree', 'sri', 'bank', 'book', 'the', 'and', 'ltd', 'limited',
  'श्री', 'पॅलेस', 'पैलेस', 'पेलेस', 'हाईट्स', 'हाइट्स', 'हाईटस', 'टॉवर', 'बिल्डिंग', 'अपार्टमेंट', 'सोसायटी',
]);

const MATCH_MIN = 0.8;

/** Tokens that identify the building (generic words removed; all tokens if nothing is left). */
function distinctiveTokens(name) {
  const raw = String(name || '').split(/[\s,.\-_/()]+/).filter((t) => t.length > 1);
  const kept = raw.filter((t) => !GENERIC.has(t.toLowerCase()));
  return kept.length ? kept : raw;
}

/** Similarity (0..1) of two building names on their distinctive tokens only. */
function buildingSimilarity(a, b) {
  const ta = distinctiveTokens(a).join(' ');
  const tb = distinctiveTokens(b).join(' ');
  if (!ta || !tb) return 0;
  return fuzzyNameSimilarity(ta, tb);
}

/** Title-cased English for a name that has no configured spelling. */
function englishName(raw) {
  const latin = HAS_DEVANAGARI.test(raw) ? transliterateName(raw) : String(raw || '');
  return latin.split(/\s+/).filter(Boolean).map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
}

/**
 * Canonical building name for text read from an agreement. A configured building
 * whose name / alias matches wins; otherwise the transliterated text is used.
 * @param {string} raw        e.g. "मैत्री हाईट्स"
 * @param {object} buildings  overrides.buildings
 */
function canonicalBuilding(raw, buildings = {}) {
  const text = String(raw || '').trim();
  if (!text) return '';
  let best = null;
  for (const [name, aliases] of Object.entries(buildings)) {
    if (name.startsWith('_')) continue;
    for (const spelling of [name, ...(Array.isArray(aliases) ? aliases : [])]) {
      const score = buildingSimilarity(text, spelling);
      if (!best || score > best.score) best = { name, score };
    }
  }
  return best && best.score >= MATCH_MIN ? best.name : englishName(text);
}

/**
 * The known building a free text (a Tally book title) names, or ''. Only tokens
 * of 5+ letters are compared, so a bank number or "HDFC" never matches.
 */
function buildingInText(text, names, buildings = {}) {
  const tokens = nameTokens(String(text || '')).filter((t) => t.length >= 5);
  if (!tokens.length) return '';
  let best = null;
  for (const name of names) {
    const spellings = [name, ...(Array.isArray(buildings[name]) ? buildings[name] : [])];
    for (const spelling of spellings) {
      const key = distinctiveTokens(spelling).filter((t) => t.length >= 5);
      for (const k of key) {
        for (const t of tokens) {
          const score = fuzzyNameSimilarity(k, t);
          if (score >= MATCH_MIN && (!best || score > best.score)) best = { name, score };
        }
      }
    }
  }
  return best ? best.name : '';
}

module.exports = {
  canonicalBuilding, buildingInText, buildingSimilarity, englishName, distinctiveTokens,
};
