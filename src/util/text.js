'use strict';

const TITLES = new Set([
  'mr', 'mrs', 'ms', 'miss', 'smt', 'shri', 'sri', 'm/s', 'dr', 'kum',
]);

/** Parse a numeric string that may contain commas, currency symbols, spaces. */
function parseAmount(raw) {
  if (raw == null) return null;
  const cleaned = String(raw).replace(/[^\d.]/g, '');
  if (!cleaned) return null;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : null;
}

/** Normalize a person name into comparable lowercase tokens (titles removed). */
function nameTokens(name) {
  if (!name) return [];
  return String(name)
    .toLowerCase()
    .replace(/[^a-z\s]/g, ' ')
    .split(/\s+/)
    .map((t) => t.replace(/\./g, '').trim())
    .filter((t) => t && !TITLES.has(t));
}

/** Score similarity between two names (0..1) by shared tokens + first/last bonus. */
function nameSimilarity(a, b) {
  const ta = nameTokens(a);
  const tb = nameTokens(b);
  if (!ta.length || !tb.length) return 0;
  const setB = new Set(tb);
  // single-letter tokens (initials) match by first character of any token in b
  let shared = 0;
  for (const t of ta) {
    if (setB.has(t)) shared += 1;
    else if (t.length === 1 && tb.some((x) => x[0] === t)) shared += 0.5;
    else if (tb.some((x) => x.length === 1 && x[0] === t[0])) shared += 0.5;
  }
  const denom = Math.max(ta.length, tb.length);
  let score = shared / denom;
  // bonus if surnames (last token) match
  if (ta[ta.length - 1] === tb[tb.length - 1]) score += 0.15;
  return Math.min(1, score);
}

/** Ordinal floor label from a numeric floor (0 -> "GR", 1 -> "1st Floor"...). */
function floorLabel(floorNum) {
  if (floorNum == null) return '';
  if (floorNum === 0) return 'GR';
  const n = floorNum;
  const s = ['th', 'st', 'nd', 'rd'];
  const v = n % 100;
  const suffix = s[(v - 20) % 10] || s[v] || s[0];
  return `${n}${suffix} Floor`;
}

/**
 * Derive floor + unit from a flat number.
 * 3-digit (e.g. 105) -> floor 1; 4-digit (e.g. 1804) -> floor 18.
 */
function floorFromFlat(flatNo) {
  const digits = String(flatNo).replace(/\D/g, '');
  if (digits.length <= 2) return null;
  const floorPart = digits.slice(0, digits.length - 2);
  const n = parseInt(floorPart, 10);
  return Number.isFinite(n) ? n : null;
}

const DEVANAGARI_DIGITS = '०१२३४५६७८९';

/** Replace Devanagari digits ०-९ with 0-9. */
function normalizeDigits(str) {
  return String(str == null ? '' : str).replace(/[०-९]/g, (d) => String(DEVANAGARI_DIGITS.indexOf(d)));
}

const HAS_DEVANAGARI = /[\u0900-\u097F]/;

const DEV_CONSONANTS = {
  क: 'k', ख: 'kh', ग: 'g', घ: 'gh', च: 'ch', छ: 'chh', ज: 'j', झ: 'jh', ट: 't', ठ: 'th',
  ड: 'd', ढ: 'dh', ण: 'n', त: 't', थ: 'th', द: 'd', ध: 'dh', न: 'n', प: 'p', फ: 'f',
  ब: 'b', भ: 'bh', म: 'm', य: 'y', र: 'r', ल: 'l', व: 'v', श: 'sh', ष: 'sh', स: 's',
  ह: 'h', ळ: 'l',
};
const DEV_VOWELS = {
  अ: 'a', आ: 'a', इ: 'i', ई: 'i', उ: 'u', ऊ: 'u', ए: 'e', ऐ: 'ai', ओ: 'o', औ: 'au', ऑ: 'o',
};
const DEV_MATRAS = {
  'ा': 'a', 'ि': 'i', 'ी': 'i', 'ु': 'u', 'ू': 'u', 'े': 'e', 'ै': 'ai', 'ो': 'o', 'ौ': 'au', 'ॉ': 'o',
};

/** Rough Devanagari -> Latin transliteration (enough for fuzzy name matching). */
function transliterateDevanagari(str) {
  // ज्ञ is "dny" in Marathi names (प्रज्ञा = Pradnya), not "jn"
  const chars = Array.from(String(str || '').replace(/ज्ञ/g, 'dny'));
  let out = '';
  for (let i = 0; i < chars.length; i += 1) {
    const c = chars[i];
    if (DEV_CONSONANTS[c]) {
      out += DEV_CONSONANTS[c];
      const next = chars[i + 1];
      if (next === '्') i += 1; // virama: no inherent vowel
      else if (!DEV_MATRAS[next]) out += 'a';
    } else if (DEV_MATRAS[c]) out += DEV_MATRAS[c];
    else if (DEV_VOWELS[c]) out += DEV_VOWELS[c];
    else if (c === 'ं' || c === 'ँ') out += 'n';
    else if (/\s/.test(c)) out += ' ';
    else if (/[A-Za-z]/.test(c)) out += c;
  }
  return out;
}

function levenshtein(a, b) {
  const dp = Array.from({ length: a.length + 1 }, (_, i) => [i]);
  for (let j = 1; j <= b.length; j += 1) dp[0][j] = j;
  for (let i = 1; i <= a.length; i += 1) {
    for (let j = 1; j <= b.length; j += 1) {
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
  }
  return dp[a.length][b.length];
}

/** Consonant skeleton of a Latin token ("sinha" / "singh" -> "sng"). */
function skeleton(token) {
  const t = token
    .replace(/nh/g, 'ng')
    .replace(/([bcdgjkpstv])h/g, '$1')
    .replace(/(.)\1+/g, '$1');
  return t.replace(/[aeiou]/g, '') || t;
}

/**
 * Fuzzy name similarity (0..1) that also works when one name is Devanagari
 * (transliterated) and the other Latin. Each token of the shorter name is
 * matched to its closest token in the other by consonant skeleton.
 */
function fuzzyNameSimilarity(a, b) {
  const prep = (n) => {
    const latin = HAS_DEVANAGARI.test(String(n || '')) ? transliterateDevanagari(n) : String(n || '');
    return nameTokens(latin).filter((t) => t.length > 1).map(skeleton).filter(Boolean);
  };
  const ta = prep(a);
  const tb = prep(b);
  if (!ta.length || !tb.length) return 0;
  const [small, big] = ta.length <= tb.length ? [ta, tb] : [tb, ta];
  let total = 0;
  for (const t of small) {
    let best = 0;
    for (const u of big) {
      let r = 1 - levenshtein(t, u) / Math.max(t.length, u.length);
      // Marathi surnames carry extra suffixes (Prabhukar / Prabhukarvatkar): a long
      // token that is the start of the other still counts as a close match.
      if (r < 0.8 && Math.min(t.length, u.length) >= 4 && (u.startsWith(t) || t.startsWith(u))) r = 0.8;
      if (r > best) best = r;
    }
    total += best >= 0.7 ? best : 0;
  }
  return total / small.length;
}

module.exports = {
  normalizeDigits,
  transliterateDevanagari,
  fuzzyNameSimilarity,
  HAS_DEVANAGARI,
  parseAmount,
  nameTokens,
  nameSimilarity,
  floorLabel,
  floorFromFlat,
};
