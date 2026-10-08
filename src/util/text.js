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

module.exports = {
  parseAmount,
  nameTokens,
  nameSimilarity,
  floorLabel,
  floorFromFlat,
};
