'use strict';

const path = require('path');
const { extractPageImages } = require('./extractPages');
const { ocrImage } = require('./ocr');
const { parseAmount } = require('../util/text');
const { parseDate, formatMDY } = require('../util/dates');

const VALUE_PAGE = 1; // registration receipt (मोबदला / consideration)
const NAME_PAGES = [6, 7]; // parties page — check both (varies by agreement)

function getHeadBeforePayments(text) {
  const lines = text.split(/\r?\n/);
  let cutoff = lines.length;
  for (let i = 0; i < lines.length; i += 1) {
    if (/^\s*1\s*[)\].:]/.test(lines[i])) { cutoff = i; break; }
  }
  return lines.slice(0, cutoff).join('\n');
}

/** Parse Indian-style grouped numbers: 1,23,03,000 or 12303000 */
function parseIndianAmount(raw) {
  if (raw == null) return null;
  const digits = String(raw).replace(/[^\d]/g, '');
  if (!digits) return null;
  let n = Number(digits);
  if (!Number.isFinite(n)) return null;
  // OCR may append stray trailing digits (e.g. 123030008 -> 12303000)
  if (digits.length >= 9) {
    for (let trim = 1; trim <= 2; trim += 1) {
      const t = Number(digits.slice(0, -trim));
      if (t >= 500000 && t <= 500000000 && t % 1000 === 0) return t;
    }
  }
  return n;
}

function collectSlashAmounts(head) {
  const amounts = [];
  const re = /([\d][\d,.\s]{2,}?)\s*\/\s*-+/g;
  let m;
  while ((m = re.exec(head)) !== null) {
    const v = parseIndianAmount(m[1]);
    if (v != null && v >= 1000) amounts.push(v);
  }
  return amounts;
}

function collectLargeAmounts(head) {
  const amounts = [];
  const re = /(?:^|[^\d])((?:\d{1,2}(?:,\d{2}){0,3},\d{3}|\d{7,9}))(?:[^\d]|$)/g;
  let m;
  while ((m = re.exec(head)) !== null) {
    const v = parseIndianAmount(m[1]);
    if (v != null && v >= 100000 && v <= 500000000) amounts.push(v);
  }
  return amounts;
}

/**
 * From the registration-receipt OCR text, pull the consideration value.
 * Primary: 2nd "amount /-" before the payment block (market, consideration, stamp duty).
 * Fallback: ordered large amounts in the header when OCR garbles the "/-" markers.
 */
function extractAgreementValue(text) {
  const head = getHeadBeforePayments(text);

  const slashAmounts = collectSlashAmounts(head);
  let candidates = slashAmounts.slice();
  let value = null;

  if (slashAmounts.length >= 2) {
    value = slashAmounts[1];
  } else if (slashAmounts.length === 1) {
    value = slashAmounts[0];
  }

  if (value == null) {
    const large = collectLargeAmounts(head);
    candidates = candidates.concat(large);
    if (large.length >= 3) {
      // receipt order: market value, consideration, stamp duty
      value = large[1];
    } else if (large.length === 2) {
      value = Math.max(large[0], large[1]);
    } else if (large.length === 1) {
      value = large[0];
    }
  }

  return { value, candidates: [...new Set(candidates)] };
}

/** Pull the registration / agreement date (dd/mm/yyyy on the receipt). */
function extractAgreementDate(text) {
  const m = text.match(/(\d{1,2})\s*\/\s*(\d{1,2})\s*\/\s*(\d{4})/);
  if (!m) return null;
  return parseDate(`${m[1]}/${m[2]}/${m[3]}`, { dayFirst: true });
}

/**
 * Extract allottee names from the parties page.
 * Supports:
 *   (1) MR. NAME, aged about ...
 *   (1) MR. NAME, (PAN ...), aged about ...
 */
function extractNames(text) {
  const names = [];
  const normalized = text.replace(/\s+/g, ' ');

  const re = /\(\s*(\d)\s*\)\s*((?:MR|MRS|MS|MISS|SMT|SHRI|SRI|DR|KUM)\.?\s+[A-Z][A-Z\s.]+?)(?=,\s*(?:\(PAN|aged\s+about))/gi;
  let m;
  while ((m = re.exec(normalized)) !== null) {
    const cleaned = m[2].replace(/\s+/g, ' ').trim();
    if (cleaned && !names.includes(cleaned)) names.push(cleaned);
  }

  // fallback: older pattern without PAN block
  if (!names.length) {
    const re2 = /\(\s*(\d)\s*\)\s*((?:MR|MRS|MS|MISS|SMT|SHRI|SRI|DR|KUM)?\.?\s*[A-Z][A-Z\s.]+?)\s*,\s*aged\s+about/gi;
    while ((m = re2.exec(normalized)) !== null) {
      const cleaned = m[2].replace(/\s+/g, ' ').trim();
      if (cleaned && !names.includes(cleaned)) names.push(cleaned);
    }
  }

  return names;
}

function flatNoFromFilename(base) {
  // Strip multer-generated prefix only when it looks like a timestamp:
  // e.g. "1749647476296-6958-FLAT NO.904.pdf" → timestamp has ≥10 digits
  // Do NOT strip "1804-CHHAYA KADAM.pdf" (1804 is the actual flat number)
  let stripped = base;
  if (/^\d{8,}-/.test(base)) {
    // Remove all leading numeric-hyphen segments (the timestamp + random suffix)
    stripped = base.replace(/^(\d{6,}-)+/i, '');
  }

  // Priority 1: explicit "flat no. NNN" or "flat NNN" or "shop no. NNN" label
  const explicit = stripped.match(/(?:flat\s*no\.?\s*|shop\s*no\.?\s*)(\d{3,4})/i);
  if (explicit) return explicit[1];

  // Priority 2: last 3-4 digit group in the stripped name (real flat numbers are usually last)
  const allGroups = [...stripped.matchAll(/(\d{3,4})/g)];
  if (allGroups.length) return allGroups[allGroups.length - 1][1];

  return null;
}

/**
 * OCR pages 6 and 7 for allottee names; prefer the page that yields more names.
 */
async function extractNamesFromPages(byPage, pages = NAME_PAGES) {
  let best = [];
  for (const page of pages) {
    if (!byPage[page]) continue;
    const text = await ocrImage(byPage[page]);
    const found = extractNames(text);
    if (found.length > best.length) best = found;
  }
  return best;
}

/**
 * Process one agreement PDF.
 */
async function processAgreement(pdfPath, opts = {}) {
  const base = path.basename(pdfPath);
  const flatNoHint = flatNoFromFilename(base);
  const namePages = opts.namePages || NAME_PAGES;

  const workDir = opts.workDir || path.join(path.dirname(pdfPath), '_pages');
  const pagesToExtract = [VALUE_PAGE, ...namePages];
  const imgs = extractPageImages(pdfPath, {
    pages: [...new Set(pagesToExtract)],
    outDir: workDir,
  });
  const byPage = Object.fromEntries(imgs.map((i) => [i.page, i.file]));

  let agreementValue = null;
  let valueCandidates = [];
  let agreementDate = null;
  if (byPage[VALUE_PAGE]) {
    const text = await ocrImage(byPage[VALUE_PAGE]);
    const r = extractAgreementValue(text);
    agreementValue = r.value;
    valueCandidates = r.candidates;
    agreementDate = extractAgreementDate(text);
  }

  const names = await extractNamesFromPages(byPage, namePages);

  return {
    pdf: base,
    flatNoHint,
    agreementValue,
    valueCandidates,
    agreementDate,
    agreementDateText: agreementDate ? formatMDY(agreementDate) : '',
    names,
    primaryName: names.length ? names.join(' / ') : null,
    namePagesUsed: namePages,
  };
}

module.exports = {
  processAgreement,
  extractAgreementValue,
  extractNames,
  extractAgreementDate,
  flatNoFromFilename,
  NAME_PAGES,
  VALUE_PAGE,
};
