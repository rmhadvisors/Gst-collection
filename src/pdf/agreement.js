'use strict';

const path = require('path');
const { extractPageImages } = require('./extractPages');
const { ocrImage, ocrImageVariants, MIN_OCR_CHARS } = require('./ocr');
const { renderPages, countPages } = require('./renderPages');
const { parseIndexIIPage, looksLikeIndexII, amountAfter } = require('./indexII');
const { retryAgreementDate } = require('./dateRetry');
const { flatKeyFromFilename, makeFlatKey, wingOf } = require('../util/flatKey');
const { parseAmount, normalizeDigits } = require('../util/text');
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
function extractAgreementValue(text, opts = {}) {
  // A labelled consideration ("मोबदला" / "Consideration") always beats
  // market value ("बाजार मुल्य" / "Market value"), whatever the amount order.
  const labelled = labelledConsideration(text);
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

  if (labelled != null) {
    return { value: labelled, candidates: [...new Set([labelled, ...candidates])] };
  }
  // English agreements: only a labelled "Consideration" is the value; other
  // amounts stay as candidates for the CA to look at but are never picked.
  if (opts.considerationOnly) return { value: null, candidates: [...new Set(candidates)] };
  return { value, candidates: [...new Set(candidates)] };
}

function labelledConsideration(text) {
  const re = /(?:मो|शि|मे)\S{0,3}दला|\bconsideration\b/gi;
  const norm = normalizeDigits(text);
  let m;
  while ((m = re.exec(norm)) !== null) {
    const v = amountAfter(norm, m.index + m[0].length);
    if (v != null) return v;
  }
  return null;
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

/** Flat key ("B-603" / "1804") from an uploaded filename; see util/flatKey. */
function flatNoFromFilename(base) {
  return flatKeyFromFilename(base);
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

const SQFT_TO_SQM = 0.09290304;
const MAX_FULL_SCAN_PAGES = 10; // longer PDFs: only the first page is Index-II

/**
 * Process one agreement PDF into one record per flat found.
 *
 * Every page is OCR'd (mar+eng) and parsed separately as a Marathi Index-II
 * page. Pages that yield different flats become separate records; the OCR flat
 * is preferred over the filename. Index-II pages with no readable value are
 * still returned (agreementValue null -> "manual needed"). PDFs with no
 * Index-II page fall back to the legacy English registration-receipt flow.
 *
 * @returns {Promise<object[]>} agreement objects (never empty)
 */
async function processAgreements(pdfPath, opts = {}) {
  const base = path.basename(pdfPath);
  const fileKey = flatKeyFromFilename(base);
  const fileWing = wingOf(fileKey);
  const namePages = opts.namePages || NAME_PAGES;

  const workDir = opts.workDir || path.join(path.dirname(pdfPath), '_pages');
  const imgs = extractPageImages(pdfPath, {
    pages: [...new Set([...Array.from({ length: MAX_FULL_SCAN_PAGES }, (_, n) => n + 1), ...namePages])],
    outDir: workDir,
  });
  const wanted = [...new Set([...Array.from({ length: MAX_FULL_SCAN_PAGES }, (_, n) => n + 1), ...namePages])];
  // Embedded image missing or unreadable -> render the page at 300 dpi instead.
  let totalPages = imgs.totalPages;
  let toRender = imgs.failedPages || [];
  if (!imgs.length) {
    totalPages = await countPages(pdfPath).catch(() => 0);
    toRender = wanted.filter((p) => p <= totalPages);
  }
  if (toRender.length) {
    try {
      imgs.push(...await renderPages(pdfPath, toRender, workDir));
    } catch (e) {
      console.error(`[agreement] ${base}: page render failed: ${e.message}`);
    }
  }
  const byPage = Object.fromEntries(imgs.map((i) => [i.page, i.file]));
  const scanPages = totalPages <= MAX_FULL_SCAN_PAGES
    ? [...new Set(imgs.map((i) => i.page))].sort((a, b) => a - b)
    : [VALUE_PAGE];

  // A page with almost no OCR text is reported, never silently left blank.
  const ocrFailedPages = scanPages.filter((p) => !byPage[p]);
  const ocrCache = {};
  const readPage = async (page) => {
    if (!(page in ocrCache)) {
      ocrCache[page] = await ocrImage(byPage[page]);
      if (ocrCache[page].trim().length < MIN_OCR_CHARS) ocrFailedPages.push(page);
    }
    return ocrCache[page];
  };
  const withOcrStatus = (list) => list.map((r) => ({
    ...r,
    ocrFailed: ocrFailedPages.length > 0,
    ocrFailedPages: [...new Set(ocrFailedPages)].sort((a, b) => a - b),
  }));

  // ── Marathi Index-II pages ──
  const parsed = [];
  for (const page of scanPages) {
    const first = await readPage(page);
    if (!looksLikeIndexII(first)) continue;
    const variants = [first, ...(await ocrImageVariants(byPage[page], ['4', '6', '11']))];
    const r = parseIndexIIPage(variants);
    r.dateRetried = false;
    if ((r.flat || r.value) && !r.date) {
      // rows 9/10 unreadable: deskew + upscale and retry just that region
      const retry = await retryAgreementDate(byPage[page]);
      r.dateRetried = true;
      if (retry) r.date = retry.date;
    }
    if (r.flat || r.value) parsed.push({ page, ...r });
  }

  if (parsed.length) {
    const usedKeys = new Set();
    const records = new Map(); // key -> record (merge pages that give the same flat)
    for (const pg of parsed) {
      let flatSource = 'ocr';
      let key;
      if (pg.flat) {
        key = makeFlatKey(pg.wing || fileWing, pg.flat);
      } else if (fileKey && !usedKeys.has(fileKey) && !records.has(fileKey)) {
        key = fileKey;
        flatSource = 'filename';
      } else {
        key = '';
        flatSource = 'unknown';
      }
      usedKeys.add(key);

      const rec = records.get(key || `?p${pg.page}`) || {
        pdf: base,
        flatNoHint: key || null,
        flatSource,
        agreementValue: null,
        valueCandidates: [],
        valueMethod: null,
        agreementDate: null,
        names: [],
        areaSqft: null,
        areaSqm: null,
        page: pg.page,
      };
      if (rec.agreementValue == null && pg.value != null) {
        rec.agreementValue = pg.value;
        rec.valueMethod = pg.valueMethod;
      }
      rec.valueCandidates = [...new Set([...rec.valueCandidates, ...pg.valueCandidates])];
      if (!rec.agreementDate && pg.date) {
        rec.agreementDate = pg.date;
        rec.dateSource = pg.dateRetried ? 'ocr-retry (deskew)' : 'ocr';
      }
      if (pg.purchaser && !rec.names.includes(pg.purchaser)) rec.names.push(pg.purchaser);
      if (rec.areaSqft == null && pg.area != null) rec.areaSqft = pg.area;
      if (rec.areaSqm == null && pg.areaSqm != null) rec.areaSqm = pg.areaSqm;
      records.set(key || `?p${pg.page}`, rec);
    }

    return withOcrStatus([...records.values()].map((rec) => ({
      ...rec,
      agreementDateText: rec.agreementDate ? formatMDY(rec.agreementDate) : '',
      // an area printed in sq.m is used as is; sq.ft is converted
      carpetAreaSqm: rec.areaSqm != null ? rec.areaSqm
        : (rec.areaSqft != null ? Math.round(rec.areaSqft * SQFT_TO_SQM * 100) / 100 : null),
      primaryName: rec.names.length ? rec.names.join(' / ') : null,
      namePagesUsed: [rec.page],
    })));
  }

  // ── Legacy English registration-receipt flow ──
  const legacyPages = [VALUE_PAGE, ...namePages].filter((p) => byPage[p]);
  let agreementValue = null;
  let valueCandidates = [];
  let agreementDate = null;
  if (byPage[VALUE_PAGE]) {
    const text = await readPage(VALUE_PAGE);
    const r = extractAgreementValue(text, { considerationOnly: true });
    agreementValue = r.value;
    valueCandidates = r.candidates;
    agreementDate = extractAgreementDate(text);
  }
  const names = await extractNamesFromPages(byPage, namePages);
  if (!byPage[VALUE_PAGE]) ocrFailedPages.push(VALUE_PAGE);

  return withOcrStatus([{
    pdf: base,
    flatNoHint: fileKey,
    flatSource: 'filename',
    agreementValue,
    valueCandidates,
    agreementDate,
    agreementDateText: agreementDate ? formatMDY(agreementDate) : '',
    names,
    primaryName: names.length ? names.join(' / ') : null,
    namePagesUsed: legacyPages,
  }]);
}

/** Back-compat wrapper: the first record of processAgreements(). */
async function processAgreement(pdfPath, opts = {}) {
  return (await processAgreements(pdfPath, opts))[0];
}

module.exports = {
  processAgreement,
  processAgreements,
  extractAgreementValue,
  extractNames,
  extractAgreementDate,
  flatNoFromFilename,
  NAME_PAGES,
  VALUE_PAGE,
};
