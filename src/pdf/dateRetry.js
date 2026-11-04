'use strict';

/**
 * Second-chance OCR for the Index-II "दस्त नोंदणी केल्याचा दिनांक" (row 9/10)
 * date. Phone photos of the page are skewed, shadowed and low resolution, so
 * the generic page OCR often drops or mangles the date. Here the value column
 * is cropped in a few vertical windows, deskewed by -3..+3 degrees, upscaled
 * (higher effective DPI) and read with an English digits-only worker; every
 * dd/mm/yyyy or dd-mm-yyyy hit is a vote and the most common valid date wins.
 */

const path = require('path');
const sharp = require('sharp');
const { createWorker } = require('tesseract.js');

const LANG_PATH = path.join(__dirname, '..', '..');
const ANGLES = [0, -3, 3, -1.5, 1.5];
const SCALES = [3, 4];
const WINDOWS = [0.38, 0.41, 0.44, 0.47, 0.50, 0.53, 0.56]; // top of the crop, fraction of page height
const DATE_RE = /(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(20\d{2})/g;
const MIN_VOTES = 2;

let workerPromise = null;
function digitsWorker() {
  if (!workerPromise) {
    workerPromise = createWorker('eng', 1, { langPath: LANG_PATH, gzip: false }).then(async (w) => {
      await w.setParameters({ tessedit_char_whitelist: '0123456789/-. ', tessedit_pageseg_mode: '6' });
      return w;
    });
  }
  return workerPromise;
}

async function closeDateWorker() {
  if (workerPromise) {
    const w = await workerPromise;
    await w.terminate();
    workerPromise = null;
  }
}

/** Valid calendar dates only (dd/mm, 2020-2035); returns a Date or null. */
function toDate(d, m, y) {
  const dt = new Date(y, m - 1, d);
  const ok = dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
  return ok && y >= 2020 && y <= 2035 ? dt : null;
}

/** All dd/mm/yyyy | dd-mm-yyyy dates in OCR text, as "dd/mm/yyyy" strings. */
function findDates(text) {
  const out = [];
  for (const m of String(text || '').matchAll(DATE_RE)) {
    if (toDate(+m[1], +m[2], +m[3])) out.push(`${m[1].padStart(2, '0')}/${m[2].padStart(2, '0')}/${m[3]}`);
  }
  return out;
}

/** Most-voted date text (>= MIN_VOTES; first seen wins ties) or null. */
function pickDate(texts) {
  const votes = new Map();
  for (const t of texts) for (const d of findDates(t)) votes.set(d, (votes.get(d) || 0) + 1);
  let best = null;
  let bestN = 0;
  for (const [d, n] of votes) if (n > bestN) { best = d; bestN = n; }
  return bestN >= MIN_VOTES ? best : null;
}

/**
 * @param {string} imagePath page image
 * @returns {Promise<{date: Date, text: string}|null>}
 */
async function retryAgreementDate(imagePath) {
  const meta = await sharp(imagePath).metadata();
  const left = Math.round(meta.width * 0.32);
  const width = Math.round(meta.width * 0.37);
  const height = Math.round(meta.height * 0.075);
  const worker = await digitsWorker();
  const texts = [];
  for (const angle of ANGLES) {
    const rotated = await sharp(imagePath).rotate(angle, { background: '#fff' }).toBuffer();
    for (const t of WINDOWS) {
      for (const scale of SCALES) {
        const img = await sharp(rotated)
          .extract({ left, top: Math.round(meta.height * t), width, height })
          .resize({ width: width * scale, kernel: 'lanczos3' })
          .grayscale().clahe({ width: 8, height: 8, maxSlope: 3 }).normalise()
          .extend({ top: 30, bottom: 30, left: 30, right: 30, background: '#fff' })
          .png().toBuffer();
        texts.push((await worker.recognize(img)).data.text);
      }
    }
  }
  const text = pickDate(texts);
  if (!text) return null;
  const [d, m, y] = text.split('/').map(Number);
  return { date: toDate(d, m, y), text };
}

module.exports = { retryAgreementDate, closeDateWorker, findDates, pickDate };
