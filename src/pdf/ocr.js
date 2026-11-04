'use strict';

const fs = require('fs');
const path = require('path');
const { createWorker } = require('tesseract.js');

// Marathi (Index-II) + English. mar.traineddata / eng.traineddata live in the
// project root and are loaded from there (no network access at run time).
const LANGS = 'mar+eng';
const LANG_PATH = path.join(__dirname, '..', '..');

const MIN_OCR_CHARS = 50; // fewer characters than this on a page = the OCR failed

/**
 * Fail early and clearly when a Tesseract language file is missing, instead of
 * a cryptic worker error half-way through an upload. Call once at startup.
 */
function assertOcrReady() {
  const missing = LANGS.split('+')
    .map((l) => `${l}.traineddata`)
    .filter((f) => !fs.existsSync(path.join(LANG_PATH, f)));
  if (missing.length) {
    throw new Error(
      `OCR language data missing: ${missing.join(', ')}. Place ${missing.join(' and ')} in ${LANG_PATH} `
      + '(download from https://github.com/tesseract-ocr/tessdata) and restart.',
    );
  }
}

let workerPromise = null;

function getWorker() {
  assertOcrReady();
  if (!workerPromise) {
    workerPromise = createWorker(LANGS, 1, { langPath: LANG_PATH, gzip: false });
  }
  return workerPromise;
}

/** OCR an image file to plain text. `psm` is the Tesseract page-segmentation mode. */
async function ocrImage(imagePath, psm = '3') {
  const worker = await getWorker();
  await worker.setParameters({ tessedit_pageseg_mode: String(psm) });
  const { data } = await worker.recognize(imagePath);
  return data.text;
}

/**
 * OCR one page under several segmentation modes. Marathi Index-II tables lose
 * different cells under each mode (e.g. the consideration row), so callers
 * combine the variants. Modes: 3 auto, 4 single column, 6 single block, 11 sparse text.
 */
async function ocrImageVariants(imagePath, modes = ['3', '4', '11']) {
  const texts = [];
  for (const m of modes) texts.push(await ocrImage(imagePath, m));
  return texts;
}

async function closeOcr() {
  await require('./dateRetry').closeDateWorker();
  if (workerPromise) {
    const worker = await workerPromise;
    await worker.terminate();
    workerPromise = null;
  }
}

module.exports = {
  ocrImage, ocrImageVariants, closeOcr, assertOcrReady, MIN_OCR_CHARS,
};
