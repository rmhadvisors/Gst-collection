'use strict';

const fs = require('fs');
const path = require('path');
const { createWorker } = require('tesseract.js');
const { withTimeout } = require('../util/timeout');

// Marathi (Index-II) + English. mar.traineddata / eng.traineddata live in the
// project root and are loaded from there (no network access at run time).
const LANGS = 'mar+eng';
const LANG_PATH = path.join(__dirname, '..', '..');

const MIN_OCR_CHARS = 50; // fewer characters than this on a page = the OCR failed
const CALL_TIMEOUT_MS = Number(process.env.OCR_CALL_TIMEOUT_MS) || 4 * 60 * 1000; // one page, one mode
const TERMINATE_TIMEOUT_MS = 15 * 1000;

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

// ── worker state, reported through an optional logger (the server adds the job id) ──
let logger = () => {};
let state = 'none'; // none | created | busy | idle | terminated | failed
function setState(next, detail) {
  state = next;
  logger(`ocr worker ${next}${detail ? ` (${detail})` : ''}`);
}
const setOcrLogger = (fn) => { logger = typeof fn === 'function' ? fn : () => {}; };
const ocrState = () => state;

let workerPromise = null;

function getWorker() {
  assertOcrReady();
  if (!workerPromise) {
    // a failed start must not be cached: the next call has to try again
    workerPromise = createWorker(LANGS, 1, { langPath: LANG_PATH, gzip: false }).then((w) => {
      setState('created');
      return w;
    }, (e) => {
      workerPromise = null;
      setState('failed', `could not start: ${e.message}`);
      throw e;
    });
  }
  return workerPromise;
}

/** Terminate the worker (never waits longer than TERMINATE_TIMEOUT_MS) and forget it. */
async function killWorker() {
  const pending = workerPromise;
  workerPromise = null;
  if (!pending) return;
  try {
    const worker = await withTimeout(pending, TERMINATE_TIMEOUT_MS, 'OCR worker start');
    await withTimeout(worker.terminate(), TERMINATE_TIMEOUT_MS, 'OCR worker terminate');
    setState('terminated');
  } catch (e) {
    setState('terminated', `forced: ${e.message}`);
  }
}

/**
 * OCR an image file to plain text. `psm` is the Tesseract page-segmentation mode.
 * A call that hangs or crashes is abandoned after CALL_TIMEOUT_MS, the worker is
 * replaced and the call is tried once more, so a stuck worker can never stall a job.
 */
async function ocrImage(imagePath, psm = '3') {
  let lastError;
  for (let attempt = 1; attempt <= 2; attempt += 1) {
    try {
      const worker = await withTimeout(getWorker(), CALL_TIMEOUT_MS, 'OCR worker start');
      setState('busy', `psm ${psm}${attempt > 1 ? ', retry' : ''}`);
      await worker.setParameters({ tessedit_pageseg_mode: String(psm) });
      const { data } = await withTimeout(worker.recognize(imagePath), CALL_TIMEOUT_MS, `OCR of ${path.basename(imagePath)}`);
      setState('idle');
      return data.text;
    } catch (e) {
      lastError = e;
      setState('failed', e.message);
      await killWorker();
    }
  }
  throw lastError;
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
  try {
    await require('./dateRetry').closeDateWorker();
  } catch (_) { /* best effort: never let teardown hang or throw */ }
  await killWorker();
}

module.exports = {
  ocrImage, ocrImageVariants, closeOcr, assertOcrReady, MIN_OCR_CHARS, setOcrLogger, ocrState,
};
