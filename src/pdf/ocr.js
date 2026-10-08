'use strict';

const { createWorker } = require('tesseract.js');

let workerPromise = null;

function getWorker() {
  if (!workerPromise) {
    workerPromise = createWorker('eng');
  }
  return workerPromise;
}

/** OCR an image file to plain text. */
async function ocrImage(imagePath) {
  const worker = await getWorker();
  const { data } = await worker.recognize(imagePath);
  return data.text;
}

async function closeOcr() {
  if (workerPromise) {
    const worker = await workerPromise;
    await worker.terminate();
    workerPromise = null;
  }
}

module.exports = { ocrImage, closeOcr };
