'use strict';

const fs = require('fs');
const path = require('path');
const { pathToFileURL } = require('url');

const RENDER_DPI = 300;
const PDF_POINTS_PER_INCH = 72;

let pdfjsPromise = null;

/** pdf.js is ESM-only; load it lazily so the CLI starts without it. */
function loadPdfjs() {
  if (!pdfjsPromise) pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs');
  return pdfjsPromise;
}

async function openPdf(pdfPath) {
  const pdfjs = await loadPdfjs();
  const standardFontDataUrl = `${pathToFileURL(path.join(
    path.dirname(require.resolve('pdfjs-dist/package.json')), 'standard_fonts',
  ))}/`;
  const task = pdfjs.getDocument({
    data: new Uint8Array(fs.readFileSync(pdfPath)),
    standardFontDataUrl,
    useSystemFonts: false,
    verbosity: 0,
  });
  const doc = await task.promise;
  return { doc, close: () => task.destroy() };
}

/** Number of pages in a PDF. */
async function countPages(pdfPath) {
  const { doc, close } = await openPdf(pdfPath);
  try { return doc.numPages; } finally { await close(); }
}

/**
 * Render the given 1-based pages at 300 dpi to PNG files. Used when the embedded
 * page image cannot be extracted (unsupported filter, vector/text PDFs, ...).
 *
 * @returns {Promise<{page:number, file:string, bytes:number}[]>} (carries `totalPages`)
 */
async function renderPages(pdfPath, pages, outDir) {
  const { createCanvas } = require('@napi-rs/canvas');
  fs.mkdirSync(outDir, { recursive: true });
  const { doc, close } = await openPdf(pdfPath);
  const results = [];
  try {
    for (const n of pages) {
      if (n < 1 || n > doc.numPages) continue;
      const page = await doc.getPage(n);
      const viewport = page.getViewport({ scale: RENDER_DPI / PDF_POINTS_PER_INCH });
      const canvas = createCanvas(Math.ceil(viewport.width), Math.ceil(viewport.height));
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;
      const file = path.join(outDir, `page-${String(n).padStart(2, '0')}.png`);
      const png = canvas.toBuffer('image/png');
      fs.writeFileSync(file, png);
      results.push({ page: n, file, bytes: png.length, rendered: true });
      page.cleanup();
    }
    results.totalPages = doc.numPages;
  } finally {
    await close();
  }
  return results;
}

module.exports = { renderPages, countPages, RENDER_DPI };
