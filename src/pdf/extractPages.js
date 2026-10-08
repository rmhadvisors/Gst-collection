'use strict';

const fs = require('fs');
const path = require('path');

/**
 * Extract embedded JPEG (DCTDecode) image streams from a scanned PDF.
 * These registration PDFs store exactly one full-page scanned JPEG per page,
 * so the Nth image stream corresponds to page N.
 *
 * @param {string} pdfPath
 * @param {object} [opts]
 * @param {number[]} [opts.pages] 1-based page numbers to extract (default: all)
 * @param {string} [opts.outDir] directory to write JPEGs into
 * @returns {{page:number, file:string, bytes:number}[]}
 */
function extractPageImages(pdfPath, opts = {}) {
  const buf = fs.readFileSync(pdfPath);
  const text = buf.latin1Slice(0, buf.length);
  const outDir = opts.outDir || path.join(path.dirname(pdfPath), '_pages');
  fs.mkdirSync(outDir, { recursive: true });

  const objRe = /<<[^>]*?\/Subtype\s*\/Image[^>]*?>>\s*stream\r?\n/g;
  const results = [];
  let m;
  let pageIdx = 0;
  while ((m = objRe.exec(text)) !== null) {
    const dict = m[0];
    if (!/DCTDecode/.test(dict)) continue; // only handle JPEG-encoded pages
    pageIdx += 1;
    if (opts.pages && !opts.pages.includes(pageIdx)) continue;

    const lenMatch = dict.match(/\/Length\s+(\d+)/);
    const start = m.index + m[0].length;
    let end;
    if (lenMatch) {
      end = start + parseInt(lenMatch[1], 10);
    } else {
      end = text.indexOf('endstream', start);
    }
    const slice = buf.subarray(start, end);
    const file = path.join(outDir, `page-${String(pageIdx).padStart(2, '0')}.jpg`);
    fs.writeFileSync(file, slice);
    results.push({ page: pageIdx, file, bytes: slice.length });
  }
  return results;
}

module.exports = { extractPageImages };
