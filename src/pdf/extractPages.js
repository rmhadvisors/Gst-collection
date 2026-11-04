'use strict';

const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i += 1) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

/** Wrap raw 8-bit pixel rows (no filter bytes unless `filtered`) into a PNG. */
function encodePng(raw, width, height, channels, filtered) {
  const rowBytes = width * channels;
  let scan = raw;
  if (!filtered) {
    scan = Buffer.alloc((rowBytes + 1) * height);
    for (let y = 0; y < height; y += 1) {
      raw.copy(scan, y * (rowBytes + 1) + 1, y * rowBytes, (y + 1) * rowBytes);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = channels === 1 ? 0 : channels === 4 ? 6 : 2;
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    pngChunk('IDAT', zlib.deflateSync(scan)),
    pngChunk('IEND', Buffer.alloc(0)),
  ]);
}

/**
 * Extract the embedded page images from a scanned PDF. Handles JPEG
 * (DCTDecode) streams as-is and 8-bit Gray/RGB FlateDecode streams (e.g. phone
 * "scan to PDF" exports) by re-encoding them as PNG. These PDFs store exactly
 * one full-page image per page, so the Nth image stream is page N.
 *
 * @param {string} pdfPath
 * @param {object} [opts]
 * @param {number[]} [opts.pages] 1-based page numbers to extract (default: all)
 * @param {string} [opts.outDir] directory to write images into
 * @returns {{page:number, file:string, bytes:number}[]} (array also carries `totalPages`)
 */
function extractPageImages(pdfPath, opts = {}) {
  const buf = fs.readFileSync(pdfPath);
  const text = buf.latin1Slice(0, buf.length);
  const outDir = opts.outDir || path.join(path.dirname(pdfPath), '_pages');
  fs.mkdirSync(outDir, { recursive: true });

  const objRe = /<<(?:(?!endobj)[\s\S])*?\/Subtype\s*\/Image(?:(?!endobj)[\s\S])*?>>\s*stream\r?\n/g;
  const results = [];
  const failedPages = [];
  let m;
  let pageIdx = 0;
  while ((m = objRe.exec(text)) !== null) {
    const dict = m[0];
    const isDct = /DCTDecode/.test(dict);
    const isFlate = /FlateDecode/.test(dict);
    if (!isDct && !isFlate) continue;
    pageIdx += 1;
    if (opts.pages && !opts.pages.includes(pageIdx)) continue;

    const lenMatch = dict.match(/\/Length\s+(\d+)(?!\s+\d+\s+R)/);
    const start = m.index + m[0].length;
    const end = lenMatch ? start + parseInt(lenMatch[1], 10) : text.indexOf('endstream', start);
    const slice = buf.subarray(start, end);
    const stem = path.join(outDir, `page-${String(pageIdx).padStart(2, '0')}`);

    let file;
    let out = slice;
    try {
      if (isDct) {
        file = `${stem}.jpg`;
        // /Filter [/FlateDecode /DCTDecode]: inflate first to get the JPEG bytes
        if (isFlate) out = zlib.inflateSync(slice);
      } else {
        const width = parseInt((dict.match(/\/Width\s+(\d+)/) || [])[1], 10);
        const height = parseInt((dict.match(/\/Height\s+(\d+)/) || [])[1], 10);
        const bpc = parseInt((dict.match(/\/BitsPerComponent\s+(\d+)/) || [])[1], 10);
        const channels = /DeviceRGB/.test(dict) ? 3 : /DeviceGray/.test(dict) ? 1 : 0;
        if (!width || !height || bpc !== 8 || !channels) throw new Error('unsupported image format');
        const predictor = parseInt((dict.match(/\/Predictor\s+(\d+)/) || [])[1] || '1', 10);
        out = encodePng(zlib.inflateSync(slice), width, height, channels, predictor >= 10);
        file = `${stem}.png`;
      }
      if (!out || !out.length) throw new Error('empty image stream');
      fs.writeFileSync(file, out);
      results.push({ page: pageIdx, file, bytes: out.length });
    } catch (_) {
      // corrupt / unsupported stream: the caller renders this page at 300 dpi instead
      failedPages.push(pageIdx);
    }
  }
  results.totalPages = pageIdx;
  results.failedPages = failedPages;
  return results;
}

module.exports = { extractPageImages };
