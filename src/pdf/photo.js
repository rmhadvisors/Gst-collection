'use strict';

const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const IMAGE_RE = /\.(jpe?g|png)$/i;
const MIN_WIDTH = 2400; // phone photos are ~1600 px wide: too small for the Marathi OCR

const isImageFile = (file) => IMAGE_RE.test(String(file || ''));

/**
 * A phone photo of an Index-II page, made OCR-ready: EXIF rotation applied,
 * greyscale, contrast stretched, scaled up and sharpened. Written as page 1 of
 * `outDir`, so one image goes through the same pipeline as one PDF page.
 *
 * @returns {Promise<{page:number, file:string, bytes:number}>}
 */
async function preparePhoto(src, outDir) {
  fs.mkdirSync(outDir, { recursive: true });
  const out = path.join(outDir, 'page-01.png');
  try {
    const meta = await sharp(src).metadata();
    const rotated = meta.orientation && meta.orientation >= 5; // width/height swap after EXIF rotation
    const width = rotated ? meta.height : meta.width;
    let img = sharp(src).rotate().grayscale().normalise();
    if (!width || width < MIN_WIDTH) img = img.resize({ width: MIN_WIDTH, kernel: 'lanczos3' });
    await img.sharpen().png().toFile(out);
  } catch (e) {
    // unreadable by sharp: OCR the original bytes and let the 50-character check speak up
    fs.copyFileSync(src, out.replace(/\.png$/, path.extname(src) || '.jpg'));
    return { page: 1, file: out.replace(/\.png$/, path.extname(src) || '.jpg'), bytes: fs.statSync(src).size };
  }
  return { page: 1, file: out, bytes: fs.statSync(out).size };
}

module.exports = { preparePhoto, isImageFile };
