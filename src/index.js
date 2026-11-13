#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { assertOcrReady } = require('./pdf/ocr');
const { buildRecords, loadOverrides } = require('./services/build-records');
const { writeStyledCollectionWorkbook } = require('./output/workbook-styled');

function parseArgs(argv) {
  const args = { agreements: [], tally: [] };
  for (let i = 2; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--tally') args.tally.push(argv[++i]);
    else if (a === '--sheet') args.sheet = argv[++i];
    else if (a === '--out') args.out = argv[++i];
    else if (a === '--overrides') args.overrides = argv[++i];
    else if (a === '--keep-pages') args.keepPages = true;
    else if (a === '--help' || a === '-h') args.help = true;
    else args.agreements.push(a);
  }
  return args;
}

function usage() {
  console.log(`
Collection Sheet Automation
===========================
Reads scanned Agreement PDF(s) + a Tally receipts workbook and produces a
collection sheet (== summary + Wing B detail) with GST computed per the
45-lakh rule (>45L => 5%, <=45L => 1%).

Usage:
  node src/index.js --tally <tally.xlsx> [--tally <more.xlsx> ...] [--out <out.xlsx>] <agreement1.pdf> [agreement2.pdf | folder ...]

Options:
  --tally <file>       Tally receipts workbook (required; repeat for several books).
                       Every sheet of every file is read.
  --out <file>         Output workbook (default: collection-output.xlsx)
  --overrides <file>   JSON of per-flat corrections (default: config/overrides.json if present)
  --keep-pages         Keep the extracted page JPEGs for inspection
  -h, --help           Show this help

Agreements may be .pdf, .jpg, .jpeg or .png files (a photo is one Index-II page), or folders of them.
The flat number is read from the agreement; the filename (e.g. "1804-CHHAYA KADAM.pdf") is only a fallback.
The building comes from the Index-II text; config/overrides.json can name each Tally book's building.
The old note follows. The flat number is taken from the PDF filename (e.g. "1804-CHHAYA KADAM.pdf").
`);
}

const IMAGE_OR_PDF = /\.(pdf|jpe?g|png)$/i; // a photo of an Index-II page is one page

function collectPdfs(inputs) {
  const pdfs = [];
  for (const inp of inputs) {
    if (!fs.existsSync(inp)) { console.warn(`! not found: ${inp}`); continue; }
    const stat = fs.statSync(inp);
    if (stat.isDirectory()) {
      fs.readdirSync(inp)
        .filter((f) => IMAGE_OR_PDF.test(f))
        .forEach((f) => pdfs.push(path.join(inp, f)));
    } else if (IMAGE_OR_PDF.test(inp)) {
      pdfs.push(inp);
    }
  }
  return pdfs;
}

function fmt(n) {
  if (n == null || n === '') return '';
  return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help || !args.tally.length || args.agreements.length === 0) {
    usage();
    process.exit(args.help ? 0 : 1);
  }

  const outPath = args.out || 'collection-output.xlsx';
  const overrides = loadOverrides(args.overrides);
  assertOcrReady();

  const pdfs = collectPdfs(args.agreements);
  if (!pdfs.length) {
    console.error('No agreement PDFs found.');
    process.exit(1);
  }

  const workDir = path.join(path.dirname(path.resolve(outPath)), '_agreement_pages');
  const asFile = (p) => ({ path: p, originalname: path.basename(p) });

  // Same pipeline and same workbook as the web upload.
  const {
    records, tally, tallyTotal, tallyFiles,
  } = await buildRecords({
    tallyFiles: args.tally.map(asFile),
    agreementFiles: pdfs.map(asFile),
    overrides,
    workDir,
    onProgress: (msg) => console.log(msg),
  });

  for (const f of tallyFiles) {
    const note = f.duplicateUpload ? 'duplicate upload - ignored' : (f.receipts ? '' : 'NOT READ');
    console.log(`Tally ${f.file}: ${f.receipts} receipts, ${fmt(f.total)} ${note}`.trim());
  }
  for (const rec of records.filter((r) => r.hasAgreement)) {
    console.log(
      `flat ${rec.flatNo} | ${rec.name} | agr ${fmt(rec.agreementValue)} [${rec.valueSource}] `
      + `| GST@${rec.gstRate * 100}% ${fmt(rec.gstOnAgreement)} | recd ${fmt(rec.totalReceived)} `
      + `| bal ${fmt(rec.balance)} [${rec.matchReason}]`,
    );
  }

  await writeStyledCollectionWorkbook(records, outPath, {
    projectName: tally.projectName || undefined,
    reconciliation: { tallyTotal, files: tallyFiles },
  });
  console.log(`\nWrote ${records.length} flat/party block(s) to: ${outPath}`);

  if (!args.keepPages) {
    fs.rmSync(workDir, { recursive: true, force: true });
  } else {
    console.log(`Extracted pages kept in: ${workDir}`);
  }
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
