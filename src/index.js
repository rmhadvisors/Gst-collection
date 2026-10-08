#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { processAgreement } = require('./pdf/agreement');
const { closeOcr } = require('./pdf/ocr');
const { parseTally } = require('./tally/parse');
const { buildRecord } = require('./engine/collection');
const { buildTallyOnlyRecords } = require('./engine/tallyOnly');
const { writeCollectionWorkbook } = require('./output/workbook');

function parseArgs(argv) {
  const args = { agreements: [] };
  for (let i = 2; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--tally') args.tally = argv[++i];
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
  node src/index.js --tally <tally.xlsx> [--out <out.xlsx>] <agreement1.pdf> [agreement2.pdf | folder ...]

Options:
  --tally <file>       Tally receipts workbook (required)
  --sheet <name>       Tally sheet name (default: first sheet)
  --out <file>         Output workbook (default: collection-output.xlsx)
  --overrides <file>   JSON of per-flat corrections (default: config/overrides.json if present)
  --keep-pages         Keep the extracted page JPEGs for inspection
  -h, --help           Show this help

Agreements may be individual .pdf files or folders containing .pdf files.
The flat number is taken from the PDF filename (e.g. "1804-CHHAYA KADAM.pdf").
`);
}

function collectPdfs(inputs) {
  const pdfs = [];
  for (const inp of inputs) {
    if (!fs.existsSync(inp)) { console.warn(`! not found: ${inp}`); continue; }
    const stat = fs.statSync(inp);
    if (stat.isDirectory()) {
      fs.readdirSync(inp)
        .filter((f) => f.toLowerCase().endsWith('.pdf'))
        .forEach((f) => pdfs.push(path.join(inp, f)));
    } else if (inp.toLowerCase().endsWith('.pdf')) {
      pdfs.push(inp);
    }
  }
  return pdfs;
}

function loadOverrides(file) {
  const target = file || path.join(__dirname, '..', 'config', 'overrides.json');
  if (!fs.existsSync(target)) return {};
  try {
    return JSON.parse(fs.readFileSync(target, 'utf8'));
  } catch (e) {
    console.warn(`! could not parse overrides ${target}: ${e.message}`);
    return {};
  }
}

function fmt(n) {
  if (n == null || n === '') return '';
  return n.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

async function main() {
  const args = parseArgs(process.argv);
  if (args.help || !args.tally || args.agreements.length === 0) {
    usage();
    process.exit(args.help ? 0 : 1);
  }

  const outPath = args.out || 'collection-output.xlsx';
  const overrides = loadOverrides(args.overrides);

  console.log(`Reading tally: ${args.tally}`);
  const tally = parseTally(args.tally, args.sheet);
  console.log(`  -> ${tally.rows.length} receipt rows across ${tally.byFlat.size} flats`);

  const pdfs = collectPdfs(args.agreements);
  if (!pdfs.length) {
    console.error('No agreement PDFs found.');
    process.exit(1);
  }

  const workDir = path.join(path.dirname(outPath), '_agreement_pages');
  const records = [];

  for (const pdf of pdfs) {
    process.stdout.write(`Processing ${path.basename(pdf)} ... `);
    try {
      const agreement = await processAgreement(pdf, { workDir });
      const ov = overrides[agreement.flatNoHint] || {};
      const rec = buildRecord(agreement, tally, ov);
      rec.pdf = agreement.pdf;
      records.push(rec);
      console.log(
        `flat ${rec.flatNo} | ${rec.name} | agr ${fmt(rec.agreementValue)} `
        + `| GST@${rec.gstRate * 100}% ${fmt(rec.gstOnAgreement)} | recd ${fmt(rec.totalReceived)} `
        + `| bal ${fmt(rec.balance)} [${rec.matchReason}]`,
      );
    } catch (e) {
      console.log(`FAILED: ${e.message}`);
    }
  }

  await closeOcr();

  // Include every other flat present in the Tally receipts (no agreement yet):
  // their receipts are listed with GST computed at the default 1% rate.
  const matched = new Set(records.map((r) => r.flatNo));
  const tallyOnly = buildTallyOnlyRecords(tally, matched, overrides, {});
  if (tallyOnly.length) {
    console.log(`Adding ${tallyOnly.length} flat(s) from Tally without an agreement (GST @ 1% default)`);
    records.push(...tallyOnly);
  }

  // sort records by flat number for a tidy sheet
  records.sort((a, b) => Number(a.flatNo) - Number(b.flatNo));

  writeCollectionWorkbook(records, outPath, { projectName: tally.projectName || undefined });
  console.log(`\nWrote ${records.length} flat(s) to: ${outPath}`);

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
