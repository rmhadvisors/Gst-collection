'use strict';

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { processAgreement } = require('./src/pdf/agreement');
const { closeOcr } = require('./src/pdf/ocr');
const { parseTally, mergeTallies } = require('./src/tally/parse');
const { buildRecord } = require('./src/engine/collection');
const { buildTallyOnlyRecords } = require('./src/engine/tallyOnly');
const { writeCollectionWorkbook }   = require('./src/output/workbook');
const { writeStyledCollectionWorkbook } = require('./src/output/workbook-styled');
const { writeGstWorkbook }     = require('./src/output/gstWorkbook');
const { updateCollectionSheet } = require('./src/services/collection-update.service');
const { updateGstWorkbook }     = require('./src/services/gst-update.service');

const app = express();
const PORT = process.env.PORT || 3000; // <-- Change this on line 19

app.use((req, res, next) => {
  const basePath = (req.get('x-forwarded-prefix') || '').replace(/\/$/, '');
  res.locals.basePath = basePath;
  next();
});

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));

const UPLOAD_DIR = path.join(__dirname, 'uploads');
const OUTPUT_DIR = path.join(__dirname, 'output');
fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOAD_DIR),
    filename: (req, file, cb) => {
      const unique = Date.now() + '-' + Math.round(Math.random() * 1e4);
      cb(null, unique + '-' + file.originalname);
    },
  }),
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (file.fieldname === 'tally' && ['.xlsx', '.xls'].includes(ext)) cb(null, true);
    else if (file.fieldname === 'agreements' && ext === '.pdf') cb(null, true);
    else if (file.fieldname === 'existingWorkbook' && ['.xlsx'].includes(ext)) cb(null, true);
    else cb(new Error(`Invalid file type: ${file.originalname}`));
  },
  limits: { fileSize: 100 * 1024 * 1024 },
});

const uploadFields = upload.fields([
  { name: 'tally',            maxCount: 50 },
  { name: 'agreements',       maxCount: 50 },
  { name: 'existingWorkbook', maxCount:  1 },
]);

function fmtINR(n) {
  if (n == null || n === '') return '-';
  return Number(n).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function loadOverrides() {
  const file = path.join(__dirname, 'config', 'overrides.json');
  if (!fs.existsSync(file)) return {};
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

/**
 * Shared processing: parse tally + OCR agreements → sorted records array.
 * agreementFiles may be empty (tally-only run).
 */
async function buildRecords(tallyFiles, agreementFiles, sheetName, overrides) {
  const tallies = tallyFiles.map((tf) => parseTally(tf.path, sheetName || undefined));
  const tally = mergeTallies(tallies);
  const workDir = path.join(OUTPUT_DIR, '_pages_' + Date.now());
  const records = [];

  for (const af of agreementFiles) {
    const agreement = await processAgreement(af.path, { workDir });
    const rec = buildRecord(agreement, tally, overrides[agreement.flatNoHint] || {});
    rec.pdf = agreement.pdf;
    records.push(rec);
  }

  await closeOcr();

  // Add every other flat in the Tally receipts (no agreement) with GST @ 1% default.
  const matched = new Set(records.map((r) => r.flatNo));
  records.push(...buildTallyOnlyRecords(tally, matched, overrides, {}));
  records.sort((a, b) => Number(a.flatNo) - Number(b.flatNo));

  // cleanup
  fs.rmSync(workDir, { recursive: true, force: true });
  for (const f of [...tallyFiles, ...agreementFiles]) {
    try { fs.unlinkSync(f.path); } catch (_) { /* ignore */ }
  }

  return { records, tally };
}

// ─── Routes ──────────────────────────────────────────────

app.get('/', (req, res) => {
  res.render('index');
});

// ── Collection Sheet ──────────────────────────────────────
app.post('/process', (req, res, next) => {
  uploadFields(req, res, async (err) => {
    if (err) return res.render('index', { error: err.message });
    try {
      const tallyFiles = req.files['tally'] || [];
      const agreementFiles = req.files['agreements'] || [];

      if (!tallyFiles.length) return res.render('index', { error: 'Please upload at least one Tally receipts file (.xlsx)' });
      if (!agreementFiles.length) return res.render('index', { error: 'Please upload at least one Agreement PDF' });

      const overrides = loadOverrides();
      const { records, tally } = await buildRecords(tallyFiles, agreementFiles, req.body.sheet, overrides);

      const outFilename = `collection-${Date.now()}.xlsx`;
      const outPath = path.join(OUTPUT_DIR, outFilename);
      await writeStyledCollectionWorkbook(records, outPath, { projectName: tally.projectName || undefined });

      res.render('result', {
        records, fmtINR,
        downloadFile: outFilename,
        gstDownloadFile: null,
        projectName: tally.projectName || 'Collection Sheet',
      });
    } catch (e) {
      console.error(e);
      res.render('index', { error: 'Processing failed: ' + e.message });
    }
  });
});

// ── GST Calculation Workbook ──────────────────────────────
app.post('/process-gst', (req, res, next) => {
  uploadFields(req, res, async (err) => {
    if (err) return res.render('index', { error: err.message });
    try {
      const tallyFiles = req.files['tally'] || [];
      const agreementFiles = req.files['agreements'] || [];

      if (!tallyFiles.length) return res.render('index', { error: 'Please upload at least one Tally receipts file (.xlsx)' });
      // Agreement PDFs are OPTIONAL for the GST workbook
      // (tally-only → all flats treated as Case 3)

      const overrides = loadOverrides();
      const { records, tally } = await buildRecords(tallyFiles, agreementFiles, req.body.sheet, overrides);

      const outFilename = `gst-workbook-${Date.now()}.xlsx`;
      const outPath = path.join(OUTPUT_DIR, outFilename);
      writeGstWorkbook(records, outPath, { projectName: tally.projectName || undefined });

      res.render('result', {
        records, fmtINR,
        downloadFile: null,
        gstDownloadFile: outFilename,
        projectName: tally.projectName || 'GST Workbook',
      });
    } catch (e) {
      console.error(e);
      res.render('index', { error: 'GST workbook generation failed: ' + e.message });
    }
  });
});

// ── Update Collection Sheet ───────────────────────────────
app.post('/update-collection', (req, res, next) => {
  uploadFields(req, res, async (err) => {
    if (err) return res.render('index', { error: err.message });
    try {
      const tallyFiles    = req.files['tally']            || [];
      const agrFiles      = req.files['agreements']       || [];
      const existingFiles = req.files['existingWorkbook'] || [];

      if (!tallyFiles.length)    return res.render('index', { error: 'Please upload the latest Tally receipts file.' });
      if (!existingFiles.length) return res.render('index', { error: 'Please upload your existing Collection Sheet.' });

      const overrides = loadOverrides();

      // Parse tally (agreements are handled inside updateCollectionSheet via OCR)
      const tallies = tallyFiles.map((tf) => parseTally(tf.path, req.body.sheet || undefined));
      const tally   = mergeTallies(tallies);

      const agrPaths    = agrFiles.map(f => f.path);
      const outFilename = `updated-collection-${Date.now()}.xlsx`;
      const outPath     = path.join(OUTPUT_DIR, outFilename);
      const workDir     = path.join(OUTPUT_DIR, '_pages_' + Date.now());
      fs.mkdirSync(workDir, { recursive: true });

      const result = await updateCollectionSheet({
        existingPath:   existingFiles[0].path,
        outPath,
        tally,
        agreementPaths: agrPaths,   // ← PDFs now passed for OCR
        overrides,
        workDir,
        projectName:    tally.projectName || 'YASHWANTH COUNTY',
      });

      fs.rmSync(workDir, { recursive: true, force: true });
      for (const f of [...tallyFiles, ...agrFiles, ...existingFiles]) {
        try { fs.unlinkSync(f.path); } catch (_) {}
      }

      res.render('result', {
        records: [],
        fmtINR,
        downloadFile:    null,
        gstDownloadFile: null,
        updatedCollectionFile: outFilename,
        updatedGstFile:        null,
        updateSummary: {
          newPayments:     result.newPayments,
          reconciledFlats: result.reconciledFlats,
          newFlats:        result.newFlats,
          pdfWarnings:     result.pdfWarnings,
          newMonths:       [],
          updatedMonths:   [],
        },
        projectName: tally.projectName || 'Collection Sheet (Updated)',
      });
    } catch (e) {
      console.error(e);
      res.render('index', { error: 'Update failed: ' + e.message });
    }
  });
});

// ── Update GST Workbook ───────────────────────────────────
app.post('/update-gst', (req, res, next) => {
  uploadFields(req, res, async (err) => {
    if (err) return res.render('index', { error: err.message });
    try {
      const tallyFiles    = req.files['tally']            || [];
      const agrFiles      = req.files['agreements']       || [];
      const existingFiles = req.files['existingWorkbook'] || [];

      if (!tallyFiles.length)    return res.render('index', { error: 'Please upload the latest Tally receipts file.' });
      if (!existingFiles.length) return res.render('index', { error: 'Please upload your existing GST Workbook.' });

      const overrides = loadOverrides();
      const { records } = await buildRecords(tallyFiles, agrFiles, req.body.sheet, overrides);

      const outFilename = `updated-gst-${Date.now()}.xlsx`;
      const outPath     = path.join(OUTPUT_DIR, outFilename);

      const result = await updateGstWorkbook({
        existingPath: existingFiles[0].path,
        outPath,
        records,
        projectName: 'SPACE HOME',
      });

      for (const f of [...tallyFiles, ...agrFiles, ...existingFiles]) {
        try { fs.unlinkSync(f.path); } catch (_) {}
      }

      res.render('result', {
        records: [],
        fmtINR,
        downloadFile:    null,
        gstDownloadFile: null,
        updatedCollectionFile: null,
        updatedGstFile:        outFilename,
        updateSummary: {
          newPayments:    0,
          reconciledFlats: [],
          newMonths:      result.newMonths,
          updatedMonths:  result.updatedMonths,
        },
        projectName: 'GST Workbook (Updated)',
      });
    } catch (e) {
      console.error(e);
      res.render('index', { error: 'GST update failed: ' + e.message });
    }
  });
});

// ── Downloads ─────────────────────────────────────────────
app.get('/download/:file', (req, res) => {
  const file = path.join(OUTPUT_DIR, req.params.file);
  if (!fs.existsSync(file)) return res.status(404).send('File not found');
  res.download(file, 'Collection Sheet.xlsx');
});

app.get('/download-gst/:file', (req, res) => {
  const file = path.join(OUTPUT_DIR, req.params.file);
  if (!fs.existsSync(file)) return res.status(404).send('File not found');
  res.download(file, 'GST Calculation Workbook.xlsx');
});

app.get('/download-updated-collection/:file', (req, res) => {
  const file = path.join(OUTPUT_DIR, req.params.file);
  if (!fs.existsSync(file)) return res.status(404).send('File not found');
  res.download(file, 'Collection Sheet (Updated).xlsx');
});

app.get('/download-updated-gst/:file', (req, res) => {
  const file = path.join(OUTPUT_DIR, req.params.file);
  if (!fs.existsSync(file)) return res.status(404).send('File not found');
  res.download(file, 'GST Workbook (Updated).xlsx');
});


app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n  CA Office Automation`);
  console.log(`  Collection Sheet + GST Workbook`);
  console.log(`  Running live on port: ${PORT}`);
  console.log(`  Open: http://localhost:${PORT}\n`);
});

