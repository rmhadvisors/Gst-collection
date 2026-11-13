'use strict';

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const {
  assertOcrReady, closeOcr, setOcrLogger, ocrState,
} = require('./src/pdf/ocr');
const { withTimeout } = require('./src/util/timeout');
const { buildRecords, prepareTally, loadOverrides } = require('./src/services/build-records');
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
    else if (file.fieldname === 'agreements' && ['.pdf', '.jpg', '.jpeg', '.png'].includes(ext)) cb(null, true); // PDF or a photo of an Index-II page
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

/**
 * Background jobs. A large agreement PDF takes minutes of OCR, so the upload
 * request returns at once with a job id; the page polls /job/:id for progress and
 * opens the result when it is done. Jobs run one at a time because the Tesseract
 * worker is shared.
 */
const jobs = new Map();
let jobQueue = Promise.resolve();
let waiting = 0;
const JOB_TTL_MS = 60 * 60 * 1000;
const JOB_TIMEOUT_MS = Number(process.env.JOB_TIMEOUT_MS) || 30 * 60 * 1000; // a job that runs longer is failed

/** One log line per step, always with the job id (first 8 characters). */
const jlog = (job, msg) => console.log(`${new Date().toISOString()} [job ${job.id.slice(0, 8)}] ${msg}`);

// A crash anywhere (e.g. inside an OCR worker thread) is logged instead of silently killing a job.
process.on('uncaughtException', (e) => console.error(`${new Date().toISOString()} UNCAUGHT EXCEPTION: ${e.stack || e}`));
process.on('unhandledRejection', (e) => console.error(`${new Date().toISOString()} UNHANDLED REJECTION: ${(e && e.stack) || e}`));

setInterval(() => {
  const cutoff = Date.now() - JOB_TTL_MS;
  for (const [id, job] of jobs) if (job.finished && job.finished < cutoff) jobs.delete(id);
}, 10 * 60 * 1000).unref();

function removeFiles(files) {
  for (const f of files) {
    try { fs.unlinkSync(f.path); } catch (_) { /* already gone */ }
  }
}

/** Queue `work(progress)`; it returns the locals for the result page. Answers 202 + job id. */
function startJob(req, res, files, work) {
  const id = crypto.randomUUID();
  const job = {
    id, status: 'queued', message: 'Queued', started: Date.now(), finished: null,
  };
  jobs.set(id, job);
  waiting += 1;
  if (waiting > 1) job.message = `Queued behind ${waiting - 1} other job(s)`;
  jlog(job, `queued: ${req.path}, ${files.length} file(s), ${waiting - 1} job(s) ahead`);
  jobQueue = jobQueue.then(async () => {
    waiting -= 1;
    job.status = 'running';
    job.message = 'Starting';
    setOcrLogger((m) => jlog(job, m)); // OCR worker created / busy / idle / terminated, with this job's id
    jlog(job, `started (ocr worker ${ocrState()})`);
    try {
      job.locals = await withTimeout(work((msg) => { job.message = msg; jlog(job, msg); }), JOB_TIMEOUT_MS, 'Job');
      job.status = 'done';
      job.message = 'Done';
    } catch (e) {
      job.status = 'failed';
      job.error = e.message || String(e);
      job.message = 'Failed';
      jlog(job, `FAILED: ${e.stack || e.message}`);
    } finally {
      job.finished = Date.now();
      removeFiles(files);
      // whatever happened, no OCR worker may outlive its job (a leftover one could stall the next)
      try { await closeOcr(); } catch (_) { /* logged by the ocr module */ }
      jlog(job, `${job.status} after ${Math.round((job.finished - job.started) / 1000)}s (ocr worker ${ocrState()})`);
      setOcrLogger(null);
    }
  });
  res.status(202).json({ jobId: id, statusUrl: `${res.locals.basePath}/job/${id}` });
}

app.get('/job/:id', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job) return res.status(404).json({ status: 'failed', error: 'Unknown or expired job' });
  res.json({
    jobId: job.id,
    status: job.status,
    message: job.message,
    error: job.error || null,
    elapsedSec: Math.round(((job.finished || Date.now()) - job.started) / 1000),
    resultUrl: job.status === 'done' ? `${res.locals.basePath}/job/${job.id}/result` : null,
  });
});

app.get('/job/:id/result', (req, res) => {
  const job = jobs.get(req.params.id);
  if (!job || job.status !== 'done') return res.redirect(`${res.locals.basePath}/`);
  res.render('result', job.locals);
});

const allFiles = (req) => Object.values(req.files || {}).flat();

/** Shared by the four POST routes: validate the upload, then run `work` as a job. */
function jobRoute(route, check, work) {
  app.post(route, (req, res) => {
    uploadFields(req, res, async (err) => {
      if (err) return res.render('index', { error: err.message });
      const files = {
        tally: req.files.tally || [],
        agreements: req.files.agreements || [],
        existing: req.files.existingWorkbook || [],
      };
      const fail = (msg) => { removeFiles(allFiles(req)); return res.render('index', { error: msg }); };
      const problem = check(files);
      if (problem) return fail(problem);
      let overrides;
      try { overrides = loadOverrides(); } catch (e) { return fail(e.message); }

      return startJob(req, res, allFiles(req), async (progress) => {
        const workDir = path.join(OUTPUT_DIR, '_pages_' + Date.now() + '_' + Math.round(Math.random() * 1e4));
        fs.mkdirSync(workDir, { recursive: true });
        try {
          return await work({ files, overrides, workDir, progress });
        } finally {
          fs.rmSync(workDir, { recursive: true, force: true });
        }
      });
    });
  });
}

// ─── Routes ──────────────────────────────────────────────

app.get('/', (req, res) => {
  res.render('index');
});

// ── Collection Sheet ──────────────────────────────────────
jobRoute('/process', (f) => {
  if (!f.tally.length) return 'Please upload at least one Tally receipts file (.xlsx)';
  if (!f.agreements.length) return 'Please upload at least one Agreement (PDF or photo)';
  return null;
}, async ({
  files, overrides, workDir, progress,
}) => {
  const {
    records, tally, tallyTotal, tallyFiles,
  } = await buildRecords({
    tallyFiles: files.tally, agreementFiles: files.agreements, overrides, workDir, onProgress: progress,
  });
  progress('Writing the workbook');
  const outFilename = `collection-${Date.now()}.xlsx`;
  await writeStyledCollectionWorkbook(records, path.join(OUTPUT_DIR, outFilename), {
    projectName: tally.projectName || undefined, reconciliation: { tallyTotal, files: tallyFiles },
  });
  return {
    records, fmtINR, downloadFile: outFilename, gstDownloadFile: null,
    projectName: tally.projectName || 'Collection Sheet',
  };
});

// ── GST Calculation Workbook ──────────────────────────────
// Agreement PDFs are OPTIONAL here (tally-only -> all flats treated as Case 3).
jobRoute('/process-gst', (f) => (
  f.tally.length ? null : 'Please upload at least one Tally receipts file (.xlsx)'
), async ({
  files, overrides, workDir, progress,
}) => {
  const { records, tally } = await buildRecords({
    tallyFiles: files.tally, agreementFiles: files.agreements, overrides, workDir, onProgress: progress,
  });
  progress('Writing the workbook');
  const outFilename = `gst-workbook-${Date.now()}.xlsx`;
  writeGstWorkbook(records,path.join(OUTPUT_DIR, outFilename), { projectName: tally.projectName || undefined });
  return {
    records, fmtINR, downloadFile: null, gstDownloadFile: outFilename,
    projectName: tally.projectName || 'GST Workbook',
  };
});

// ── Update Collection Sheet ───────────────────────────────
jobRoute('/update-collection', (f) => {
  if (!f.tally.length) return 'Please upload the latest Tally receipts file.';
  if (!f.existing.length) return 'Please upload your existing Collection Sheet.';
  return null;
}, async ({
  files, overrides, workDir, progress,
}) => {
  progress('Reading Tally files');
  const tally = prepareTally(files.tally, overrides);
  const outFilename = `updated-collection-${Date.now()}.xlsx`;
  progress(`Reading ${files.agreements.length} agreement PDF(s) and updating the sheet`);
  const result = await updateCollectionSheet({
    existingPath: files.existing[0].path,
    outPath: path.join(OUTPUT_DIR, outFilename),
    tally,
    agreementPaths: files.agreements.map((f) => f.path), // PDFs are OCR'd inside the service
    overrides,
    workDir,
    projectName: tally.projectName || 'YASHWANTH COUNTY',
  });
  return {
    records: [],
    fmtINR,
    downloadFile: null,
    gstDownloadFile: null,
    updatedCollectionFile: outFilename,
    updatedGstFile: null,
    updateSummary: {
      newPayments: result.newPayments,
      reconciledFlats: result.reconciledFlats,
      newFlats: result.newFlats,
      pdfWarnings: result.pdfWarnings,
      newMonths: [],
      updatedMonths: [],
    },
    projectName: tally.projectName || 'Collection Sheet (Updated)',
  };
});

// ── Update GST Workbook ───────────────────────────────────
jobRoute('/update-gst', (f) => {
  if (!f.tally.length) return 'Please upload the latest Tally receipts file.';
  if (!f.existing.length) return 'Please upload your existing GST Workbook.';
  return null;
}, async ({
  files, overrides, workDir, progress,
}) => {
  const { records } = await buildRecords({
    tallyFiles: files.tally, agreementFiles: files.agreements, overrides, workDir, onProgress: progress,
  });
  progress('Updating the GST workbook');
  const outFilename = `updated-gst-${Date.now()}.xlsx`;
  const result = await updateGstWorkbook({
    existingPath: files.existing[0].path,
    outPath: path.join(OUTPUT_DIR, outFilename),
    records,
    projectName: 'SPACE HOME',
  });
  return {
    records: [],
    fmtINR,
    downloadFile: null,
    gstDownloadFile: null,
    updatedCollectionFile: null,
    updatedGstFile: outFilename,
    updateSummary: {
      newPayments: 0,
      reconciledFlats: [],
      newMonths: result.newMonths,
      updatedMonths: result.updatedMonths,
    },
    projectName: 'GST Workbook (Updated)',
  };
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


try {
  assertOcrReady();
} catch (e) {
  console.error(`
  STARTUP ERROR: ${e.message}
`);
  process.exit(1);
}

const server = app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n  CA Office Automation`);
  console.log(`  Collection Sheet + GST Workbook`);
  console.log(`  Running live on port: ${PORT}`);
  console.log(`  Open: http://localhost:${PORT}\n`);
});
// Uploads of many large scanned PDFs can take longer than Node's default 5-minute
// request timeout; processing itself runs in a background job (see startJob).
server.requestTimeout = 0;
server.keepAliveTimeout = 65 * 1000;
server.headersTimeout = 66 * 1000; // must exceed keepAliveTimeout

