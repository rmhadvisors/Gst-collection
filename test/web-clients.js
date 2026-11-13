'use strict';

/**
 * Web upload, client by client: start the real server, upload each client's
 * fixture files over HTTP, follow the background job, download the Excel and
 * check it is not empty and carries the numbers in expected.json.
 *
 *   node test/web-clients.js            fixtures only
 *   node test/web-clients.js --big      also uploads the 43 MB Space Home 1804 agreement
 */

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const ExcelJS = require('exceljs');
const { clientDirs } = require('./clients');

const BIG = process.argv.includes('--big');
const ONLY = (process.argv.find((a) => a.startsWith('--only=')) || '').slice(7);
// The server applies the production config/overrides.json. Its extraPayments for a
// client's flats are receipts the CA typed in, so they belong in the expected total.
const PROD = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config', 'overrides.json'), 'utf8'));
const extraFor = (flats) => flats.reduce(
  (s, k) => s + ((PROD[k] && PROD[k].extraPayments) || []).reduce((t, p) => t + p.amount, 0), 0,
);
const BIG_PDF = path.join(__dirname, '..', 'config', 'samples', '1804-CHHAYA KADAM.pdf');
const JOB_TIMEOUT_MS = 15 * 60 * 1000;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
// WEB_TEST_CONNECTION_CLOSE=1: no keep-alive sockets (used to rule keep-alive in or out)
const CONN = process.env.WEB_TEST_CONNECTION_CLOSE === '1' ? { connection: 'close' } : {};
const money = (n) => Number(n).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

/** A TCP port that is free right now. */
function freePort() {
  return new Promise((resolve, reject) => {
    const srv = require('net').createServer();
    srv.once('error', reject);
    srv.listen(0, '127.0.0.1', () => { const { port } = srv.address(); srv.close(() => resolve(port)); });
  });
}

async function startServer() {
  const port = await freePort(); // the OS picks one nobody else is using
  const base = `http://127.0.0.1:${port}`;
  const proc = spawn(process.execPath, [path.join(__dirname, '..', 'server.js')], {
    env: { ...process.env, PORT: String(port) }, stdio: ['ignore', 'pipe', 'pipe'],
  });
  const state = { log: '', exit: null };
  proc.stdout.on('data', (d) => { state.log += d; });
  proc.stderr.on('data', (d) => { state.log += d; });
  proc.on('exit', (code, signal) => { state.exit = `exited code=${code} signal=${signal}`; });
  for (let i = 0; i < 60; i += 1) {
    try { if ((await fetch(base, { headers: CONN })).ok) return { base, proc, state }; } catch (_) { await sleep(500); }
  }
  proc.kill();
  throw new Error(`server did not start:\n${state.log}`);
}

/** Upload, follow the job, return { seconds, messages, file (Buffer), kind }. */
async function uploadAndDownload(base, route, files) {
  const fd = new FormData();
  for (const [field, filePath, name] of files) {
    fd.append(field, new Blob([fs.readFileSync(filePath)]), name || path.basename(filePath));
  }
  const started = Date.now();
  const res = await fetch(`${base}${route}`, { method: 'POST', body: fd, headers: CONN });
  assert.strictEqual(res.status, 202, `${route} answered ${res.status} (expected 202 + job id)`);
  const answeredAfter = Date.now() - started;
  const { statusUrl } = await res.json();

  const messages = [];
  let job;
  while (Date.now() - started < JOB_TIMEOUT_MS) {
    job = await (await fetch(`${base}${statusUrl}`, { headers: CONN })).json();
    if (messages[messages.length - 1] !== job.message) messages.push(job.message);
    if (job.status === 'done' || job.status === 'failed') break;
    await sleep(1000);
  }
  assert.strictEqual(job.status, 'done', `job ${job.status}: ${job.error || 'timed out'}`);
  const html = await (await fetch(`${base}${job.resultUrl}`, { headers: CONN })).text();
  const m = html.match(/href="[^"]*\/(download(?:-gst)?)\/([^"]+\.xlsx)"/);
  assert.ok(m, 'result page has no Excel download link');
  const dl = await fetch(`${base}/${m[1]}/${m[2]}`, { headers: CONN });
  assert.strictEqual(dl.status, 200);
  return {
    seconds: Math.round((Date.now() - started) / 1000),
    answeredMs: answeredAfter,
    messages,
    kind: m[1] === 'download' ? 'collection' : 'gst',
    file: Buffer.from(await dl.arrayBuffer()),
  };
}

/** All cell texts of a workbook, plus its non-empty row count. */
async function inspect(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  let rows = 0;
  const text = new Set();
  const sheets = wb.worksheets.map((ws) => {
    let n = 0;
    ws.eachRow((row) => {
      n += 1;
      row.eachCell((c) => { if (c.value != null && c.value !== '') text.add(String(c.value.result != null ? c.value.result : c.value)); });
    });
    rows += n;
    return { name: ws.name, rows: n };
  });
  return { wb, sheets, rows, text };
}

(async () => {
  const { base, proc, state } = await startServer();
  const results = [];
  let failed = 0;
  const report = (client, route, ok, detail) => {
    if (!ok) failed += 1;
    results.push({ client, route, ok, detail });
    console.log(`  ${ok ? 'ok  ' : 'FAIL'}- ${client} ${route}: ${detail}`);
  };

  try {
    for (const dir of clientDirs()) {
      if (ONLY && path.basename(dir) !== ONLY) continue;
      const spec = JSON.parse(fs.readFileSync(path.join(dir, 'expected.json'), 'utf8'));
      const exp = spec.expected;
      const tally = spec.inputs.tally.map((n) => ['tally', path.join(dir, n)]);
      const agreements = (spec.inputs.agreements || []).map((n) => ['agreements', path.join(dir, n)]);
      const flatCount = Object.keys(exp.flats).length;
      console.log(`\n${spec.client}`);

      const runs = [['/process-gst', [...tally, ...agreements]]];
      if (agreements.length) runs.unshift(['/process', [...tally, ...agreements]]);
      if (BIG && path.basename(dir) === 'space-home') runs.push(['/process', [...tally, ['agreements', BIG_PDF]]]);

      for (const [route, files] of runs) {
        const big = files.some(([, f]) => f === BIG_PDF);
        const label = `${route}${big ? ' (43 MB agreement)' : ''}`;
        try {
          const r = await uploadAndDownload(base, route, files);
          const x = await inspect(r.file);
          assert.ok(r.file.length > 5000, `file is only ${r.file.length} bytes`);
          assert.ok(x.rows > flatCount, `only ${x.rows} rows for ${flatCount} flats`);
          // every flat of the client appears in the workbook
          const missing = Object.keys(exp.flats).filter((k) => !x.text.has(k));
          assert.deepStrictEqual(missing, [], `flats missing from the Excel: ${missing.join(', ')}`);

          let extra = '';
          if (r.kind === 'collection') {
            assert.deepStrictEqual(x.sheets.map((s) => s.name), exp.sheets);
            const audit = x.wb.getWorksheet('Audit');
            let check = '';
            audit.eachRow((row) => { if (row.getCell(1).value === 'Check') check = String(row.getCell(4).value); });
            assert.strictEqual(check, 'MATCHED', `Audit check says "${check}"`);
            // Summary: sum of 'Total received' over the flats equals the sum in expected.json
            const sum = Object.values(exp.flats).reduce((s, f) => s + f.totalReceived, 0)
              + extraFor(Object.keys(exp.flats));
            let got = 0;
            x.wb.getWorksheet('Summary').eachRow((row) => {
              if (typeof row.getCell(1).value === 'number') got += Number(row.getCell(10).value) || 0;
            });
            assert.ok(Math.abs(got - sum) < 0.01, `Summary received ${money(got)} != expected ${money(sum)}`);
            extra = ` | Summary received ${money(got)} | Audit: ${check}`;
          }
          report(spec.client, label, true,
            `${(r.file.length / 1024).toFixed(0)} KB, ${x.rows} rows (${x.sheets.map((s) => `${s.name}:${s.rows}`).join(', ')}) | ${flatCount} flats all present${extra} | POST answered in ${r.answeredMs} ms, job ${r.seconds}s | progress: ${r.messages.join(' > ')}`);
        } catch (e) {
          // say what the server itself was doing when the upload failed
          const tail = state.log.split('\n')
            .filter((l) => l && !/Estimating|too small|cannot be rec|diacritics/.test(l)).slice(-12).join(' | ');
          const cause = e.cause ? ` (cause: ${e.cause.code || ''} ${e.cause.message || e.cause})` : '';
          report(spec.client, label, false, `${e.message}${cause} [server ${state.exit || 'still running'}; log tail: ${tail || '(empty)'}]`);
        }
      }
    }
  } finally {
    proc.kill();
  }

  console.log(`\n${results.filter((r) => r.ok).length} of ${results.length} web uploads produced a valid, non-empty Excel`);
  process.exit(failed ? 1 : 0);
})();
