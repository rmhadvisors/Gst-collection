'use strict';

/**
 * Web upload as a background job: the POST answers at once with a job id, the
 * page polls /job/:id for progress, and the finished workbook downloads.
 * No OCR - runs in seconds.
 */

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const S = (f) => path.join(__dirname, '..', 'config', 'samples', f);

let passed = 0;
let failed = 0;
async function check(name, fn) {
  try { await fn(); passed += 1; console.log(`  ok  - ${name}`); } catch (e) {
    failed += 1; console.log(`  FAIL- ${name}\n       ${e.message}`);
  }
}

(async () => {
  console.log('Web upload: background job with progress, no timeout');
  await runServerTests();
  console.log(`\n${passed} web-job checks passed${failed ? `, ${failed} FAILED` : ''}`);
  process.exit(failed ? 1 : 0);
})();

/** A TCP port that is free right now. */
function freePort() {
  return new Promise((resolve, reject) => {
    const srv = require('net').createServer();
    srv.once('error', reject);
    srv.listen(0, '127.0.0.1', () => { const { port } = srv.address(); srv.close(() => resolve(port)); });
  });
}

async function runServerTests() {
  const port = await freePort(); // the OS picks one nobody else is using
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'server.js')], {
    env: { ...process.env, PORT: String(port) }, stdio: ['ignore', 'pipe', 'pipe'],
  });
  let log = '';
  server.stdout.on('data', (d) => { log += d; });
  server.stderr.on('data', (d) => { log += d; });
  try {
    for (let i = 0; i < 60; i += 1) {
      try { if ((await fetch(base)).ok) break; } catch (_) { await new Promise((r) => setTimeout(r, 500)); }
      if (i === 59) throw new Error(`server did not start:\n${log}`);
    }

    const form = () => {
      const fd = new FormData();
      fd.append('tally', new Blob([fs.readFileSync(S('mahira-federal-bank.xlsx'))]), 'fed.xlsx');
      return fd;
    };

    await check('POST answers immediately with 202 + job id; progress is reported; result downloads', async () => {
      const res = await fetch(`${base}/process-gst`, { method: 'POST', body: form() });
      assert.strictEqual(res.status, 202);
      assert.ok((res.headers.get('content-type') || '').includes('application/json'));
      const { jobId, statusUrl } = await res.json();
      assert.ok(jobId && statusUrl.endsWith(`/job/${jobId}`));
      const seen = new Set();
      let job;
      for (let i = 0; i < 120; i += 1) {
        job = await (await fetch(`${base}${statusUrl}`)).json();
        seen.add(job.status);
        if (job.status === 'done' || job.status === 'failed') break;
        await new Promise((r) => setTimeout(r, 250));
      }
      assert.strictEqual(job.status, 'done', job.error);
      assert.ok(job.resultUrl);
      const page = await fetch(`${base}${job.resultUrl}`);
      assert.strictEqual(page.status, 200);
      const html = await page.text();
      const m = html.match(/download-gst\/([^"']+\.xlsx)/);
      assert.ok(m, 'result page links the workbook');
      const dl = await fetch(`${base}/download-gst/${m[1]}`);
      assert.strictEqual(dl.status, 200);
      assert.ok((await dl.arrayBuffer()).byteLength > 1000);
    });
    await check('a rejected upload is answered at once with an HTML error (no job)', async () => {
      const res = await fetch(`${base}/process`, { method: 'POST', body: form() }); // no agreement PDF
      assert.strictEqual(res.status, 200);
      assert.ok((res.headers.get('content-type') || '').includes('text/html'));
      assert.ok(/Please upload at least one Agreement PDF/.test(await res.text()));
    });
    await check('unknown job id is a 404', async () => {
      assert.strictEqual((await fetch(`${base}/job/does-not-exist`)).status, 404);
    });
    await check('a broken upload always reaches done or failed (never hangs)', async () => {
      const fd = new FormData();
      fd.append('tally', new Blob([Buffer.from('not really an xlsx')]), 'broken.xlsx');
      const { statusUrl } = await (await fetch(`${base}/process-gst`, { method: 'POST', body: fd })).json();
      let job;
      for (let i = 0; i < 80; i += 1) {
        job = await (await fetch(`${base}${statusUrl}`)).json();
        if (job.status === 'done' || job.status === 'failed') break;
        await new Promise((r) => setTimeout(r, 250));
      }
      assert.ok(job.status === 'failed' || job.status === 'done', job.status);
    });
  } finally {
    server.kill();
  }
  await timeLimitTest();
}

/** A job that runs past its limit is FAILED with the reason, and the server carries on. */
async function timeLimitTest() {
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'server.js')], {
    env: { ...process.env, PORT: String(port), JOB_TIMEOUT_MS: '3000' }, stdio: ['ignore', 'pipe', 'pipe'],
  });
  let log = '';
  server.stdout.on('data', (d) => { log += d; });
  server.stderr.on('data', (d) => { log += d; });
  const poll = async (statusUrl) => {
    let job;
    for (let i = 0; i < 120; i += 1) {
      job = await (await fetch(`${base}${statusUrl}`)).json();
      if (job.status === 'done' || job.status === 'failed') break;
      await new Promise((r) => setTimeout(r, 250));
    }
    return job;
  };
  try {
    for (let i = 0; i < 60; i += 1) {
      try { if ((await fetch(base)).ok) break; } catch (_) { await new Promise((r) => setTimeout(r, 500)); }
    }
    await check('a job past its time limit ends as "failed" with the reason (it never stays running)', async () => {
      const fd = new FormData();
      fd.append('tally', new Blob([fs.readFileSync(S('mahira-federal-bank.xlsx'))]), 'fed.xlsx');
      // the agreement PDF takes minutes of OCR: far longer than the 3 s limit
      fd.append('agreements', new Blob([fs.readFileSync(path.join(__dirname, 'fixtures', 'mahira-tower', 'B-603_1.pdf'))]), 'B-603_1.pdf');
      const { statusUrl, jobId } = await (await fetch(`${base}/process-gst`, { method: 'POST', body: fd })).json();
      const job = await poll(statusUrl);
      assert.strictEqual(job.status, 'failed');
      assert.ok(/timed out/.test(job.error), job.error);
      assert.ok(log.includes(`[job ${jobId.slice(0, 8)}] FAILED`), 'the failure is logged with the job id');
    });
    await check('the next job still runs normally after a failed one', async () => {
      const fd = new FormData();
      fd.append('tally', new Blob([fs.readFileSync(S('mahira-federal-bank.xlsx'))]), 'fed.xlsx');
      const { statusUrl } = await (await fetch(`${base}/process-gst`, { method: 'POST', body: fd })).json();
      const job = await poll(statusUrl);
      assert.strictEqual(job.status, 'done', job.error);
    });
    await check('every step of a job is logged with its job id, including the OCR worker state', async () => {
      assert.ok(/\[job [0-9a-f]{8}\] queued/.test(log), 'queued line');
      assert.ok(/\[job [0-9a-f]{8}\] started \(ocr worker/.test(log), 'started line');
      assert.ok(/\[job [0-9a-f]{8}\] ocr worker (created|busy|terminated)/.test(log), 'ocr worker state line');
    });
  } finally {
    server.kill();
  }
}
