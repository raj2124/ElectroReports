// HTTP-level checks of the review, PDF export and client-logo routes (Session 6C), against a real
// server process on a temporary database. No AI provider is called: the Anthropic key is blank,
// so the narrative takes the deterministic fallback.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const net = require('net');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { devEnv, createCaller } = require('./lib/dev-auth');
const { openPdf, pageCount, pageTextLines } = require('../src/pdfRaster');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});

const png = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.alloc(200, 7)]);
const dataUrl = (mime, buffer) => `data:${mime};base64,${buffer.toString('base64')}`;

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'review-routes-'));
  const dataFile = path.join(dir, 'reports.json');
  const store = createReportStore({ filePath: dataFile });
  const project = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' });
  const payload = buildPayload();
  const base = normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', clientId: 'Z-77', reportDate: '2026-09-14' } });
  const daily = store.addReport({ ...base, kind: 'daily', projectId: project.id, collectionDate: '2026-09-14' });
  const { report: final } = store.assembleFinal(project.id);
  const legacy = store.addReport(normalizeReportInput(payload)); // an old-style standalone report

  const port = await freePort();
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'src', 'server.js')], {
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: dataFile, ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ANTHROPIC_API_KEY: '', OPENAI_API_KEY: '', GEMINI_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    server.on('exit', (code) => reject(new Error(`server exited early (${code})`)));
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });

  const base_url = `http://127.0.0.1:${port}`;
  const call = createCaller(base_url); // signs in as the admin unless a user is named
  const pdfText = async (url) => {
    const bytes = await fs.promises.readFile(path.join(dir, 'pdfs', path.basename(url)));
    assert.strictEqual(bytes.subarray(0, 5).toString(), '%PDF-');
    const pdf = await openPdf(bytes);
    const pages = [];
    for (let i = 0; i < pageCount(pdf); i += 1) pages.push(pageTextLines(pdf, i).map((f) => f.text).join(' '));
    return pages;
  };

  try {
    await check('an old saved (standalone) report still exports through the legacy layout; it has no review workflow', async () => {
      const exported = await call('POST', `/api/reports/${legacy.id}/pdf`);
      assert.strictEqual(exported.status, 200);
      const pages = await pdfText(exported.body.pdfUrl);
      const text = pages.join(' ');
      assert.ok(text.includes('Survey Overview') && !text.includes('Document Control & Distribution'));
      assert.strictEqual((await call('GET', `/api/reports/${legacy.id}/review`)).status, 400);
      assert.strictEqual((await call('POST', `/api/reports/${legacy.id}/review/submit`, {})).status, 400);
    });

    await check('daily report exports with the new renderer: one page, document number DPR_01', async () => {
      const exported = await call('POST', `/api/reports/${daily.id}/pdf`);
      assert.strictEqual(exported.status, 200);
      const pages = await pdfText(exported.body.pdfUrl);
      assert.strictEqual(pages.length, 1);
      assert.ok(pages[0].includes('Daily Progress Report') && pages[0].includes('024_FI_ESHA_DPR_01'));
    });

    await check('review state: checklist, sanity check not run, warnings', async () => {
      const { status, body } = await call('GET', `/api/reports/${final.id}/review`);
      assert.strictEqual(status, 200);
      assert.strictEqual(body.status, 'draft');
      assert.strictEqual(body.sanityCheck, 'not run');
      assert.ok(body.checklist.length === 5 && body.consistency.ok);
      assert.ok(body.warnings.some((warning) => /Planned quantity not set/.test(warning)));
    });

    await check('steps are recorded against the signed-in user; check needs ticks; approval needs the check and the acknowledgement', async () => {
      // admin = H. S. Patil prepares, reviewer = A. P. checks, approver = J. C. Shah approves
      const who = { submit: 'admin@elegrow.com', check: 'reviewer@elegrow.com', approve: 'approver@elegrow.com' };
      const step = (name, body) => call('POST', `/api/reports/${final.id}/review/${name}`, body, who[name]);
      assert.strictEqual((await call('POST', `/api/reports/${final.id}/review/submit`, {}, 'engineer@elegrow.com')).status, 403, 'an engineer cannot submit a report that is not theirs');
      assert.strictEqual((await step('submit', {})).status, 200);
      assert.strictEqual((await step('submit', {})).status, 409, 'cannot submit twice');
      const early = await step('check', {});
      assert.deepStrictEqual([early.status, early.body.code], [409, 'CHECKLIST_INCOMPLETE']);
      const notChecked = await step('approve', {});
      assert.deepStrictEqual([notChecked.status, notChecked.body.code], [409, 'NOT_CHECKED']);
      const ticks = Object.fromEntries((await call('GET', `/api/reports/${final.id}/review`)).body.checklist.map((item) => [item.id, true]));
      assert.strictEqual((await call('PUT', `/api/reports/${final.id}/review/ticks`, { ticks }, 'reviewer@elegrow.com')).status, 200);
      const checked = await step('check', { note: 'ok' });
      assert.strictEqual(checked.status, 200);
      assert.strictEqual(checked.body.checked.by, 'A. P.');
      const noAck = await step('approve', {});
      assert.deepStrictEqual([noAck.status, noAck.body.code], [409, 'ACKNOWLEDGEMENT_REQUIRED']);
      assert.ok(/Sanity check not run/.test(noAck.body.message) && /No photo reference for P1 pits/.test(noAck.body.message));
      const approved = await step('approve', { acknowledgements: { sanityCheck: 'Provider unavailable; narrative read by the reviewer.', p1Photos: 'Photos to follow in Annexure B.' } });
      assert.strictEqual(approved.status, 200);
      assert.strictEqual(approved.body.status, 'approved');
      assert.deepStrictEqual(approved.body.initials, { prepared: 'HSP', checked: 'AP', approved: 'JCS' });
      assert.deepStrictEqual(approved.body.transitions.map((row) => row.action), ['submit', 'check', 'approve']);
      assert.strictEqual(approved.body.transitions[2].designation, 'Director', 'the name and designation come from the profile');
    });

    await check('final export: deterministic narrative note without Claude, cover initials, approved (no draft marker); reused on re-export', async () => {
      const first = await call('POST', `/api/reports/${final.id}/pdf`);
      assert.strictEqual(first.status, 200, JSON.stringify(first.body));
      assert.strictEqual(first.body.narrative.aiWritten, false);
      const pages = await pdfText(first.body.pdfUrl);
      const text = pages.join(' ');
      assert.ok(text.includes('Narrative not AI-written') && text.includes('HSP') && text.includes('JCS') && text.includes('024_FI_ESHA'));
      assert.ok(!text.includes('DRAFT – NOT APPROVED'));
      const stored = (await call('GET', `/api/reports/${final.id}`)).body.reportNarrative;
      assert.ok(stored && stored.aiWritten === false && stored.generatedAt, 'the narrative used is kept with the report');
      await call('POST', `/api/reports/${final.id}/pdf`);
      assert.strictEqual((await call('GET', `/api/reports/${final.id}`)).body.reportNarrative.generatedAt, stored.generatedAt, 'an approved report keeps the narrative that was reviewed');
    });

    await check('re-issue over HTTP creates R1 as a draft', async () => {
      assert.strictEqual((await call('POST', `/api/reports/${final.id}/reissue`, {}, 'engineer@elegrow.com')).status, 403, 'only an approver re-issues');
      const reissued = await call('POST', `/api/reports/${final.id}/reissue`, { reason: 'Client comments' }, 'approver@elegrow.com');
      assert.strictEqual(reissued.status, 201);
      assert.deepStrictEqual([reissued.body.revision, reissued.body.reviewStatus, reissued.body.revisionOf], [1, 'draft', final.id]);
      assert.strictEqual((await call('POST', `/api/reports/${final.id}/reissue`, {}, 'approver@elegrow.com')).body.code, 'ALREADY_REISSUED');
      const exported = await call('POST', `/api/reports/${reissued.body.id}/pdf`);
      const text = (await pdfText(exported.body.pdfUrl)).join(' ');
      assert.ok(text.includes('R1') && text.includes('Client comments') && text.includes('DRAFT – NOT APPROVED'));
    });

    await check('client logo: validated on upload, stored per client, served back, deleted; report override validated too', async () => {
      assert.strictEqual((await call('PUT', '/api/client-logo', { clientName: 'Acme', dataUrl: dataUrl('image/gif', png) })).status, 400);
      assert.strictEqual((await call('PUT', '/api/client-logo', { clientName: 'Acme', dataUrl: dataUrl('image/png', Buffer.from('nope')) })).status, 400);
      const big = await call('PUT', '/api/client-logo', { clientName: 'Acme', dataUrl: dataUrl('image/png', Buffer.concat([png, Buffer.alloc(1024 * 1024)])) });
      assert.strictEqual(big.status, 400);
      assert.ok(/1 MB/.test(big.body.message));
      assert.strictEqual((await call('PUT', '/api/client-logo', { dataUrl: dataUrl('image/png', png) })).status, 400, 'needs a client');
      assert.strictEqual((await call('GET', '/api/client-logo?clientName=Acme')).status, 404);
      assert.strictEqual((await call('PUT', '/api/client-logo', { clientId: 'Z-9', clientName: 'Acme', dataUrl: dataUrl('image/png', png) })).status, 200);
      const served = await call('GET', '/api/client-logo?clientName=ACME');
      assert.strictEqual(served.status, 200);
      assert.strictEqual(served.type, 'image/png');
      assert.ok(served.body.equals(png));
      assert.strictEqual((await call('GET', '/api/client-logo?clientId=Z-9')).status, 200);
      const override = await call('PUT', `/api/reports/${daily.id}/client-logo`, { dataUrl: dataUrl('image/png', png) });
      assert.deepStrictEqual([override.status, override.body.hasOverride], [200, true]);
      assert.strictEqual((await call('PUT', `/api/reports/${daily.id}/client-logo`, { dataUrl: dataUrl('image/png', Buffer.from('x')) })).status, 400);
      assert.strictEqual((await call('PUT', `/api/reports/${daily.id}/client-logo`, { dataUrl: null })).body.hasOverride, false);
      assert.strictEqual((await call('DELETE', '/api/client-logo?clientId=Z-9&clientName=Acme')).status, 204);
      assert.strictEqual((await call('GET', '/api/client-logo?clientName=Acme')).status, 404);
    });
  } finally {
    server.kill();
  }
  console.log(`\n${passed} review-route checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
