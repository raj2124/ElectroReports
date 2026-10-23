// Session 6I over HTTP on the mocks: who may upload, read and change a scan, the upload -> read -> edit -> accept flow with the
// overwrite confirmation, a failing provider and its retry, discard, and the admin purge. The reader is the mock; its "reading" is
// a truth file written here, and the pages are real printed sheets (QR codes included) filled in by the test.
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { devEnv, createCaller, login } = require('./lib/dev-auth');
const { fillIssued, renderSheetsPdf } = require('./lib/scan-fixtures');

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
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'scan-routes-'));
const truthFile = path.join(dir, 'truth.json');
const setTruth = (truth) => fs.writeFileSync(truthFile, JSON.stringify(truth));
const PROJECT = { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' };
const ROWS = [
  { tag: 'P-1', location: 'Yard', electrodeType: 'Rod', materialType: 'GI', length: '3', diameter: '20', resistanceWithGrid: '0.52', method: '61.8%', conditionCodes: ['C1'] },
  { tag: 'P-2', location: 'Gate', electrodeType: 'Plate', materialType: 'Cu', length: '0.6', diameter: '3', resistanceWithGrid: '1.20', method: '61.8%' }
];

(async () => {
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: path.join(__dirname, '..'),
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', GRAPH_USE_MOCK: 'true', OCR_USE_MOCK: 'true', OCR_MOCK_TRUTH_FILE: truthFile, AI_CHECKS_USE_MOCK: 'true', GEMINI_API_KEY: '', OPENAI_API_KEY: '', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  try {
    await new Promise((resolve, reject) => {
      server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
      setTimeout(() => reject(new Error('server did not start')), 15000);
    });
    const call = createCaller(base);
    const cookies = {};
    const asCookie = async (email) => (cookies[email] = cookies[email] || (await login(base, email)));
    const upload = async (files, as = 'engineer@elegrow.com') => {
      const form = new FormData();
      files.forEach(({ name, type, bytes }) => form.append('files', new Blob([bytes], { type }), name));
      const response = await fetch(`${base}/api/scans`, { method: 'POST', body: form, headers: as ? { Cookie: await asCookie(as) } : {} });
      return { status: response.status, body: await response.json().catch(() => ({})) };
    };
    const settle = async (as = 'engineer@elegrow.com') => {
      for (let i = 0; i < 80; i += 1) {
        const { scans } = (await call('GET', '/api/scans', undefined, as)).body;
        if (!scans.some((scan) => ['queued', 'reading'].includes(scan.status))) return scans;
        await sleep(100);
      }
      throw new Error('scans did not settle');
    };

    // a workspace, and two real printed pages (electrode + loop) issued by the server
    const workspace = (await call('POST', '/api/projects', PROJECT, 'engineer@elegrow.com')).body.projectId;
    const issueOne = async (sheetKey, workDate) => {
      const response = await fetch(`${base}/api/sheets/print`, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: await asCookie('engineer@elegrow.com') }, body: JSON.stringify({ project: PROJECT, workDate, sheets: [sheetKey] }) });
      assert.strictEqual(response.status, 200);
      const [id] = response.headers.get('x-sheet-ids').split(',');
      return (await call('GET', `/api/sheets/issued/${id}`, undefined, 'engineer@elegrow.com')).body;
    };
    const electrode = await fillIssued(await issueOne('electrodeResistance', '2026-10-12'), { instrument: 'mi3290', serialNo: 'SN77', rows: ROWS });
    const loop = await fillIssued(await issueOne('loopImpedanceTest', '2026-10-12'), { rows: [{ location: 'MCC-1', feederTag: 'F-1', deviceType: 'MCCB', deviceRating: '63', loopImpedance: '0.31', voltage: '230' }, { loopImpedance: '0.33' }, { loopImpedance: '0.35' }] });
    const both = (await renderSheetsPdf([electrode.ctx, loop.ctx])).buffer;
    const truthFor = (sheet, extra = {}) => ({ primary: { ...sheet.truth, ...extra.primary }, ...(extra.escalation ? { escalation: { ...sheet.truth, ...extra.escalation } } : {}) });
    const baseTruth = () => ({ [electrode.issued.sheetId]: truthFor(electrode, { primary: { 'r0.resistanceWithGrid': { value: '0.5?', confidence: 0.6 } }, escalation: {} }), [loop.issued.sheetId]: truthFor(loop) });
    setTruth(baseTruth());

    await check('who may upload: signed-in engineers only; wrong files are refused with a clear message', async () => {
      const files = [{ name: 'visit.pdf', type: 'application/pdf', bytes: both }];
      assert.strictEqual((await upload(files, null)).status, 401);
      assert.strictEqual((await upload(files, 'approver@elegrow.com')).status, 403);
      assert.strictEqual((await upload(files, 'reviewer@elegrow.com')).status, 403);
      assert.strictEqual((await call('POST', '/api/scans', {}, 'engineer@elegrow.com')).status, 400);
      const text = await upload([{ name: 'notes.txt', type: 'text/plain', bytes: Buffer.from('hello') }]);
      assert.deepStrictEqual([text.status, /Upload a PDF, PNG or JPEG/.test(text.body.message)], [400, true]);
      const broken = await upload([{ name: 'bad.pdf', type: 'application/pdf', bytes: Buffer.from('%PDF-1.4 nonsense') }]);
      assert.deepStrictEqual([broken.status, /could not be opened/.test(broken.body.message)], [400, true]);
      assert.deepStrictEqual((await call('GET', '/api/scans', undefined, 'engineer@elegrow.com')).body.scans, []);
    });

    let first;
    let second;
    await check('a two-page PDF becomes two scans, each identified by its QR and read; the flagged cells and the second read are reported', async () => {
      const result = await upload([{ name: 'visit.pdf', type: 'application/pdf', bytes: both }]);
      assert.strictEqual(result.status, 202);
      assert.deepStrictEqual(result.body.scans.map((scan) => [scan.status, scan.pageIndex]), [['queued', 0], ['queued', 1]]);
      const scans = await settle();
      assert.strictEqual(scans.length, 2);
      assert.ok(scans.every((scan) => scan.status === 'read'));
      first = scans.find((scan) => scan.sheetKey === 'electrodeResistance');
      second = scans.find((scan) => scan.sheetKey === 'loopImpedanceTest');
      assert.deepStrictEqual([first.sheetCode, first.projectNo, first.workDate, first.flagged], ['MS-02', PROJECT.projectNo, '2026-10-12', 1]);
      assert.strictEqual(second.flagged, 0);
      const detail = (await call('GET', `/api/scans/${first.id}`, undefined, 'engineer@elegrow.com')).body;
      assert.deepStrictEqual([detail.escalated, detail.models.escalation.length > 0, detail.identified.startsWith('QR'), detail.canChange, detail.threshold], [true, true, true, true, 0.8]);
      assert.strictEqual(detail.cells['r0.resistanceWithGrid'].value, '0.52');
      assert.deepStrictEqual(detail.cells['r0.resistanceWithGrid'].flags.map((flag) => flag.code), ['READS_DIFFER']);
      assert.ok(detail.form.cells.length > 100 && detail.form.columns.some((col) => col.key === 'tag'));
      const image = await fetch(`${base}/api/scans/${first.id}/image`, { headers: { Cookie: await asCookie('engineer@elegrow.com') } });
      assert.deepStrictEqual([image.status, image.headers.get('content-type'), image.headers.get('cache-control')], [200, 'image/png', 'private, no-store']);
      assert.strictEqual((await fetch(`${base}/api/scans/${first.id}/image`)).status, 401);
    });

    await check('visibility: another engineer cannot see a scan; a reviewer can see it but not change it', async () => {
      assert.strictEqual((await call('GET', `/api/scans/${first.id}`, undefined, 'engineer2@elegrow.com')).status, 404);
      assert.strictEqual((await call('GET', `/api/scans/${first.id}/image`, undefined, 'engineer2@elegrow.com')).status, 404);
      assert.deepStrictEqual((await call('GET', '/api/scans', undefined, 'engineer2@elegrow.com')).body.scans, []);
      assert.strictEqual((await call('GET', '/api/scans?all=1', undefined, 'reviewer@elegrow.com')).body.scans.length, 2);
      assert.strictEqual((await call('GET', '/api/scans', undefined, 'reviewer@elegrow.com')).body.scans.length, 0, 'without all=1 it is your own');
      const seen = await call('GET', `/api/scans/${first.id}`, undefined, 'reviewer@elegrow.com');
      assert.deepStrictEqual([seen.status, seen.body.canChange], [200, false]);
      assert.strictEqual((await call('PUT', `/api/scans/${first.id}`, { edits: {} }, 'reviewer@elegrow.com')).status, 403, 'a reviewer holds no engineer role');
      assert.strictEqual((await call('PUT', `/api/scans/${first.id}`, { edits: {} }, 'engineer2@elegrow.com')).status, 404);
      assert.strictEqual((await call('POST', `/api/scans/${first.id}/accept`, {}, 'engineer2@elegrow.com')).status, 404);
    });

    await check('editing a flagged cell re-checks it; Accept writes the day (created, credited to the engineer), the audit log records it, and the scan is closed', async () => {
      const edited = await call('PUT', `/api/scans/${first.id}`, { edits: { 'r0.resistanceWithGrid': { value: '0,53' }, 'r1.length': { value: '-2' } } }, 'engineer@elegrow.com');
      assert.strictEqual(edited.status, 200);
      assert.deepStrictEqual([edited.body.cells['r0.resistanceWithGrid'].value, edited.body.cells['r0.resistanceWithGrid'].flags.length, edited.body.cells['r1.length'].flags.map((flag) => flag.code)], ['0.53', 0, ['OUT_OF_RANGE']]);
      assert.strictEqual(edited.body.flagged, 1);
      // a bad length is a validation error: nothing is written, the reasons are returned
      const refused = await call('POST', `/api/scans/${first.id}/accept`, {}, 'engineer@elegrow.com');
      assert.deepStrictEqual([refused.status, refused.body.code], [422, 'INVALID_REPORT'], JSON.stringify(refused.body).slice(0, 200));
      assert.strictEqual((await call('GET', `/api/projects/${workspace}`, undefined, 'engineer@elegrow.com')).body.days.length, 0);
      await call('PUT', `/api/scans/${first.id}`, { edits: { 'r1.length': { value: '0.6' } } }, 'engineer@elegrow.com');
      const accepted = await call('POST', `/api/scans/${first.id}/accept`, {}, 'engineer@elegrow.com');
      assert.strictEqual(accepted.status, 200, JSON.stringify(accepted.body).slice(0, 200));
      assert.deepStrictEqual([accepted.body.createdDay, accepted.body.overwrites, accepted.body.projectId, accepted.body.workDate], [true, 0, workspace, '2026-10-12']);
      const day = (await call('GET', `/api/reports/${accepted.body.reportId}`, undefined, 'engineer@elegrow.com')).body;
      assert.deepStrictEqual([day.kind, day.reviewStatus, day.collectionDate, day.createdBy.length > 0], ['daily', 'draft', '2026-10-12', true]);
      assert.deepStrictEqual(day.electrodeResistance.map((row) => [row.sheetRow, row.tag, row.resistanceWithGrid, row.enteredBy.email]), [[0, 'P-1', '0.53', 'engineer@elegrow.com'], [1, 'P-2', '1.20', 'engineer@elegrow.com']]);
      assert.deepStrictEqual([day.sheetHeaders.electrodeResistance.instrument, day.sheetHeaders.electrodeResistance.serialNo], ['mi3290', 'SN77']);
      assert.strictEqual((await call('GET', `/api/projects/${workspace}`, undefined, 'engineer@elegrow.com')).body.days.length, 1);
      const audit = (await call('GET', '/api/admin/audit?action=scan.', undefined, 'admin@elegrow.com')).body.entries.map((entry) => entry.action);
      assert.ok(audit.includes('scan.accept') && audit.includes('scan.upload'));
      assert.strictEqual((await call('POST', `/api/scans/${first.id}/accept`, {}, 'engineer@elegrow.com')).status, 409, 'a scan is accepted once');
      assert.strictEqual((await call('PUT', `/api/scans/${first.id}`, { edits: {} }, 'engineer@elegrow.com')).status, 409);
      // the loop page joins the same day
      const loopAccepted = await call('POST', `/api/scans/${second.id}/accept`, {}, 'engineer@elegrow.com');
      assert.deepStrictEqual([loopAccepted.status, loopAccepted.body.createdDay, loopAccepted.body.reportId === accepted.body.reportId], [200, false, true], 'one day per date');
      const merged = (await call('GET', `/api/reports/${accepted.body.reportId}`, undefined, 'engineer@elegrow.com')).body;
      assert.deepStrictEqual(merged.loopImpedanceTest.map((row) => [row.sheetRow, row.measuredPoints, row.location, row.loopImpedance]), [[0, 'R-E', 'MCC-1', '0.31'], [1, 'Y-E', 'MCC-1', '0.33'], [2, 'B-E', 'MCC-1', '0.35']]);
      assert.strictEqual(merged.electrodeResistance.length, 2);
    });

    await check('scanning the same page again overwrites only after the engineer confirms, and says exactly what it replaces', async () => {
      setTruth({ [electrode.issued.sheetId]: { primary: { ...electrode.truth, 'r0.resistanceWithGrid': { value: '0.47' }, 'r1.location': { value: 'Main gate' } } } });
      await upload([{ name: 'again.png', type: 'application/pdf', bytes: electrode.pdf }]);
      const scans = await settle();
      const again = scans.find((scan) => scan.status === 'read');
      const asked = await call('POST', `/api/scans/${again.id}/accept`, {}, 'engineer@elegrow.com');
      assert.deepStrictEqual([asked.status, asked.body.code, asked.body.overwrites.map((entry) => [entry.key, entry.before, entry.after])], [409, 'NEEDS_CONFIRM', [['resistanceWithGrid', '0.53', '0.47'], ['location', 'Gate', 'Main gate']]]);
      const day = (await call('GET', `/api/projects/${workspace}`, undefined, 'engineer@elegrow.com')).body.days[0];
      assert.strictEqual((await call('GET', `/api/reports/${day.reportId}`, undefined, 'engineer@elegrow.com')).body.electrodeResistance[0].resistanceWithGrid, '0.53', 'nothing changed without the confirmation');
      const done = await call('POST', `/api/scans/${again.id}/accept`, { confirm: true }, 'engineer@elegrow.com');
      assert.deepStrictEqual([done.status, done.body.overwrites], [200, 2]);
      assert.strictEqual((await call('GET', `/api/reports/${day.reportId}`, undefined, 'engineer@elegrow.com')).body.electrodeResistance[0].resistanceWithGrid, '0.47');
    });

    await check('a day that has gone to review cannot be scanned into until it is returned; the scan stays readable meanwhile', async () => {
      const day = (await call('GET', `/api/projects/${workspace}`, undefined, 'engineer@elegrow.com')).body.days[0];
      assert.strictEqual((await call('POST', `/api/reports/${day.reportId}/review/submit`, {}, 'engineer@elegrow.com')).status, 200);
      setTruth({ [electrode.issued.sheetId]: { primary: { ...electrode.truth, 'r0.tag': { value: 'ZZ' } } } });
      await upload([{ name: 'late.pdf', type: 'application/pdf', bytes: electrode.pdf }]);
      const late = (await settle()).find((scan) => scan.status === 'read');
      const refused = await call('POST', `/api/scans/${late.id}/accept`, { confirm: true }, 'engineer@elegrow.com');
      assert.deepStrictEqual([refused.status, refused.body.code], [409, 'NOT_DRAFT']);
      assert.strictEqual((await call('POST', `/api/scans/${late.id}/discard`, {}, 'engineer@elegrow.com')).body.status, 'discarded');
    });

    await check('a failing provider leaves the scan failed with the reason, Try again reads it once the provider is back, and the status screen shows the pause', async () => {
      setTruth({ [loop.issued.sheetId]: { fail: 'QUOTA' } });
      const ids = [];
      for (let i = 0; i < 4; i += 1) {
        await upload([{ name: `q${i}.pdf`, type: 'application/pdf', bytes: loop.pdf }]);
        const scans = await settle();
        ids.push(scans.find((scan) => scan.status === 'failed' && !ids.includes(scan.id)).id);
      }
      const states = await Promise.all(ids.map(async (id) => (await call('GET', '/api/scans', undefined, 'engineer@elegrow.com')).body.scans.find((scan) => scan.id === id)));
      assert.deepStrictEqual(states.slice(0, 3).map((scan) => scan.error.code), ['QUOTA', 'QUOTA', 'QUOTA']);
      assert.strictEqual(states[3].error.code, 'COOLING');
      assert.strictEqual((await call('GET', '/api/ai/status')).body.sheetReader.guard.state, 'cooling');
      assert.strictEqual((await call('POST', `/api/scans/${ids[0]}/retry`, {}, 'engineer@elegrow.com')).body.status, 'queued');
      await settle();
      assert.strictEqual((await call('GET', '/api/scans', undefined, 'engineer@elegrow.com')).body.scans.find((scan) => scan.id === ids[0]).status, 'failed', 'still paused, so still failed (the provider is not hammered)');
      assert.strictEqual((await call('POST', `/api/scans/${first.id}/retry`, {}, 'engineer@elegrow.com')).status, 409);
      const usage = (await call('GET', '/api/admin/ai-usage', undefined, 'admin@elegrow.com')).body.usage;
      assert.deepStrictEqual(usage.filter((row) => row.provider === 'gemini').map((row) => [row.kind, row.ok < row.calls]), [['ocr', true]]);
    });

    await check('admin purge deletes a date range of scans and their images, is audited, and is for admins only', async () => {
      const today = new Date().toISOString().slice(0, 10);
      assert.strictEqual((await call('POST', '/api/admin/scans/purge', { from: today, to: today }, 'engineer@elegrow.com')).status, 403);
      assert.strictEqual((await call('POST', '/api/admin/scans/purge', { from: today, to: '2001-01-01' }, 'admin@elegrow.com')).status, 400);
      const before = (await call('GET', '/api/scans', undefined, 'engineer@elegrow.com')).body.scans.length;
      assert.ok(before >= 4);
      assert.deepStrictEqual((await call('POST', '/api/admin/scans/purge', { from: '2001-01-01', to: '2001-12-31' }, 'admin@elegrow.com')).body.removed, 0);
      const purged = await call('POST', '/api/admin/scans/purge', { from: today, to: today }, 'admin@elegrow.com');
      assert.deepStrictEqual([purged.status, purged.body.removed], [200, before]);
      assert.deepStrictEqual((await call('GET', '/api/scans', undefined, 'engineer@elegrow.com')).body.scans, []);
      assert.strictEqual((await call('GET', `/api/scans/${first.id}/image`, undefined, 'engineer@elegrow.com')).status, 404);
      assert.ok((await call('GET', '/api/admin/audit?action=scan.', undefined, 'admin@elegrow.com')).body.entries.some((entry) => entry.action === 'scan.purge'));
      // the report that was accepted is not touched by a purge
      assert.strictEqual((await call('GET', `/api/projects/${workspace}`, undefined, 'engineer@elegrow.com')).body.days.length, 1);
    });
  } finally {
    server.kill();
  }
  console.log(`${passed} checks passed`);
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
