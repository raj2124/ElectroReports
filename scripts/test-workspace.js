// Project workspace (Session 6C.4), over HTTP against a real server on a temporary database:
// one workspace per Zoho project, one daily per work date, statuses, cumulative progress against
// the planned quantities, what a new day carries over, assembling the final report, and the
// Active / Completed split.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const net = require('net');
const { spawn } = require('child_process');
const { buildPayload } = require('./fixtures/golden-esha');

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

const ZOHO = { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej, Bharuch', workOrder: '1002-1261201769', engineerName: 'H. S. Patil', zohoProjectId: 'Z-1001', zohoProjectName: 'GFL ESHA', unit: 'FI Plant', unitCode: 'FI' };

// A day's draft: the golden payload, with the given pits and (optionally) no soil of its own.
function draftFor(pits, { soil = true, nextDayPlan = '' } = {}) {
  const payload = buildPayload();
  const draft = { ...payload, project: { ...payload.project, ...ZOHO, nextDayPlan, groundCondition: 'Dry' }, electrodeResistance: payload.electrodeResistance.slice(...pits) };
  if (!soil) {
    draft.tests = { ...draft.tests, soilResistivity: false };
    draft.soilResistivity = { locations: [] };
  }
  draft.tests = { soilResistivity: soil, electrodeResistance: true };
  delete draft.loopImpedanceTest;
  delete draft.prospectiveFaultCurrent;
  delete draft.riserIntegrityTest;
  return draft;
}

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'workspace-'));
  const port = await freePort();
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'src', 'server.js')], {
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ANTHROPIC_API_KEY: '', OPENAI_API_KEY: '', GEMINI_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    server.on('exit', (code) => reject(new Error(`server exited early (${code})`)));
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const call = async (method, url, body) => {
    const response = await fetch(`http://127.0.0.1:${port}${url}`, { method, headers: { 'Content-Type': 'application/json' }, body: body === undefined ? undefined : JSON.stringify(body) });
    return { status: response.status, body: await response.json().catch(() => null) };
  };
  const step = async (id, name, body) => (await call('POST', `/api/reports/${id}/review/${name}`, body)).body;
  const tickAll = async (id) => {
    const ticks = Object.fromEntries((await call('GET', `/api/reports/${id}/review`)).body.checklist.map((item) => [item.id, true]));
    await call('PUT', `/api/reports/${id}/review/ticks`, { ticks });
  };

  try {
    let projectId;
    await check('one workspace per Zoho project: created once, found again; a project needs a number and a client', async () => {
      const first = await call('POST', '/api/projects', { project: ZOHO });
      assert.deepStrictEqual([first.status, first.body.created], [201, true]);
      projectId = first.body.projectId;
      const again = await call('POST', '/api/projects', { project: { ...ZOHO, engineerName: 'Someone else' } });
      assert.deepStrictEqual([again.status, again.body.created, again.body.projectId], [200, false, projectId]);
      const other = await call('POST', '/api/projects', { project: { ...ZOHO, zohoProjectId: 'Z-2002', projectNo: '025-D03' } });
      assert.strictEqual(other.status, 201);
      assert.notStrictEqual(other.body.projectId, projectId);
      assert.strictEqual((await call('POST', '/api/projects', { project: { projectNo: 'X' } })).status, 400);
      assert.strictEqual((await call('GET', '/api/projects/missing')).status, 404);
    });

    await check('an empty workspace: no days, active, nothing to carry but the project details', async () => {
      const { body } = await call('GET', `/api/projects/${projectId}`);
      assert.deepStrictEqual([body.days.length, body.state, body.finals.length, body.assembly.totalDailyReports], [0, 'active', 0, 0]);
      assert.strictEqual(body.dayTemplate.project.clientName, ZOHO.clientName);
      assert.strictEqual(body.dayTemplate.project.zohoProjectId, 'Z-1001');
    });

    let day1;
    let day2;
    await check('planned quantities are set on the project; progress starts at zero', async () => {
      const put = await call('PUT', `/api/projects/${projectId}`, { project: { plannedQuantities: { electrodeResistance: '10', soilResistivity: 2, loopImpedanceTest: 'x', riserIntegrityTest: -1 } } });
      assert.strictEqual(put.status, 200);
      assert.deepStrictEqual(put.body.planned, { electrodeResistance: 10, soilResistivity: 2 });
      assert.deepStrictEqual(put.body.progress.map((row) => [row.testId, row.planned, row.done, row.percent]), [['electrodeResistance', 10, 0, 0], ['soilResistivity', 2, 0, 0]]);
    });

    await check('a day is recorded once per work date; a second save for the date is refused with the report to reopen', async () => {
      const first = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-14', report: draftFor([0, 6], { nextDayPlan: 'Earth pits EP-07 to EP-10.' }) });
      assert.strictEqual(first.status, 201, JSON.stringify(first.body));
      day1 = first.body;
      assert.deepStrictEqual([day1.kind, day1.projectId, day1.collectionDate, day1.reviewStatus, day1.project.reportDate], ['daily', projectId, '2026-09-14', 'draft', '2026-09-14']);
      const duplicate = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-14', report: draftor() });
      assert.deepStrictEqual([duplicate.status, duplicate.body.code, duplicate.body.reportId], [409, 'DAY_EXISTS', day1.id]);
      assert.strictEqual((await call('POST', `/api/projects/${projectId}/days`, { workDate: 'tomorrow', report: draftFor([0, 1]) })).status, 400);
      assert.strictEqual((await call('POST', '/api/projects/missing/days', { workDate: '2026-09-15', report: draftFor([0, 1]) })).status, 404);
      function draftor() { return draftFor([0, 2]); }
    });

    await check('a new day carries the project details over, but NOT the previous day\'s plan, date or ground condition', async () => {
      const { body } = await call('GET', `/api/projects/${projectId}`);
      const carried = body.dayTemplate.project;
      assert.ok(!('nextDayPlan' in carried), 'next-day plan is not carried');
      assert.ok(!('reportDate' in carried) && !('groundCondition' in carried));
      assert.deepStrictEqual([carried.clientName, carried.siteLocation, carried.workOrder, carried.engineerName, carried.unitCode], [ZOHO.clientName, ZOHO.siteLocation, ZOHO.workOrder, ZOHO.engineerName, 'FI']);
      assert.deepStrictEqual(carried.plannedQuantities, { electrodeResistance: 10, soilResistivity: 2 });
      assert.deepStrictEqual(body.dayTemplate.tests.electrodeResistance, true);
      assert.ok(!('id' in carried) && !('createdAt' in carried));
    });

    await check('days list with their status, and the cumulative progress counts each point once', async () => {
      const second = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-15', report: draftFor([6, 10], { soil: false }) });
      assert.strictEqual(second.status, 201, JSON.stringify(second.body));
      day2 = second.body;
      const { body } = await call('GET', `/api/projects/${projectId}`);
      assert.deepStrictEqual(body.days.map((day) => [day.date, day.status]), [['2026-09-14', 'draft'], ['2026-09-15', 'draft']]);
      const progress = Object.fromEntries(body.progress.map((row) => [row.testId, row]));
      assert.deepStrictEqual([progress.electrodeResistance.done, progress.electrodeResistance.percent], [10, 100]);
      assert.deepStrictEqual([progress.soilResistivity.done, progress.soilResistivity.percent], [2, 100]);
    });

    await check('statuses follow the review: draft, in review, checked, approved', async () => {
      const status = async (id) => (await call('GET', `/api/projects/${projectId}`)).body.days.find((day) => day.reportId === id).status;
      assert.strictEqual(await status(day1.id), 'draft');
      await step(day1.id, 'submit', { actor: 'H. S. Patil' });
      assert.strictEqual(await status(day1.id), 'in_review');
      await tickAll(day1.id);
      await step(day1.id, 'check', { actor: 'A. P.' });
      assert.strictEqual(await status(day1.id), 'checked');
      await step(day1.id, 'approve', { actor: 'J. C. Shah' });
      assert.strictEqual(await status(day1.id), 'approved');
      assert.strictEqual(await status(day2.id), 'draft');
    });

    await check('a draft day can be edited (PUT); an approved day, a changed date and a final report cannot', async () => {
      const edited = draftFor([6, 10], { soil: false });
      edited.electrodeResistance[0].observation = 'Re-measured after cleaning the clamp';
      const ok = await call('PUT', `/api/reports/${day2.id}`, { ...edited, project: { ...edited.project, reportDate: '2026-09-15' } });
      assert.strictEqual(ok.status, 200, JSON.stringify(ok.body));
      assert.deepStrictEqual([ok.body.kind, ok.body.projectId, ok.body.collectionDate, ok.body.reviewStatus], ['daily', projectId, '2026-09-15', 'draft']);
      assert.strictEqual(ok.body.electrodeResistance[0].observation, 'Re-measured after cleaning the clamp');
      const moved = await call('PUT', `/api/reports/${day2.id}`, { ...edited, project: { ...edited.project, reportDate: '2026-09-20' } });
      assert.strictEqual(moved.status, 400);
      const approved = await call('PUT', `/api/reports/${day1.id}`, draftFor([0, 6]));
      assert.deepStrictEqual([approved.status, approved.body.code], [409, 'NOT_DRAFT']);
      assert.strictEqual((await call('PUT', '/api/reports/missing', draftFor([0, 1]))).status, 404);
    });

    let firstFinal;
    await check('assemble: every day is stitched, the days not yet approved are listed and flagged, the result is a draft', async () => {
      const { body: before } = await call('GET', `/api/projects/${projectId}`);
      assert.deepStrictEqual([before.assembly.includedDates, before.assembly.flaggedDates, before.assembly.hasUnapproved], [['2026-09-14', '2026-09-15'], ['2026-09-15'], true]);
      const result = await call('POST', `/api/projects/${projectId}/final/assemble`, { actor: 'H. S. Patil' });
      assert.strictEqual(result.status, 201, JSON.stringify(result.body));
      assert.deepStrictEqual([result.body.includedDates, result.body.flaggedDates, result.body.hasUnapproved], [['2026-09-14', '2026-09-15'], ['2026-09-15'], true]);
      assert.deepStrictEqual([result.body.report.kind, result.body.report.reviewStatus, result.body.report.projectId], ['final', 'draft', projectId]);
      assert.strictEqual(result.body.report.electrodeResistance.length, 10, 'all ten pits from both days');
      firstFinal = result.body.report;
      assert.strictEqual((await call('POST', '/api/projects/missing/final/assemble', {})).status, 404);
    });

    await check('re-assemble makes a NEW draft final and keeps the old one', async () => {
      const again = await call('POST', `/api/projects/${projectId}/final/assemble`, { actor: 'H. S. Patil' });
      assert.strictEqual(again.status, 201);
      assert.notStrictEqual(again.body.report.id, firstFinal.id);
      const { body } = await call('GET', `/api/projects/${projectId}`);
      assert.deepStrictEqual(body.finals.map((entry) => entry.status), ['draft', 'draft']);
      assert.ok(body.finals.some((entry) => entry.id === firstFinal.id), 'the earlier final is still there');
      assert.strictEqual((await call('GET', `/api/reports/${firstFinal.id}`)).body.reviewStatus, 'draft');
    });

    await check('a project with no daily report cannot be assembled', async () => {
      const other = await call('POST', '/api/projects', { project: { ...ZOHO, zohoProjectId: 'Z-3003', projectNo: '026-D03' } });
      const result = await call('POST', `/api/projects/${other.body.projectId}/final/assemble`, {});
      assert.strictEqual(result.status, 400);
      assert.ok(/no daily reports/.test(result.body.message));
    });

    await check('Active until the final is approved, then Completed; a day recorded afterwards makes it active again', async () => {
      const summary = async () => (await call('GET', '/api/projects')).body.find((entry) => entry.id === projectId);
      assert.deepStrictEqual([(await summary()).state, (await summary()).days, (await summary()).approvedDays, (await summary()).finalStatus], ['active', 2, 1, 'draft']);
      // approve the second day, re-assemble, approve that final
      await step(day2.id, 'submit', { actor: 'H. S. Patil' });
      await tickAll(day2.id);
      await step(day2.id, 'check', { actor: 'A. P.' });
      await step(day2.id, 'approve', { actor: 'J. C. Shah' });
      const { body: assembled } = await call('POST', `/api/projects/${projectId}/final/assemble`, { actor: 'H. S. Patil' });
      assert.strictEqual(assembled.hasUnapproved, false);
      const final = assembled.report;
      assert.strictEqual((await call('PUT', `/api/reports/${final.id}`, draftFor([0, 2]))).status, 400, 'a final report is assembled, never edited by hand');
      await step(final.id, 'submit', { actor: 'H. S. Patil' });
      await tickAll(final.id);
      await step(final.id, 'check', { actor: 'A. P.' });
      const approval = await step(final.id, 'approve', { actor: 'J. C. Shah', acknowledgements: { sanityCheck: 'Read by hand.', p1Photos: 'Photos follow.' } });
      assert.strictEqual(approval.status, 'approved', JSON.stringify(approval));
      assert.deepStrictEqual([(await summary()).state, (await summary()).finalStatus], ['completed', 'approved']);
      const third = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-22', report: draftFor([0, 2], { soil: false }) });
      assert.strictEqual(third.status, 201, JSON.stringify(third.body));
      assert.strictEqual((await summary()).state, 'active');
    });

    await check('the home list carries every workspace with its counts and link data', async () => {
      const { body } = await call('GET', '/api/projects');
      assert.strictEqual(body.length, 3);
      const mine = body.find((entry) => entry.id === projectId);
      assert.deepStrictEqual([mine.projectNo, mine.clientName, mine.zohoProjectId, mine.lastDay], [ZOHO.projectNo, ZOHO.clientName, 'Z-1001', '2026-09-22']);
    });
  } finally {
    server.kill();
  }
  console.log(`\n${passed} workspace checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
