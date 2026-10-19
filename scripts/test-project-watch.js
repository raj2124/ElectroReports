// New D03 project watch (Session 6D): which projects get a notice and who gets it, one notice per
// project and recipient forever, the two answers and the default data-collection mode, and the
// optional Work Order. Store-level with a stubbed Zoho for the rules, then over HTTP on the mock
// portal (ZOHO_USE_MOCK, 056-D03_TS-26-CLL and friends).
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { pollNewProjects, answerNotice, dismissNotice, dataCollectionModeFor } = require('../src/projectWatch');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const cfg = (extra = {}) => ({ zoho: { notifyStatuses: ['Planning', 'Active'], ...extra } });
const ms = (iso) => String(Date.parse(iso));
const project = (id, name, stage, created, owner = ['Nidhi Shah', 'nidhi@elegrow.com']) => ({ id, name, projectNumber: name, clientName: name.split('-').pop(), stage, createdAt: ms(created), ownerName: owner[0], ownerEmail: owner[1] });
const onsiteTask = (id, ownerId, ownerName, email) => ({ id, name: `[ESHA] Earth Pit Resistance Testing - ${id}`, taskListName: 'Onsite', owners: [{ id: ownerId, name: ownerName.split(' ')[0].toLowerCase(), email }], customFields: [{ label: 'Work Location', apiName: 'UDF_CHAR81', value: 'Onsite' }] });

const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});

(async () => {
  await check('the default mode: phones allowed AND mobile data -> on-site entry; anything else -> printed sheets', () => {
    assert.strictEqual(dataCollectionModeFor({ phonesAllowed: true, mobileData: true }), 'On-site entry');
    assert.strictEqual(dataCollectionModeFor({ phonesAllowed: true, mobileData: false }), 'Printed sheets');
    assert.strictEqual(dataCollectionModeFor({ phonesAllowed: false, mobileData: true }), 'Printed sheets');
    assert.strictEqual(dataCollectionModeFor({ phonesAllowed: false, mobileData: false }), 'Printed sheets');
  });

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'watch-'));
  const store = createReportStore({ filePath: path.join(dir, 'watch.db') });
  const users = [{ id: 'u1', displayName: 'Anup Parekh', email: 'anup@elegrow.com' }, { id: 'u2', displayName: 'Karishma Prajapati', email: 'karishma@elegrow.com' }];
  const stub = (projects, tasks) => ({ getProjects: async () => projects, getProjectTasks: async (id) => tasks[id] || [], getProjectUsers: async () => users });
  const portal = [
    project('P1', '056-D03_TS-26-CLL', 'Planning', '2026-10-13T08:00:00Z'),
    project('P2', '060-D03_TS-26-XYZ', 'Active', '2026-10-15T08:00:00Z'),
    project('P3', '024-D03_TS-26-GFL', 'Client Review', '2026-07-01T08:00:00Z'),
    project('P4', '004-D03_TS-26-APL', 'Invoicing', '2026-05-05T08:00:00Z'),
    project('P5', '011-D03_TS-25-OLD', 'Active', '2025-12-20T08:00:00Z'),
    project('P6', 'Corporate Website Revamp', 'Planning', '2026-10-15T08:00:00Z'),
    project('P7', '070-D03_TS-26-NODATE', 'Planning', '')
  ];
  const tasks = { P1: [onsiteTask('a', '1', 'Anup Parekh', 'anup@elegrow.com'), onsiteTask('b', '2', 'Karishma Prajapati', 'karishma@elegrow.com'), { id: 'c', name: '[ESHA] Draft report', owners: [{ id: '9', name: 'zed', email: 'zed@elegrow.com' }], customFields: [{ label: 'Work Location', value: 'Offsite' }] }, { id: 'd', name: '[IRT] Thermography', owners: [{ id: '8', name: 'irt', email: 'irt@elegrow.com' }], customFields: [{ label: 'Work Location', value: 'Onsite' }] }] };

  await check('notices only for D03 projects in Planning/Active, created on or after 2026-01-01 (a known date), with no workspace', async () => {
    const result = await pollNewProjects({ store, zoho: stub(portal, tasks), config: cfg() });
    assert.deepStrictEqual(result.candidates, ['056-D03_TS-26-CLL', '060-D03_TS-26-XYZ']);
    assert.strictEqual(result.checked, 7);
  });

  await check('recipients are the owners of the ESHA Onsite tasks (not Offsite, not IRT); with none, the project owner', () => {
    const open = store.listNotices({ status: 'open' });
    const forP1 = open.filter((notice) => notice.zohoProjectId === 'P1');
    assert.deepStrictEqual(forP1.map((notice) => [notice.recipientKey, notice.recipientName, notice.payload.reason]).sort(), [['anup@elegrow.com', 'Anup Parekh', 'onsite_owner'], ['karishma@elegrow.com', 'Karishma Prajapati', 'onsite_owner']]);
    const forP2 = open.filter((notice) => notice.zohoProjectId === 'P2');
    assert.deepStrictEqual(forP2.map((notice) => [notice.recipientKey, notice.payload.reason]), [['nidhi@elegrow.com', 'project_owner']]);
    assert.strictEqual(open.length, 3);
    assert.strictEqual(forP1[0].payload.projectName, '056-D03_TS-26-CLL');
  });

  await check('the status list is configurable', async () => {
    const other = createReportStore({ filePath: path.join(dir, 'other.db') });
    const result = await pollNewProjects({ store: other, zoho: stub(portal, tasks), config: cfg({ notifyStatuses: ['Client Review'] }) });
    assert.deepStrictEqual(result.candidates, ['024-D03_TS-26-GFL']);
  });

  await check('a second poll creates nothing: a notice exists once per project and recipient', async () => {
    const again = await pollNewProjects({ store, zoho: stub(portal, tasks), config: cfg() });
    assert.deepStrictEqual(again.created, []);
    assert.strictEqual(store.listNotices().length, 3);
  });

  await check('the two answers are required; the first answer creates the workspace with the answers and the default mode', () => {
    const notice = store.listNotices({ recipientKey: 'anup@elegrow.com' })[0];
    assert.throws(() => answerNotice(store, notice.id, { phonesAllowed: true }), (error) => error.code === 'ANSWERS_REQUIRED');
    assert.throws(() => answerNotice(store, 'missing', { phonesAllowed: true, mobileData: true }), (error) => error.code === 'NOT_FOUND');
    const result = answerNotice(store, notice.id, { phonesAllowed: true, mobileData: false, by: 'Anup Parekh' });
    assert.strictEqual(result.dataCollectionMode, 'Printed sheets');
    const saved = store.getProject(result.projectId);
    assert.deepStrictEqual([saved.zohoProjectId, saved.projectNo, saved.dataCollectionMode], ['P1', '056-D03_TS-26-CLL', 'Printed sheets']);
    assert.deepStrictEqual([saved.siteAccess.phonesAllowed, saved.siteAccess.mobileData, saved.siteAccess.answeredBy], [true, false, 'Anup Parekh']);
  });

  await check('the other owners\' notices for that project close; an answered notice cannot be answered again', () => {
    const karishma = store.listNotices({ recipientKey: 'karishma@elegrow.com' })[0];
    assert.strictEqual(karishma.status, 'answered');
    assert.throws(() => answerNotice(store, karishma.id, { phonesAllowed: true, mobileData: true }), (error) => error.code === 'NOT_OPEN');
    assert.strictEqual(store.listNotices({ recipientKey: 'anup@elegrow.com' })[0].answers.phonesAllowed, true);
  });

  await check('a dismissed notice does not come back; a project with a workspace is not a candidate', async () => {
    const nidhi = store.listNotices({ recipientKey: 'nidhi@elegrow.com' })[0];
    dismissNotice(store, nidhi.id, 'Nidhi Shah');
    assert.strictEqual(store.getNotice(nidhi.id).status, 'dismissed');
    assert.throws(() => dismissNotice(store, nidhi.id), (error) => error.code === 'NOT_OPEN');
    const after = await pollNewProjects({ store, zoho: stub(portal, tasks), config: cfg() });
    assert.deepStrictEqual(after.candidates, ['060-D03_TS-26-XYZ'], 'P1 has a workspace now; P2 is still a candidate');
    assert.deepStrictEqual(after.created, [], 'but nobody is asked again');
    assert.strictEqual(store.listNotices({ status: 'open' }).length, 0);
  });

  await check('the project owner is the fallback only when there is no Onsite ESHA task owner', async () => {
    const third = createReportStore({ filePath: path.join(dir, 'third.db') });
    await pollNewProjects({ store: third, zoho: stub([portal[0]], { P1: [] }), config: cfg() });
    assert.deepStrictEqual(third.listNotices().map((notice) => notice.payload.reason), ['project_owner']);
  });

  // ---- over HTTP on the mock portal -----------------------------------------------------------
  const port = await freePort();
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'src', 'server.js')], {
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'http.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ZOHO_POLL_MINUTES: '0', ANTHROPIC_API_KEY: '' },
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
  try {
    await check('HTTP: the poll finds the mock D03 projects; notices are listed for the person (by e-mail or name) and only for them', async () => {
      const poll = await call('POST', '/api/zoho/poll');
      assert.strictEqual(poll.status, 200);
      assert.deepStrictEqual(poll.body.candidates, ['056-D03_TS-26-CLL', '060-D03_TS-26-XYZ']);
      assert.strictEqual(poll.body.created.length, 2);
      assert.strictEqual((await call('POST', '/api/zoho/poll')).body.created.length, 0);
      assert.strictEqual((await call('GET', '/api/notices?email=anup@elegrow.com')).body.length, 1);
      assert.strictEqual((await call('GET', '/api/notices?name=Anup%20Parekh')).body.length, 0, 'matched by the key it was addressed to (the e-mail)');
      assert.strictEqual((await call('GET', '/api/notices?email=karishma@elegrow.com')).body.length, 0, 'she owns only an IRT task');
      assert.strictEqual((await call('GET', '/api/notices')).body.length, 0);
    });

    let projectId;
    await check('HTTP: answering opens the workspace with the default mode, which new site days start from', async () => {
      const [notice] = (await call('GET', '/api/notices?email=anup@elegrow.com')).body;
      assert.strictEqual((await call('POST', `/api/notices/${notice.id}/answer?email=anup@elegrow.com`, { phonesAllowed: true })).status, 400);
      const answered = await call('POST', `/api/notices/${notice.id}/answer?email=anup@elegrow.com&name=Anup%20Parekh`, { phonesAllowed: true, mobileData: true });
      assert.deepStrictEqual([answered.status, answered.body.dataCollectionMode], [200, 'On-site entry']);
      projectId = answered.body.projectId;
      const workspace = (await call('GET', `/api/projects/${projectId}`)).body;
      assert.strictEqual(workspace.project.dataCollectionMode, 'On-site entry');
      assert.strictEqual(workspace.dayTemplate.project.dataCollectionMode, 'On-site entry');
      assert.strictEqual(workspace.project.siteAccess.answeredBy, 'Anup Parekh');
      assert.strictEqual((await call('POST', `/api/notices/${notice.id}/answer?email=anup@elegrow.com`, { phonesAllowed: true, mobileData: true })).status, 409);
      assert.strictEqual((await call('GET', '/api/notices?email=anup@elegrow.com')).body.length, 0);
    });

    await check('HTTP: dismissing removes the notice for good', async () => {
      const [notice] = (await call('GET', '/api/notices?email=nidhi@elegrow.com')).body;
      assert.strictEqual((await call('POST', `/api/notices/${notice.id}/dismiss?email=nidhi@elegrow.com`, {})).status, 200);
      assert.strictEqual((await call('POST', `/api/notices/${notice.id}/dismiss?email=nidhi@elegrow.com`, {})).status, 409);
      await call('POST', '/api/zoho/poll');
      assert.strictEqual((await call('GET', '/api/notices?email=nidhi@elegrow.com')).body.length, 0);
    });

    await check('HTTP: the task plan of a D03 project (mock) and "Refresh from Zoho" fills only clear quantities', async () => {
      const plan = (await call('GET', '/api/zoho/projects/27127000009850033/plan')).body;
      assert.strictEqual(plan.plan.d03, true);
      assert.deepStrictEqual([plan.plan.tests.soilResistivity.planned, plan.plan.tests.electrodeResistance.planned], [4, 170]);
      assert.deepStrictEqual(plan.plan.tasks.filter((task) => task.selected).map((task) => task.ownerName), ['Anup Parekh', 'Anup Parekh']);
      const plain = (await call('GET', '/api/zoho/projects/1001/plan')).body;
      assert.deepStrictEqual([plain.plan.d03, plain.plan.tasks.filter((task) => task.selected).length], [false, 0]);
      const refreshed = await call('POST', `/api/projects/${projectId}/zoho-refresh`, {});
      assert.strictEqual(refreshed.status, 200, JSON.stringify(refreshed.body));
      assert.deepStrictEqual(refreshed.body.workspace.planned, { soilResistivity: 4, electrodeResistance: 170 });
      assert.strictEqual((await call('POST', '/api/projects/missing/zoho-refresh', {})).status, 404);
    });

    await check('HTTP: Work Order is optional on the project and in a site day; an empty value clears it', async () => {
      const put = await call('PUT', `/api/projects/${projectId}`, { project: { workOrder: '1002-1261201769' } });
      assert.strictEqual(put.body.project.workOrder, '1002-1261201769');
      assert.strictEqual(put.body.dayTemplate.project.workOrder, '1002-1261201769');
      const cleared = await call('PUT', `/api/projects/${projectId}`, { project: { workOrder: '' } });
      assert.strictEqual(cleared.body.project.workOrder, '');
      assert.strictEqual(cleared.body.dayTemplate.project.workOrder, '');
      assert.strictEqual((await call('PUT', `/api/projects/${projectId}`, { project: { dataCollectionMode: 'Carrier pigeon' } })).status, 400);
      const mode = await call('PUT', `/api/projects/${projectId}`, { project: { dataCollectionMode: 'Printed sheets' } });
      assert.strictEqual(mode.body.project.dataCollectionMode, 'Printed sheets');
      const { buildPayload } = require('./fixtures/golden-esha');
      const payload = buildPayload();
      payload.project = { ...payload.project, workOrder: '', unit: 'FI Plant', unitCode: 'FI', nextDayPlan: '' };
      const day = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-14', report: payload });
      assert.strictEqual(day.status, 201, JSON.stringify(day.body));
      assert.strictEqual(day.body.project.workOrder, '');
    });
  } finally {
    server.kill();
  }
  console.log(`\n${passed} project-watch checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
