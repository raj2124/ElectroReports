// Session 6H, Microsoft Graph, all on the in-memory Graph (nothing here reaches Microsoft):
//   6H.1 the client: certificate / secret token, retry on throttling, errors without tokens or addresses, and the
//        HARD RULE that SharePoint is never written;
//   6H.2 the Outlook calendar sync: the events, idempotence, changes, deletes, never touching an event the app did
//        not create, mailboxes outside the group, the toggle and the dry run;
//   6H.3 / 6H.4 the knowledge base: which files are indexed, delta refresh, search, and who finds what;
//   6H.5 references on a final report (reviewer-picked, client-private) and the HTTP routes.
const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { createGraphClient, GraphError, SharePointReadOnlyError } = require('../src/graphClient');
const { createMockGraph } = require('../src/graphMock');
const { makePdf, makeDocx, makeXlsx } = require('../src/graphFixtures');
const { extractText, chunkText } = require('../src/kbExtract');
const { createCalendarSync, PROPERTY_ID } = require('../src/calendarSync');
const { createKnowledgeBase, projectKeyOf, classify } = require('../src/knowledgeBase');
const { generateFinalReportPdf } = require('../src/finalReport/pdf');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { openPdf, pageTextLines, pageCount } = require('../src/pdfRaster');
const { devEnv, createCaller } = require('./lib/dev-auth');

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

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-'));
const newStore = () => createReportStore({ filePath: path.join(dir, `${Math.random().toString(36).slice(2)}.db`) });
const BASE_CONFIG = { baseUrl: 'https://graph.microsoft.com/v1.0', authority: 'https://login.test', tenantId: 'tenant-1', clientId: 'client-1', clientSecret: 'shh-secret', certPath: '', certThumbprintSha256: '', timezone: 'India Standard Time', appBaseUrl: 'https://reports.example.test' };
const mockClient = (mock, extra = {}) => createGraphClient({ http: mock.http, config: { ...BASE_CONFIG, ...extra }, sleep: async () => {} });

(async () => {
  // ======== 6H.1 the client ===================================================================================
  await check('token by certificate: a signed assertion with the thumbprint header, the right claims, and a valid RS256 signature', async () => {
    const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
    const keyFile = path.join(dir, 'test.key');
    fs.writeFileSync(keyFile, privateKey.export({ type: 'pkcs8', format: 'pem' }));
    const thumbprint = crypto.createHash('sha256').update('not a real certificate').digest('hex');
    let body;
    const http = { post: async (url, text) => { body = { url, params: new URLSearchParams(text) }; return { data: { access_token: 'tok-1', expires_in: 3600 } }; }, request: async () => ({ status: 200, data: { value: [] }, headers: {} }) };
    const client = createGraphClient({ http, config: { ...BASE_CONFIG, clientSecret: '', certPath: keyFile, certThumbprintSha256: thumbprint } });
    await client.get('/users/a@x.test/events');
    assert.strictEqual(body.url, 'https://login.test/tenant-1/oauth2/v2.0/token');
    assert.strictEqual(body.params.get('grant_type'), 'client_credentials');
    assert.strictEqual(body.params.get('scope'), 'https://graph.microsoft.com/.default');
    assert.ok(!body.params.has('client_secret'));
    assert.strictEqual(body.params.get('client_assertion_type'), 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer');
    const [header, payload, signature] = body.params.get('client_assertion').split('.');
    const head = JSON.parse(Buffer.from(header, 'base64url'));
    const claims = JSON.parse(Buffer.from(payload, 'base64url'));
    assert.deepStrictEqual([head.alg, head['x5t#S256']], ['RS256', Buffer.from(thumbprint, 'hex').toString('base64url')]);
    assert.deepStrictEqual([claims.aud, claims.iss, claims.sub], ['https://login.test/tenant-1/oauth2/v2.0/token', 'client-1', 'client-1']);
    assert.ok(claims.exp > claims.nbf && claims.jti);
    assert.ok(crypto.createVerify('RSA-SHA256').update(`${header}.${payload}`).verify(publicKey, Buffer.from(signature, 'base64url')));
  });

  await check('token by secret, cached until it expires; an unconfigured client says what is missing', async () => {
    let tokens = 0;
    let clock = 0;
    const http = { post: async (url, text) => { tokens += 1; assert.ok(new URLSearchParams(text).get('client_secret') === 'shh-secret'); return { data: { access_token: `t${tokens}`, expires_in: 3600 } }; }, request: async (config) => ({ status: 200, data: { value: [], auth: config.headers.Authorization }, headers: {} }) };
    const client = createGraphClient({ http, config: BASE_CONFIG, now: () => clock });
    assert.strictEqual((await client.get('/users/a@x.test/events')).data.auth, 'Bearer t1');
    await client.get('/users/a@x.test/events');
    assert.strictEqual(tokens, 1);
    clock += 3600 * 1000;
    assert.strictEqual((await client.get('/users/a@x.test/events')).data.auth, 'Bearer t2');
    await assert.rejects(() => createGraphClient({ http, config: { ...BASE_CONFIG, clientSecret: '' } }).get('/x'), /GRAPH_TENANT_ID.*GRAPH_CLIENT_ID.*certificate/);
  });

  await check('429 / 503 are retried after Retry-After; a 401 refreshes the token once; other failures stop at once', async () => {
    const sleeps = [];
    let attempts = 0;
    const http = {
      post: async () => ({ data: { access_token: 'tok', expires_in: 3600 } }),
      request: async () => {
        attempts += 1;
        if (attempts === 1) throw Object.assign(new Error('x'), { response: { status: 429, headers: { 'retry-after': '7' }, data: {} } });
        if (attempts === 2) throw Object.assign(new Error('x'), { response: { status: 503, headers: {}, data: {} } });
        return { status: 200, data: { ok: true }, headers: {} };
      }
    };
    const client = createGraphClient({ http, config: BASE_CONFIG, sleep: async (ms) => sleeps.push(ms) });
    assert.deepStrictEqual((await client.get('/users/a@x.test/events')).data, { ok: true });
    assert.deepStrictEqual(sleeps, [7000, 2000]);
    let n = 0;
    const denied = createGraphClient({ http: { post: async () => ({ data: { access_token: `tok${++n}`, expires_in: 3600 } }), request: async (c) => { if (c.headers.Authorization === 'Bearer tok1') throw Object.assign(new Error('x'), { response: { status: 401, data: {}, headers: {} } }); return { status: 200, data: {}, headers: {} }; } }, config: BASE_CONFIG });
    await denied.get('/users/a@x.test/events');
    assert.strictEqual(n, 2, 'one refresh');
    let tries = 0;
    const gone = createGraphClient({ http: { post: async () => ({ data: { access_token: 't', expires_in: 3600 } }), request: async () => { tries += 1; throw Object.assign(new Error('x'), { response: { status: 404, data: { error: { code: 'ErrorItemNotFound' } }, headers: {} } }); } }, config: BASE_CONFIG });
    await assert.rejects(() => gone.get('/x'), (error) => error instanceof GraphError && error.status === 404 && error.code === 'ErrorItemNotFound');
    assert.strictEqual(tries, 1);
  });

  await check('errors carry the status and code only: no token, no secret, no address from the message', async () => {
    const http = {
      post: async () => { throw Object.assign(new Error('boom shh-secret'), { response: { status: 401, data: { error: 'invalid_client', error_description: 'secret shh-secret for anna.k@elegrow.com' } } }); },
      request: async () => { throw Object.assign(new Error('boom'), { response: { status: 403, headers: {}, data: { error: { code: 'ErrorAccessDenied', message: 'Access is denied for anna.k@elegrow.com' } } } }); }
    };
    await assert.rejects(() => createGraphClient({ http, config: BASE_CONFIG }).get('/x'), (error) => error.message === 'Microsoft Graph request failed (401: invalid_client)' && !/shh|anna/.test(error.message + JSON.stringify(error)));
    const ok = createGraphClient({ http: { ...http, post: async () => ({ data: { access_token: 'secret-token-value', expires_in: 3600 } }) }, config: BASE_CONFIG });
    await assert.rejects(() => ok.get('/x'), (error) => error.status === 403 && error.code === 'ErrorAccessDenied' && !/anna|secret-token/.test(error.message + JSON.stringify(error)));
  });

  await check('HARD RULE: SharePoint is never written. POST, PUT, PATCH and DELETE on /sites and /drives are refused before any request; calendar writes are not', async () => {
    const mock = createMockGraph({ group: ['a@x.test'] });
    mock.sp.addSite('S', 'https://t.sharepoint.com/sites/S');
    const client = mockClient(mock);
    const paths = ['/sites/t.sharepoint.com:/sites/S', '/sites/abc/drives', '/drives/drive-S/root/children', '/drives/drive-S/items/item-1', '/drives/drive-S/items/item-1/content', 'https://graph.microsoft.com/v1.0/drives/drive-S/root/delta'];
    for (const verb of ['POST', 'PUT', 'PATCH', 'DELETE']) {
      for (const target of paths) {
        await assert.rejects(() => client.request(verb, target, { body: {} }), (error) => error instanceof SharePointReadOnlyError, `${verb} ${target}`);
      }
    }
    assert.strictEqual(mock.state.log.length, 0, 'not one request was sent');
    assert.strictEqual(mock.state.sharePointWrites, 0);
    assert.strictEqual((await client.request('POST', '/users/a@x.test/events', { body: { subject: 'x' } })).status, 201);
  });

  await check('paging: getAll follows @odata.nextLink', async () => {
    const mock = createMockGraph({ group: ['a@x.test'], pageSize: 2 });
    for (let i = 0; i < 5; i += 1) await mockClient(mock).request('POST', '/users/a@x.test/events', { body: { subject: `e${i}`, singleValueExtendedProperties: [{ id: PROPERTY_ID, value: `k${i}` }] } });
    const { items } = await mockClient(mock).getAll('/users/a@x.test/events', { params: { $filter: `singleValueExtendedProperties/Any(ep: ep/id eq '${PROPERTY_ID}')` } });
    assert.strictEqual(items.length, 5);
  });

  // ======== 6H.2 calendar =====================================================================================
  const owner = (name, email) => ({ id: name, name: name.toLowerCase(), fullName: name, email });
  const task = (id, extra = {}) => ({ id, name: '[ESHA] Soil Resistivity - Locations', owners: [owner('Anup Parekh', 'anup@elegrow.com')], startDate: '10-26-2026', dueDate: '10-28-2026', status: 'Planned', closed: false, testId: 'soilResistivity', ...extra });
  const project = (id, tasks, extra = {}) => ({ id, name: `0${id}-D03_TS-26-AAA`, clientName: 'AAA Ltd', tasks, ...extra });
  function calendarSetup({ tasks = [task('T1')], members = ['anup@elegrow.com', 'nidhi@elegrow.com'], appBaseUrl = 'https://reports.example.test', pageSize } = {}) {
    const store = newStore();
    const mock = createMockGraph({ group: members, pageSize });
    const graph = mockClient(mock);
    const sync = createCalendarSync({ store, graph, config: { ...BASE_CONFIG, appBaseUrl, hiddenTaskStatuses: ['Closed'] }, log: { error() {} } });
    const setTasks = (list) => store.setScheduleCache('tasks', { projects: list });
    setTasks([project('11', tasks)]);
    return { store, mock, graph, sync, setTasks };
  }
  const writes = (mock) => mock.state.log.filter(([method]) => method !== 'GET');
  const events = (mock, upn) => mock.calendar.events(upn);

  await check('events: subject, all-day date range (end is the day after), site as location, project link, workingElsewhere, the app\'s tag; one per owner', async () => {
    const t = calendarSetup({ tasks: [task('T1', { owners: [owner('Anup Parekh', 'anup@elegrow.com'), owner('Nidhi Shah', 'Nidhi@elegrow.com')] })] });
    t.store.addProject({ projectNo: '011-D03_TS-26-AAA', clientName: 'AAA Ltd', siteLocation: 'Dahej plant', zohoProjectId: '11' });
    const summary = await t.sync.run({ dryRun: false });
    assert.strictEqual(summary.counts.create, 2);
    const workspace = t.store.listProjects()[0];
    for (const upn of ['anup@elegrow.com', 'nidhi@elegrow.com']) {
      const [event] = events(t.mock, upn);
      assert.strictEqual(events(t.mock, upn).length, 1);
      assert.strictEqual(event.subject, 'Site: 011-D03_TS-26-AAA – Soil Resistivity Test');
      assert.deepStrictEqual([event.isAllDay, event.showAs], [true, 'workingElsewhere']);
      assert.deepStrictEqual([event.start, event.end], [{ dateTime: '2026-10-26T00:00:00', timeZone: 'India Standard Time' }, { dateTime: '2026-10-29T00:00:00', timeZone: 'India Standard Time' }]);
      assert.strictEqual(event.location.displayName, 'Dahej plant');
      assert.ok(event.body.content.includes(`https://reports.example.test/?project=${workspace.id}`));
      assert.deepStrictEqual(event.singleValueExtendedProperties, [{ id: PROPERTY_ID, value: `T1|${upn}` }]);
    }
    assert.strictEqual(t.store.listCalendarSync().length, 2);
  });

  await check('with no workspace the link carries the Zoho project id and the location falls back to the client; no APP_BASE_URL means no link', async () => {
    const t = calendarSetup();
    await t.sync.run({ dryRun: false });
    const [event] = events(t.mock, 'anup@elegrow.com');
    assert.strictEqual(event.location.displayName, 'AAA Ltd');
    assert.ok(event.body.content.endsWith('/?project=11'));
    const bare = calendarSetup({ appBaseUrl: '' });
    await bare.sync.run({ dryRun: false });
    assert.strictEqual(events(bare.mock, 'anup@elegrow.com')[0].body.content, '');
  });

  await check('off and dry-run by default: nothing is sent; Preview plans and writes nothing; the hook does nothing while off', async () => {
    const t = calendarSetup();
    assert.deepStrictEqual(t.sync.getSettings(), { enabled: false, dryRun: true });
    assert.strictEqual(await t.sync.onTasksRefreshed(), null);
    assert.strictEqual(t.mock.state.log.length, 0, 'off: not one request');
    const preview = await t.sync.run({ dryRun: true });
    assert.deepStrictEqual([preview.dryRun, preview.counts.create, preview.items[0].action, preview.items[0].project], [true, 1, 'create', '011-D03_TS-26-AAA']);
    assert.deepStrictEqual(writes(t.mock), []);
    assert.deepStrictEqual(t.store.listCalendarSync(), []);
    t.sync.saveSettings({ enabled: true });
    assert.strictEqual((await t.sync.onTasksRefreshed()).dryRun, true, 'on but still dry run: plans only');
    assert.deepStrictEqual(writes(t.mock), []);
    t.sync.saveSettings({ dryRun: false }, 'admin');
    assert.strictEqual((await t.sync.onTasksRefreshed()).counts.create, 1);
    assert.strictEqual(events(t.mock, 'anup@elegrow.com').length, 1);
  });

  await check('idempotent: a second run sends no write; changed dates update (PATCH); a user-deleted event comes back', async () => {
    const t = calendarSetup();
    await t.sync.run({ dryRun: false });
    const before = writes(t.mock).length;
    const again = await t.sync.run({ dryRun: false });
    assert.deepStrictEqual([again.counts.unchanged, again.counts.create + again.counts.update + again.counts.delete], [1, 0]);
    assert.strictEqual(writes(t.mock).length, before);
    t.setTasks([project('11', [task('T1', { startDate: '10-27-2026', dueDate: '10-27-2026' })])]);
    const moved = await t.sync.run({ dryRun: false });
    assert.strictEqual(moved.counts.update, 1);
    const [event] = events(t.mock, 'anup@elegrow.com');
    assert.deepStrictEqual([event.start.dateTime, event.end.dateTime], ['2026-10-27T00:00:00', '2026-10-28T00:00:00']);
    assert.ok(writes(t.mock).slice(before).every(([method]) => method === 'PATCH'));
    t.mock.state.events.set('anup@elegrow.com', []);
    assert.strictEqual((await t.sync.run({ dryRun: false })).counts.create, 1, 'recreated');
  });

  await check('removed, closed, hidden-status, dates cleared, or a new owner: the old event is deleted (and only that one)', async () => {
    const t = calendarSetup({ tasks: [task('T1'), task('T2', { owners: [owner('Nidhi Shah', 'nidhi@elegrow.com')] }), task('T3'), task('T4')] });
    await t.sync.run({ dryRun: false });
    assert.strictEqual(events(t.mock, 'anup@elegrow.com').length, 3);
    t.setTasks([project('11', [
      task('T1', { closed: true }),                                // closed
      task('T2', { owners: [owner('Anup Parekh', 'anup@elegrow.com')] }), // owner changed
      task('T3', { startDate: '', dueDate: '' }),                  // dates cleared
      task('T4', { status: 'Closed' })                              // hidden status
    ])]);
    const summary = await t.sync.run({ dryRun: false });
    assert.deepStrictEqual([summary.counts.delete, summary.counts.create], [4, 1]);
    assert.deepStrictEqual(events(t.mock, 'nidhi@elegrow.com'), [], 'deleted from the old owner');
    assert.deepStrictEqual(events(t.mock, 'anup@elegrow.com').map((event) => event.singleValueExtendedProperties[0].value), ['T2|anup@elegrow.com'], 'created for the new owner');
    t.setTasks([]);
    await t.sync.run({ dryRun: false });
    assert.deepStrictEqual(events(t.mock, 'anup@elegrow.com'), []);
    assert.deepStrictEqual(t.store.listCalendarSync(), []);
  });

  await check('NEVER touches an event the app did not create: not listed, not updated, not deleted', async () => {
    const t = calendarSetup();
    t.mock.calendar.addForeignEvent('anup@elegrow.com', { subject: 'Site: 011-D03_TS-26-AAA – Soil Resistivity Test', start: { dateTime: '2026-10-26T00:00:00' } }); // same subject, no tag
    t.mock.calendar.addForeignEvent('anup@elegrow.com', { subject: 'Dentist' });
    const foreignIds = events(t.mock, 'anup@elegrow.com').map((event) => event.id);
    await t.sync.run({ dryRun: false });
    t.setTasks([project('11', [task('T1', { startDate: '10-30-2026', dueDate: '10-30-2026' })])]);
    await t.sync.run({ dryRun: false });
    t.setTasks([]);
    await t.sync.run({ dryRun: false });
    const remaining = events(t.mock, 'anup@elegrow.com');
    assert.deepStrictEqual(remaining.map((event) => event.id), foreignIds, 'both foreign events are still there, unchanged');
    assert.strictEqual(remaining[0].start.dateTime, '2026-10-26T00:00:00');
    const touched = t.mock.state.log.filter(([method, p]) => ['PATCH', 'DELETE'].includes(method) && foreignIds.some((id) => p.endsWith(`/${id}`)));
    assert.deepStrictEqual(touched, []);
    assert.ok(t.mock.state.log.filter(([method]) => method === 'GET').length > 0);
  });

  await check('a delete is read back first: an event that no longer carries the app\'s tag is refused, not deleted', async () => {
    const t = calendarSetup({ tasks: [] });
    t.mock.calendar.addForeignEvent('anup@elegrow.com', { subject: 'Not ours' });
    const [foreign] = events(t.mock, 'anup@elegrow.com');
    // a server that ignored the property filter would hand us this event as "ours"
    const lying = { ...t.graph, getAll: async () => ({ items: [{ id: foreign.id, singleValueExtendedProperties: [{ id: PROPERTY_ID, value: 'T9|anup@elegrow.com' }] }] }) };
    const sync = createCalendarSync({ store: t.store, graph: lying, config: { ...BASE_CONFIG, hiddenTaskStatuses: ['Closed'] }, log: { error() {} } });
    t.store.saveCalendarSync('T9|anup@elegrow.com', { taskId: 'T9', ownerEmail: 'anup@elegrow.com', eventId: foreign.id, hash: 'x', state: 'done' });
    const summary = await sync.run({ dryRun: false });
    assert.deepStrictEqual([summary.counts.delete, summary.counts.failed], [0, 1]);
    assert.match(summary.failures[0].error, /not created by this app/);
    assert.strictEqual(events(t.mock, 'anup@elegrow.com').length, 1);
    assert.deepStrictEqual(writes(t.mock), [], 'no DELETE was sent');
  });

  await check('a mailbox outside the Exchange-scoped group (403) or unknown (404): skipped with a warning, nothing deleted there, the rest still synced', async () => {
    const t = calendarSetup({ tasks: [task('T1'), task('T2', { owners: [owner('Outside Person', 'outside@elegrow.com')] }), task('T3', { owners: [owner('Ghost', 'ghost@elegrow.com')] })] });
    t.mock.calendar.addOutsideMailbox('outside@elegrow.com');
    const summary = await t.sync.run({ dryRun: false });
    assert.deepStrictEqual([summary.counts.create, summary.counts.skip], [1, 2]);
    assert.ok(summary.warnings.some((warning) => /403: ErrorAccessDenied/.test(warning)) && summary.warnings.some((warning) => /404: ErrorInvalidUser/.test(warning)));
    assert.strictEqual(events(t.mock, 'anup@elegrow.com').length, 1);
    assert.deepStrictEqual(t.mock.state.log.filter(([method, p]) => method === 'POST' && /outside|ghost/.test(p)), []);
    // an owner without an e-mail is skipped with a warning, not an error
    const none = calendarSetup({ tasks: [task('T1', { owners: [{ id: '1', name: 'x', fullName: 'No Mail' }] })] });
    const result = await none.sync.run({ dryRun: true });
    assert.ok(result.warnings.some((warning) => /owner without an e-mail/.test(warning)) && result.counts.create === 0);
  });

  await check('paging and table loss: with the local table gone, the app\'s own events are still found by their tag and not duplicated', async () => {
    const t = calendarSetup({ tasks: [task('T1'), task('T2'), task('T3'), task('T4'), task('T5')], pageSize: 2 });
    await t.sync.run({ dryRun: false });
    t.store.listCalendarSync().forEach((row) => t.store.deleteCalendarSync(row.key));
    const summary = await t.sync.run({ dryRun: false });
    assert.strictEqual(events(t.mock, 'anup@elegrow.com').length, 5);
    assert.deepStrictEqual([summary.counts.create, summary.counts.update], [0, 5], 'found and brought back in line, not recreated');
  });

  await check('reversed dates and a single date: a reversed task is left out with a warning, one date is a one-day event', async () => {
    const t = calendarSetup({ tasks: [task('T1', { startDate: '10-28-2026', dueDate: '10-26-2026' }), task('T2', { startDate: '', dueDate: '10-27-2026' })] });
    const summary = await t.sync.run({ dryRun: false });
    assert.strictEqual(summary.counts.create, 1);
    assert.ok(summary.warnings.some((warning) => /ends before it starts/.test(warning)));
    assert.deepStrictEqual([events(t.mock, 'anup@elegrow.com')[0].start.dateTime, events(t.mock, 'anup@elegrow.com')[0].end.dateTime], ['2026-10-27T00:00:00', '2026-10-28T00:00:00']);
  });

  // ======== 6H.3 / 6H.4 knowledge base ===============================================================
  await check('both project folder styles are matched to one key; section rules; Working is never indexed', () => {
    assert.strictEqual(projectKeyOf('056-D03_TS-26-CLL'), '056-D03_TS-26-CLL');
    assert.strictEqual(projectKeyOf('056-d03_ts-26-cll Dahej'), '056-D03_TS-26-CLL');
    assert.strictEqual(projectKeyOf('294_04_23_TS_APL_ESHA'), '294_04_23_TS_APL_ESHA');
    assert.strictEqual(projectKeyOf('Misc'), '');
    const client = { scope: 'client', folders: [] };
    const at = (...segments) => classify(client, ['Shared Documents', ...segments]);
    ['Submitted', 'Reference', 'Input', 'submitted'].forEach((section) => assert.strictEqual(at('056-D03_TS-26-CLL', section, 'a.pdf').include, true, section));
    assert.deepStrictEqual(at('056-D03_TS-26-CLL', 'Submitted', 'Sub', 'deep', 'a.docx'), { include: true, ext: 'docx', projectKey: '056-D03_TS-26-CLL', section: 'submitted' });
    assert.strictEqual(at('056-D03_TS-26-CLL', 'Working', 'a.pdf').include, false);
    assert.strictEqual(at('056-D03_TS-26-CLL', 'Submitted', 'Working', 'a.pdf').include, false, 'Working anywhere');
    assert.strictEqual(at('056-D03_TS-26-CLL', 'a.pdf').include, false, 'a file directly in the project folder');
    assert.strictEqual(at('056-D03_TS-26-CLL', 'Other', 'a.pdf').include, false);
    assert.strictEqual(at('294_04_23_TS_APL_ESHA', 'Reference', 'a.xlsx').projectKey, '294_04_23_TS_APL_ESHA');
    assert.deepStrictEqual(at('Misc folder', 'a.pdf'), { include: true, ext: 'pdf', projectKey: '', section: '' }, 'unmatched: indexed with no project');
    assert.strictEqual(at('056-D03_TS-26-CLL', 'Submitted', 'a.png').include, false, 'only pdf, docx, xlsx');
    const general = { scope: 'general', folders: ['Shared Documents/ETPL Knowledge Hub/02 Technical Knowledge'] };
    assert.strictEqual(classify(general, ['Shared Documents', 'ETPL Knowledge Hub', '02 Technical Knowledge', 'IS 3043', 'a.pdf']).include, true);
    assert.strictEqual(classify(general, ['Shared Documents', 'ETPL Knowledge Hub', '01 Other', 'a.pdf']).include, false);
    assert.strictEqual(classify(general, ['Shared Documents', 'a.pdf']).include, false);
  });

  await check('text comes out of real PDF, DOCX and XLSX bytes; chunks are bounded and overlap', async () => {
    assert.match(await extractText('pdf', await makePdf(['Earth electrode resistance limit', 'Second paragraph'])), /Earth electrode resistance limit\nSecond paragraph/);
    assert.strictEqual(await extractText('docx', makeDocx(['Hello & <goodbye>', 'Two'])), 'Hello & <goodbye>\nTwo');
    assert.strictEqual(await extractText('xlsx', makeXlsx([['Pit', 'Location'], ['P-1', 12.5]])), 'Pit Location\nP-1 12.5');
    assert.strictEqual(await extractText('pdf', await makePdf('')), '', 'a PDF with no text layer');
    await assert.rejects(() => extractText('docx', Buffer.from('not a zip')));
    const chunks = chunkText('alpha '.repeat(900));
    assert.ok(chunks.length > 3 && chunks.every((chunk) => chunk.length <= 1000));
    assert.ok(chunks[1].startsWith('alpha') && chunks.join(' ').length > 5400);
  });

  const fixtureKb = async ({ pageSize, configure } = {}) => {
    const store = newStore();
    const mock = createMockGraph({ group: [], pageSize });
    const sites = [
      { name: 'OP', url: 'https://t.sharepoint.com/sites/OP', folders: [], scope: 'client' },
      { name: 'OLD', url: 'https://t.sharepoint.com/sites/OLD', folders: [], scope: 'client' },
      { name: 'HUB', url: 'https://t.sharepoint.com/sites/HUB', folders: ['Shared Documents/Hub/02 Technical Knowledge'], scope: 'general' }
    ];
    sites.forEach((site) => mock.sp.addSite(site.name, site.url));
    const pdf = (text) => () => makePdf(text);
    mock.sp.addFile('OP', 'Shared Documents/024-D03_TS-26-GFL/Submitted/GFL final report.pdf', pdf('GFL Dahej pit resistance and continuity results'));
    mock.sp.addFile('OP', 'Shared Documents/024-D03_TS-26-GFL/Working/GFL draft.pdf', pdf('GFL draft only zebra'));
    mock.sp.addFile('OP', 'Shared Documents/056-D03_TS-26-CLL/Input/CLL data.xlsx', () => makeXlsx([['Pit', 'Yard'], ['P-1', 'Substation']]));
    mock.sp.addFile('OP', 'Shared Documents/Loose folder/loose note.docx', () => makeDocx('loose note about pit resistance'));
    mock.sp.addFile('OP', 'Shared Documents/056-D03_TS-26-CLL/Submitted/scan.pdf', pdf(''));
    mock.sp.addFile('OLD', 'Shared Documents/294_04_23_TS_APL_ESHA/Reference/APL reference.pdf', pdf('APL older assessment pit resistance'));
    mock.sp.addFile('HUB', 'Shared Documents/Hub/02 Technical Knowledge/Earthing practice.pdf', pdf('Earth electrode resistance practice pit resistance'));
    mock.sp.addFile('HUB', 'Shared Documents/Hub/01 Other/not indexed.pdf', pdf('pit resistance outside the folder'));
    const kb = createKnowledgeBase({ graph: mockClient(mock), config: { kb: { sites, refreshMinutes: 0, maxMb: 40, dbPath: ':memory:' }, ...(configure || {}) }, store, log: { error() {} } });
    return { store, mock, kb, sites };
  };
  const titles = (hits) => hits.map((hit) => hit.title).sort();
  const everything = { all: true };

  await check('refresh indexes only what the rules allow (Working, other folders, folders outside the hub path are not), flags scans, and never writes to SharePoint', async () => {
    const { mock, kb } = await fixtureKb({ pageSize: 3 });
    const summary = await kb.refresh();
    assert.deepStrictEqual(summary.sites.map((site) => [site.name, site.added, site.noText, site.errors, site.error]), [['OP', 3, 1, 0, ''], ['OLD', 1, 0, 0, ''], ['HUB', 1, 0, 0, '']]);
    assert.deepStrictEqual(titles(kb.search({ query: 'pit resistance', access: everything })), ['APL reference', 'Earthing practice', 'GFL final report', 'loose note']);
    assert.deepStrictEqual(kb.search({ query: 'zebra', access: everything }), [], 'Working is not in the index');
    assert.deepStrictEqual(kb.search({ query: 'outside the folder', access: everything }), []);
    const hit = kb.search({ query: 'dahej', access: everything })[0];
    assert.deepStrictEqual([hit.title, hit.source, hit.projectKey, hit.scope], ['GFL final report', 'OP', '024-D03_TS-26-GFL', 'client']);
    assert.match(hit.webUrl, /^https:\/\/t\.sharepoint\.com\/sites\/OP\/Shared%20Documents\/024-D03_TS-26-GFL\/Submitted\//);
    assert.ok(hit.snippet.includes('Dahej') || hit.snippet.toLowerCase().includes('gfl'));
    assert.strictEqual(kb.search({ query: 'substation', access: everything })[0].title, 'CLL data', 'xlsx text is searchable');
    assert.strictEqual(kb.status().sites.find((site) => site.name === 'OP').noText, 1);
    assert.deepStrictEqual(kb.search({ query: '"; DROP TABLE kb_files; --', access: everything }), [], 'the query is tokenised, not interpreted');
    assert.strictEqual(mock.state.sharePointWrites, 0);
    assert.ok(mock.state.log.filter(([method, p]) => /^\/(sites|drives)/.test(p) && method !== 'GET').length === 0);
  });

  await check('incremental refresh by delta: nothing downloaded when nothing changed; an edit re-reads that file only; a rename keeps the text; a deletion removes it; a move into Working removes it', async () => {
    const { mock, kb } = await fixtureKb();
    await kb.refresh();
    const downloads = () => mock.state.log.filter(([, p]) => p.endsWith('/content')).length;
    const base = downloads();
    const quiet = await kb.refresh();
    assert.deepStrictEqual(quiet.sites.map((site) => [site.added, site.updated, site.removed]), [[0, 0, 0], [0, 0, 0], [0, 0, 0]]);
    assert.strictEqual(downloads(), base, 'no file was downloaded again');
    assert.ok(mock.state.log.some(([, p]) => p.endsWith('/root/delta')) && mock.state.log.filter(([, p]) => p.endsWith('/root/delta')).length >= 6);
    mock.sp.addFile('OP', 'Shared Documents/024-D03_TS-26-GFL/Submitted/GFL final report.pdf', () => makePdf('GFL revised report with earth mat findings'), { modified: '2026-10-20T09:00:00Z' });
    const edited = await kb.refresh();
    assert.deepStrictEqual([edited.sites[0].updated, downloads() - base], [1, 1]);
    assert.strictEqual(kb.search({ query: 'earth mat', access: everything })[0].title, 'GFL final report');
    assert.deepStrictEqual(kb.search({ query: 'dahej', access: everything }), [], 'the old text is gone');
    mock.sp.moveFile('OP', 'Shared Documents/024-D03_TS-26-GFL/Submitted/GFL final report.pdf', 'Shared Documents/024-D03_TS-26-GFL/Reference/GFL moved.pdf');
    const moved = await kb.refresh();
    assert.deepStrictEqual([moved.sites[0].updated, downloads() - base], [1, 1], 'metadata only, no new download');
    assert.strictEqual(kb.search({ query: 'earth mat', access: everything })[0].path, 'Shared Documents/024-D03_TS-26-GFL/Reference/GFL moved.pdf');
    mock.sp.moveFile('OP', 'Shared Documents/024-D03_TS-26-GFL/Reference/GFL moved.pdf', 'Shared Documents/024-D03_TS-26-GFL/Working/GFL moved.pdf');
    assert.strictEqual((await kb.refresh()).sites[0].removed, 1);
    assert.deepStrictEqual(kb.search({ query: 'earth mat', access: everything }), []);
    mock.sp.removeFile('OLD', 'Shared Documents/294_04_23_TS_APL_ESHA/Reference/APL reference.pdf');
    assert.strictEqual((await kb.refresh()).sites[1].removed, 1);
    assert.deepStrictEqual(kb.search({ query: 'APL older', access: everything }), []);
    assert.strictEqual(mock.state.sharePointWrites, 0);
  });

  await check('an expired delta token (410) triggers a full resync that also sweeps what vanished meanwhile', async () => {
    const { mock, kb } = await fixtureKb();
    await kb.refresh();
    mock.sp.removeFile('OP', 'Shared Documents/Loose folder/loose note.docx');
    mock.sp.expireDelta();
    const summary = await kb.refresh();
    assert.strictEqual(summary.sites[0].error, '');
    assert.deepStrictEqual(kb.search({ query: 'loose note', access: everything }), []);
    assert.ok(kb.search({ query: 'dahej', access: everything }).length === 1, 'the rest was re-indexed');
  });

  await check('a site that cannot be read keeps its old index and records why; the other sites go on', async () => {
    const { mock, kb } = await fixtureKb();
    await kb.refresh();
    mock.state.sites.get('OP').granted = false;
    const summary = await kb.refresh();
    assert.strictEqual(summary.sites[0].error, '403 accessDenied');
    assert.strictEqual(summary.sites[2].error, '');
    assert.ok(kb.search({ query: 'dahej', access: everything }).length === 1, 'old index kept');
    assert.strictEqual(kb.status().sites[0].error, '403 accessDenied');
  });

  await check('who finds what: reviewer, approver and admin search everything; an engineer finds the general documents and the projects they worked on; unmatched client files are reviewer and above only', async () => {
    const { store, kb } = await fixtureKb();
    await kb.refresh();
    const engineer = store.addUser({ email: 'eng@elegrow.com', name: 'Eng One', roles: ['engineer'] });
    const other = store.addUser({ email: 'other@elegrow.com', name: 'Eng Two', roles: ['engineer'] });
    const reviewer = { id: 'r', email: 'r@elegrow.com', roles: ['reviewer'] };
    assert.deepStrictEqual(kb.accessFor(reviewer), { all: true });
    assert.deepStrictEqual(kb.accessFor({ id: 'a', roles: ['approver'] }), { all: true });
    assert.deepStrictEqual(kb.accessFor({ id: 'x', roles: ['admin'] }), { all: true });
    assert.deepStrictEqual(kb.accessFor(engineer), { projectKeys: [] });
    const find = (user) => titles(kb.search({ query: 'pit resistance', access: kb.accessFor(user) }));
    assert.deepStrictEqual(find(engineer), ['Earthing practice'], 'no project yet: general only');
    // worked on GFL (created a day there)
    const gfl = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'GFL' });
    store.addReport({ ...normalizeReportInput(buildPayload()), kind: 'daily', projectId: gfl.id, collectionDate: '2026-10-01', createdBy: engineer.id });
    assert.deepStrictEqual(find(engineer), ['Earthing practice', 'GFL final report']);
    assert.deepStrictEqual(find(other), ['Earthing practice']);
    // credited on a row of a day someone else created
    const stamped = normalizeReportInput(buildPayload());
    stamped.electrodeResistance[0].enteredBy = { userId: other.id, name: 'Eng Two', email: other.email };
    const cll = store.addProject({ projectNo: '056-D03_TS-26-CLL', clientName: 'CLL' });
    store.addReport({ ...stamped, kind: 'daily', projectId: cll.id, collectionDate: '2026-10-02', createdBy: engineer.id });
    assert.ok(kb.search({ query: 'substation', access: kb.accessFor(other) }).length === 1, 'row credit counts as having worked on it');
    // owner of an Onsite task in Zoho
    store.setScheduleCache('tasks', { projects: [{ id: '5', name: '294_04_23_TS_APL_ESHA', clientName: 'APL', tasks: [{ id: 't', owners: [{ email: 'Other@elegrow.com' }] }] }] });
    assert.ok(find(other).includes('APL reference'), 'task ownership counts, by e-mail');
    assert.ok(!find(engineer).includes('loose note') && !find(other).includes('loose note'), 'an unmatched client file is not for engineers');
    assert.ok(find(reviewer).includes('loose note'));
    // the same rules for opening one file
    const looseId = kb.search({ query: 'loose note', access: everything })[0].id;
    assert.strictEqual(kb.getFile(looseId, kb.accessFor(engineer)), null);
    assert.strictEqual(kb.getFile(looseId, kb.accessFor(reviewer)).title, 'loose note');
  });

  await check('the knowledge-base code has no write call to SharePoint at all (static scan) and imports nothing that could', () => {
    ['knowledgeBase.js', 'kbExtract.js'].forEach((file) => {
      const source = fs.readFileSync(path.join(__dirname, '..', 'src', file), 'utf8');
      assert.ok(!/\.(post|put|patch|delete|request)\(/.test(source), `${file} makes only get calls`);
      assert.ok(!/method\s*:\s*['"](POST|PUT|PATCH|DELETE)/i.test(source), `${file} names no write method`);
      assert.ok(!/require\(['"]axios['"]\)|fs\.(write|append|unlink|rename|copyFile)/.test(source.replace(/fs\.mkdirSync/g, '')), `${file} writes no file and talks only through the client`);
    });
    const client = fs.readFileSync(path.join(__dirname, '..', 'src', 'graphClient.js'), 'utf8');
    assert.ok(/SHAREPOINT_PATH\.test\(path\) && verb !== 'GET'/.test(client), 'the guard is in the one client');
  });

  // ======== 6H.5 references and routes ==================================================================
  await check('the final PDF lists the references (title, location, date) and nothing else from them; without references it has no such section', async () => {
    const store = newStore();
    const proj = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'GFL', unit: 'FI', unitCode: 'FI' });
    const payload = buildPayload();
    store.addReport({ ...normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: '2026-10-01' } }), kind: 'daily', projectId: proj.id, collectionDate: '2026-10-01' });
    const { report } = store.assembleFinal(proj.id, { actor: 'lead' });
    const text = async (r) => {
      const doc = await openPdf((await generateFinalReportPdf(r, { status: 'draft' })).buffer);
      let all = '';
      for (let page = 0; page < pageCount(doc); page += 1) all += ` ${pageTextLines(doc, page).map((fragment) => fragment.text).join(' ')}`;
      return all;
    };
    const plain = await text(report);
    assert.ok(!/References/.test(plain.replace(/Reference ?[A-Za-z]+/g, '')) || !/Documents the reviewer selected/.test(plain));
    const withRefs = await text({ ...report, references: [{ fileId: 1, title: 'Earthing practice notes', source: 'HUB', path: 'Shared Documents/Hub/02 Technical Knowledge/Earthing practice notes.pdf', modified: '2026-09-30T10:00:00Z', webUrl: 'https://x', snippet: 'SECRET SNIPPET 9.99 ohm' }] });
    assert.ok(/Documents the reviewer selected as further reading/.test(withRefs) && /Earthing practice notes/.test(withRefs) && /2026-09-30/.test(withRefs));
    assert.ok(!/SECRET SNIPPET/.test(withRefs), 'no passage text ever reaches the report');
  });

  // ---- HTTP, on the mock portal and the mock Graph ----
  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: path.join(__dirname, '..'),
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'http', 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', GRAPH_USE_MOCK: 'true', GRAPH_KB_SITES: '', GRAPH_KB_REFRESH_MINUTES: '0', APP_BASE_URL: 'https://reports.example.test', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  try {
    await new Promise((resolve, reject) => {
      server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
      setTimeout(() => reject(new Error('server did not start')), 15000);
    });
    const call = createCaller(`http://127.0.0.1:${port}`);

    await check('routes: the admin panels are admin only; the knowledge base refresh indexes the demo sites', async () => {
      for (const [method, url] of [['GET', '/api/admin/kb'], ['POST', '/api/admin/kb/refresh'], ['GET', '/api/admin/calendar-sync'], ['PUT', '/api/admin/calendar-sync'], ['POST', '/api/admin/calendar-sync/preview'], ['POST', '/api/admin/calendar-sync/run']]) {
        for (const as of ['engineer@elegrow.com', 'reviewer@elegrow.com', 'approver@elegrow.com']) assert.strictEqual((await call(method, url, method === 'GET' ? undefined : {}, as)).status, 403, `${method} ${url} as ${as}`);
      }
      const before = (await call('GET', '/api/admin/kb')).body;
      assert.deepStrictEqual(before.sites.map((site) => site.files), [0, 0, 0]);
      const refreshed = await call('POST', '/api/admin/kb/refresh', {});
      assert.strictEqual(refreshed.status, 200);
      assert.ok(refreshed.body.sites.every((site) => site.error === '') && refreshed.body.sites.reduce((sum, site) => sum + site.added, 0) >= 4);
      assert.ok((await call('GET', '/api/admin/audit?action=kb.')).body.entries.some((entry) => entry.action === 'kb.refresh'));
    });

    await check('routes: search is filtered on the server by role and by the projects the engineer worked on', async () => {
      const find = async (as, q = 'resistance') => (await call('GET', `/api/kb/search?q=${encodeURIComponent(q)}`, undefined, as)).body.hits.map((hit) => hit.title).sort();
      assert.deepStrictEqual(await find('engineer@elegrow.com'), ['Earthing practice notes'], 'general only until they work on a project');
      const created = await call('POST', '/api/projects', { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' }, 'engineer@elegrow.com');
      const payload = buildPayload();
      const day = await call('POST', `/api/projects/${created.body.projectId}/days`, { workDate: '2026-10-12', report: { ...payload, project: { ...payload.project, reportDate: '2026-10-12' } } }, 'engineer@elegrow.com');
      assert.strictEqual(day.status, 201);
      assert.deepStrictEqual(await find('engineer@elegrow.com'), ['Earthing practice notes', 'GFL ESHA report']);
      assert.deepStrictEqual(await find('engineer2@elegrow.com'), ['Earthing practice notes'], 'another engineer does not see it');
      const all = await find('reviewer@elegrow.com', 'assessment');
      assert.ok(all.includes('APL reference') && all.includes('GFL ESHA report'));
      assert.deepStrictEqual((await call('GET', '/api/kb/search?q=')).body.hits, []);
      const draft = await find('admin@elegrow.com', 'draft notes');
      assert.deepStrictEqual(draft, [], 'Working is not in the index');
    });

    await check('routes: references are picked by a reviewer for a final report, only from the report\'s own project and the general base, and are locked after approval', async () => {
      const projects = (await call('GET', '/api/projects')).body;
      const gfl = projects.find((entry) => entry.projectNo === '024-D03_TS-26-GFL');
      const assembled = await call('POST', `/api/projects/${gfl.id}/final/assemble`, {}, 'engineer@elegrow.com');
      assert.strictEqual(assembled.status, 201, JSON.stringify(assembled.body).slice(0, 300));
      const finalId = assembled.body.report.id;
      const hits = (await call('GET', `/api/reports/${finalId}/kb-search?q=resistance`, undefined, 'reviewer@elegrow.com')).body.hits;
      assert.deepStrictEqual(hits.map((hit) => hit.title).sort(), ['Earthing practice notes', 'GFL ESHA report'], 'another client\'s documents are not even offered');
      assert.strictEqual((await call('GET', `/api/reports/${finalId}/kb-search?q=resistance`, undefined, 'engineer@elegrow.com')).status, 403);
      const general = hits.find((hit) => hit.title === 'Earthing practice notes');
      const own = hits.find((hit) => hit.title === 'GFL ESHA report');
      const foreign = (await call('GET', '/api/kb/search?q=reference', undefined, 'reviewer@elegrow.com')).body.hits.find((hit) => hit.title === 'APL reference');
      assert.ok(foreign);
      assert.strictEqual((await call('PUT', `/api/reports/${finalId}/references`, { fileIds: [general.id] }, 'engineer@elegrow.com')).status, 403, 'engineers cannot pick');
      const rejected = await call('PUT', `/api/reports/${finalId}/references`, { fileIds: [general.id, foreign.id] }, 'reviewer@elegrow.com');
      assert.deepStrictEqual([rejected.status, rejected.body.code], [400, 'REFERENCE_NOT_ALLOWED']);
      assert.strictEqual((await call('GET', `/api/reports/${finalId}`)).body.references, undefined, 'a rejected request changes nothing');
      const saved = await call('PUT', `/api/reports/${finalId}/references`, { fileIds: [general.id, own.id, own.id] }, 'reviewer@elegrow.com');
      assert.strictEqual(saved.status, 200);
      assert.deepStrictEqual(saved.body.references.map((entry) => entry.title), ['Earthing practice notes', 'GFL ESHA report']);
      assert.ok(saved.body.references.every((entry) => entry.addedBy && entry.webUrl && entry.path));
      assert.deepStrictEqual((await call('GET', `/api/reports/${finalId}`)).body.references.map((entry) => entry.fileId), [general.id, own.id]);
      assert.strictEqual((await call('PUT', `/api/reports/${finalId}/references`, { fileIds: [own.id] }, 'reviewer@elegrow.com')).body.references.length, 1, 'a removal');
      const daily = (await call('GET', `/api/projects/${gfl.id}`)).body.days[0];
      assert.strictEqual((await call('PUT', `/api/reports/${daily.reportId || daily.id}/references`, { fileIds: [] }, 'reviewer@elegrow.com')).status, 400, 'only a final report carries references');
    });

    await check('routes: the project link of a calendar event resolves a workspace id or a Zoho project id', async () => {
      const projects = (await call('GET', '/api/projects')).body;
      const gfl = projects.find((entry) => entry.projectNo === '024-D03_TS-26-GFL');
      assert.deepStrictEqual((await call('GET', `/api/projects/lookup/${gfl.id}`)).body.workspaceId, gfl.id);
      const unknown = (await call('GET', '/api/projects/lookup/987654')).body;
      assert.deepStrictEqual([unknown.workspaceId, unknown.zohoProjectId], ['', '987654']);
    });

    await check('routes: calendar sync settings are saved and audited; Run needs it switched on; Preview plans the mock portal\'s tasks without writing', async () => {
      assert.deepStrictEqual((await call('GET', '/api/admin/calendar-sync')).body.settings, { enabled: false, dryRun: true });
      assert.strictEqual((await call('POST', '/api/admin/calendar-sync/run', {})).status, 409);
      await call('POST', '/api/schedule/refresh', {}, 'engineer@elegrow.com'); // reads the mock Zoho tasks into the cache
      const preview = await call('POST', '/api/admin/calendar-sync/preview', {});
      assert.strictEqual(preview.status, 200);
      assert.ok(preview.body.dryRun && preview.body.counts.create + preview.body.counts.skip >= 1, JSON.stringify(preview.body.counts));
      assert.strictEqual((await call('GET', '/api/admin/calendar-sync')).body.owned, 0, 'a preview writes nothing');
      const on = await call('PUT', '/api/admin/calendar-sync', { enabled: true, dryRun: false });
      assert.deepStrictEqual(on.body.settings, { enabled: true, dryRun: false });
      const run = await call('POST', '/api/admin/calendar-sync/run', {});
      assert.deepStrictEqual([run.status, run.body.dryRun], [200, false]);
      assert.strictEqual((await call('GET', '/api/admin/calendar-sync')).body.owned, run.body.counts.create);
      const audit = (await call('GET', '/api/admin/audit?action=calendar.')).body.entries.map((entry) => entry.action);
      assert.ok(audit.includes('calendar.settings') && audit.includes('calendar.run'));
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
