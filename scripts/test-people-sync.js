// Session 6G, Site Work Log in Zoho People, all on mocks (nothing here reaches Zoho). One People record per
// engineer + work date + test, written on approval only: who is credited (first enterer, set once; lastEditedBy
// for audit only), the aggregate and its figures (completed, planned, returns, data gaps, on-time, first-time
// approved), no write on submit / return / re-issue, no duplicate by Record Key, a later revision updating the
// same record, a quantity that falls to 0, a day split between two engineers, the field discovery by label,
// failures that never block, the Backfill, the totals, the People client's write calls, and the HTTP routes.
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { buildReviewState } = require('../src/review');
const { buildFinalReportData } = require('../src/finalReport/data');
const { stampReport, listUnits } = require('../src/entryCredit');
const { createPeopleSync, isSampleReport } = require('../src/peopleSync');
const { createPeopleClient, lookupAttribute, normalizeEmployee, normalizeFormField, WORKLOG_LABELS } = require('../src/zohoPeople');
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

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'people-sync-'));
const config = { zoho: { useMock: true, workLog: { form: 'Site_Work_Log', fields: {} } } };
const apiName = (label) => `f_${label.replace(/[^A-Za-z0-9]+/g, '_')}`; // the double's API names: nothing is assumed about them
const F = Object.fromEntries(Object.entries(WORKLOG_LABELS).map(([key, label]) => [key, apiName(label)]));

// A People double that records every call. `formExists`, the form's fields and the employee list can be changed per test.
function fakePeople() {
  const records = new Map();
  const calls = [];
  let seq = 0;
  const people = {
    calls,
    records,
    formExists: true,
    lookup: null,
    labels: Object.values(WORKLOG_LABELS),
    employees: [
      { email: 'anup@elegrow.com', recordId: 'P-A', employeeId: 'ELE-A', name: 'Anup Parekh', location: 'Surat' },
      { email: 'nidhi@elegrow.com', recordId: 'P-N', employeeId: 'ELE-N', name: 'Nidhi Shah', location: 'Surat' }
    ],
    failNext: 0,
    async hasForm() { calls.push('hasForm'); return people.formExists; },
    async listEmployees() { return people.employees; },
    fields: null,
    async getFormFields() { calls.push('getFormFields'); if (people.fields) return people.fields; return people.labels.map((label) => ({ label, apiName: apiName(label), type: label === 'Engineer' ? 'lookup' : 'text', lookup: label === 'Engineer' ? people.lookup : null })); },
    async findRecord(_form, field, key) { calls.push('find'); const hit = [...records].find(([, value]) => value[field] === key); return hit ? hit[0] : null; },
    async insertRecord(_form, data) { calls.push('insert'); if (people.failNext > 0) { people.failNext -= 1; throw new Error('People is down'); } const id = `R${++seq}`; records.set(id, { ...data }); return id; },
    async updateRecord(_form, id, data) { calls.push('update'); if (people.failNext > 0) { people.failNext -= 1; throw new Error('People is down'); } records.set(id, { ...records.get(id), ...data }); return id; },
    mockRecords: () => []
  };
  return people;
}
const verbs = (people) => people.calls.filter((call) => call === 'insert' || call === 'update');
const rec = (people) => [...people.records.values()];
const byTest = (people, testLabel) => rec(people).filter((record) => record[F.test] === testLabel);

function setup() {
  const store = createReportStore({ filePath: path.join(dir, `${Math.random().toString(36).slice(2)}.db`) });
  const anup = store.addUser({ email: 'anup@elegrow.com', name: 'Anup Parekh', roles: ['engineer'] });
  const nidhi = store.addUser({ email: 'nidhi@elegrow.com', name: 'Nidhi Shah', roles: ['engineer'] });
  const project = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E', zohoProjectName: '024-D03_TS-26-GFL Dahej earthing', plannedQuantities: { electrodeResistance: 20, soilResistivity: 4 } });
  const people = fakePeople();
  const sync = createPeopleSync({ store, people, config });
  const actor = (user) => ({ id: user.id, name: user.name, email: user.email });
  function daily(date, by = anup, { projectId = project.id, projectExtra = {} } = {}) {
    const payload = buildPayload();
    const normalized = normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: date, clientId: 'Z-77', ...projectExtra } });
    return store.addReport({ ...stampReport(normalized, null, actor(by)), kind: 'daily', projectId, collectionDate: date, createdBy: by.id });
  }
  function review(id, steps) {
    const ticks = () => Object.fromEntries(buildReviewState(store, store.getReport(id)).checklist.filter((item) => item.manual).map((item) => [item.id, true]));
    const status = store.getReport(id).reviewStatus;
    if (status === 'draft') store.transition(id, 'in_review', { actor: 'P' });
    store.setReviewTicks(id, ticks());
    store.markChecked(id, { actor: 'C' });
    for (let i = 0; i < steps.returns; i += 1) {
      store.transition(id, 'draft', { actor: 'C', note: 'redo' });
      store.transition(id, 'in_review', { actor: 'P' });
      store.setReviewTicks(id, ticks());
      store.markChecked(id, { actor: 'C' });
    }
    return store.transition(id, 'approved', { actor: 'A' }).report;
  }
  const approve = (report, steps = { returns: 0 }) => review(report.id, steps);
  // Saves `mutate(payload)` over the stored report the way the PUT route does, as `by`.
  function edit(id, by, mutate) {
    const payload = JSON.parse(JSON.stringify(store.getReport(id)));
    mutate(payload);
    return store.updateReport(id, (current) => stampReport(normalizeReportInput({ ...payload, project: { ...payload.project } }), current, actor(by)));
  }
  const records = (people_) => rec(people_ || people);
  return { store, anup, nidhi, project, people, sync, actor, daily, approve, edit, records };
}
const istToday = () => new Date(Date.now() + 5.5 * 3600 * 1000).toISOString().slice(0, 10);
const distinct = (report) => new Set(listUnits(report).map((entry) => entry.testId)).size;

(async () => {
// ---- credit -----------------------------------------------------------------------------------
  await check('credit: the first enterer of a unit keeps it; lastEditedBy moves only when the content changes', () => {
    const { store, anup, nidhi, daily, actor } = setup();
    const report = daily('2026-10-01');
    const units = listUnits(report);
    assert.ok(units.length >= 3, 'the sample day has several units');
    assert.ok(units.every((entry) => entry.unit.enteredBy && entry.unit.enteredBy.email === 'anup@elegrow.com' && !entry.unit.lastEditedBy));

    // Nidhi edits one electrode row and adds one; the browser tries to forge a stamp on another row.
    const edited = JSON.parse(JSON.stringify(report));
    edited.electrodeResistance[0].resistanceWithGrid = '0.77';
    edited.electrodeResistance[1].enteredBy = { email: 'forged@elegrow.com' };
    edited.electrodeResistance.push({ ...edited.electrodeResistance[0], rowId: 'electrode-new-1', tag: 'PIT-NEW' });
    const renormalized = normalizeReportInput({ ...edited, project: { ...edited.project } });
    assert.ok(renormalized.electrodeResistance.every((row) => !row.enteredBy), 'the normaliser drops stamps sent by the browser');
    const saved = store.updateReport(report.id, () => stampReport(renormalized, store.getReport(report.id), actor(nidhi), '2026-10-02T10:00:00.000Z'));
    const byId = Object.fromEntries(saved.electrodeResistance.map((row) => [row.rowId, row]));
    const first = report.electrodeResistance[0].rowId;
    assert.strictEqual(byId[first].enteredBy.email, 'anup@elegrow.com', 'the first enterer keeps the credit after an edit by someone else');
    assert.strictEqual(byId[first].lastEditedBy.email, 'nidhi@elegrow.com');
    assert.strictEqual(byId[report.electrodeResistance[1].rowId].enteredBy.email, 'anup@elegrow.com', 'a forged stamp is ignored');
    assert.ok(!byId[report.electrodeResistance[1].rowId].lastEditedBy, 'an unchanged row has no editor');
    assert.strictEqual(byId['electrode-new-1'].enteredBy.email, 'nidhi@elegrow.com', 'a new row is credited to whoever entered it');

    // Saving the same content again moves nothing.
    const again = stampReport(normalizeReportInput({ ...saved, project: { ...saved.project } }), saved, actor(anup), '2026-10-03T10:00:00.000Z');
    const row = again.electrodeResistance.find((entry) => entry.rowId === first);
    assert.strictEqual(row.lastEditedBy.email, 'nidhi@elegrow.com');
    assert.strictEqual(row.enteredBy.at, report.electrodeResistance[0].enteredBy.at);
  });

  await check('credit: a unit stored before 6G stays unstamped (the creator is credited at sync time); soil is per location, tower per group', () => {
    const { anup, nidhi, daily, actor } = setup();
    const report = daily('2026-10-01');
    const legacy = JSON.parse(JSON.stringify(report));
    legacy.electrodeResistance.forEach((row) => { delete row.enteredBy; });
    const next = stampReport(JSON.parse(JSON.stringify(legacy)), legacy, actor(nidhi));
    assert.ok(next.electrodeResistance.every((row) => !row.enteredBy), 'an old row is not handed to whoever saves next');
    const soil = listUnits(report).filter((entry) => entry.testId === 'soilResistivity');
    assert.ok(soil.length && soil.every((entry) => entry.unitId === entry.unit.locationId));
    assert.ok(anup && listUnits(report).some((entry) => entry.testId === 'towerFootingResistance' ? entry.unitId === entry.unit.groupId : true));
  });

  // ---- what is written when -----------------------------------------------------------------------
  await check('approval writes ONE record per engineer + date + test, with the agreed fields; nothing is written before', async () => {
    const { store, people, sync, daily, approve, project } = setup();
    const report = daily('2026-10-01');
    store.transition(report.id, 'in_review', { actor: 'P' });
    assert.deepStrictEqual(await sync.afterTransition('submit', store.getReport(report.id)), { skipped: 'step' });
    assert.deepStrictEqual(await sync.afterTransition('return', report), { skipped: 'step' });
    assert.deepStrictEqual(people.calls, [], 'no People call at all on submit or return, not even a form check');
    store.transition(report.id, 'draft', { actor: 'C', note: 'redo' });
    const approved = approve(report);
    await sync.afterTransition('approve', approved);
    const tests = new Set(listUnits(report).map((entry) => entry.testId));
    assert.strictEqual(people.records.size, tests.size, 'one record per test for the one engineer, not one per row');
    const electrodes = byTest(people, 'Earth Electrode Resistance Test');
    assert.strictEqual(electrodes.length, 1);
    const record = electrodes[0];
    assert.deepStrictEqual(Object.keys(record).sort(), Object.values(F).sort(), 'every field of the spec, each under the API name found by its label');
    assert.strictEqual(record[F.recordKey], `${project.id}|2026-10-01|anup@elegrow.com|electrodeResistance`);
    // (the day was returned once before approval above: Returned Count 1, so not first-time approved; late: not on time)
    assert.deepStrictEqual(
      [record[F.engineer], record[F.workDate], record[F.client], record[F.test], record[F.unit], record[F.plannedQty], record[F.completedQty], record[F.reportStatus], record[F.returnedCount], record[F.onTime], record[F.firstTimeApproved], record[F.reportId]],
      ['P-A', '01-Oct-2026', 'Gujarat Fluorochemicals Ltd', 'Earth Electrode Resistance Test', 'pits', 20, report.electrodeResistance.length, 'Approved', 1, 'No', 'No', report.id]
    );
    assert.strictEqual(record[F.project], '024-D03_TS-26-GFL Dahej earthing');
    assert.strictEqual(byTest(people, 'Soil Resistivity Test')[0][F.unit], 'locations');
    assert.strictEqual(byTest(people, 'Soil Resistivity Test')[0][F.completedQty], report.soilResistivity.locations.length);
  });

  await check('no duplicates: approving again, a Backfill, or a lost local row all land on the same People record', async () => {
    const { store, people, sync, daily, approve } = setup();
    const approved = approve(daily('2026-10-01'));
    await sync.afterTransition('approve', approved);
    const ids = [...people.records.keys()].sort();
    const writes = verbs(people).length;
    await sync.afterTransition('approve', approved);
    assert.strictEqual(verbs(people).length, writes, 'unchanged: nothing is written again');
    await sync.backfill({ from: '2026-10-01', to: '2026-10-01' });
    assert.deepStrictEqual([...people.records.keys()].sort(), ids);
    store.listPeopleSync().forEach((row) => store.savePeopleSync(row.key, { peopleRecordId: '', state: 'failed', hash: '' }));
    await sync.retry();
    assert.deepStrictEqual([...people.records.keys()].sort(), ids, 'with the record id forgotten the Record Key search finds the record: no second one');
    assert.ok(people.calls.includes('find'));
  });

  await check('re-issue: its submit writes nothing; approving R1 updates the SAME record (by Record Key) and R1 is not first-time approved', async () => {
    const { store, people, sync, daily, approve, edit, anup } = setup();
    const r0 = daily('2026-10-01');
    await sync.afterTransition('approve', approve(r0));
    const idsBefore = [...people.records.keys()].sort();
    const before = byTest(people, 'Earth Electrode Resistance Test')[0];
    assert.strictEqual(before[F.firstTimeApproved], 'Yes');
    const r1 = store.reissueReport(r0.id, { actor: 'J', userId: '', reason: 'Client comments' });
    assert.strictEqual(sync.rootIdOf(r1), r0.id);
    people.calls.length = 0;
    store.transition(r1.id, 'in_review', { actor: 'P' });
    assert.deepStrictEqual(await sync.afterTransition('submit', store.getReport(r1.id)), { skipped: 'step' });
    assert.deepStrictEqual(people.calls, [], 'People keeps the last approved values while R1 is under review');
    edit(r1.id, anup, (payload) => { payload.electrodeResistance.pop(); });
    await sync.afterTransition('approve', approve(store.getReport(r1.id)));
    assert.deepStrictEqual([...people.records.keys()].sort(), idsBefore, 'no record was added');
    const after = byTest(people, 'Earth Electrode Resistance Test')[0];
    assert.deepStrictEqual([after[F.completedQty], after[F.reportId], after[F.firstTimeApproved], after[F.reportStatus]], [r0.electrodeResistance.length - 1, r1.id, 'No', 'Approved']);
  });

  await check('a day split between two engineers: one record each, Completed Qty by first enterer, the credit never moves on an edit', async () => {
    const { store, people, sync, daily, approve, edit, anup, nidhi } = setup();
    const report = daily('2026-10-01');
    const anupRows = report.electrodeResistance.length;
    const saved = edit(report.id, nidhi, (payload) => {
      payload.electrodeResistance.push({ ...payload.electrodeResistance[0], rowId: 'electrode-n-1', tag: 'N-1' }, { ...payload.electrodeResistance[0], rowId: 'electrode-n-2', tag: 'N-2' });
      payload.electrodeResistance[0].resistanceWithGrid = '0.66'; // Nidhi corrects one of Anup's rows: the credit stays Anup's
    });
    assert.ok(anup && saved.electrodeResistance[0].lastEditedBy.email === 'nidhi@elegrow.com');
    await sync.afterTransition('approve', approve(store.getReport(report.id)));
    const pits = byTest(people, 'Earth Electrode Resistance Test');
    assert.strictEqual(pits.length, 2);
    const of = (email) => pits.find((record) => record[F.recordKey].includes(`|${email}|`));
    assert.deepStrictEqual([of('anup@elegrow.com')[F.completedQty], of('anup@elegrow.com')[F.engineer]], [anupRows, 'P-A']);
    assert.deepStrictEqual([of('nidhi@elegrow.com')[F.completedQty], of('nidhi@elegrow.com')[F.engineer]], [2, 'P-N']);
    assert.strictEqual(byTest(people, 'Soil Resistivity Test').length, 1, 'tests only Anup worked on stay a single record');
    assert.ok(nidhi);
  });

  await check('a quantity that drops to 0 in a later approved revision: Completed Qty 0, still Approved, with an audit entry', async () => {
    const { store, people, sync, daily, approve, edit, nidhi, anup } = setup();
    const r0 = daily('2026-10-01');
    edit(r0.id, nidhi, (payload) => { payload.electrodeResistance.push({ ...payload.electrodeResistance[0], rowId: 'electrode-n-1', tag: 'N-1' }); });
    await sync.afterTransition('approve', approve(store.getReport(r0.id)));
    assert.strictEqual(byTest(people, 'Earth Electrode Resistance Test').find((record) => record[F.recordKey].includes('|nidhi@'))[F.completedQty], 1);
    const r1 = store.reissueReport(r0.id, { actor: 'J', reason: 'Row was Anup\'s' });
    edit(r1.id, anup, (payload) => { payload.electrodeResistance = payload.electrodeResistance.filter((row) => row.rowId !== 'electrode-n-1'); });
    const approvedR1 = approve(store.getReport(r1.id));
    const tally = await sync.afterTransition('approve', approvedR1);
    assert.strictEqual(tally.zeroed, 1);
    const nidhiRecord = byTest(people, 'Earth Electrode Resistance Test').find((record) => record[F.recordKey].includes('|nidhi@'));
    assert.deepStrictEqual([nidhiRecord[F.completedQty], nidhiRecord[F.reportStatus], nidhiRecord[F.reportId]], [0, 'Approved', r1.id]);
    const audit = store.listAudit({ action: 'people.' }).entries.filter((entry) => entry.action === 'people.qty_zero');
    assert.strictEqual(audit.length, 1);
    assert.deepStrictEqual([audit[0].detail.engineer, audit[0].detail.test, audit[0].detail.was], ['nidhi@elegrow.com', 'electrodeResistance', 1]);
    assert.strictEqual((await sync.afterTransition('approve', approvedR1)).zeroed, 0, 'a record already at 0 is not zeroed again, nor audited twice');
    assert.strictEqual(store.listAudit({ action: 'people.' }).entries.filter((entry) => entry.action === 'people.qty_zero').length, 1);
  });

  await check('returned count counts return transitions across the revision chain; first-time approved needs none and revision 0', async () => {
    const { store, people, sync, daily, approve, edit, anup } = setup();
    const r0 = daily('2026-10-01');
    await sync.afterTransition('approve', approve(r0, { returns: 2 }));
    let record = byTest(people, 'Earth Electrode Resistance Test')[0];
    assert.deepStrictEqual([record[F.returnedCount], record[F.firstTimeApproved]], [2, 'No']);
    const r1 = store.reissueReport(r0.id, { actor: 'J', reason: 'x' });
    edit(r1.id, anup, (payload) => { payload.electrodeResistance[0].resistanceWithGrid = '0.9'; });
    await sync.afterTransition('approve', approve(store.getReport(r1.id), { returns: 1 }));
    record = byTest(people, 'Earth Electrode Resistance Test')[0];
    assert.strictEqual(record[F.returnedCount], 3, '2 on R0 and 1 on R1');
    const clean = daily('2026-10-02');
    await sync.afterTransition('approve', approve(clean));
    assert.strictEqual(rec(people).find((entry) => entry[F.workDate] === '02-Oct-2026')[F.firstTimeApproved], 'Yes');
  });

  await check('on-time: first submit on or before the next working day (6F working days, holidays, Asia/Kolkata)', async () => {
    const { store, sync, daily } = setup();
    assert.strictEqual(sync.submitDeadline('2026-10-23'), '2026-10-24', 'Friday: Saturday is a working day');
    assert.strictEqual(sync.submitDeadline('2026-10-24'), '2026-10-26', 'Saturday: Sunday is not');
    store.setScheduleCache('cal:2026-10', { leave: [], holidays: [{ id: 'h', name: 'Diwali', locations: [], days: [{ date: '2026-10-26', weight: 1 }] }, { id: 'h2', name: 'Surat day', locations: ['Surat'], days: [{ date: '2026-10-27', weight: 1 }] }, { id: 'h3', name: 'Half', locations: [], days: [{ date: '2026-10-28', weight: 0.5 }] }] });
    assert.strictEqual(sync.submitDeadline('2026-10-24'), '2026-10-27', 'a holiday for everyone is skipped');
    assert.strictEqual(sync.submitDeadline('2026-10-24', 'Surat'), '2026-10-28', 'a location holiday is skipped for people there only');
    assert.strictEqual(sync.submitDeadline('2026-10-27'), '2026-10-28', 'a half day is a working day');
    const today = daily(istToday());
    store.transition(today.id, 'in_review', { actor: 'P' });
    assert.strictEqual(sync.buildEntries(store.getReport(today.id))[0].onTime, true, 'submitted the day itself');
    const late = daily('2026-10-01');
    store.transition(late.id, 'in_review', { actor: 'P' });
    assert.strictEqual(sync.buildEntries(store.getReport(late.id))[0].onTime, false, 'submitted weeks later');
    const never = daily('2026-10-03');
    assert.strictEqual(sync.buildEntries(never)[0].onTime, false, 'never submitted');
  });

  await check('data gaps are the CAP data-gap lines of that test; planned comes from the project', async () => {
    const { people, sync, daily, approve } = setup();
    const report = daily('2026-10-01');
    const expected = {};
    buildFinalReportData(report).cap.filter((line) => line.kind === 'gap').forEach((line) => { expected[line.testId] = (expected[line.testId] || 0) + 1; });
    const entries = sync.buildEntries(report);
    entries.forEach((entry) => assert.strictEqual(entry.dataGaps, expected[entry.testId] || 0, entry.testId));
    assert.ok(entries.find((entry) => entry.testId === 'soilResistivity').planned === 4 && entries.find((entry) => entry.testId === 'loopImpedanceTest').planned === null, 'a test with no planned quantity sends no Planned Qty');
    await sync.afterTransition('approve', approve(report));
    assert.ok(byTest(people, 'Loop Impedance Test').every((record) => !(F.plannedQty in record)), 'no planned quantity: the field is left out, not sent as 0');
  });

  await check('finals, sample and test reports write nothing', async () => {
    const { store, project, people, sync, daily, approve } = setup();
    const day = daily('2026-10-01');
    const final = store.addReport({ ...JSON.parse(JSON.stringify(day)), id: undefined, kind: 'final', projectId: project.id, reviewStatus: 'draft' });
    assert.deepStrictEqual(await sync.afterTransition('approve', { ...final, reviewStatus: 'approved' }), { skipped: 'not-daily' });
    const sampleProject = store.addProject({ projectNo: 'SAMPLE-001', clientName: 'Acme' });
    assert.deepStrictEqual(await sync.afterTransition('approve', approve(daily('2026-10-02', undefined, { projectId: sampleProject.id }))), { skipped: 'sample' });
    const named = daily('2026-10-03', undefined, { projectExtra: { projectNo: 'Test run 7' } });
    assert.deepStrictEqual(await sync.afterTransition('approve', approve(named)), { skipped: 'sample' });
    assert.ok(isSampleReport({ project: { projectNo: 'X', clientName: 'Demo Client' } }) && !isSampleReport({ project: { projectNo: '024-D03', clientName: 'Contest Ltd' } }));
    assert.strictEqual(people.records.size, 0);
    assert.deepStrictEqual(people.calls, []);
  });

  // ---- failures never block, field discovery ---------------------------------------------------------
  await check('form missing: entries wait as pending, nothing is lost, and Retry sends them once the form exists', async () => {
    const { store, people, sync, daily, approve } = setup();
    people.formExists = false;
    const report = daily('2026-10-01');
    const result = await sync.afterTransition('approve', approve(report));
    assert.strictEqual(result.blocked, distinct(report));
    assert.strictEqual(people.records.size, 0);
    assert.ok(store.listPeopleSync().every((row) => row.state === 'pending' && /does not exist yet/.test(row.error)));
    people.formExists = true;
    const tally = await sync.retry();
    assert.strictEqual(tally.written, distinct(report));
    assert.deepStrictEqual(store.peopleSyncCounts(), { done: distinct(report) });
  });

  await check('fields are found by label at write time: matching ignores case and punctuation, a missing label waits and names it, an env override wins', async () => {
    const one = setup();
    one.people.labels = one.people.labels.map((label) => label.toLowerCase().replace(/-/g, ' ')).filter((label) => label !== 'data gaps' && label !== 'unit');
    const report = one.daily('2026-10-01');
    const result = await one.sync.afterTransition('approve', one.approve(report));
    assert.strictEqual(result.blocked, distinct(report));
    const waiting = one.store.listPeopleSync();
    assert.ok(waiting.every((row) => row.state === 'pending' && /no field labelled: Unit, Data Gaps\./.test(row.error)), waiting[0].error);
    assert.strictEqual(one.people.records.size, 0);
    one.people.labels.push('Data Gaps', 'Unit');
    assert.strictEqual((await one.sync.retry()).written, distinct(report));
    assert.ok(rec(one.people).every((record) => record.f_Data_Gaps !== undefined && record.f_Unit !== undefined), 'the fields found later are sent');
    // an override by API name takes the field without a label match
    const two = setup();
    two.people.labels = two.people.labels.filter((label) => label !== 'Completed Qty');
    const overridden = createPeopleSync({ store: two.store, people: two.people, config: { zoho: { useMock: true, workLog: { form: 'Site_Work_Log', fields: { completedQty: 'custom_completed_1' } } } } });
    const report2 = two.daily('2026-10-01');
    await overridden.afterTransition('approve', two.approve(report2));
    assert.strictEqual(byTest(two.people, 'Soil Resistivity Test')[0].custom_completed_1, report2.soilResistivity.locations.length);
  });

  await check('a People failure is kept as failed with its reason and the review step is not affected; Retry recovers', async () => {
    const { store, people, sync, daily, approve } = setup();
    const report = daily('2026-10-01');
    people.failNext = 2;
    const approved = approve(report);
    assert.strictEqual(approved.reviewStatus, 'approved');
    const result = await sync.afterTransition('approve', approved);
    assert.strictEqual(result.failed, 2);
    assert.ok(store.listPeopleSync({ state: 'failed' }).every((row) => row.error === 'People is down' && row.attempts === 1));
    const tally = await sync.retry();
    assert.deepStrictEqual([tally.written, tally.failed], [2, 0]);
    assert.strictEqual(people.records.size, distinct(report));
  });

  await check('an engineer with no People employee fails with a clear reason; a row with no stamp is credited to the report creator', async () => {
    const { store, nidhi, people, sync, daily, approve } = setup();
    people.employees = [people.employees[0]];
    const report = daily('2026-10-01', nidhi);
    await sync.afterTransition('approve', approve(report));
    const failed = store.listPeopleSync({ state: 'failed' });
    assert.strictEqual(failed.length, distinct(report));
    assert.ok(failed.every((row) => /No Zoho People employee has the e-mail nidhi@elegrow\.com/.test(row.error)));
    const legacy = JSON.parse(JSON.stringify(store.getReport(report.id)));
    legacy.electrodeResistance.forEach((row) => { delete row.enteredBy; });
    const entry = sync.buildEntries(legacy).find((item) => item.testId === 'electrodeResistance');
    assert.deepStrictEqual([entry.engineer.email, entry.creditedByCreator, entry.completed], ['nidhi@elegrow.com', legacy.electrodeResistance.length, legacy.electrodeResistance.length]);
  });

  await check('the Engineer lookup value comes from the field metadata and defaults to the record id', async () => {
    assert.strictEqual(lookupAttribute({}), 'recordId');
    assert.strictEqual(lookupAttribute({ lookup: { displayField: 'EmployeeID' } }), 'employeeId');
    assert.strictEqual(lookupAttribute({ lookup: { fieldName: 'Email address' } }), 'email');
    assert.strictEqual(lookupAttribute({ lookup: { labelName: 'Full name' } }), 'name');
    assert.strictEqual(lookupAttribute({ lookup: { form: 'employee' } }), 'recordId');
    const { people, sync, daily, approve } = setup();
    people.lookup = { displayField: 'EmployeeID' };
    await sync.afterTransition('approve', approve(daily('2026-10-01')));
    assert.ok(rec(people).every((record) => record[F.engineer] === 'ELE-A'));
    assert.strictEqual(normalizeFormField({ labelName: 'Engineer', componentName: 'Engineer1', type: 'Lookup', lookupDetails: { displayField: 'Email' } }).apiName, 'Engineer1');
    assert.strictEqual(normalizeEmployee({ EmployeeID: 'E1', recordId: '55', EmailID: 'A@x.com' }).recordId, '55');
  });

  // ---- the live form (shapes and quirks seen on the real "Site Work Log" form) ---------------------------
  const fixture = JSON.parse(fs.readFileSync(path.join(__dirname, 'fixtures', 'people-site-work-log-components.json'), 'utf8'));

  await check('live metadata: label -> API name -> type and the dropdown / radio options are read from the real components shape', async () => {
    const fields = await createPeopleClient({ http: { get: async (url) => ({ data: url.includes('/components') ? fixture : { access_token: 'x' } }), post: async () => ({ data: { access_token: 'tok' } }) }, config: { zoho: { useMock: false, peopleBaseUrl: 'https://people.test', accountsBaseUrl: 'https://accounts.test', refreshToken: 'r', clientId: 'c', clientSecret: 's' } } }).getFormFields('site_work_log');
    const by = Object.fromEntries(fields.map((field) => [field.label, field]));
    assert.deepStrictEqual([by['Work Date'].apiName, by['Work Date'].type, by['Completed Qty.'].apiName, by['Completed Qty.'].type, by['Engineer'].type], ['work_date', 'date', 'completed_qty', 'number', 'lookup']);
    assert.deepStrictEqual(by.Unit.options, ['Points', 'Panels', 'Locations', 'Pits', 'Towers']);
    assert.deepStrictEqual(by['On-time Submission'].options, ['No', 'Yes']);
    assert.strictEqual(by.Test.options.length, 9);
    assert.deepStrictEqual(by.Engineer.lookupOptions[0], { value: 'ELE0001 - Anup - Parekh', id: '9001' });
    assert.strictEqual(by.Engineer.options, null, 'a lookup\'s options are records, not values');
    // every label of the spec is found on the form ("Planned Qty." carries a trailing dot)
    assert.deepStrictEqual(Object.values(WORKLOG_LABELS).filter((label) => !fields.some((field) => field.label.toLowerCase().replace(/[^a-z0-9]/g, '') === label.toLowerCase().replace(/[^a-z0-9]/g, ''))), []);
  });

  await check('live values: dropdowns get the form\'s own spelling (short test names, capitalised units), the Engineer is the lookup option Id, snake_case API names; an unknown option fails with the options listed', async () => {
    const live = await createPeopleClient({ http: { get: async () => ({ data: fixture }), post: async () => ({ data: { access_token: 'tok' } }) }, config: { zoho: { useMock: false, peopleBaseUrl: 'https://people.test', accountsBaseUrl: 'https://accounts.test', refreshToken: 'r', clientId: 'c', clientSecret: 's' } } }).getFormFields('site_work_log');
    const { people, sync, daily, approve, store } = setup();
    people.fields = live;
    people.employees = [{ email: 'anup@elegrow.com', recordId: 'ZID-1', employeeId: 'ELE0001', name: 'Anup Parekh' }];
    const report = daily('2026-10-01');
    await sync.afterTransition('approve', approve(report));
    assert.strictEqual(people.records.size, distinct(report));
    const names = rec(people).map((record) => record.test).sort();
    assert.ok(names.includes('Soil Resistivity') && names.includes('Earth Electrode Resistance') && names.every((name) => !/ Test$/.test(name)), names.join());
    const pits = rec(people).find((record) => record.test === 'Earth Electrode Resistance');
    assert.deepStrictEqual([pits.unit, pits.engineer, pits.work_date, pits.report_status, pits.on_time_submission, typeof pits.completed_qty], ['Pits', '9001', '01-Oct-2026', 'Approved', 'No', 'number']);
    assert.ok(Object.keys(pits).every((key) => /^[a-z_]+$/.test(key)), 'API names are the labelname of the form');
    // a form whose Report Status dropdown has no "Approved" fails the entry with a reason, it does not send a value People would refuse
    const other = setup();
    other.people.fields = live.map((field) => (field.label === 'Report Status' ? { ...field, options: ['Draft', 'Done'] } : field));
    other.people.employees = people.employees;
    await other.sync.afterTransition('approve', other.approve(other.daily('2026-10-01')));
    assert.strictEqual(other.people.records.size, 0);
    assert.ok(other.store.listPeopleSync({ state: 'failed' }).every((row) => /"Approved" is not an option of the Report Status field \(options: Draft, Done\)/.test(row.error)));
    assert.ok(store);
  });

  await check('People quirks: a raw "|" is sent as \\u007c, "No records found (7024)" means not found, and the id of a record is its row key (not Zoho_ID)', async () => {
    const seen = [];
    const rows = { '555': [{ Zoho_ID: 554, record_key: 'k' }] };
    let mode = 'empty';
    const http = {
      post: async (url, body) => ({ data: url.includes('/oauth/') ? { access_token: 'tok' } : { response: { result: { pkId: '1' }, status: 0 } } }),
      get: async (url, opts) => {
        seen.push(opts.params);
        if (url.includes('/employee/getRecords')) return { data: { response: { result: [{ '777': [{ EmployeeID: 'E1', Zoho_ID: 776, EmailID: 'a@x.com' }] }] } } };
        return { data: mode === 'empty' ? { response: { status: 1, errors: { code: 7024, message: 'No records found' } } } : { response: { result: [rows], status: 0 } } };
      }
    };
    const client = createPeopleClient({ http, config: { zoho: { useMock: false, peopleBaseUrl: 'https://people.test', accountsBaseUrl: 'https://accounts.test', refreshToken: 'r', clientId: 'c', clientSecret: 's' } } });
    assert.strictEqual(await client.findRecord('site_work_log', 'record_key', 'a|b|c'), null);
    assert.ok(!seen[0].searchParams.includes('|') && seen[0].searchParams.includes('\\u007c'), seen[0].searchParams);
    assert.deepStrictEqual(JSON.parse(seen[0].searchParams), { searchField: 'record_key', searchOperator: 'Is', searchText: 'a|b|c' }, 'the escape is still valid JSON of the same value');
    mode = 'found';
    assert.strictEqual(await client.findRecord('site_work_log', 'record_key', 'k'), '555');
    assert.strictEqual((await client.listEmployees())[0].recordId, '777');
  });

  // ---- backfill and totals ---------------------------------------------------------------------------
  await check('backfill: approved days in the range, latest revision only, sample skipped, finals ignored, safe to run twice', async () => {
    const { store, people, sync, daily, approve, edit, anup } = setup();
    const inRange = approve(daily('2026-10-05'));
    approve(daily('2026-09-05'));                       // out of range
    daily('2026-10-06');                                // still a draft
    const sampleProject = store.addProject({ projectNo: 'TEST-9', clientName: 'Acme' });
    approve(daily('2026-10-07', undefined, { projectId: sampleProject.id }));
    const r1 = store.reissueReport(inRange.id, { actor: 'J', reason: 'x' });
    edit(r1.id, anup, (payload) => { payload.electrodeResistance[0].resistanceWithGrid = '0.31'; });
    approve(store.getReport(r1.id));                    // R1 approved
    await assert.rejects(() => sync.backfill({ from: '2026-10-09', to: '2026-10-01' }), (error) => error.code === 'BAD_RANGE');
    await assert.rejects(() => sync.backfill({ from: 'x', to: '2026-10-01' }), (error) => error.code === 'BAD_RANGE');
    const first = await sync.backfill({ from: '2026-10-01', to: '2026-10-31' });
    assert.deepStrictEqual([first.reports, first.skippedSample, first.written, first.failed], [1, 1, distinct(inRange), 0]);
    assert.ok(rec(people).every((record) => record[F.reportId] === r1.id), 'the latest approved revision is what is sent');
    const second = await sync.backfill({ from: '2026-10-01', to: '2026-10-31' });
    assert.deepStrictEqual([second.written, second.unchanged], [0, distinct(inRange)]);
    assert.strictEqual(store.getScheduleSetting('peopleSyncLastBackfill').unchanged, second.unchanged);
    const status = await sync.status();
    assert.deepStrictEqual([status.formExists, status.problems.length, status.counts.done], [true, 0, distinct(inRange)]);
  });

  await check('totals per engineer: completed vs planned per test, on-time %, first-time approved %, returns, data gaps', async () => {
    const { store, project, people, sync, daily, approve, edit, nidhi } = setup();
    store.setScheduleSetting('x', 1);
    const d1 = daily(istToday());                       // submitted on time, first time
    store.transition(d1.id, 'in_review', { actor: 'P' });
    await sync.afterTransition('approve', approve(d1));
    const d2 = daily('2026-10-02');                     // late, returned once
    edit(d2.id, nidhi, (payload) => { payload.electrodeResistance.push({ ...payload.electrodeResistance[0], rowId: 'electrode-n-1', tag: 'N-1' }); });
    await sync.afterTransition('approve', approve(store.getReport(d2.id), { returns: 1 }));
    const totals = sync.totals({ from: '2026-10-01', to: '2026-10-31' });
    const anupTotals = totals.engineers.find((who) => who.email === 'anup@elegrow.com');
    const pits = anupTotals.tests.find((test) => test.testId === 'electrodeResistance');
    assert.deepStrictEqual([pits.completed, pits.planned, pits.unit], [d1.electrodeResistance.length + d2.electrodeResistance.length, 20, 'pits'], 'planned counts the project once');
    assert.deepStrictEqual([anupTotals.engineerDays, anupTotals.onTimePct, anupTotals.firstTimeApprovedPct, anupTotals.returns], [2, istToday() >= '2026-10-01' ? 50 : 0, 50, 1]);
    const nidhiTotals = totals.engineers.find((who) => who.email === 'nidhi@elegrow.com');
    assert.deepStrictEqual([nidhiTotals.tests.length, nidhiTotals.tests[0].completed, nidhiTotals.engineerDays, nidhiTotals.returns], [1, 1, 1, 1]);
    assert.strictEqual(sync.totals({ from: '2026-11-01', to: '2026-11-30' }).engineers.length, 0);
    assert.throws(() => sync.totals({ from: '2026-11-30', to: '2026-11-01' }), (error) => error.code === 'BAD_RANGE');
    assert.ok(project && people);
  });

  // ---- the People client ----------------------------------------------------------------------------
  await check('People client: insert and update are form-encoded POSTs with inputData, a 401 refreshes the token once, an API error is surfaced', async () => {
    const posts = [];
    let denied = true;
    const http = {
      post: async (url, body) => {
        posts.push({ url, body });
        if (url.includes('/oauth/v2/token')) return { data: { access_token: `tok${posts.length}` } };
        if (denied) { denied = false; const error = new Error('401'); error.response = { status: 401, data: {} }; throw error; }
        if (url.endsWith('/insertRecord')) return { data: { response: { result: { pkId: '4242', message: 'ok' }, status: 0 } } };
        if (url.endsWith('/boom/updateRecord')) return { data: { response: { status: 1, errors: { code: 7012, message: 'Invalid value' } } } };
        return { data: { response: { result: {}, status: 0 } } };
      },
      get: async (url) => ({ data: { response: { result: [{ displayName: 'Site Work Log', formLinkName: 'Site_Work_Log' }] } } })
    };
    const live = { zoho: { useMock: false, peopleBaseUrl: 'https://people.test', accountsBaseUrl: 'https://accounts.test', refreshToken: 'r', clientId: 'c', clientSecret: 's' } };
    const client = createPeopleClient({ http, config: live });
    assert.strictEqual(await client.hasForm('Site_Work_Log'), true);
    assert.strictEqual(await client.hasForm('Nope'), false);
    assert.strictEqual(await client.insertRecord('Site_Work_Log', { Record_Key: 'k', Quantity: 1 }), '4242');
    const insert = posts.filter((call) => call.url.endsWith('/people/api/forms/json/Site_Work_Log/insertRecord'));
    assert.strictEqual(insert.length, 2, 'retried once after the 401');
    assert.deepStrictEqual(JSON.parse(new URLSearchParams(insert[1].body).get('inputData')), { Record_Key: 'k', Quantity: 1 });
    assert.strictEqual(posts.filter((call) => call.url.includes('/oauth/v2/token')).length, 2, 'the token was refreshed once');
    await client.updateRecord('Site_Work_Log', '4242', { Status: 'Approved' });
    const update = posts[posts.length - 1];
    assert.ok(update.url.endsWith('/Site_Work_Log/updateRecord') && new URLSearchParams(update.body).get('recordId') === '4242');
    await assert.rejects(() => client.updateRecord('boom', '1', {}), /Invalid value \(7012\)/);
  });

  // ---- HTTP routes on the mock portal -----------------------------------------------------------------
  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: path.join(__dirname, '..'),
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'http', 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  try {
    await new Promise((resolve, reject) => {
      server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
      setTimeout(() => reject(new Error('server did not start')), 15000);
    });
    const call = createCaller(`http://127.0.0.1:${port}`);
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const people = async () => (await call('GET', '/api/admin/people-sync')).body;

    await check('routes: admin only; status reports the form and counts', async () => {
      assert.strictEqual((await call('GET', '/api/admin/people-sync', undefined, 'engineer@elegrow.com')).status, 403);
      assert.strictEqual((await call('POST', '/api/admin/people-sync/backfill', { from: '2026-10-01', to: '2026-10-31' }, 'approver@elegrow.com')).status, 403);
      assert.strictEqual((await call('POST', '/api/admin/people-sync/retry', {}, 'engineer@elegrow.com')).status, 403);
      const status = await people();
      assert.deepStrictEqual([status.form, status.formExists, status.counts, status.mock], ['site_work_log', true, {}, []]);
      assert.strictEqual((await call('POST', '/api/admin/people-sync/backfill', { from: '2026-10-09', to: '2026-10-01' })).status, 400);
    });

    await check('routes: a day entered by Anup, edited by an admin, approved by the reviewers: nothing in People until approval, then one record per test, credit to Anup', async () => {
      const project = (await call('POST', '/api/projects', { projectNo: '031-D03_TS-26-AAA', clientName: 'AAA Ltd', siteLocation: 'Surat', workOrder: 'W1', engineerName: 'Anup Parekh' }, 'anup@elegrow.com')).body.projectId;
      const payload = buildPayload();
      const created = await call('POST', `/api/projects/${project}/days`, { workDate: '2026-10-12', report: { ...payload, project: { ...payload.project, reportDate: '2026-10-12' } } }, 'anup@elegrow.com');
      assert.strictEqual(created.status, 201, JSON.stringify(created.body));
      const id = created.body.id;
      const tests = new Set(listUnits(created.body).map((entry) => entry.testId)).size;
      assert.ok(created.body.electrodeResistance.every((row) => row.enteredBy.email === 'anup@elegrow.com'));

      // Nidhi cannot edit Anup's day; an admin edit stamps lastEditedBy but never moves the credit.
      const edit = JSON.parse(JSON.stringify(created.body));
      delete edit.assessment;
      edit.electrodeResistance[0].resistanceWithGrid = '0.55';
      assert.strictEqual((await call('PUT', `/api/reports/${id}`, edit, 'nidhi@elegrow.com')).status, 403);
      const adminEdit = await call('PUT', `/api/reports/${id}`, edit, 'admin@elegrow.com');
      assert.strictEqual(adminEdit.status, 200, JSON.stringify(adminEdit.body));
      const row = adminEdit.body.electrodeResistance[0];
      assert.deepStrictEqual([row.enteredBy.email, row.lastEditedBy.email], ['anup@elegrow.com', 'admin@elegrow.com']);

      assert.strictEqual((await call('POST', `/api/reports/${id}/review/submit`, {}, 'anup@elegrow.com')).status, 200);
      await sleep(300);
      assert.deepStrictEqual((await people()).mock, [], 'a submit writes nothing to People');

      const ticks = Object.fromEntries((await call('GET', `/api/reports/${id}/review`)).body.checklist.map((item) => [item.id, true]));
      await call('PUT', `/api/reports/${id}/review/ticks`, { ticks }, 'reviewer@elegrow.com');
      assert.strictEqual((await call('POST', `/api/reports/${id}/review/check`, {}, 'reviewer@elegrow.com')).status, 200);
      assert.strictEqual((await call('POST', `/api/reports/${id}/review/approve`, {}, 'approver@elegrow.com')).status, 200);
      for (let i = 0; i < 40 && (await people()).mock.length < tests; i += 1) await sleep(50);
      const status = await people();
      assert.strictEqual(status.mock.length, tests, 'one record per test, not per row');
      assert.ok(status.mock.every((record) => record.Report_Status === 'Approved' && record.Work_Date === '12-Oct-2026' && record.Engineer === 'P-1001' && record.Client === 'Gujarat Fluorochemicals Ltd'), 'credited to Anup (his People record id), not to the admin who edited');
      const pits = status.mock.find((record) => record.Test === 'Earth Electrode Resistance Test');
      assert.strictEqual(pits.Completed_Qty, created.body.electrodeResistance.length);
      assert.ok(pits.Record_Key.endsWith('|2026-10-12|anup@elegrow.com|electrodeResistance'));
      assert.deepStrictEqual(status.counts, { done: tests });
      const totals = (await call('GET', '/api/admin/people-sync/totals?from=2026-10-01&to=2026-10-31')).body;
      assert.deepStrictEqual([totals.engineers.length, totals.engineers[0].email, totals.engineers[0].tests.length], [1, 'anup@elegrow.com', tests]);
      assert.strictEqual((await call('GET', '/api/admin/people-sync/totals?from=2026-10-31&to=2026-10-01')).status, 400);
      assert.strictEqual((await call('GET', '/api/admin/people-sync/totals?from=2026-10-01&to=2026-10-31', undefined, 'engineer@elegrow.com')).status, 403);
    });

    await check('routes: backfill and retry run, are audited, and a second backfill changes nothing', async () => {
      const first = await call('POST', '/api/admin/people-sync/backfill', { from: '2026-10-01', to: '2026-10-31' });
      assert.strictEqual(first.status, 200);
      assert.deepStrictEqual([first.body.written, first.body.failed, first.body.skippedSample], [0, 0, 0]);
      assert.ok(first.body.unchanged > 0 && first.body.reports === 1);
      const retry = await call('POST', '/api/admin/people-sync/retry', {});
      assert.deepStrictEqual([retry.status, retry.body.written], [200, 0]);
      const audit = (await call('GET', '/api/admin/audit?action=people.')).body.entries.map((entry) => entry.action);
      assert.ok(audit.includes('people.backfill') && audit.includes('people.retry'));
      assert.strictEqual((await people()).lastBackfill.from, '2026-10-01');
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
