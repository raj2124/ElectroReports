// Session 6F, scheduling: the calendar rules (working days, holidays, clashes, leave, quantity vs days),
// the Zoho People client (paging, token retry, and that it only ever reads), the cache (age, stale data
// kept on a failure, the one-a-minute refresh limit) and the HTTP routes on the mock portal.
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { devEnv, createCaller } = require('./lib/dev-auth');
const { parseZohoDate, monthsOf, monthRange, clampWindow, defaultSettings, mergeSettings, buildSchedule } = require('../src/schedule');
const { createScheduleService } = require('../src/scheduleService');
const { createPeopleClient, fromPeopleDate, toPeopleDate, leaveState, normalizeLeave, normalizeHoliday } = require('../src/zohoPeople');

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

const config = { schedule: { refreshMinutes: 0, manualRefreshSeconds: 60, hiddenTaskStatuses: ['Closed'], workingDays: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'] }, zoho: { useMock: false, peopleBaseUrl: 'https://people.test', accountsBaseUrl: 'https://accounts.test', refreshToken: 'r', clientId: 'c', clientSecret: 's' } };
const settings = () => defaultSettings(config);
const user = (id, email, name, roles = ['engineer'], active = true) => ({ id, email, name, roles, active });
const anup = user('U1', 'anup@elegrow.com', 'Anup Parekh');
const karishma = user('U2', 'karishma@elegrow.com', 'Karishma Prajapati');
const nidhi = user('U3', 'nidhi@elegrow.com', 'Nidhi Shah');
const admin = user('U4', 'admin@elegrow.com', 'H. S. Patil', ['admin']);
const owner = (u) => ({ id: u.id, name: u.name.split(' ')[0].toLowerCase(), fullName: u.name, email: u.email });
const task = (id, extra = {}) => ({ id, name: `[ESHA] Task ${id}`, owners: [owner(anup)], startDate: '10-19-2026', dueDate: '10-23-2026', status: 'Planned', closed: false, duration: '', eshaTest: '', testId: '', plannedRaw: '', ...extra });
const project = (id, tasks, extra = {}) => ({ id, name: `0${id}-D03_TS-26-AAA`, clientName: 'AAA', tasks, ...extra });
const build = (input) => buildSchedule({ from: '2026-10-19', to: '2026-10-25', projects: [], employees: [], leave: [], holidays: [], users: [anup, karishma, nidhi, admin], me: nidhi, settings: settings(), hiddenStatuses: ['Closed'], workspaces: {}, ...input });
const kinds = (result) => result.warnings.map((warning) => warning.kind).sort();
const rowOf = (result, name) => result.rows.find((row) => row.name === name);

(async () => {
  // ---- dates and People normalisers --------------------------------------------------------------
  await check('Zoho dates: the v1 "MM-DD-YYYY" (portal local) and ISO both read as the same calendar day', () => {
    assert.strictEqual(parseZohoDate('10-26-2026'), '2026-10-26');
    assert.strictEqual(parseZohoDate('2026-10-26'), '2026-10-26');
    assert.strictEqual(parseZohoDate('2026-10-26T00:00:00+05:30'), '2026-10-26');
    assert.strictEqual(parseZohoDate(''), '');
    assert.strictEqual(parseZohoDate('13-45-2026'), '');
    assert.strictEqual(fromPeopleDate('26-Oct-2026'), '2026-10-26');
    assert.strictEqual(toPeopleDate('2026-10-06'), '06-Oct-2026');
    assert.deepStrictEqual(monthsOf('2026-10-28', '2026-12-02'), ['2026-10', '2026-11', '2026-12']);
    assert.deepStrictEqual(monthRange('2026-02'), { from: '2026-02-01', to: '2026-02-28' });
    assert.deepStrictEqual(clampWindow('2026-10-01', '2030-01-01'), { from: '2026-10-01', to: '2027-01-01' });
  });

  await check('leave: approved and pending count; rejected and cancelled do not; an Absent record is an approved absence', () => {
    assert.strictEqual(leaveState('Approved', 'PAID'), 'approved');
    assert.strictEqual(leaveState('Pending', 'UNPAID'), 'pending');
    assert.strictEqual(leaveState('Rejected', 'PAID'), '');
    assert.strictEqual(leaveState('Cancelled', 'PAID'), '');
    assert.strictEqual(leaveState('', 'ABSENT'), 'approved');
    assert.strictEqual(leaveState('', 'PAID'), '');
    const leave = normalizeLeave('L9', { From: '21-Oct-2026', To: '24-Oct-2026', ApprovalStatus: 'Approved', Type: 'PAID', Leavetype: 'Casual Leave', Employee: 'K', EmployeeId: 'E1', ZUID: 5, Days: { '21-Oct-2026': { LeaveCount: '1.0' }, '22-Oct-2026': { LeaveCount: '0.0' }, '24-Oct-2026': { LeaveCount: '1.0' } } });
    assert.deepStrictEqual(leave.days.map((day) => day.date), ['2026-10-21', '2026-10-24'], 'days that do not count (weekend, holiday) are left out');
    assert.strictEqual(normalizeLeave('L1', { From: '21-Oct-2026', To: '21-Oct-2026', ApprovalStatus: 'Rejected', Type: 'PAID' }), null);
  });

  await check('holidays: a restricted (optional) holiday is not a day off; a half day weighs 0.5; locations are kept', () => {
    assert.strictEqual(normalizeHoliday({ id: 'h', name: 'Optional', isRestricted: true, from: '01-Nov-2026', dayDetails: [] }), null);
    const half = normalizeHoliday({ id: 'h2', name: 'Half', from: '12-Nov-2026', dayDetails: [{ '12-Nov-2026': { ishalfday: true } }], locations: [{ name: 'Surat' }] });
    assert.deepStrictEqual(half.days, [{ date: '2026-11-12', weight: 0.5 }]);
    assert.deepStrictEqual(half.locations, ['Surat']);
  });

  // ---- the calendar rules ----------------------------------------------------------------------
  await check('settings: Monday to Saturday by default; configurable; validated', () => {
    assert.deepStrictEqual(settings().workingDays, [1, 2, 3, 4, 5, 6]);
    assert.deepStrictEqual(defaultSettings({ schedule: { workingDays: ['Mon', 'Tue', 'Sun'] } }).workingDays, [0, 1, 2]);
    assert.ok(mergeSettings(settings(), { workingDays: [] }).error);
    assert.ok(mergeSettings(settings(), { workingDays: [7] }).error);
    assert.ok(mergeSettings(settings(), { rates: { nope: 3 } }).error);
    assert.ok(mergeSettings(settings(), { rates: { electrodeResistance: 0 } }).error);
    const merged = mergeSettings(settings(), { workingDays: [1, 2, 3, 4, 5], rates: { electrodeResistance: 40 } }).settings;
    assert.deepStrictEqual(merged.workingDays, [1, 2, 3, 4, 5]);
    assert.strictEqual(merged.rates.electrodeResistance.perDay, 40);
    assert.strictEqual(settings().rates.electrodeResistance.perDay, 35, 'the defaults are not changed by an edit');
  });

  await check('the window shows only tasks that overlap it, never closed ones, whatever the project status', () => {
    const result = build({ projects: [project('1', [task('in'), task('before', { startDate: '10-01-2026', dueDate: '10-05-2026' }), task('after', { startDate: '11-01-2026', dueDate: '11-05-2026' }), task('edge', { startDate: '10-25-2026', dueDate: '10-30-2026' }), task('closed', { status: 'Closed', closed: true }), task('closed2', { status: 'closed' })], { stage: 'Invoicing' })] });
    const bars = rowOf(result, 'Anup Parekh').bars.map((bar) => bar.taskId).sort();
    assert.deepStrictEqual(bars, ['edge', 'in']);
    const edge = rowOf(result, 'Anup Parekh').bars.find((bar) => bar.taskId === 'edge');
    assert.ok(edge.clippedEnd && edge.to === '2026-10-25' && edge.end === '2026-10-30');
  });

  await check('rows: me first, then the engineers; a person who never signed in gets a row from the Zoho name; no owner has its own row', () => {
    const mohana = { id: 'z9', name: 'mohana', fullName: 'Mohana Vamsee', email: 'mohana@elegrow.com' };
    const result = build({ me: nidhi, projects: [project('1', [task('a'), task('b', { owners: [mohana], startDate: '10-20-2026', dueDate: '10-20-2026' }), task('c', { owners: [], startDate: '10-21-2026', dueDate: '10-21-2026' }), task('d', { owners: [{ name: 'Rakesh Patel' }], startDate: '10-21-2026', dueDate: '10-21-2026' })])] });
    assert.deepStrictEqual(result.rows.map((row) => row.name), ['Nidhi Shah', 'Anup Parekh', 'Karishma Prajapati', 'Mohana Vamsee', 'Rakesh Patel', 'No owner']);
    assert.strictEqual(result.rows[0].isMe, true);
    assert.strictEqual(rowOf(result, 'Mohana Vamsee').kind, 'zoho');
    assert.strictEqual(rowOf(result, 'Rakesh Patel').kind, 'zoho', 'an owner with no e-mail is matched by name');
    assert.strictEqual(rowOf(result, 'No owner').bars[0].taskId, 'c');
    const asAdmin = build({ me: admin, projects: [] });
    assert.deepStrictEqual(asAdmin.rows.map((row) => row.name), ['H. S. Patil', 'Anup Parekh', 'Karishma Prajapati', 'Nidhi Shah'], 'a non-engineer sees their own row first; everyone else is the engineers');
    const reviewer = user('U5', 'rev@elegrow.com', 'Reviewer', ['reviewer']);
    assert.ok(!build({ me: nidhi, users: [anup, nidhi, reviewer], projects: [] }).rows.some((row) => row.name === 'Reviewer'), 'a reviewer with no task is not a row');
    assert.ok(build({ me: nidhi, users: [anup, nidhi, reviewer], projects: [project('1', [task('r', { owners: [owner(reviewer)] })])] }).rows.some((row) => row.name === 'Reviewer'), 'but owning a task puts them on the calendar');
  });

  await check('two Onsite tasks on one day for one person warn; adjacent days do not; same person on a multi-owner task counts once', () => {
    const clash = build({ projects: [project('1', [task('a', { startDate: '10-19-2026', dueDate: '10-21-2026' }), task('b', { startDate: '10-21-2026', dueDate: '10-22-2026' })])] });
    const warning = clash.warnings.find((entry) => entry.kind === 'double_booked');
    assert.deepStrictEqual(warning.dates, ['2026-10-21', '2026-10-21']);
    assert.deepStrictEqual(rowOf(clash, 'Anup Parekh').clashDates, ['2026-10-21']);
    assert.ok(rowOf(clash, 'Anup Parekh').bars.every((bar) => bar.warnings.includes(warning.id)));
    const adjacent = build({ projects: [project('1', [task('a', { startDate: '10-19-2026', dueDate: '10-20-2026' }), task('b', { startDate: '10-21-2026', dueDate: '10-22-2026' })])] });
    assert.ok(!kinds(adjacent).includes('double_booked'));
    const twoOwners = build({ projects: [project('1', [task('a', { owners: [owner(anup), owner(karishma)] }), task('b', { owners: [owner(karishma)], startDate: '10-22-2026', dueDate: '10-22-2026' })])] });
    assert.strictEqual(twoOwners.warnings.filter((entry) => entry.kind === 'double_booked').length, 1);
    assert.strictEqual(twoOwners.warnings[0].personName, 'Karishma Prajapati', 'only the person who is on both');
    const lanes = rowOf(clash, 'Anup Parekh').bars.map((bar) => bar.lane).sort();
    assert.deepStrictEqual(lanes, [0, 1], 'overlapping bars take separate lanes');
  });

  await check('leave: approved and pending both warn (pending is marked as such); inclusive dates; a leave day Zoho left out does not', () => {
    const employees = [{ employeeId: 'E1', zuid: '1', email: 'anup@elegrow.com', name: 'Anup Parekh', location: 'Surat' }, { employeeId: 'E2', zuid: '2', email: 'karishma@elegrow.com', name: 'Karishma Prajapati', location: 'Surat' }];
    const leave = [
      { id: 'L1', employeeId: 'E1', from: '2026-10-23', to: '2026-10-23', days: [{ date: '2026-10-23', count: 1 }], state: 'approved', type: 'Casual Leave' },
      { id: 'L2', employeeId: 'E2', from: '2026-10-21', to: '2026-10-22', days: [{ date: '2026-10-21', count: 1 }, { date: '2026-10-22', count: 1 }], state: 'pending', type: 'Privilege Leave' },
      { id: 'L3', employeeId: 'E1', from: '2026-10-24', to: '2026-10-25', days: [{ date: '2026-10-24', count: 1 }], state: 'approved', type: 'Casual Leave' }
    ];
    const result = build({ employees, leave, projects: [project('1', [task('a', { startDate: '10-19-2026', dueDate: '10-23-2026' }), task('k', { owners: [owner(karishma)], startDate: '10-22-2026', dueDate: '10-22-2026' })])] });
    const found = result.warnings.filter((entry) => entry.kind === 'leave_overlap');
    assert.strictEqual(found.length, 2);
    assert.ok(found.some((entry) => entry.personName === 'Anup Parekh' && entry.state === 'approved' && entry.dates[0] === '2026-10-23'), 'leave on the last day of the task counts');
    assert.ok(found.some((entry) => entry.personName === 'Karishma Prajapati' && entry.state === 'pending'));
    assert.deepStrictEqual(rowOf(result, 'Karishma Prajapati').leave.map((entry) => [entry.state, entry.dates]), [['pending', ['2026-10-21', '2026-10-22']]]);
    const none = build({ employees, leave: [leave[2]], projects: [project('1', [task('a', { startDate: '10-19-2026', dueDate: '10-23-2026' })])] });
    assert.ok(!kinds(none).includes('leave_overlap'));
    const unmatched = build({ employees: [], leave: [{ id: 'L9', employeeId: 'X', zuid: '', name: 'Anup Parekh', from: '2026-10-19', to: '2026-10-19', days: null, state: 'approved', type: 'x' }], projects: [project('1', [task('a')])] });
    assert.ok(kinds(unmatched).includes('leave_overlap'), 'without an employee record the leave is matched by name');
  });

  await check('working days: Sunday is off by default and a holiday is off too; a location holiday only for that location', () => {
    const result = build({ from: '2026-10-19', to: '2026-10-25', holidays: [{ id: 'h', name: 'Dussehra', locations: [], days: [{ date: '2026-10-20', weight: 1 }] }, { id: 'h2', name: 'Local', locations: ['Surat'], days: [{ date: '2026-10-22', weight: 1 }] }], employees: [{ employeeId: 'E1', email: 'anup@elegrow.com', name: 'Anup Parekh', location: 'Surat' }, { employeeId: 'E2', email: 'karishma@elegrow.com', name: 'Karishma Prajapati', location: 'Envox' }], projects: [project('1', [task('a', { owners: [owner(anup), owner(karishma)] })])] });
    assert.deepStrictEqual(result.days.map((day) => day.working), [true, true, true, true, true, true, false]);
    assert.deepStrictEqual(result.days[1].holidays.map((holiday) => holiday.name), ['Dussehra']);
    assert.deepStrictEqual(rowOf(result, 'Anup Parekh').holidayDates, ['2026-10-20', '2026-10-22']);
    assert.deepStrictEqual(rowOf(result, 'Karishma Prajapati').holidayDates, ['2026-10-20']);
    const sundays = build({ from: '2026-10-19', to: '2026-10-25', settings: { ...settings(), workingDays: [0, 1, 2, 3, 4, 5, 6] } });
    assert.ok(sundays.days.every((day) => day.working));
  });

  await check('quantity vs days: short (holiday and approved leave count against the days), long, and the checks that are skipped', () => {
    const pits = (plannedRaw, extra = {}) => task('p', { testId: 'electrodeResistance', eshaTest: 'Earth Electrode Resistance', plannedRaw, ...extra });
    // 170 pits at 35/day needs 4.86 days; 5 working days (Mon-Fri) is enough, a holiday makes it short
    assert.ok(!kinds(build({ projects: [project('1', [pits('170')])] })).includes('quantity_short'));
    const holiday = build({ holidays: [{ id: 'h', name: 'Dussehra', locations: [], days: [{ date: '2026-10-20', weight: 1 }] }], projects: [project('1', [pits('170')])] });
    assert.strictEqual(holiday.warnings.find((entry) => entry.kind === 'quantity_short').kind, 'quantity_short');
    assert.match(holiday.warnings.find((entry) => entry.kind === 'quantity_short').message, /170 pits at about 35 a day needs 4\.9 working days of one engineer; 4 are available/);
    // approved leave takes a day away, pending leave does not
    const employees = [{ employeeId: 'E1', email: 'anup@elegrow.com', name: 'Anup Parekh', location: '' }];
    const approved = build({ employees, leave: [{ id: 'L', employeeId: 'E1', from: '2026-10-21', to: '2026-10-21', days: null, state: 'approved', type: 'x' }], projects: [project('1', [pits('170')])] });
    assert.ok(kinds(approved).includes('quantity_short'));
    const pending = build({ employees, leave: [{ id: 'L', employeeId: 'E1', from: '2026-10-21', to: '2026-10-21', days: null, state: 'pending', type: 'x' }], projects: [project('1', [pits('170')])] });
    assert.ok(!kinds(pending).includes('quantity_short'));
    // two engineers share the work
    assert.ok(!kinds(build({ holidays: [{ id: 'h', name: 'x', locations: [], days: [{ date: '2026-10-20', weight: 1 }] }], projects: [project('1', [pits('170', { owners: [owner(anup), owner(karishma)] })])] })).includes('quantity_short'));
  });

  await check('quantity vs days: the long and skipped cases', () => {
    const pits = (plannedRaw, extra = {}) => task('p', { testId: 'electrodeResistance', eshaTest: 'Earth Electrode Resistance', plannedRaw, ...extra });
    assert.ok(kinds(build({ projects: [project('1', [pits('10')])] })).includes('quantity_long'), '10 pits in 5 working days: needs 1, booked 5');
    assert.ok(!kinds(build({ projects: [project('1', [pits('100', { startDate: '10-19-2026', dueDate: '10-23-2026' })])] })).includes('quantity_long'), 'close to the need: no warning');
    assert.ok(!kinds(build({ projects: [project('1', [pits('100', { startDate: '10-19-2026', dueDate: '10-22-2026' })])] })).includes('quantity_long'));
    assert.ok(!kinds(build({ projects: [project('1', [pits('about 170')])] })).includes('quantity_short'), 'an unclear quantity is not checked');
    assert.ok(!kinds(build({ projects: [project('1', [pits(''), pits('', { id: 'q' })])] })).includes('quantity_short'), 'no quantity: not checked');
    assert.ok(!kinds(build({ projects: [project('1', [pits('500', { id: 'a' }), pits('500', { id: 'b', startDate: '10-24-2026', dueDate: '10-24-2026' })])] })).includes('quantity_short'), 'two Onsite tasks of one test: the quantity cannot be split, so it is not checked');
    assert.ok(!kinds(build({ projects: [project('1', [task('n', { plannedRaw: '500' })])] })).includes('quantity_short'), 'no ESHA Test: no rate');
    const edited = build({ settings: mergeSettings(settings(), { rates: { electrodeResistance: 20 } }).settings, projects: [project('1', [pits('170')])] });
    assert.ok(kinds(edited).includes('quantity_short'), 'the admin-edited rate is used');
  });

  await check('no owner, no dates and a single date are warned about, never hidden or blocked', () => {
    const result = build({ projects: [project('1', [task('o', { owners: [] }), task('d', { startDate: '', dueDate: '' }), task('s', { owners: [owner(karishma)], startDate: '10-21-2026', dueDate: '' }), task('r', { owners: [owner(nidhi)], startDate: '10-23-2026', dueDate: '10-21-2026' })])] });
    assert.deepStrictEqual(kinds(result), ['dates_reversed', 'no_dates', 'no_owner', 'one_date']);
    const all = result.rows.flatMap((row) => row.bars.map((bar) => bar.taskId));
    assert.ok(!all.includes('d'), 'a task with no dates cannot be placed');
    assert.ok(all.includes('s') && all.includes('r'));
    assert.strictEqual(rowOf(result, 'Karishma Prajapati').bars.find((bar) => bar.taskId === 's').end, '2026-10-21');
  });

  // ---- the People client: only reads ---------------------------------------------------------------
  const recorder = (handlers = {}) => {
    const calls = [];
    return {
      calls,
      get: async (url, options) => {
        calls.push({ method: 'GET', url, params: options.params, auth: options.headers && options.headers.Authorization });
        return handlers.get(url, options.params, calls);
      },
      post: async (url) => {
        calls.push({ method: 'POST', url });
        return { data: { access_token: `tok-${calls.filter((call) => call.method === 'POST').length}` } };
      }
    };
  };
  const leaveRecord = (n) => ({ From: '21-Oct-2026', To: '21-Oct-2026', ApprovalStatus: 'Approved', Type: 'PAID', Leavetype: 'Casual Leave', Employee: `P${n}`, EmployeeId: `E${n}`, Days: { '21-Oct-2026': { LeaveCount: '1.0' } } });

  await check('People client: leave is paged (200 a call) and de-duplicated; a server that ignores the page index does not loop', async () => {
    const http = recorder({
      get: (_url, params) => ({ data: { records: Object.fromEntries(Array.from({ length: params.startIndex === 1 ? 200 : 50 }, (_, i) => [`R${params.startIndex + i}`, leaveRecord(params.startIndex + i)])) } })
    });
    const people = createPeopleClient({ http, config });
    const leave = await people.listLeave('2026-10-01', '2026-10-31');
    assert.strictEqual(leave.length, 250);
    assert.deepStrictEqual(http.calls.filter((call) => call.method === 'GET').map((call) => call.params.startIndex), [1, 201]);
    assert.strictEqual(http.calls.find((call) => call.method === 'GET').params.from, '01-Oct-2026', 'the dates go out in the format People requires');
    const stuck = recorder({ get: () => ({ data: { records: Object.fromEntries(Array.from({ length: 200 }, (_, i) => [`R${i}`, leaveRecord(i)])) } }) });
    assert.strictEqual((await createPeopleClient({ http: stuck, config }).listLeave('2026-10-01', '2026-10-31')).length, 200);
    assert.strictEqual(stuck.calls.filter((call) => call.method === 'GET').length, 2);
  });

  await check('People client: employees (paged), holidays, forms; the token is reused and refreshed once on a 401', async () => {
    let expired = true;
    const http = recorder({
      get: (url, params) => {
        if (expired && url.endsWith('/people/api/forms')) { expired = false; throw Object.assign(new Error('x'), { response: { status: 401, data: {} } }); }
        if (url.endsWith('/people/api/forms')) return { data: { response: { result: [{ displayName: 'Leave', formLinkName: 'leave' }] } } };
        if (url.includes('/employee/getRecords')) {
          const n = params.sIndex === 1 ? 200 : 7;
          return { data: { response: { result: Array.from({ length: n }, (_, i) => ({ [`${params.sIndex + i}`]: [{ EmployeeID: `E${params.sIndex + i}`, EmailID: `P${i}@x.com`, FirstName: 'A', LastName: 'B', ZUID: '9', Employeestatus: 'Active', LocationName: 'Surat' }] })) } } };
        }
        return { data: { holidays: [{ id: 'h', name: 'Dussehra', from: '20-Oct-2026', to: '20-Oct-2026', locations: [], isRestricted: false, dayDetails: [{ '20-Oct-2026': { session: 0, ishalfday: false } }] }] } };
      }
    });
    const people = createPeopleClient({ http, config });
    assert.deepStrictEqual(await people.listForms(), [{ displayName: 'Leave', formLinkName: 'leave' }]);
    assert.strictEqual((await people.listEmployees()).length, 207);
    assert.strictEqual((await people.listHolidays('2026-10-01', '2026-10-31'))[0].name, 'Dussehra');
    const posts = http.calls.filter((call) => call.method === 'POST');
    assert.strictEqual(posts.length, 2, 'one token, one refresh after the 401, then reused');
    assert.ok(posts.every((call) => call.url === 'https://accounts.test/oauth/v2/token'));
    assert.ok(http.calls.filter((call) => call.method === 'GET').every((call) => /^Zoho-oauthtoken tok-/.test(call.auth)));
    const failing = createPeopleClient({ http: recorder({ get: () => { throw Object.assign(new Error('boom'), { response: { status: 401, data: { error: { message: 'Invalid OAuth Scope' } } } }); } }), config });
    await assert.rejects(() => failing.listLeave('2026-10-01', '2026-10-31'), (error) => /401/.test(error.message) && /People scopes/.test(error.message) && !/tok-|boom/.test(error.message), 'the error says what to check and carries no token');
  });

  await check('read only: the only non-GET Zoho call in the schedule code is the OAuth token refresh', () => {
    const src = (file) => fs.readFileSync(path.join(__dirname, '..', 'src', file), 'utf8');
    ['schedule.js', 'scheduleService.js'].forEach((file) => assert.ok(!/\b(axios|http|zoho|people)\.(post|put|patch|delete|request)\(|method:\s*['"](POST|PUT|PATCH|DELETE)/i.test(src(file)), `${file} makes no write call`));
    const people = src('zohoPeople.js');
    assert.strictEqual((people.match(/http\.(post|put|patch|delete)\(/g) || []).length, 1, 'zohoPeople.js: one POST');
    assert.ok(/http\.post\(`\$\{String\(config\.zoho\.accountsBaseUrl\)[^`]*\/oauth\/v2\/token`/.test(people), '... and it is the token refresh');
    const server = src('server.js');
    assert.ok(!/zoho[A-Za-z.]*\.(createTask|updateTask|updateProject)/.test(server));
  });

  // ---- the cache service ------------------------------------------------------------------------
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'schedule-'));
  const mkStore = (name) => createReportStore({ filePath: path.join(dir, `${name}.db`) });
  const zProject = (id, name) => ({ id, name, projectNumber: name, clientName: 'C', ownerName: 'O', stage: 'Planning', workOrderNo: '' });
  const rawOnsite = (id, name) => ({ id, name: `[ESHA] ${name}`, owners: [owner(anup)], startDate: '10-20-2026', dueDate: '10-21-2026', status: 'Planned', customFields: [{ label: 'Work Location', apiName: 'category', value: 'Onsite' }, { label: 'ESHA Test', apiName: 'UDF_CHAR90', value: 'Soil Resistivity' }] });
  const fakeZoho = (state) => ({
    calls: [],
    getProjects: async () => { if (state.projectsFail) throw new Error('Zoho API request failed (500)'); return state.projects; },
    getPortalUsers: async () => [{ id: anup.id, displayName: 'Anup Parekh', email: anup.email }],
    getProjectTasks: async (id) => { if (state.failFor === id) throw new Error('Zoho API request failed (500)'); return state.tasks[id] || []; }
  });
  const fakePeople = (state) => ({
    employees: 0,
    listEmployees: async () => { if (state.peopleFail) throw new Error('Zoho People request failed (401). Check the People scopes of the grant.'); return [{ employeeId: 'E1', email: anup.email, name: 'Anup Parekh', location: 'Surat' }]; },
    listLeave: async () => { if (state.peopleFail) throw new Error('Zoho People request failed (401). Check the People scopes of the grant.'); return state.leave || []; },
    listHolidays: async () => state.holidays || []
  });
  const quiet = { error() {} };

  await check('cache: the first view reads Zoho and stores it with its age; later views come from the cache, not from Zoho', async () => {
    const state = { projects: [zProject('P1', '001-D03_TS-26-AAA'), zProject('P2', 'Corporate Website')], tasks: { P1: [rawOnsite('t1', 'Soil Resistivity Testing')], P2: [rawOnsite('x', 'Not D03')] } };
    const zoho = fakeZoho(state);
    let calls = 0;
    const counting = { ...zoho, getProjectTasks: async (id) => { calls += 1; return zoho.getProjectTasks(id); } };
    const store = mkStore('a');
    let clock = Date.parse('2026-10-19T10:00:00Z');
    const service = createScheduleService({ store, zoho: counting, people: fakePeople(state), config, now: () => clock, log: quiet });
    const first = await service.snapshot({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(calls, 1, 'only the D03 project is read');
    assert.strictEqual(rowOf(first, 'Anup Parekh').bars.length, 1);
    assert.ok(first.zoho.tasks.fetchedAt);
    clock += 5 * 60 * 1000;
    const again = await service.snapshot({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(calls, 1, 'served from the cache');
    assert.ok(again.zoho.tasks.ageSeconds >= 300, `age is ${again.zoho.tasks.ageSeconds}`);
    const other = await service.snapshot({ from: '2026-11-02', to: '2026-11-08', user: anup });
    assert.ok(other.zoho.leave.loaded && other.zoho.leave.fetchedAt, 'a month not read yet is read when it is shown');
    assert.strictEqual(calls, 1);
  });

  await check('cache: a failed refresh keeps the last data with its age and the reason; a project that fails keeps its older tasks', async () => {
    const state = { projects: [zProject('P1', '001-D03_TS-26-AAA'), zProject('P2', '002-D03_TS-26-BBB')], tasks: { P1: [rawOnsite('t1', 'Soil Resistivity Testing')], P2: [rawOnsite('t2', 'Soil Resistivity Testing')] } };
    const people = fakePeople(state);
    const store = mkStore('b');
    let clock = Date.parse('2026-10-19T10:00:00Z');
    const service = createScheduleService({ store, zoho: fakeZoho(state), people, config: { ...config, schedule: { ...config.schedule, manualRefreshSeconds: 0 } }, now: () => clock, log: quiet });
    await service.snapshot({ from: '2026-10-19', to: '2026-10-25', user: anup });
    clock += 3600 * 1000;
    state.failFor = 'P2';
    state.tasks.P1 = [rawOnsite('t1', 'Soil Resistivity Testing'), rawOnsite('t3', 'Soil Resistivity Testing')];
    let result = await service.refresh({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(rowOf(result, 'Anup Parekh').bars.length, 3, 'P1 updated, P2 kept from before');
    assert.match(result.zoho.tasks.error, /1 project could not be read/);
    state.projectsFail = true;
    state.peopleFail = true;
    clock += 3600 * 1000;
    result = await service.refresh({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(rowOf(result, 'Anup Parekh').bars.length, 3, 'the whole task list kept when Zoho fails');
    assert.match(result.zoho.tasks.error, /500/);
    assert.ok(result.zoho.tasks.ageSeconds >= 3600, 'and shown with its age');
    assert.match(result.zoho.leave.error, /People scopes/);
    assert.ok(result.zoho.leave.loaded, 'old leave data is still used');
    state.projectsFail = false;
    state.peopleFail = false;
    state.failFor = '';
    clock += 3600 * 1000;
    result = await service.refresh({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(result.zoho.tasks.error, '', 'a good read clears the error');
  });

  await check('refresh is limited to once a minute (the data of the first is returned, flagged)', async () => {
    const state = { projects: [zProject('P1', '001-D03_TS-26-AAA')], tasks: { P1: [rawOnsite('t1', 'Soil Resistivity Testing')] } };
    let reads = 0;
    const zoho = { ...fakeZoho(state), getProjects: async () => { reads += 1; return state.projects; } };
    let clock = Date.parse('2026-10-19T10:00:00Z');
    const service = createScheduleService({ store: mkStore('c'), zoho, people: fakePeople(state), config, now: () => clock, log: quiet });
    const first = await service.refresh({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(first.throttled, false);
    assert.strictEqual(reads, 1);
    clock += 20 * 1000;
    const second = await service.refresh({ from: '2026-10-19', to: '2026-10-25', user: anup });
    assert.strictEqual(second.throttled, true);
    assert.strictEqual(second.retryInSeconds, 40);
    assert.strictEqual(reads, 1, 'Zoho was not read again');
    assert.ok(second.rows.length, 'the cached schedule is still returned');
    clock += 41 * 1000;
    assert.strictEqual((await service.refresh({ from: '2026-10-19', to: '2026-10-25', user: anup })).throttled, false);
    assert.strictEqual(reads, 2);
  });

  await check('settings are stored (admin edits) and used; the defaults come back with an unusable saved value', async () => {
    const store = mkStore('d');
    const service = createScheduleService({ store, zoho: fakeZoho({ projects: [] }), people: fakePeople({}), config, log: quiet });
    assert.deepStrictEqual(service.getSettings().workingDays, [1, 2, 3, 4, 5, 6]);
    assert.ok(service.saveSettings({ workingDays: [] }, 'a').error);
    service.saveSettings({ workingDays: [1, 2, 3, 4, 5], rates: { soilResistivity: 5 } }, 'a');
    assert.deepStrictEqual(service.getSettings().workingDays, [1, 2, 3, 4, 5]);
    assert.strictEqual(service.getSettings().rates.soilResistivity.perDay, 5);
    store.setScheduleSetting('settings', { workingDays: 'x' }, 'a');
    assert.deepStrictEqual(service.getSettings().workingDays, [1, 2, 3, 4, 5, 6], 'a damaged saved value falls back to the defaults');
  });

  // ---- over HTTP on the mock portal ----------------------------------------------------------------
  const port = await freePort();
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'src', 'server.js')], {
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'http.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ANTHROPIC_API_KEY: '', OPENAI_API_KEY: '', GEMINI_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    server.on('exit', (code) => reject(new Error(`server exited early (${code})`)));
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const call = createCaller(`http://127.0.0.1:${port}`);
  const WINDOW = 'from=2026-10-19&to=2026-11-15';

  try {
    await check('HTTP: signed-out requests are refused', async () => {
      assert.strictEqual((await call('GET', `/api/schedule?${WINDOW}`, undefined, null)).status, 401);
      assert.strictEqual((await call('POST', '/api/schedule/refresh', {}, null)).status, 401);
    });

    let schedule;
    await check('HTTP: every signed-in person sees the whole team, their own row first (an engineer with no task too)', async () => {
      const mine = await call('GET', `/api/schedule?${WINDOW}`, undefined, 'nidhi@elegrow.com');
      assert.strictEqual(mine.status, 200);
      schedule = mine.body;
      const names = schedule.rows.map((row) => row.name);
      assert.strictEqual(names[0], 'Nidhi Shah');
      assert.ok(schedule.rows[0].isMe);
      ['Anup Parekh', 'Karishma Prajapati', 'Mohana Vamsee', 'No owner'].forEach((name) => assert.ok(names.includes(name), `${name} is on the calendar`));
      const other = (await call('GET', `/api/schedule?${WINDOW}`, undefined, 'anup@elegrow.com')).body;
      assert.strictEqual(other.rows[0].name, 'Anup Parekh');
      assert.deepStrictEqual([...other.rows.map((row) => row.name)].sort(), [...names].sort(), 'an engineer sees the same team');
      const reviewer = (await call('GET', `/api/schedule?${WINDOW}`, undefined, 'reviewer@elegrow.com')).body;
      assert.strictEqual(reviewer.rows[0].name, 'A. P.');
      assert.ok(reviewer.rows.some((row) => row.name === 'Anup Parekh'), 'a reviewer sees the team too');
      assert.strictEqual(rowOf(schedule, 'Mohana Vamsee').kind, 'zoho');
    });

    await check('HTTP: the Onsite ESHA tasks, not the closed, IRT or offsite ones, with their warnings', async () => {
      const bars = schedule.rows.flatMap((row) => row.bars.map((bar) => `${row.name}:${bar.taskId}`)).sort();
      assert.deepStrictEqual(bars, ['Anup Parekh:g1', 'Anup Parekh:r3', 'Anup Parekh:r4', 'Karishma Prajapati:g2', 'Mohana Vamsee:g5', 'No owner:g3']);
      assert.deepStrictEqual(kinds(schedule), ['double_booked', 'leave_overlap', 'leave_overlap', 'leave_overlap', 'no_dates', 'no_owner', 'quantity_long', 'quantity_short', 'quantity_short']);
      const warning = (kind, personName) => schedule.warnings.find((entry) => entry.kind === kind && (!personName || entry.personName === personName));
      assert.deepStrictEqual(warning('double_booked').taskIds, ['g1', 'r4']);
      assert.strictEqual(warning('leave_overlap', 'Karishma Prajapati').state, 'approved');
      assert.strictEqual(warning('leave_overlap', 'Anup Parekh').state, 'pending');
      assert.match(schedule.warnings.find((entry) => entry.kind === 'quantity_short' && entry.taskIds[0] === 'r4').message, /Dussehra|170 pits/, '170 pits over 19-23 Oct falls short: 20 Oct is a holiday');
      assert.ok(schedule.days.find((day) => day.date === '2026-10-20').holidays.some((holiday) => holiday.name === 'Dussehra'));
      assert.strictEqual(schedule.projects.length, 2);
      assert.strictEqual(schedule.zoho.mock, true);
      const surat = rowOf(schedule, 'Anup Parekh').holidayDates;
      assert.ok(surat.includes('2026-10-30') && surat.includes('2026-10-20'), 'a Surat-only holiday applies to Surat people');
      assert.ok(!rowOf(schedule, 'Mohana Vamsee').holidayDates.includes('2026-10-30'), 'but not to someone at another location');
    });

    await check('HTTP: refresh works for everyone and is limited to once a minute', async () => {
      const first = await call('POST', '/api/schedule/refresh', { from: '2026-10-19', to: '2026-10-25' }, 'anup@elegrow.com');
      assert.strictEqual(first.status, 200);
      assert.strictEqual(first.body.throttled, false);
      const second = await call('POST', '/api/schedule/refresh', { from: '2026-10-19', to: '2026-10-25' }, 'nidhi@elegrow.com');
      assert.strictEqual(second.status, 200);
      assert.strictEqual(second.body.throttled, true);
      assert.ok(second.body.retryInSeconds > 0 && second.body.rows.length > 0);
    });

    await check('HTTP: only an admin edits the settings; the change is audited and used by the warnings', async () => {
      assert.strictEqual((await call('PUT', '/api/admin/schedule-settings', { rates: { electrodeResistance: 50 } }, 'engineer@elegrow.com')).status, 403);
      assert.strictEqual((await call('PUT', '/api/admin/schedule-settings', { workingDays: [] }, 'admin@elegrow.com')).status, 400);
      const saved = await call('PUT', '/api/admin/schedule-settings', { workingDays: [1, 2, 3, 4, 5, 6], rates: { electrodeResistance: 50 } }, 'admin@elegrow.com');
      assert.strictEqual(saved.status, 200);
      assert.strictEqual(saved.body.rates.electrodeResistance.perDay, 50);
      const after = (await call('GET', `/api/schedule?${WINDOW}`, undefined, 'nidhi@elegrow.com')).body;
      assert.ok(!after.warnings.some((entry) => entry.kind === 'quantity_short' && entry.taskIds[0] === 'r4'), '170 pits at 50 a day fits in 4 working days');
      assert.strictEqual(after.settings.rates.electrodeResistance.perDay, 50);
      const audit = (await call('GET', '/api/admin/audit?action=schedule.', undefined, 'admin@elegrow.com')).body;
      assert.strictEqual(audit.entries.length, 1);
      assert.strictEqual(audit.entries[0].userEmail, 'admin@elegrow.com');
    });

    await check('HTTP: the window is limited and a bad date falls back (the page cannot ask for years of data)', async () => {
      const big = (await call('GET', '/api/schedule?from=2026-01-01&to=2030-01-01', undefined, 'nidhi@elegrow.com')).body;
      assert.ok(big.days.length <= 93);
      const bad = await call('GET', '/api/schedule?from=nonsense&to=', undefined, 'nidhi@elegrow.com');
      assert.strictEqual(bad.status, 200);
      assert.ok(bad.body.days.length === 7);
    });
  } finally {
    server.kill();
  }

  console.log(`\n${passed} schedule checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
