// Scheduling (Session 6F): who is on which site, when, with clashes and leave flagged. Read only
// toward Zoho. Zoho Projects is the source of truth for the dates: they are cached with a time, never
// copied into our own tables.
//
// Rules (agreed):
//   - Every Onsite ESHA task of a D03 project that is not closed shows, whatever the status of the
//     project, when it overlaps the visible window. Tasks without dates cannot be placed: they are
//     listed in the warnings.
//   - Rows: the signed-in user first, then the engineers of our app, then anyone else who owns a task
//     in the window. People who never signed in get a row from their Zoho name. Tasks nobody owns
//     go in a "No owner" row. Everyone signed in sees the whole team.
//   - Leave (approved and pending) and holidays come from Zoho People. Holidays are non-working days,
//     as are the days that are not working days in the settings (default Mon-Sat). A holiday with
//     locations applies to the people of those locations.
//   - Warnings are shown, never blocking: two onsite tasks on one day, an onsite task over leave, no
//     owner, no dates, and a planned quantity that looks unrealistic against the days available.
//     The quantity check uses a rate per test per engineer per working day: an estimate that an admin edits.

const { parseQuantity, ESHA_TESTS } = require('./zohoTasks');

const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const lower = (value) => text(value).toLowerCase();
const DAY_MS = 24 * 60 * 60 * 1000;
const MAX_WINDOW_DAYS = 93;
const WEEKDAYS = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];

// ---- Dates (ISO yyyy-mm-dd, all arithmetic in UTC so a date never shifts) ----------------------
const isIso = (value) => /^\d{4}-\d{2}-\d{2}$/.test(text(value)) && !Number.isNaN(Date.parse(`${value}T00:00:00Z`));
const toMs = (iso) => Date.parse(`${iso}T00:00:00Z`);
const fromMs = (ms) => new Date(ms).toISOString().slice(0, 10);
const addDays = (iso, count) => fromMs(toMs(iso) + count * DAY_MS);
const weekdayOf = (iso) => new Date(`${iso}T00:00:00Z`).getUTCDay();
function eachDay(from, to) {
  const days = [];
  for (let ms = toMs(from); ms <= toMs(to); ms += DAY_MS) days.push(fromMs(ms));
  return days;
}

// Zoho Projects v1 dates are "MM-DD-YYYY" (portal local, no time zone); the mock and v3 use ISO.
function parseZohoDate(value) {
  const raw = text(value);
  const v1 = /^(\d{2})-(\d{2})-(\d{4})$/.exec(raw);
  const iso = v1 ? `${v1[3]}-${v1[1]}-${v1[2]}` : (/^(\d{4}-\d{2}-\d{2})/.exec(raw) || [])[1] || '';
  return isIso(iso) && fromMs(toMs(iso)) === iso ? iso : '';
}

const monthKey = (iso) => iso.slice(0, 7);
function monthsOf(from, to) {
  const months = [];
  for (let key = monthKey(from); key <= monthKey(to); ) {
    months.push(key);
    const [year, month] = key.split('-').map(Number);
    key = month === 12 ? `${year + 1}-01` : `${year}-${String(month + 1).padStart(2, '0')}`;
  }
  return months;
}
const monthRange = (key) => {
  const [year, month] = key.split('-').map(Number);
  return { from: `${key}-01`, to: fromMs(Date.UTC(year, month, 0)) };
};

function clampWindow(from, to) {
  const start = isIso(from) ? from : fromMs(Date.now());
  let end = isIso(to) && to >= start ? to : addDays(start, 6);
  if (toMs(end) - toMs(start) > (MAX_WINDOW_DAYS - 1) * DAY_MS) end = addDays(start, MAX_WINDOW_DAYS - 1);
  return { from: start, to: end };
}

// ---- Settings ----------------------------------------------------------------------------------
// Rates: units of work one engineer does in one working day. Only the pit rate is backed by a real
// project (056-D03 plans 170 pits in 5 days, about 34 a day); the others are estimates to be edited.
const DEFAULT_RATES = {
  electrodeResistance: { perDay: 35, unit: 'pits' },
  soilResistivity: { perDay: 4, unit: 'locations' },
  continuityTest: { perDay: 60, unit: 'points' },
  loopImpedanceTest: { perDay: 25, unit: 'panels' },
  prospectiveFaultCurrent: { perDay: 25, unit: 'panels' },
  riserIntegrityTest: { perDay: 40, unit: 'points' },
  groundResistance: { perDay: 6, unit: 'points' },
  stepTouchVoltage: { perDay: 15, unit: 'points' },
  towerFootingResistance: { perDay: 6, unit: 'towers' }
};

function defaultSettings(config) {
  const names = (config && config.schedule && config.schedule.workingDays) || ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  const days = [...new Set(names.map((name) => WEEKDAYS.indexOf(lower(name).slice(0, 3))).filter((index) => index >= 0))].sort();
  return { workingDays: days.length ? days : [1, 2, 3, 4, 5, 6], rates: JSON.parse(JSON.stringify(DEFAULT_RATES)) };
}

// Validates an admin edit against the current settings. Returns { settings } or { error }.
function mergeSettings(current, patch) {
  const next = { workingDays: [...current.workingDays], rates: JSON.parse(JSON.stringify(current.rates)) };
  const input = patch && typeof patch === 'object' ? patch : {};
  if (input.workingDays !== undefined) {
    const days = Array.isArray(input.workingDays) ? [...new Set(input.workingDays.map(Number))] : [];
    if (!days.length || days.some((day) => !Number.isInteger(day) || day < 0 || day > 6)) return { error: 'Choose at least one working day.' };
    next.workingDays = days.sort();
  }
  if (input.rates !== undefined) {
    for (const [testId, value] of Object.entries(input.rates || {})) {
      if (!next.rates[testId]) return { error: `Unknown test "${testId}".` };
      const perDay = Number(value && typeof value === 'object' ? value.perDay : value);
      if (!Number.isFinite(perDay) || perDay <= 0 || perDay > 10000) return { error: `The rate for ${testId} must be a number above 0.` };
      next.rates[testId].perDay = perDay;
    }
  }
  return { settings: next };
}

// ---- The schedule ------------------------------------------------------------------------------
function colorOf(projectId) {
  let hash = 0;
  for (const char of text(projectId)) hash = (hash * 31 + char.charCodeAt(0)) % 1000;
  return hash % 10;
}

const dateLabel = (iso) => `${Number(iso.slice(8))} ${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][Number(iso.slice(5, 7)) - 1]}`;
const rangeLabel = (from, to) => (from === to ? dateLabel(from) : `${dateLabel(from)} to ${dateLabel(to)}`);
const fixed = (value) => String(Math.round(value * 10) / 10);

// input: { from, to, projects (the cached Onsite ESHA tasks per D03 project), employees, leave, holidays,
//          users (our users), me (the signed-in user), settings, hiddenStatuses, workspaces ({ zohoProjectId: id }) }
function buildSchedule(input) {
  const { from, to } = clampWindow(input.from, input.to);
  const settings = input.settings;
  const users = (input.users || []).filter((user) => user);
  const me = input.me || null;
  const hidden = (input.hiddenStatuses || []).map(lower);

  // -- people ---------------------------------------------------------------------------------
  const employees = input.employees || [];
  const employeeByEmail = new Map(employees.filter((e) => e.email).map((e) => [e.email, e]));
  const employeeById = new Map(employees.filter((e) => e.employeeId).map((e) => [e.employeeId, e]));
  const employeeByZuid = new Map(employees.filter((e) => e.zuid).map((e) => [e.zuid, e]));
  const userByEmail = new Map(users.map((user) => [lower(user.email), user]));
  const nameCount = new Map();
  users.forEach((user) => nameCount.set(lower(user.name), (nameCount.get(lower(user.name)) || 0) + 1));
  const userByName = new Map(users.filter((user) => nameCount.get(lower(user.name)) === 1).map((user) => [lower(user.name), user]));

  const people = new Map(); // key -> { key, name, email, user, location }
  function personFor({ email = '', name = '' }) {
    const mail = lower(email);
    const user = (mail && userByEmail.get(mail)) || (name && userByName.get(lower(name))) || null;
    const key = user ? lower(user.email) : mail || (name ? `name:${lower(name)}` : '');
    if (!key) return null;
    if (!people.has(key)) {
      const employee = employeeByEmail.get(key) || (name && employees.find((e) => lower(e.name) === lower(name))) || null;
      people.set(key, { key, name: (user && user.name) || text(name) || (employee && employee.name) || mail, email: (user && lower(user.email)) || mail, user, location: employee ? employee.location : '' });
    }
    return people.get(key);
  }

  // -- working days and holidays ------------------------------------------------------------------
  const holidays = [];
  const seenHoliday = new Set();
  (input.holidays || []).forEach((holiday) => {
    if (seenHoliday.has(holiday.id || holiday.name + holiday.days.map((d) => d.date).join())) return;
    seenHoliday.add(holiday.id || holiday.name + holiday.days.map((d) => d.date).join());
    holidays.push(holiday);
  });
  function weightFor(date, location) {
    if (!settings.workingDays.includes(weekdayOf(date))) return 0;
    let weight = 1;
    holidays.forEach((holiday) => {
      const applies = !holiday.locations.length || holiday.locations.some((entry) => lower(entry) === lower(location));
      const day = applies && holiday.days.find((entry) => entry.date === date);
      if (day) weight = Math.min(weight, 1 - day.weight);
    });
    return weight;
  }

  const days = eachDay(from, to).map((date) => ({
    date,
    weekday: WEEKDAYS[weekdayOf(date)],
    working: settings.workingDays.includes(weekdayOf(date)),
    holidays: holidays.filter((holiday) => holiday.days.some((entry) => entry.date === date)).map((holiday) => ({ name: holiday.name, locations: holiday.locations }))
  }));

  // -- leave per person ---------------------------------------------------------------------------
  const leaveByKey = new Map();
  const seenLeave = new Set();
  (input.leave || []).forEach((leave) => {
    if (leave.id && seenLeave.has(leave.id)) return;
    seenLeave.add(leave.id);
    const employee = employeeById.get(leave.employeeId) || employeeByZuid.get(leave.zuid) || null;
    const person = personFor({ email: employee ? employee.email : '', name: employee ? employee.name : leave.name });
    if (!person) return;
    const dates = new Set((leave.days || eachDay(leave.from, leave.to).map((date) => ({ date }))).map((entry) => entry.date));
    if (!leaveByKey.has(person.key)) leaveByKey.set(person.key, []);
    leaveByKey.get(person.key).push({ id: leave.id, from: leave.from, to: leave.to, state: leave.state, type: leave.type, dates });
  });

  // -- tasks ----------------------------------------------------------------------------------
  const warnings = [];
  const projectsOut = new Map();
  const taskItems = []; // one per task
  (input.projects || []).forEach((project) => {
    const onsiteByTest = new Map();
    (project.tasks || []).forEach((task) => onsiteByTest.set(task.testId, (onsiteByTest.get(task.testId) || 0) + 1));
    (project.tasks || []).forEach((task) => {
      if (task.closed || hidden.includes(lower(task.status))) return;
      let start = parseZohoDate(task.startDate);
      let end = parseZohoDate(task.dueDate);
      const noDates = !start && !end;
      if (!noDates && (!start || !end)) {
        start = start || end;
        end = end || start;
      }
      if (start && end && end < start) [start, end] = [end, start];
      const placed = !noDates && start <= to && end >= from;
      if (!noDates && !placed) return;
      taskItems.push({ project, task, start, end, noDates, oneDate: !noDates && (!parseZohoDate(task.startDate) || !parseZohoDate(task.dueDate)), reversed: !noDates && parseZohoDate(task.startDate) && parseZohoDate(task.dueDate) && parseZohoDate(task.dueDate) < parseZohoDate(task.startDate), singleTest: onsiteByTest.get(task.testId) === 1 });
    });
  });

  const warn = (kind, message, extra = {}) => {
    const entry = { id: `${kind}:${extra.taskIds ? extra.taskIds.join('+') : ''}:${extra.personKey || ''}:${(extra.dates || []).join('..')}`, kind, message, taskIds: [], personKey: '', personName: '', projectName: '', dates: [], ...extra };
    if (!warnings.some((existing) => existing.id === entry.id)) warnings.push(entry);
    return entry.id;
  };

  const barsByPerson = new Map();
  const noOwnerBars = [];
  const bars = [];
  taskItems.forEach((item) => {
    const { project, task, start, end } = item;
    const owners = (task.owners || []).map((owner) => personFor({ email: owner.email, name: owner.fullName || owner.name })).filter(Boolean);
    const uniqueOwners = [...new Map(owners.map((person) => [person.key, person])).values()];
    const warningIds = [];
    const taskLabel = `${task.name}${project.name ? ` (${project.name})` : ''}`;
    const about = { taskIds: [task.id], projectName: project.name };

    if (item.noDates) warningIds.push(warn('no_dates', `${taskLabel} has no start or end date, so it is not on the calendar.`, about));
    else {
      if (item.oneDate) warningIds.push(warn('one_date', `${taskLabel} has only one date: shown as a single day.`, { ...about, dates: [start] }));
      if (item.reversed) warningIds.push(warn('dates_reversed', `${taskLabel} ends before it starts in Zoho: shown from ${dateLabel(start)} to ${dateLabel(end)}.`, { ...about, dates: [start, end] }));
    }
    if (!uniqueOwners.length) warningIds.push(warn('no_owner', `${taskLabel} has no owner.`, about));

    // planned quantity against the days available
    const planned = parseQuantity(task.plannedRaw).value;
    const rate = settings.rates[task.testId];
    if (!item.noDates && planned && rate && item.singleTest && uniqueOwners.length) {
      const nominal = (person) => eachDay(start, end).reduce((sum, date) => sum + weightFor(date, person.location), 0);
      const available = uniqueOwners.reduce((sum, person) => {
        const leave = (leaveByKey.get(person.key) || []).filter((entry) => entry.state === 'approved');
        return sum + eachDay(start, end).reduce((total, date) => total + (leave.some((entry) => entry.dates.has(date)) ? 0 : weightFor(date, person.location)), 0);
      }, 0);
      const needed = planned / rate.perDay;
      if (available + 1e-9 < needed) {
        warningIds.push(warn('quantity_short', `${taskLabel}: ${planned} ${rate.unit} at about ${rate.perDay} a day needs ${fixed(needed)} working days of one engineer; ${fixed(available)} are available between ${dateLabel(start)} and ${dateLabel(end)}.`, { ...about, dates: [start, end] }));
      } else {
        const perPerson = uniqueOwners.reduce((sum, person) => sum + nominal(person), 0) / uniqueOwners.length;
        const neededDays = Math.ceil(planned / (rate.perDay * uniqueOwners.length));
        if (perPerson > 2 * neededDays && perPerson - neededDays > 3) {
          warningIds.push(warn('quantity_long', `${taskLabel}: ${planned} ${rate.unit} at about ${rate.perDay} a day needs about ${neededDays} working days, but ${fixed(perPerson)} are booked.`, { ...about, dates: [start, end] }));
        }
      }
    }

    if (item.noDates) return;
    const bar = {
      id: `${project.id}:${task.id}`,
      taskId: task.id,
      projectId: project.id,
      projectName: project.name,
      clientName: project.clientName || '',
      workspaceId: (input.workspaces && input.workspaces[project.id]) || '',
      color: colorOf(project.id),
      name: task.name,
      test: task.eshaTest || '',
      status: task.status,
      duration: task.duration || '',
      plannedQuantity: planned,
      start,
      end,
      from: start < from ? from : start,
      to: end > to ? to : end,
      clippedStart: start < from,
      clippedEnd: end > to,
      owners: uniqueOwners.map((person) => person.name),
      warnings: warningIds,
      lane: 0
    };
    bars.push(bar);
    item.bar = bar;
    item.owners = uniqueOwners;
    if (!uniqueOwners.length) noOwnerBars.push(bar);
    uniqueOwners.forEach((person) => {
      if (!barsByPerson.has(person.key)) barsByPerson.set(person.key, []);
      barsByPerson.get(person.key).push({ bar, item });
    });
  });

  // -- clashes and leave, per person ----------------------------------------------------------------
  const clashDatesByPerson = new Map();
  barsByPerson.forEach((entries, key) => {
    const person = people.get(key);
    for (let i = 0; i < entries.length; i += 1) {
      for (let j = i + 1; j < entries.length; j += 1) {
        const a = entries[i].bar;
        const b = entries[j].bar;
        const clashFrom = a.start > b.start ? a.start : b.start;
        const clashTo = a.end < b.end ? a.end : b.end;
        if (clashFrom > clashTo) continue;
        const id = warn('double_booked', `${person.name} is on two Onsite tasks on ${rangeLabel(clashFrom, clashTo)}: "${a.name}" (${a.projectName}) and "${b.name}" (${b.projectName}).`, { taskIds: [a.taskId, b.taskId].sort(), personKey: key, personName: person.name, dates: [clashFrom, clashTo], projectName: a.projectName });
        a.warnings.push(id);
        b.warnings.push(id);
        if (!clashDatesByPerson.has(key)) clashDatesByPerson.set(key, new Set());
        eachDay(clashFrom, clashTo).forEach((date) => clashDatesByPerson.get(key).add(date));
      }
    }
    entries.forEach(({ bar }) => {
      (leaveByKey.get(key) || []).forEach((leave) => {
        const overlap = eachDay(bar.start, bar.end).filter((date) => leave.dates.has(date));
        if (!overlap.length) return;
        const id = warn('leave_overlap', `${person.name} is on ${leave.state} leave (${leave.type || 'leave'}) on ${rangeLabel(overlap[0], overlap[overlap.length - 1])}, during "${bar.name}" (${bar.projectName}).`, { taskIds: [bar.taskId], personKey: key, personName: person.name, dates: [overlap[0], overlap[overlap.length - 1]], projectName: bar.projectName, state: leave.state });
        bar.warnings.push(id);
      });
    });
  });
  bars.forEach((bar) => { bar.warnings = [...new Set(bar.warnings)]; });

  // -- rows ---------------------------------------------------------------------------------------
  const rowFor = (person, kind) => ({
    key: person.key,
    name: person.name,
    email: person.email,
    kind,
    isMe: Boolean(me && person.user && person.user.id === me.id),
    roles: person.user ? person.user.roles : [],
    bars: [],
    leave: [],
    clashDates: [],
    holidayDates: []
  });
  const wanted = new Map();
  if (me) {
    const mine = personFor({ email: me.email, name: me.name });
    if (mine) wanted.set(mine.key, 'user');
  }
  users.filter((user) => user.active !== false && (user.roles || []).includes('engineer')).forEach((user) => {
    const person = personFor({ email: user.email, name: user.name });
    if (person && !wanted.has(person.key)) wanted.set(person.key, 'user');
  });
  barsByPerson.forEach((_entries, key) => {
    if (!wanted.has(key)) wanted.set(key, people.get(key).user ? 'user' : 'zoho');
  });

  const rows = [...wanted.entries()].map(([key, kind]) => rowFor(people.get(key), kind));
  rows.sort((a, b) => {
    const rank = (row) => (row.isMe ? 0 : row.kind === 'user' && row.roles.includes('engineer') ? 1 : row.kind === 'user' ? 2 : 3);
    return rank(a) - rank(b) || a.name.localeCompare(b.name);
  });
  if (noOwnerBars.length) rows.push({ key: 'none', name: 'No owner', email: '', kind: 'none', isMe: false, roles: [], bars: [], leave: [], clashDates: [], holidayDates: [] });

  function lanes(list) {
    const ends = [];
    [...list].sort((a, b) => a.from.localeCompare(b.from) || a.to.localeCompare(b.to)).forEach((bar) => {
      let lane = ends.findIndex((end) => end < bar.from);
      if (lane < 0) { lane = ends.length; ends.push(bar.to); } else ends[lane] = bar.to;
      bar.lane = lane;
    });
    return ends.length;
  }

  rows.forEach((row) => {
    const entries = row.key === 'none' ? noOwnerBars.map((bar) => ({ bar })) : barsByPerson.get(row.key) || [];
    // each row gets its own copy of the bar, so the lane is the lane in this row
    row.bars = entries.map(({ bar }) => ({ ...bar }));
    row.laneCount = Math.max(1, lanes(row.bars));
    const person = people.get(row.key);
    row.leave = (leaveByKey.get(row.key) || []).map((leave) => ({ id: leave.id, state: leave.state, type: leave.type, from: leave.from, to: leave.to, dates: [...leave.dates].filter((date) => date >= from && date <= to).sort() })).filter((leave) => leave.dates.length);
    row.clashDates = [...(clashDatesByPerson.get(row.key) || [])].filter((date) => date >= from && date <= to).sort();
    row.holidayDates = days.filter((day) => day.working && weightFor(day.date, person ? person.location : '') === 0).map((day) => day.date);
    if (person) row.location = person.location;
  });

  bars.forEach((bar) => {
    if (!projectsOut.has(bar.projectId)) projectsOut.set(bar.projectId, { id: bar.projectId, name: bar.projectName, clientName: bar.clientName, color: bar.color, workspaceId: bar.workspaceId });
  });
  const projectInfo = new Map((input.projects || []).map((project) => [project.id, project]));

  return {
    window: { from, to },
    days,
    rows,
    warnings,
    projects: [...projectsOut.values()].map((project) => ({ ...project, zoho: (({ projectNumber, ownerName, stage, workOrderNo }) => ({ projectNumber, ownerName, stage, workOrderNo }))(projectInfo.get(project.id) || {}) })),
    settings,
    tests: ESHA_TESTS.map((test) => ({ id: test.id, name: test.name }))
  };
}

module.exports = { MAX_WINDOW_DAYS, WEEKDAYS, DEFAULT_RATES, isIso, addDays, eachDay, weekdayOf, parseZohoDate, monthsOf, monthRange, monthKey, clampWindow, defaultSettings, mergeSettings, buildSchedule };
