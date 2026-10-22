// Site Work Log in Zoho People (Session 6G): ONE People record per engineer + work date + test, written when
// a daily report is APPROVED and at no other time (not on submit, not on return, not on re-issue).
//
//   Record Key  = projectId|workDate|engineerEmail|testKey, so every revision of a day addresses the same record
//   Completed   = distinct units of that test credited to that engineer (first enterer, src/entryCredit.js; a row
//                 from before 6G is credited to the report's creator)
//   Planned     = the project's planned quantity for that test
//   Returned    = "return" transitions across the revision chain
//   Data Gaps   = CAP data-gap lines of that test in the report
//   On-time     = first submit on or before the next working day after the work date (6F working days and
//                 holidays, Asia/Kolkata)
//   First-time  = no return and revision 0
//
// Approving a later revision updates the same record. An engineer whose quantity for the day and test falls to
// 0 in that revision keeps the record: Completed Qty becomes 0, the status stays Approved, and the audit log
// gets an entry. Finals, sample and test reports write nothing.
//
// People is never allowed to block the review: a write that fails is kept in people_sync as failed (or pending
// while the form or one of its fields does not exist yet) and goes out with the admin's Retry or Backfill.

const crypto = require('crypto');
const { getReportKind, getReviewStatus, getCollectionDate } = require('./reportCollectionModel');
const { TEST_LIBRARY } = require('./reportModel');
const { buildFinalReportData } = require('./finalReport/data');
const { transitionAction } = require('./review');
const { listUnits, stable } = require('./entryCredit');
const { toPeopleDate, lookupAttribute, WORKLOG_LABELS } = require('./zohoPeople');

const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const ISO_DAY = /^\d{4}-\d{2}-\d{2}$/;
const DAY_MS = 24 * 60 * 60 * 1000;
const IST_OFFSET_MS = 5.5 * 60 * 60 * 1000;
const EMPLOYEE_CACHE_MS = 10 * 60 * 1000;
const SAMPLE_NAME = /^\s*(sample|test|demo|dummy)\b/i;
const norm = (value) => text(value).toLowerCase().replace(/[^a-z0-9]/g, '');
const addDays = (iso, count) => new Date(Date.parse(`${iso}T00:00:00Z`) + count * DAY_MS).toISOString().slice(0, 10);
const istDate = (iso) => new Date(Date.parse(iso) + IST_OFFSET_MS).toISOString().slice(0, 10);
const yesNo = (value) => (value ? 'Yes' : 'No');
const tests = TEST_LIBRARY.filter((test) => !test.retired);

// The names of the Test dropdown of the People form (it has no "Test" / "Measurement" suffixes). The app's
// own label is tried first, then this one.
const FORM_TEST_NAME = {
  soilResistivity: 'Soil Resistivity', electrodeResistance: 'Earth Electrode Resistance', continuityTest: 'Earth Continuity & Bonding', loopImpedanceTest: 'Loop Impedance',
  prospectiveFaultCurrent: 'Prospective Fault Current', riserIntegrityTest: 'Riser / Grid Integrity', groundResistance: 'Ground (Grid) Resistance', stepTouchVoltage: 'Step & Touch Voltage',
  towerFootingResistance: 'Tower Footing Resistance'
};

// A dropdown or radio takes only its own options. Returns the form's spelling of the first candidate that
// matches one (exactly, then ignoring case and punctuation), or throws naming the options.
function formOption(field, candidates) {
  if (!field.options || !field.options.length) return candidates[0];
  for (const match of [(option, wanted) => option === wanted, (option, wanted) => norm(option) === norm(wanted)]) {
    for (const wanted of candidates) {
      const hit = field.options.find((option) => match(option, wanted));
      if (hit) return hit;
    }
  }
  throw new Error(`"${candidates[0]}" is not an option of the ${field.label} field (options: ${field.options.join(', ')}).`);
}

// A sample or test report never reaches People: the project says so (`isSample`), or its project number
// or client is named like one.
function isSampleReport(report, project = null) {
  const info = (report && report.project) || {};
  return Boolean((project && project.isSample) || report.isSample || info.isSample || SAMPLE_NAME.test(text(info.projectNo)) || SAMPLE_NAME.test(text(info.clientName)) || (project && (SAMPLE_NAME.test(text(project.projectNo)) || SAMPLE_NAME.test(text(project.clientName)))));
}

function createPeopleSync({ store, people, config, getWorkingDays = () => [1, 2, 3, 4, 5, 6] } = {}) {
  const settings = config.zoho.workLog;
  let chain = Promise.resolve();
  let employeeCache = { at: 0, list: null };

  const enqueue = (job) => {
    const run = chain.then(job, job);
    chain = run.catch(() => {});
    return run;
  };

  // The report and the revisions before it, newest first.
  function revisionsOf(report) {
    const list = [report];
    for (let guard = 0; guard < 100; guard += 1) {
      const parent = list[list.length - 1].revisionOf ? store.getReport(list[list.length - 1].revisionOf) : null;
      if (!parent) break;
      list.push(parent);
    }
    return list;
  }
  const rootIdOf = (report) => revisionsOf(report).pop().id;

  // ---- on-time ---------------------------------------------------------------------------------
  // Holidays come from the schedule cache (6F). A holiday with locations only counts for people there.
  function isHoliday(date, location) {
    const calendar = store.getScheduleCache(`cal:${date.slice(0, 7)}`);
    const holidays = (calendar && calendar.data && calendar.data.holidays) || [];
    return holidays.some((holiday) => holiday.days.some((day) => day.date === date && day.weight >= 1) && (!holiday.locations.length || (location && holiday.locations.some((entry) => text(entry).toLowerCase() === text(location).toLowerCase()))));
  }

  // The next working day after `workDate`: the last day the report may be submitted on time.
  function submitDeadline(workDate, location = '') {
    const working = getWorkingDays();
    let day = workDate;
    for (let guard = 0; guard < 30; guard += 1) {
      day = addDays(day, 1);
      if (working.includes(new Date(`${day}T00:00:00Z`).getUTCDay()) && !isHoliday(day, location)) return day;
    }
    return day;
  }

  // ---- the entries of an approved day -------------------------------------------------------------
  function dataGapLines(report) {
    try {
      const data = buildFinalReportData(report);
      const counts = {};
      data.cap.filter((line) => line.kind === 'gap').forEach((line) => { counts[line.testId] = (counts[line.testId] || 0) + 1; });
      return counts;
    } catch (_error) {
      return {};
    }
  }

  // One entry per engineer and test with work in `report`.
  function buildEntries(report) {
    const creator = report.createdBy ? store.getUser(report.createdBy) : null;
    const info = report.project || {};
    const project = report.projectId ? store.getProject(report.projectId) : null;
    const planned = (project && project.plannedQuantities) || info.plannedQuantities || {};
    const revisions = revisionsOf(report);
    const transitions = revisions.flatMap((entry) => (store.listTransitions(entry.id) || []).map((row) => ({ ...row, action: transitionAction(row) })));
    const returned = transitions.filter((row) => row.action === 'return').length;
    const firstSubmit = transitions.filter((row) => row.action === 'submit').map((row) => row.at).sort()[0] || '';
    const workDate = getCollectionDate(report);
    const gaps = dataGapLines(report);
    const revision = Number.isInteger(report.revision) ? report.revision : 0;
    const groups = new Map();
    listUnits(report).forEach(({ testId, unit }) => {
      if (!tests.some((test) => test.id === testId)) return;
      const stamp = unit.enteredBy && (unit.enteredBy.email || unit.enteredBy.userId) ? unit.enteredBy : null;
      const credited = stamp || (creator ? { userId: creator.id, name: creator.name, email: creator.email } : null);
      const email = credited ? text(credited.email).toLowerCase() : '';
      const key = `${testId}|${email}`;
      if (!groups.has(key)) groups.set(key, { testId, email, name: credited ? credited.name || '' : '', userId: credited ? credited.userId || '' : '', completed: 0, creditedByCreator: 0 });
      const group = groups.get(key);
      group.completed += 1;
      if (!stamp) group.creditedByCreator += 1;
    });
    return [...groups.values()].map((group) => {
      const test = tests.find((entry) => entry.id === group.testId);
      const location = ((employeeCache.list || []).find((employee) => employee.email === group.email) || {}).location || '';
      const plannedQty = Number.parseFloat(planned[group.testId]);
      return {
        key: `${report.projectId}|${workDate}|${group.email}|${group.testId}`,
        rootId: rootIdOf(report),
        reportId: report.id,
        projectId: text(report.projectId),
        project: text((project && project.zohoProjectName) || info.projectNo),
        client: text(info.clientName),
        workDate,
        engineer: { userId: group.userId, name: group.name, email: group.email },
        testId: group.testId,
        testLabel: test.label,
        unit: test.unitLabel,
        planned: Number.isFinite(plannedQty) ? plannedQty : null,
        completed: group.completed,
        creditedByCreator: group.creditedByCreator,
        returned,
        dataGaps: gaps[group.testId] || 0,
        onTime: Boolean(firstSubmit) && istDate(firstSubmit) <= submitDeadline(workDate, location),
        firstTimeApproved: returned === 0 && revision === 0,
        revision: `R${revision}`
      };
    });
  }

  // ---- writing ---------------------------------------------------------------------------------
  // What one run needs from People, read once and only if something is to be written.
  function createRun() {
    let formOk;
    let discovered;
    const fieldsOf = async () => {
      if (!discovered) {
        const found = await people.getFormFields(settings.form);
        const map = {};
        const missing = [];
        Object.entries(WORKLOG_LABELS).forEach(([key, label]) => {
          const override = settings.fields[key];
          const field = override ? found.find((entry) => entry.apiName === override) || { apiName: override, label } : found.find((entry) => norm(entry.label) === norm(label));
          if (field) map[key] = field; else missing.push(label);
        });
        discovered = { map, missing };
      }
      return discovered;
    };
    return {
      async formExists() {
        if (formOk === undefined) formOk = await people.hasForm(settings.form);
        return formOk;
      },
      fieldsOf,
      async engineerValue(engineer) {
        if (!engineer || !engineer.email) throw new Error('No engineer is recorded for this entry.');
        if (!employeeCache.list || Date.now() - employeeCache.at > EMPLOYEE_CACHE_MS) {
          employeeCache = { at: Date.now(), list: await people.listEmployees() };
        }
        const employee = employeeCache.list.find((entry) => entry.email === engineer.email);
        if (!employee) throw new Error(`No Zoho People employee has the e-mail ${engineer.email}.`);
        // The lookup's own options say what to send: the option that shows this employee's id, and its Id.
        const meta = (await fieldsOf()).map.engineer;
        const option = (meta.lookupOptions || []).find((entry) => entry.value.startsWith(employee.employeeId) && employee.employeeId);
        if (option) return option.id;
        const attribute = lookupAttribute(meta);
        if (!employee[attribute]) throw new Error(`The People employee ${employee.email} has no ${attribute}, which the Engineer lookup expects.`);
        return employee[attribute];
      }
    };
  }

  // The exact fields sent to People for an entry, by the API names found on the form.
  async function peopleData(entry, run) {
    const { map } = await run.fieldsOf();
    const values = {
      recordKey: entry.key,
      engineer: await run.engineerValue(entry.engineer),
      workDate: toPeopleDate(entry.workDate),
      project: entry.project,
      client: entry.client,
      test: formOption(map.test, [entry.testLabel, FORM_TEST_NAME[entry.testId] || entry.testLabel]),
      unit: formOption(map.unit, [entry.unit]),
      plannedQty: entry.planned,
      completedQty: entry.completed,
      reportStatus: formOption(map.reportStatus, ['Approved']),
      returnedCount: entry.returned,
      dataGaps: entry.dataGaps,
      onTime: formOption(map.onTime, [yesNo(entry.onTime)]),
      firstTimeApproved: formOption(map.firstTimeApproved, [yesNo(entry.firstTimeApproved)]),
      reportId: entry.reportId
    };
    const data = {};
    Object.entries(values).forEach(([key, value]) => {
      if (value !== null && value !== undefined) data[map[key].apiName] = value;
    });
    return data;
  }

  async function writeEntry(entry, run) {
    const row = store.getPeopleSync(entry.key);
    const hash = crypto.createHash('sha256').update(stable(entry)).digest('hex');
    if (row && row.state === 'done' && row.hash === hash) return 'unchanged';
    const base = { rootId: entry.rootId, reportId: entry.reportId, testId: entry.testId, workDate: entry.workDate, status: 'Approved', hash, entry, attempts: (row ? row.attempts : 0) + 1 };
    const wait = (error) => { store.savePeopleSync(entry.key, { ...base, attempts: row ? row.attempts : 0, state: 'pending', error }); return 'blocked'; };
    try {
      if (!(await run.formExists())) return wait(`The People form "${settings.form}" does not exist yet. Create it, then run Retry or Backfill.`);
      const { missing } = await run.fieldsOf();
      if (missing.length) return wait(`The People form "${settings.form}" has no field labelled: ${missing.join(', ')}.`);
      const data = await peopleData(entry, run);
      let recordId = row && row.peopleRecordId;
      if (!recordId) recordId = await people.findRecord(settings.form, (await run.fieldsOf()).map.recordKey.apiName, entry.key);
      if (recordId) await people.updateRecord(settings.form, recordId, data);
      else recordId = await people.insertRecord(settings.form, data);
      store.savePeopleSync(entry.key, { ...base, state: 'done', peopleRecordId: String(recordId), error: '', syncedAt: new Date().toISOString() });
      return 'written';
    } catch (error) {
      store.savePeopleSync(entry.key, { ...base, state: 'failed', error: error.message });
      return 'failed';
    }
  }

  const blankTally = () => ({ written: 0, unchanged: 0, blocked: 0, failed: 0, zeroed: 0 });

  // Writes the entries of an approved report, then sets Completed Qty to 0 on the records of this day that the
  // report no longer has an entry for.
  async function writeApproved(report, run = createRun()) {
    const tally = blankTally();
    const entries = buildEntries(report);
    for (const entry of entries) tally[await writeEntry(entry, run)] += 1;
    const keys = new Set(entries.map((entry) => entry.key));
    for (const row of store.listPeopleSync({ rootId: rootIdOf(report), limit: 5000 })) {
      if (keys.has(row.key) || !row.entry || row.entry.completed === 0) continue;
      const zeroed = { ...row.entry, completed: 0, reportId: report.id, revision: `R${Number.isInteger(report.revision) ? report.revision : 0}`, rootId: row.rootId };
      const outcome = await writeEntry(zeroed, run);
      if (outcome === 'written') {
        tally.zeroed += 1;
        store.addAudit({ action: 'people.qty_zero', entityType: 'people_sync', entityId: row.key, detail: { engineer: row.entry.engineer.email, test: row.entry.testId, workDate: row.entry.workDate, was: row.entry.completed, reportId: report.id } });
      } else tally[outcome] += 1;
    }
    return tally;
  }

  function eligible(report) {
    if (!report || getReportKind(report) !== 'daily') return 'not-daily';
    const project = report.projectId ? store.getProject(report.projectId) : null;
    return isSampleReport(report, project) ? 'sample' : '';
  }

  // Called after a review step went through. Only an approval writes. Returns the promise of the write; the
  // caller does not wait for it.
  function afterTransition(action, report) {
    if (action !== 'approve') return Promise.resolve({ skipped: 'step' });
    const why = eligible(report);
    if (why) return Promise.resolve({ skipped: why });
    return enqueue(() => writeApproved(report));
  }

  // Approved daily reports with a work date in [from, to]: the latest approved revision of each day. Sample
  // and test reports are skipped; finals are not daily so never selected.
  function backfill({ from, to }) {
    if (!ISO_DAY.test(text(from)) || !ISO_DAY.test(text(to)) || from > to) {
      return Promise.reject(Object.assign(new Error('Give a from and a to date (YYYY-MM-DD), from not after to.'), { code: 'BAD_RANGE' }));
    }
    return enqueue(async () => {
      const summary = { from, to, reports: 0, skippedSample: 0, noEngineer: 0, ...blankTally() };
      const latest = new Map();
      store.listReports('').forEach((report) => {
        if (getReportKind(report) !== 'daily' || getReviewStatus(report) !== 'approved') return;
        const day = getCollectionDate(report);
        if (day < from || day > to) return;
        const root = rootIdOf(report);
        const held = latest.get(root);
        if (!held || (report.revision || 0) > (held.revision || 0)) latest.set(root, report);
      });
      const run = createRun();
      for (const report of latest.values()) {
        if (eligible(report) === 'sample') {
          summary.skippedSample += 1;
          continue;
        }
        summary.reports += 1;
        summary.noEngineer += buildEntries(report).filter((entry) => !entry.engineer.email).length;
        const tally = await writeApproved(report, run);
        Object.keys(tally).forEach((key) => { summary[key] += tally[key]; });
      }
      summary.at = new Date().toISOString();
      store.setScheduleSetting('peopleSyncLastBackfill', summary);
      return summary;
    });
  }

  // Sends again every entry that is failed or pending, as it was stored.
  function retry() {
    return enqueue(async () => {
      const tally = blankTally();
      const run = createRun();
      for (const row of store.listPeopleSync({ limit: 5000 })) {
        if (row.state === 'done' || !row.entry) continue;
        tally[await writeEntry(row.entry, run)] += 1;
      }
      return tally;
    });
  }

  // Totals per engineer for [from, to] from what was approved: completed against planned per test, and the
  // quality figures. Planned is the planned quantity of each project the engineer worked on, counted once per
  // project. On-time, first-time approved and returns count each engineer-day once (a day with several tests
  // is one report), data gaps add up over the tests.
  function totals({ from, to }) {
    if (!ISO_DAY.test(text(from)) || !ISO_DAY.test(text(to)) || from > to) {
      throw Object.assign(new Error('Give a from and a to date (YYYY-MM-DD), from not after to.'), { code: 'BAD_RANGE' });
    }
    const byEngineer = new Map();
    store.listPeopleSync({ limit: 100000 }).filter((row) => row.entry && row.workDate >= from && row.workDate <= to).forEach(({ entry }) => {
      const email = entry.engineer.email || '(no engineer)';
      if (!byEngineer.has(email)) byEngineer.set(email, { email, name: entry.engineer.name, tests: new Map(), days: new Map(), dataGaps: 0 });
      const who = byEngineer.get(email);
      if (!who.tests.has(entry.testId)) who.tests.set(entry.testId, { testId: entry.testId, label: entry.testLabel, unit: entry.unit, completed: 0, planned: 0, projects: new Set() });
      const test = who.tests.get(entry.testId);
      test.completed += entry.completed;
      if (!test.projects.has(entry.projectId)) {
        test.projects.add(entry.projectId);
        test.planned += entry.planned || 0;
      }
      who.dataGaps += entry.dataGaps;
      who.days.set(`${entry.projectId}|${entry.workDate}`, entry);
    });
    const pct = (part, whole) => (whole ? Math.round((part / whole) * 1000) / 10 : null);
    return {
      from,
      to,
      engineers: [...byEngineer.values()].sort((a, b) => a.name.localeCompare(b.name) || a.email.localeCompare(b.email)).map((who) => {
        const days = [...who.days.values()];
        return {
          email: who.email,
          name: who.name,
          tests: [...who.tests.values()].map(({ projects, ...test }) => test),
          engineerDays: days.length,
          onTimePct: pct(days.filter((day) => day.onTime).length, days.length),
          firstTimeApprovedPct: pct(days.filter((day) => day.firstTimeApproved).length, days.length),
          returns: days.reduce((sum, day) => sum + day.returned, 0),
          dataGaps: who.dataGaps
        };
      })
    };
  }

  async function status() {
    let formExists = null;
    let formError = '';
    try {
      formExists = await people.hasForm(settings.form);
    } catch (error) {
      formError = error.message;
    }
    return {
      form: settings.form,
      formExists,
      formError,
      counts: store.peopleSyncCounts(),
      problems: store.listPeopleSync({ limit: 200 }).filter((row) => row.state !== 'done').slice(0, 50).map((row) => ({ key: row.key, state: row.state, error: row.error, workDate: row.workDate, testId: row.testId, attempts: row.attempts })),
      lastBackfill: store.getScheduleSetting('peopleSyncLastBackfill'),
      mock: config.zoho.useMock ? people.mockRecords() : undefined
    };
  }

  // Writes ready-made entries (the live check uses this to prove the insert and the update path).
  const writeEntries = (entries) => enqueue(async () => {
    const run = createRun();
    const out = [];
    for (const entry of entries) out.push(await writeEntry(entry, run));
    return out;
  });

  return { afterTransition, backfill, retry, totals, status, writeEntries, buildEntries, peopleData: async (entry) => peopleData(entry, createRun()), rootIdOf, submitDeadline, idle: () => chain };
}

module.exports = { createPeopleSync, isSampleReport };
