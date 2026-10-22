// Outlook calendar sync (Session 6H.2): every Onsite ESHA task of the 6F schedule that has dates and an owner
// becomes one all-day (date-range) event in each owner's Outlook calendar:
//
//   subject   Site: <project> – <test>          location  the site (else the client)
//   body      the project's link in this app    showAs    workingElsewhere
//
// The app never touches an event it did not create. Its events carry a single-value extended property
// (ElectroReportsTaskId = <task id>|<owner e-mail>) under the app's own GUID, and a run reconciles against
// what Outlook really holds for that property: it lists only events with the property, and before it deletes
// anything it reads the event back and checks the property again. An event without the property is never
// listed, updated or deleted. A task that is closed or removed, has no dates, or changed owner has its event
// deleted; changed dates, project or link update it. If a mailbox cannot be read (not in the Exchange-scoped
// group: 403) that owner is skipped with a warning and nothing is deleted for them.
//
// Off by default, and dry-run by default: a dry run (also the admin's "Preview") plans and writes nothing.

const crypto = require('crypto');
const { TEST_LIBRARY } = require('./reportModel');
const { parseZohoDate, addDays } = require('./schedule');
const { stable } = require('./entryCredit');

const PROPERTY_GUID = '6b0f8c9e-2f5d-4c3e-9a41-7d1e5a3b8c20';
const PROPERTY_NAME = 'ElectroReportsTaskId';
const PROPERTY_ID = `String {${PROPERTY_GUID}} Name ${PROPERTY_NAME}`;
const SETTINGS_KEY = 'graphCalendar';
const LAST_KEY = 'graphCalendarLast';
const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const lower = (value) => text(value).toLowerCase();

function createCalendarSync({ store, graph, config, log = console, now = () => new Date() }) {
  let running = null;

  function getSettings() {
    const saved = store.getScheduleSetting(SETTINGS_KEY) || {};
    return { enabled: saved.enabled === true, dryRun: saved.dryRun !== false };
  }

  function saveSettings(patch, by = '') {
    const next = { ...getSettings(), ...(typeof patch.enabled === 'boolean' ? { enabled: patch.enabled } : {}), ...(typeof patch.dryRun === 'boolean' ? { dryRun: patch.dryRun } : {}) };
    store.setScheduleSetting(SETTINGS_KEY, next, by);
    return next;
  }

  // ---- what should exist ------------------------------------------------------------------------
  function linkFor(id) {
    return config.appBaseUrl ? `${config.appBaseUrl}/?project=${encodeURIComponent(id)}` : '';
  }

  function desiredEvents() {
    const cache = store.getScheduleCache('tasks');
    const projects = (cache && cache.data && cache.data.projects) || [];
    const hidden = new Set(((config.hiddenTaskStatuses) || ['Closed']).map(lower));
    const workspaces = store.listProjects();
    const desired = new Map();
    const warnings = [];
    projects.forEach((project) => {
      const workspace = workspaces.find((entry) => text(entry.zohoProjectId) === text(project.id));
      (project.tasks || []).forEach((task) => {
        if (task.closed || hidden.has(lower(task.status))) return;
        const start = parseZohoDate(task.startDate);
        const end = parseZohoDate(task.dueDate);
        const from = start || end;
        const to = end || start;
        if (!from) return;
        if (to < from) {
          warnings.push(`${project.name}: "${text(task.name)}" ends before it starts, so it is not in anyone's calendar.`);
          return;
        }
        const test = TEST_LIBRARY.find((entry) => entry.id === task.testId);
        const testName = test ? test.label : text(task.name).replace(/^\[ESHA\]\s*/i, '');
        const link = linkFor(workspace ? workspace.id : project.id);
        (task.owners || []).forEach((owner) => {
          const email = lower(owner.email);
          if (!email) {
            warnings.push(`${project.name}: "${text(task.name)}" has an owner without an e-mail address, so that owner is skipped.`);
            return;
          }
          const key = `${task.id}|${email}`;
          const payload = {
            subject: `Site: ${text(project.name)} – ${testName}`,
            isAllDay: true,
            showAs: 'workingElsewhere',
            start: { dateTime: `${from}T00:00:00`, timeZone: config.timezone },
            end: { dateTime: `${addDays(to, 1)}T00:00:00`, timeZone: config.timezone },
            location: { displayName: text((workspace && workspace.siteLocation) || project.clientName) },
            body: { contentType: 'text', content: link ? `Open the project in ElectroReports: ${link}` : '' }
          };
          desired.set(key, { key, taskId: text(task.id), email, ownerName: text(owner.fullName || owner.name), project: text(project.name), test: testName, from, to, payload, hash: crypto.createHash('sha256').update(stable(payload)).digest('hex') });
        });
      });
    });
    return { desired, warnings };
  }

  // ---- what Outlook holds -----------------------------------------------------------------------
  // The app's events in one mailbox: [{ id, key }]. Only events carrying the app's property are returned.
  async function listOwned(email) {
    const { items } = await graph.getAll(`/users/${encodeURIComponent(email)}/events`, {
      params: { $filter: `singleValueExtendedProperties/Any(ep: ep/id eq '${PROPERTY_ID}' and ep/value ne null)`, $select: 'id', $expand: `singleValueExtendedProperties($filter=id eq '${PROPERTY_ID}')` }
    });
    return items
      .map((event) => ({ id: event.id, key: ((event.singleValueExtendedProperties || []).find((prop) => prop.id === PROPERTY_ID) || {}).value }))
      .filter((event) => event.id && event.key);
  }

  // ---- the plan -----------------------------------------------------------------------------------
  async function plan() {
    const { desired, warnings } = desiredEvents();
    const rows = new Map(store.listCalendarSync().map((row) => [row.key, row]));
    const owners = new Set([...desired.values()].map((entry) => entry.email));
    rows.forEach((row) => owners.add(row.ownerEmail));
    const items = [];
    const unavailable = new Set();
    const remote = new Map();
    for (const email of owners) {
      try {
        remote.set(email, await listOwned(email));
      } catch (error) {
        unavailable.add(email);
        warnings.push(`A calendar could not be read (${error.status || 'error'}${error.code ? `: ${error.code}` : ''}), so nothing is changed there: ${email}.`);
      }
    }
    desired.forEach((entry) => {
      if (unavailable.has(entry.email)) {
        items.push({ action: 'skip', key: entry.key, owner: entry.email, ownerName: entry.ownerName, project: entry.project, test: entry.test, from: entry.from, to: entry.to });
        return;
      }
      const held = (remote.get(entry.email) || []).filter((event) => event.key === entry.key);
      const base = { key: entry.key, owner: entry.email, ownerName: entry.ownerName, project: entry.project, test: entry.test, from: entry.from, to: entry.to, entry };
      if (!held.length) items.push({ action: 'create', ...base });
      else {
        const row = rows.get(entry.key);
        const same = row && row.eventId === held[0].id && row.hash === entry.hash;
        items.push({ action: same ? 'unchanged' : 'update', eventId: held[0].id, ...base });
        held.slice(1).forEach((extra) => items.push({ action: 'delete', key: entry.key, owner: entry.email, ownerName: entry.ownerName, project: entry.project, test: entry.test, eventId: extra.id, reason: 'duplicate' }));
      }
    });
    remote.forEach((events, email) => {
      events.forEach((event) => {
        if (desired.has(event.key) && desired.get(event.key).email === email) return;
        const row = rows.get(event.key);
        items.push({ action: 'delete', key: event.key, owner: email, ownerName: '', project: row ? '' : '', test: '', eventId: event.id, reason: 'no longer planned' });
      });
    });
    // rows whose event is gone from Outlook and that are no longer planned only need their row removed
    rows.forEach((row) => {
      if (desired.has(row.key) || unavailable.has(row.ownerEmail)) return;
      if (!items.some((item) => item.action === 'delete' && item.key === row.key)) items.push({ action: 'forget', key: row.key, owner: row.ownerEmail });
    });
    return { items, warnings };
  }

  // ---- doing it -----------------------------------------------------------------------------------
  async function apply(item) {
    const base = `/users/${encodeURIComponent(item.owner)}/events`;
    if (item.action === 'create') {
      const { data } = await graph.request('POST', base, { body: { ...item.entry.payload, singleValueExtendedProperties: [{ id: PROPERTY_ID, value: item.key }] } });
      store.saveCalendarSync(item.key, { taskId: item.entry.taskId, ownerEmail: item.owner, eventId: data.id, hash: item.entry.hash, state: 'done', error: '', syncedAt: now().toISOString() });
    } else if (item.action === 'update') {
      await graph.request('PATCH', `${base}/${encodeURIComponent(item.eventId)}`, { body: item.entry.payload });
      store.saveCalendarSync(item.key, { taskId: item.entry.taskId, ownerEmail: item.owner, eventId: item.eventId, hash: item.entry.hash, state: 'done', error: '', syncedAt: now().toISOString() });
    } else if (item.action === 'delete') {
      // read it back: only an event that still carries the app's property is deleted
      const { data } = await graph.request('GET', `${base}/${encodeURIComponent(item.eventId)}`, { params: { $expand: `singleValueExtendedProperties($filter=id eq '${PROPERTY_ID}')` } });
      const tagged = (data.singleValueExtendedProperties || []).some((prop) => prop.id === PROPERTY_ID && prop.value === item.key);
      if (!tagged) throw Object.assign(new Error('not created by this app'), { refused: true });
      await graph.request('DELETE', `${base}/${encodeURIComponent(item.eventId)}`);
      store.deleteCalendarSync(item.key);
    } else if (item.action === 'forget') {
      store.deleteCalendarSync(item.key);
    }
  }

  const strip = ({ entry, ...item }) => item;

  // One run. dryRun: plan only. Returns the summary that is also kept as the last result.
  function run({ dryRun = getSettings().dryRun } = {}) {
    if (running) return running;
    running = (async () => {
      const { items, warnings } = await plan();
      const counts = { create: 0, update: 0, delete: 0, unchanged: 0, skip: 0, forget: 0, failed: 0 };
      const applied = [];
      for (const item of items) {
        if (dryRun || ['unchanged', 'skip'].includes(item.action)) {
          counts[item.action] += 1;
          continue;
        }
        try {
          await apply(item);
          counts[item.action] += 1;
        } catch (error) {
          counts.failed += 1;
          applied.push({ ...strip(item), error: error.refused ? error.message : `${error.status || ''} ${error.code || 'failed'}`.trim() });
          if (item.entry) store.saveCalendarSync(item.key, { taskId: item.entry.taskId, ownerEmail: item.owner, state: 'failed', error: error.refused ? error.message : `${error.status || ''} ${error.code || 'failed'}`.trim() });
        }
      }
      const summary = { at: now().toISOString(), dryRun, counts, items: items.filter((item) => item.action !== 'unchanged').map(strip), failures: applied, warnings };
      store.setScheduleSetting(LAST_KEY, summary);
      return summary;
    })().finally(() => { running = null; });
    return running;
  }

  // The hook after a Zoho task refresh: only when switched on.
  async function onTasksRefreshed() {
    if (!getSettings().enabled) return null;
    try {
      return await run();
    } catch (error) {
      log.error(`Calendar sync failed: ${error.status || ''} ${error.code || error.message}`.trim());
      return null;
    }
  }

  const status = () => ({ settings: getSettings(), last: store.getScheduleSetting(LAST_KEY), appBaseUrl: Boolean(config.appBaseUrl), timezone: config.timezone, owned: store.listCalendarSync().length });

  return { getSettings, saveSettings, desiredEvents, plan, run, onTasksRefreshed, status, PROPERTY_ID };
}

module.exports = { createCalendarSync, PROPERTY_ID, PROPERTY_GUID };
