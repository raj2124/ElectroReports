// The schedule's data (Session 6F): what is read from Zoho Projects and Zoho People, kept in the
// database cache with the time it was read, and put together into the screen's data on request.
//
//   'tasks'      the Onsite ESHA tasks of every D03 project
//   'people'     the Zoho People employee list (matches leave to people by e-mail)
//   'cal:YYYY-MM' leave and holidays of one month
//
// A failed read keeps the last good data and records the error, so the screen shows old data with
// its age and the reason. A manual refresh is limited to one a minute; a timer refreshes in the
// background. Everything toward Zoho is a read.

const defaultConfig = require('./config');
const { isD03Name, buildTaskPlan, nameOwners } = require('./zohoTasks');
const { defaultSettings, mergeSettings, buildSchedule, clampWindow, monthsOf, monthRange, addDays } = require('./schedule');

const SETTINGS_KEY = 'settings';
const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const message = (error) => (error && error.message ? String(error.message).slice(0, 200) : 'unknown error');

function createScheduleService({ store, zoho, people, config = defaultConfig, now = () => Date.now(), log = console, onTasksRefreshed = null }) {
  const inFlight = new Map();
  const stamp = () => new Date(now()).toISOString();
  const put = (key, data, error = '') => store.setScheduleCache(key, data, { error, at: stamp() });
  let lastManualAt = 0;

  // One read at a time per cache key; a second caller waits for the first.
  function once(key, work) {
    if (!inFlight.has(key)) inFlight.set(key, work().finally(() => inFlight.delete(key)));
    return inFlight.get(key);
  }

  function getSettings() {
    const saved = store.getScheduleSetting(SETTINGS_KEY);
    const base = defaultSettings(config);
    return saved ? mergeSettings(base, saved).settings || base : base;
  }

  function saveSettings(patch, by = '') {
    const merged = mergeSettings(getSettings(), patch);
    if (merged.error) return merged;
    store.setScheduleSetting(SETTINGS_KEY, merged.settings, by);
    return merged;
  }

  // ---- Reads from Zoho ------------------------------------------------------------------------
  async function readTasks() {
    const previous = store.getScheduleCache('tasks');
    const before = new Map(((previous && previous.data && previous.data.projects) || []).map((project) => [project.id, project]));
    const projects = (await zoho.getProjects()).filter((project) => isD03Name(project.name));
    let portalUsers = [];
    try {
      portalUsers = await zoho.getPortalUsers();
    } catch (_error) {
      portalUsers = [];
    }
    const failed = [];
    const results = new Array(projects.length);
    let next = 0;
    const worker = async () => {
      while (next < projects.length) {
        const index = next;
        next += 1;
        const project = projects[index];
        try {
          const plan = nameOwners(buildTaskPlan(await zoho.getProjectTasks(project.id), { d03: true }), portalUsers);
          results[index] = {
            id: text(project.id),
            name: text(project.name),
            projectNumber: text(project.projectNumber),
            clientName: text(project.clientName),
            ownerName: text(project.ownerName),
            stage: text(project.stage),
            workOrderNo: text(project.workOrderNo),
            tasks: plan.tasks
              .filter((task) => task.esha && task.onsite)
              .map((task) => ({ id: task.id, name: task.name, owners: task.owners, startDate: task.startDate, dueDate: task.dueDate, status: task.status, closed: task.closed, duration: task.duration, eshaTest: task.eshaTest, testId: task.testId, plannedRaw: task.plannedRaw }))
          };
        } catch (error) {
          failed.push(`${project.name}: ${message(error)}`);
          const old = before.get(text(project.id));
          if (old) results[index] = { ...old, stale: true };
        }
      }
    };
    await Promise.all([worker(), worker(), worker()]);
    return { data: { projects: results.filter(Boolean) }, failed };
  }

  function refreshTasks() {
    return once('tasks', async () => {
      try {
        const { data, failed } = await readTasks();
        put('tasks', data);
        if (failed.length) put('tasks', undefined, `${failed.length} project${failed.length === 1 ? '' : 's'} could not be read (older data kept): ${failed.slice(0, 3).join('; ')}`);
        // the Outlook calendar sync follows the tasks (a no-op while it is switched off)
        if (typeof onTasksRefreshed === 'function') Promise.resolve().then(onTasksRefreshed).catch((error) => log.error(`Schedule: calendar hook failed: ${message(error)}`));
      } catch (error) {
        log.error(`Schedule: reading Zoho tasks failed: ${message(error)}`);
        put('tasks', undefined, message(error));
      }
    });
  }

  function refreshPeople() {
    return once('people', async () => {
      try {
        put('people', { employees: await people.listEmployees() });
      } catch (error) {
        log.error(`Schedule: reading Zoho People employees failed: ${message(error)}`);
        put('people', undefined, message(error));
      }
    });
  }

  function refreshMonth(month) {
    return once(`cal:${month}`, async () => {
      const { from, to } = monthRange(month);
      try {
        const [leave, holidays] = await Promise.all([people.listLeave(from, to), people.listHolidays(from, to)]);
        put(`cal:${month}`, { leave, holidays });
      } catch (error) {
        log.error(`Schedule: reading Zoho People leave for ${month} failed: ${message(error)}`);
        put(`cal:${month}`, undefined, message(error));
      }
    });
  }

  // ---- The screen's data ------------------------------------------------------------------------
  async function snapshot({ from, to, user = null } = {}) {
    const window = clampWindow(from, to);
    const months = monthsOf(window.from, window.to);
    const missing = [];
    if (!store.getScheduleCache('tasks') || !store.getScheduleCache('tasks').data) missing.push(refreshTasks());
    if (!store.getScheduleCache('people') || !store.getScheduleCache('people').data) missing.push(refreshPeople());
    months.filter((month) => !(store.getScheduleCache(`cal:${month}`) || {}).data).forEach((month) => missing.push(refreshMonth(month)));
    await Promise.all(missing);

    const tasks = store.getScheduleCache('tasks');
    const peopleCache = store.getScheduleCache('people');
    const calendars = months.map((month) => ({ month, cache: store.getScheduleCache(`cal:${month}`) }));
    const workspaces = {};
    store.listProjects().forEach((project) => { if (project.zohoProjectId) workspaces[String(project.zohoProjectId)] = project.id; });

    const built = buildSchedule({
      from: window.from,
      to: window.to,
      projects: (tasks && tasks.data && tasks.data.projects) || [],
      employees: (peopleCache && peopleCache.data && peopleCache.data.employees) || [],
      leave: calendars.flatMap(({ cache }) => (cache && cache.data && cache.data.leave) || []),
      holidays: calendars.flatMap(({ cache }) => (cache && cache.data && cache.data.holidays) || []),
      users: store.listUsers(),
      me: user,
      settings: getSettings(),
      hiddenStatuses: config.schedule.hiddenTaskStatuses,
      workspaces
    });

    const age = (cache) => (cache && cache.fetchedAt ? Math.max(0, Math.round((now() - Date.parse(cache.fetchedAt)) / 1000)) : null);
    const leaveErrors = [...(peopleCache && peopleCache.error ? [`Employee list: ${peopleCache.error}`] : []), ...calendars.filter(({ cache }) => cache && cache.error).map(({ month, cache }) => `${month}: ${cache.error}`)];
    const leaveLoaded = calendars.every(({ cache }) => cache && cache.data) && Boolean(peopleCache && peopleCache.data);
    return {
      ...built,
      serverTime: new Date(now()).toISOString(),
      zoho: {
        mock: Boolean(config.zoho.useMock),
        tasks: { fetchedAt: (tasks && tasks.fetchedAt) || '', ageSeconds: age(tasks), error: (tasks && tasks.error) || '' },
        leave: { fetchedAt: calendars.map(({ cache }) => (cache && cache.fetchedAt) || '').sort()[0] || '', ageSeconds: age(calendars[0] && calendars[0].cache), error: leaveErrors.join(' | '), loaded: leaveLoaded }
      }
    };
  }

  // Manual refresh: tasks, the employee list and the months of the window. At most once a minute.
  async function refresh({ from, to, user = null } = {}) {
    const window = clampWindow(from, to);
    const wait = config.schedule.manualRefreshSeconds * 1000 - (now() - lastManualAt);
    if (lastManualAt && wait > 0) {
      return { throttled: true, retryInSeconds: Math.ceil(wait / 1000), ...(await snapshot({ ...window, user })) };
    }
    lastManualAt = now();
    await Promise.all([refreshTasks(), refreshPeople(), ...monthsOf(window.from, window.to).map(refreshMonth)]);
    return { throttled: false, retryInSeconds: 0, ...(await snapshot({ ...window, user })) };
  }

  // Background refresh: tasks, people and the months around today.
  async function refreshAround(today = new Date(now()).toISOString().slice(0, 10)) {
    const months = monthsOf(addDays(today, -31), addDays(today, 100));
    await Promise.all([refreshTasks(), refreshPeople(), ...months.map(refreshMonth)]);
  }

  function start() {
    const minutes = config.schedule.refreshMinutes;
    if (!(minutes > 0)) return { stop() {} };
    const run = () => refreshAround().catch((error) => log.error(`Schedule refresh failed: ${message(error)}`));
    const first = setTimeout(() => {
      const cached = store.getScheduleCache('tasks');
      const fresh = cached && cached.fetchedAt && now() - Date.parse(cached.fetchedAt) < minutes * 60 * 1000;
      if (!fresh) run();
    }, 20000);
    const timer = setInterval(run, minutes * 60 * 1000);
    first.unref && first.unref();
    timer.unref && timer.unref();
    return { stop() { clearTimeout(first); clearInterval(timer); } };
  }

  return { getSettings, saveSettings, snapshot, refresh, refreshAround, refreshTasks, refreshPeople, refreshMonth, start };
}

module.exports = { createScheduleService };
