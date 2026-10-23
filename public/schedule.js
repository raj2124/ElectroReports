// Schedule (6F): who is on which site, when. Rows are people (you first), bars are the Onsite ESHA
// tasks of Zoho Projects, leave and holidays come from Zoho People. Read only: nothing here changes
// Zoho. The server decides what is shown and which warnings apply (src/schedule.js); this file draws it.

const scheduleUi = { mode: 'week', anchor: '', data: null, loading: false, error: '', note: '', settingsOpen: false, draft: null, confirm: null, busy: false };

const SCHEDULE_WEEKDAYS = [
  { index: 1, label: 'Mon' }, { index: 2, label: 'Tue' }, { index: 3, label: 'Wed' }, { index: 4, label: 'Thu' },
  { index: 5, label: 'Fri' }, { index: 6, label: 'Sat' }, { index: 0, label: 'Sun' }
];
const WARNING_TITLE = {
  double_booked: 'Two Onsite tasks on one day',
  leave_overlap: 'Onsite task over leave',
  no_owner: 'No owner',
  no_dates: 'No dates (not on the calendar)',
  one_date: 'Only one date',
  dates_reversed: 'End before start',
  quantity_short: 'Planned quantity needs more days than available',
  quantity_long: 'Far more days than the planned quantity needs'
};

const isoMs = (iso) => Date.parse(`${iso}T00:00:00Z`);
const isoFromMs = (ms) => new Date(ms).toISOString().slice(0, 10);
const isoAdd = (iso, count) => isoFromMs(isoMs(iso) + count * 86400000);

function scheduleWindow() {
  const anchor = scheduleUi.anchor || todayIso();
  if (scheduleUi.mode === 'month') {
    const first = `${anchor.slice(0, 7)}-01`;
    const next = new Date(Date.UTC(Number(anchor.slice(0, 4)), Number(anchor.slice(5, 7)), 0));
    return { from: first, to: isoFromMs(next.getTime()) };
  }
  const weekday = new Date(`${anchor}T00:00:00Z`).getUTCDay(); // Monday first
  const monday = isoAdd(anchor, -((weekday + 6) % 7));
  return { from: monday, to: isoAdd(monday, 6) };
}

function scheduleDayLabel(iso, options) {
  return new Intl.DateTimeFormat('en-GB', { timeZone: 'UTC', ...options }).format(new Date(`${iso}T00:00:00Z`));
}

async function loadSchedule() {
  const { from, to } = scheduleWindow();
  scheduleUi.loading = true;
  scheduleUi.error = '';
  render();
  try {
    scheduleUi.data = await api(`/api/schedule?from=${from}&to=${to}`);
  } catch (error) {
    scheduleUi.error = error.message;
  }
  scheduleUi.loading = false;
  render();
}

async function refreshSchedule() {
  const { from, to } = scheduleWindow();
  scheduleUi.loading = true;
  scheduleUi.note = '';
  render();
  try {
    const result = await api('/api/schedule/refresh', { method: 'POST', body: JSON.stringify({ from, to }) });
    scheduleUi.data = result;
    scheduleUi.error = '';
    scheduleUi.note = result.throttled ? `Zoho was read a moment ago. Refresh is limited to once a minute (try again in ${result.retryInSeconds} s).` : 'Refreshed from Zoho.';
  } catch (error) {
    scheduleUi.error = error.message;
  }
  scheduleUi.loading = false;
  render();
}

function ageText(seconds) {
  if (seconds === null || seconds === undefined) return 'not read yet';
  if (seconds < 60) return 'just now';
  if (seconds < 3600) return `${Math.round(seconds / 60)} min ago`;
  if (seconds < 86400) return `${Math.round(seconds / 3600)} h ago`;
  return `${Math.round(seconds / 86400)} days ago`;
}

function leaveRuns(dates) {
  const runs = [];
  dates.forEach((date) => {
    const last = runs[runs.length - 1];
    if (last && isoAdd(last.to, 1) === date) last.to = date;
    else runs.push({ from: date, to: date });
  });
  return runs;
}

function barTitle(bar, warnings) {
  const lines = [bar.name, `${bar.projectName}${bar.clientName ? ` · ${bar.clientName}` : ''}`, `${bar.start} to ${bar.end}`, `Owners: ${bar.owners.join(', ') || 'none'}`];
  if (bar.test) lines.push(`Test: ${bar.test}${bar.plannedQuantity ? ` · planned ${bar.plannedQuantity}` : ''}`);
  if (bar.status) lines.push(`Status: ${bar.status}${bar.duration ? ` · ${bar.duration}` : ''}`);
  bar.warnings.forEach((id) => {
    const warning = warnings.find((entry) => entry.id === id);
    if (warning) lines.push(`Warning: ${warning.message}`);
  });
  return lines.join('\n');
}

function renderScheduleRow(row, data) {
  const columns = data.days.length;
  const index = (date) => data.days.findIndex((day) => day.date === date) + 1;
  const today = todayIso();
  const cells = data.days
    .map((day, position) => {
      const off = !day.working;
      const holiday = row.holidayDates.includes(day.date);
      const clash = row.clashDates.includes(day.date);
      return `<div class="sched-cell${off ? ' sched-off' : ''}${holiday ? ' sched-holiday' : ''}${clash ? ' sched-clash' : ''}${day.date === today ? ' sched-today' : ''}" style="grid-column:${position + 1};grid-row:1 / span ${row.laneCount}"></div>`;
    })
    .join('');
  const leave = row.leave
    .flatMap((entry) => leaveRuns(entry.dates).map((run) => ({ entry, run })))
    .map(({ entry, run }) => `<div class="sched-leave sched-leave-${entry.state}" style="grid-column:${index(run.from)} / ${index(run.to) + 1};grid-row:1 / span ${row.laneCount}" title="${escapeHtml(`${entry.state === 'pending' ? 'Pending' : 'Approved'} leave: ${entry.type || 'leave'} (${entry.from} to ${entry.to})`)}"><span>${entry.state === 'pending' ? 'Leave?' : 'Leave'}</span></div>`)
    .join('');
  const bars = row.bars
    .map((bar) => {
      const flagged = bar.warnings.length > 0;
      return `<button type="button" class="sched-bar sched-c${bar.color}${flagged ? ' sched-bar-warn' : ''}${bar.clippedStart ? ' sched-clip-start' : ''}${bar.clippedEnd ? ' sched-clip-end' : ''}" style="grid-column:${index(bar.from)} / ${index(bar.to) + 1};grid-row:${bar.lane + 1}" data-action="schedule-open" data-project-id="${escapeHtml(bar.projectId)}" data-task-id="${escapeHtml(bar.taskId)}" title="${escapeHtml(barTitle(bar, data.warnings))}">${flagged ? '<b aria-label="warning">!</b> ' : ''}${escapeHtml(bar.projectName)} · ${escapeHtml(bar.name.replace(/^\[ESHA\]\s*/i, ''))}</button>`;
    })
    .join('');
  return `
    <div class="sched-row${row.isMe ? ' sched-me' : ''}" data-row="${escapeHtml(row.key)}">
      <div class="sched-name">
        <strong>${escapeHtml(row.name)}</strong>
        <span class="muted">${row.isMe ? 'You' : row.kind === 'zoho' ? 'Not signed in yet' : row.kind === 'none' ? 'Unassigned tasks' : ''}</span>
      </div>
      <div class="sched-track" style="--cols:${columns};--lanes:${row.laneCount}">${cells}${leave}${bars}</div>
    </div>`;
}

function renderScheduleSettings(data) {
  const draft = scheduleUi.draft;
  return `
    <section class="surface section-card sched-settings" aria-label="Schedule settings">
      <p class="section-kicker">Admin · schedule settings</p>
      <fieldset class="sched-days"><legend>Working days</legend>
        ${SCHEDULE_WEEKDAYS.map((day) => `<label class="role-box"><input type="checkbox" data-schedule-day="${day.index}" ${draft.workingDays.includes(day.index) ? 'checked' : ''} /> ${day.label}</label>`).join('')}
        <p class="muted">Holidays from Zoho People are treated as non-working days as well.</p>
      </fieldset>
      <table class="data-table">
        <thead><tr><th>Test</th><th>Per engineer per working day</th><th>Unit</th><th><span class="visually-hidden">Note</span></th></tr></thead>
        <tbody>${data.tests.map((test) => `<tr><td>${escapeHtml(test.name)}</td><td><input class="text-input" type="number" min="0.1" step="0.5" aria-label="${escapeHtml(test.name)}: per engineer per working day" data-schedule-rate="${test.id}" value="${escapeHtml(draft.rates[test.id].perDay)}" /></td><td>${escapeHtml(draft.rates[test.id].unit)}</td><td><span class="chip">estimate</span></td></tr>`).join('')}</tbody>
      </table>
      <p class="muted">These rates are estimates used only for the "planned quantity vs days" warning. They never block anything.</p>
      <div class="button-row">
        <button class="button button-primary" data-action="schedule-save-settings" ${scheduleUi.busy ? 'disabled' : ''}>Save settings</button>
        <button class="button button-secondary" data-action="schedule-toggle-settings">Close</button>
      </div>
    </section>`;
}

function scheduleView() {
  const data = scheduleUi.data;
  const window = scheduleWindow();
  const title = scheduleUi.mode === 'month' ? scheduleDayLabel(window.from, { month: 'long', year: 'numeric' }) : `${scheduleDayLabel(window.from, { day: 'numeric', month: 'short' })} to ${scheduleDayLabel(window.to, { day: 'numeric', month: 'short', year: 'numeric' })}`;
  const projectConfirm = scheduleUi.confirm
    ? `<div class="surface notice-card" role="alert"><p><strong>${escapeHtml(scheduleUi.confirm.name)}</strong> has no ElectroReports workspace yet. Create it and open it?</p><div class="button-row"><button class="button button-primary" data-action="schedule-create-workspace">Create workspace and open</button><button class="button button-ghost" data-action="schedule-cancel-create">Cancel</button></div></div>`
    : '';
  const banners = data
    ? [
        data.zoho.tasks.error ? `<p class="form-error" role="alert">Zoho Projects: ${escapeHtml(data.zoho.tasks.error)}${data.zoho.tasks.fetchedAt ? ' Showing the last data read.' : ''}</p>` : '',
        data.zoho.leave.error ? `<p class="form-error" role="alert">Zoho People (leave, holidays): ${escapeHtml(data.zoho.leave.error)}${data.zoho.leave.loaded ? ' Showing the last data read.' : ' Leave and holidays are not shown.'}</p>` : ''
      ].join('')
    : '';
  const warnings = data ? data.warnings : [];
  const grouped = Object.keys(WARNING_TITLE).map((kind) => ({ kind, items: warnings.filter((warning) => warning.kind === kind) })).filter((group) => group.items.length);
  return `
    <section class="surface section-card sched-card" aria-label="Schedule">
      <div class="sched-head">
        <div>
          <p class="section-kicker">Schedule</p>
          <h1>${escapeHtml(title)}</h1>
          <p class="muted">${data ? `Onsite ESHA tasks from Zoho Projects, tasks read ${escapeHtml(ageText(data.zoho.tasks.ageSeconds))}${data.zoho.mock ? ' (mock portal)' : ''}. Leave and holidays from Zoho People. Read only.` : 'Loading from Zoho...'}</p>
        </div>
        <div class="sched-controls">
          <div class="button-row">
            <button class="button ${scheduleUi.mode === 'week' ? 'button-primary' : 'button-secondary'}" data-action="schedule-mode" data-mode="week">Week</button>
            <button class="button ${scheduleUi.mode === 'month' ? 'button-primary' : 'button-secondary'}" data-action="schedule-mode" data-mode="month">Month</button>
          </div>
          <div class="button-row">
            <button class="button button-secondary" data-action="schedule-step" data-step="-1" aria-label="Previous">&#8249;</button>
            <button class="button button-secondary" data-action="schedule-today">Today</button>
            <button class="button button-secondary" data-action="schedule-step" data-step="1" aria-label="Next">&#8250;</button>
            <button class="button button-primary" data-action="schedule-refresh" ${scheduleUi.loading ? 'disabled' : ''}>Refresh from Zoho</button>
            ${hasRoleClient('admin') ? '<button class="button button-ghost" data-action="schedule-toggle-settings">Settings</button>' : ''}
            <button class="button button-ghost" data-action="go-dashboard">Back to Dashboard</button>
          </div>
        </div>
      </div>
      ${scheduleUi.note ? `<p class="review-message" role="status">${escapeHtml(scheduleUi.note)}</p>` : ''}
      ${scheduleUi.error ? `<p class="form-error" role="alert">${escapeHtml(scheduleUi.error)}</p>` : ''}
      ${banners}
      ${projectConfirm}
      ${scheduleUi.settingsOpen && data && scheduleUi.draft ? renderScheduleSettings(data) : ''}
      ${
        data
          ? `
      <div class="sched-scroll">
        <div class="sched-grid" style="--cols:${data.days.length};--day-min:${data.days.length > 10 ? 29 : 38}px">
          <div class="sched-row sched-header">
            <div class="sched-name"></div>
            <div class="sched-track" style="--cols:${data.days.length}">
              ${data.days.map((day, position) => `<div class="sched-day${day.working ? '' : ' sched-off'}${day.holidays.length ? ' sched-holiday' : ''}${day.date === todayIso() ? ' sched-today' : ''}" style="grid-column:${position + 1}" ${day.holidays.length ? `title="${escapeHtml(day.holidays.map((holiday) => holiday.name + (holiday.locations.length ? ` (${holiday.locations.join(', ')})` : '')).join(', '))}"` : ''}><span>${escapeHtml(day.weekday)}</span><strong>${Number(day.date.slice(8))}</strong>${day.holidays.length ? '<em>Holiday</em>' : ''}</div>`).join('')}
            </div>
          </div>
          ${data.rows.map((row) => renderScheduleRow(row, data)).join('')}
        </div>
      </div>
      <div class="sched-legend muted">
        ${data.projects.map((project) => `<span><i class="sched-swatch sched-c${project.color}"></i>${escapeHtml(project.name)}</span>`).join('')}
        <span><i class="sched-swatch sched-leave-approved"></i>Approved leave</span>
        <span><i class="sched-swatch sched-leave-pending"></i>Pending leave</span>
        <span><i class="sched-swatch sched-off"></i>Non-working day</span>
        <span><i class="sched-swatch sched-clash"></i>Two tasks one day</span>
      </div>
      <section class="sched-warnings" aria-label="Warnings">
        <h3 aria-level="2">${warnings.length ? `${warnings.length} warning${warnings.length === 1 ? '' : 's'}` : 'No warnings'}</h3>
        <p class="muted">Warnings are shown for what is on screen. They never block anything.</p>
        ${grouped.map((group) => `<details open><summary>${escapeHtml(WARNING_TITLE[group.kind])} (${group.items.length})</summary><ul>${group.items.map((warning) => `<li data-warning="${escapeHtml(warning.kind)}">${escapeHtml(warning.message)}</li>`).join('')}</ul></details>`).join('')}
      </section>`
          : scheduleUi.loading ? '<p class="muted">Reading Zoho...</p>' : ''
      }
    </section>`;
}

// Returns true when the click was one of ours.
function handleScheduleAction(action, button) {
  if (action === 'open-schedule') {
    state.view = 'schedule';
    scheduleUi.anchor = scheduleUi.anchor || todayIso();
    loadSchedule();
    return true;
  }
  if (action === 'schedule-mode') {
    scheduleUi.mode = button.dataset.mode === 'month' ? 'month' : 'week';
    loadSchedule();
    return true;
  }
  if (action === 'schedule-step') {
    const step = Number(button.dataset.step);
    const current = scheduleWindow();
    scheduleUi.anchor = scheduleUi.mode === 'month' ? isoAdd(step > 0 ? current.to : current.from, step > 0 ? 1 : -1) : isoAdd(current.from, step * 7);
    loadSchedule();
    return true;
  }
  if (action === 'schedule-today') {
    scheduleUi.anchor = todayIso();
    loadSchedule();
    return true;
  }
  if (action === 'schedule-refresh') {
    refreshSchedule();
    return true;
  }
  if (action === 'schedule-open') {
    const project = scheduleUi.data && scheduleUi.data.projects.find((entry) => entry.id === button.dataset.projectId);
    if (!project) return true;
    if (project.workspaceId) {
      openWorkspace(project.workspaceId);
    } else {
      scheduleUi.confirm = project;
      render();
    }
    return true;
  }
  if (action === 'schedule-cancel-create') {
    scheduleUi.confirm = null;
    render();
    return true;
  }
  if (action === 'schedule-create-workspace') {
    const project = scheduleUi.confirm;
    scheduleUi.confirm = null;
    openWorkspaceFromZohoProject({ id: project.id, name: project.name, projectNumber: project.zoho.projectNumber || project.name, clientName: project.clientName, workOrderNo: project.zoho.workOrderNo, ownerName: project.zoho.ownerName, stage: project.zoho.stage });
    return true;
  }
  if (action === 'schedule-toggle-settings') {
    scheduleUi.settingsOpen = !scheduleUi.settingsOpen;
    if (scheduleUi.settingsOpen && scheduleUi.data) scheduleUi.draft = JSON.parse(JSON.stringify(scheduleUi.data.settings));
    render();
    return true;
  }
  if (action === 'schedule-save-settings') {
    scheduleUi.busy = true;
    render();
    api('/api/admin/schedule-settings', { method: 'PUT', body: JSON.stringify(scheduleUi.draft) })
      .then(async () => {
        scheduleUi.note = 'Settings saved.';
        scheduleUi.settingsOpen = false;
        await loadSchedule();
      })
      .catch((error) => showToast(error.message, 'critical'))
      .finally(() => {
        scheduleUi.busy = false;
        render();
      });
    return true;
  }
  return false;
}

// Returns true when handled.
function handleScheduleChange(target) {
  const data = target.dataset || {};
  if (data.scheduleDay !== undefined && scheduleUi.draft) {
    const day = Number(data.scheduleDay);
    scheduleUi.draft.workingDays = target.checked ? [...new Set([...scheduleUi.draft.workingDays, day])].sort() : scheduleUi.draft.workingDays.filter((entry) => entry !== day);
    return true;
  }
  if (data.scheduleRate && scheduleUi.draft) {
    scheduleUi.draft.rates[data.scheduleRate].perDay = Number(target.value);
    return true;
  }
  return false;
}
