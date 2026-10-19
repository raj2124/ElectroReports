// Project workspace (6C.4): one page per Zoho project with its site days, the cumulative progress
// against the planned quantities, the final reports, "Start today's report", "Assemble final
// report" and "Print sheets for site". The home screen lists the active projects. Rendering and
// state only: what a day carries over, which days exist and what assembling does are decided by
// the server (src/workspace.js).

const workspaceUi = {
  projects: [],
  workspace: null,
  workDate: '',
  confirmAssemble: false,
  busy: false,
  planned: {}
};

const DAY_STATUS_LABEL = { draft: 'Draft', in_review: 'In review', checked: 'Checked', approved: 'Approved' };
const DAY_STATUS_TONE = { draft: 'neutral', in_review: 'warning', checked: 'warning', approved: 'healthy' };

function todayIso() {
  return new Date().toISOString().slice(0, 10);
}

async function loadProjects() {
  try {
    const list = await api('/api/projects');
    workspaceUi.projects = Array.isArray(list) ? list : [];
  } catch (_error) {
    workspaceUi.projects = [];
  }
  if (state.view === 'dashboard') {
    render();
  }
}

function projectCard(project) {
  const final = project.finalStatus ? `${DAY_STATUS_LABEL[project.finalStatus] || project.finalStatus}${project.finalRevision ? ` (R${project.finalRevision})` : ''}` : 'Not assembled';
  return `
    <article class="project-card reveal-card">
      <span class="chip">${escapeHtml(project.projectNo)}</span>
      <h4>${escapeHtml(project.clientName)}</h4>
      <p class="muted">${escapeHtml(project.siteLocation || '')}</p>
      <div class="project-card-facts">
        <span>Site days <strong>${project.days}</strong></span>
        <span>Approved <strong>${project.approvedDays}</strong></span>
        <span>Last day <strong>${escapeHtml(project.lastDay ? formatDisplayDate(project.lastDay) || project.lastDay : '-')}</strong></span>
      </div>
      <div class="project-card-facts"><span>Final report <strong>${escapeHtml(final)}</strong></span></div>
      <div class="card-actions"><button class="button button-primary" data-action="open-workspace" data-id="${escapeHtml(project.id)}">Open workspace</button></div>
    </article>`;
}

function renderHomeProjects() {
  const active = workspaceUi.projects.filter((project) => project.state === 'active');
  const completed = workspaceUi.projects.filter((project) => project.state === 'completed');
  return `
    <section class="surface zoho-overview" aria-label="Projects">
      <div class="dashboard-section-copy">
        <p class="section-kicker">Projects</p>
        <h3>Active projects</h3>
        <p class="muted">Each project has a workspace: its site days and their status, progress against the plan, and the final report.</p>
      </div>
      ${active.length ? `<div class="project-grid">${active.map(projectCard).join('')}</div>` : '<p class="muted">No active project yet. Open a workspace from a Zoho project below.</p>'}
      ${completed.length ? `<details class="completed-projects"><summary>Completed (${completed.length})</summary><div class="project-grid">${completed.map(projectCard).join('')}</div></details>` : ''}
    </section>`;
}

async function openWorkspace(projectId) {
  try {
    workspaceUi.workspace = await api(`/api/projects/${encodeURIComponent(projectId)}`);
    workspaceUi.confirmAssemble = false;
    workspaceUi.planned = Object.fromEntries(Object.entries(workspaceUi.workspace.planned || {}).map(([key, value]) => [key, String(value)]));
    if (!workspaceUi.workDate) {
      workspaceUi.workDate = todayIso();
    }
    state.editing = null;
    state.view = 'workspace';
  } catch (error) {
    showToast(error.message, 'critical');
  }
  render();
}

async function refreshWorkspace() {
  if (workspaceUi.workspace) {
    workspaceUi.workspace = await api(`/api/projects/${encodeURIComponent(workspaceUi.workspace.project.id)}`);
  }
}

async function openWorkspaceFromZoho(zohoProjectId) {
  const project = state.zoho.projects.find((item) => String(item.id) === String(zohoProjectId));
  if (!project) {
    showToast('That Zoho project is not loaded. Refresh Zoho and try again.', 'warning');
    return;
  }
  try {
    const result = await api('/api/projects', {
      method: 'POST',
      body: JSON.stringify({
        project: {
          projectNo: safeText(project.projectNumber || getZohoProjectVisibleCode(project) || project.name, ''),
          clientName: safeText(project.clientName, ''),
          workOrder: safeText(project.workOrderNo, ''),
          engineerName: safeText(project.ownerName, ''),
          zohoProjectId: safeText(project.id, ''),
          zohoProjectName: safeText(project.name, ''),
          zohoProjectOwner: safeText(project.ownerName, ''),
          zohoProjectStage: safeText(project.stage, '')
        }
      })
    });
    await loadProjects();
    await openWorkspace(result.projectId);
  } catch (error) {
    showToast(error.message, 'critical');
  }
}

function statusPill(status) {
  return pill(DAY_STATUS_TONE[status] || 'neutral', DAY_STATUS_LABEL[status] || status);
}

function renderProgress(workspace) {
  if (!workspace.progress.length) {
    return '<p class="muted">Set the planned quantities below to track progress. Nothing is measured yet.</p>';
  }
  return workspace.progress
    .map((row) => {
      const percent = row.percent === null ? null : Math.min(row.percent, 100);
      return `
      <div class="progress-row">
        <span>${escapeHtml(row.label)}</span>
        <strong>${row.done}${row.planned === null ? ' (plan not set)' : ` of ${row.planned} ${escapeHtml(row.unit)}${row.percent === null ? '' : ` · ${row.percent}%`}`}</strong>
        <div class="progress-track" title="${percent === null ? 'No planned quantity set' : `${row.percent}%`}"><div class="progress-fill ${row.percent >= 100 ? 'is-complete' : ''}" style="width:${percent === null ? 0 : percent}%"></div></div>
      </div>`;
    })
    .join('');
}

function renderPlannedForm(workspace) {
  const ids = [...new Set([...Object.keys(workspace.planned || {}), ...Object.keys(workspace.dayTemplate.tests || {}).filter((id) => workspace.dayTemplate.tests[id])])];
  const tests = ids
    .map((id) => (id === 'soilResistivity' ? { id, label: 'Soil resistivity (locations)' } : state.catalog.find((test) => test.id === id)))
    .filter(Boolean);
  if (!tests.length) {
    return '';
  }
  return `
    <details class="workspace-planned-form">
      <summary>Planned quantities</summary>
      <div class="workspace-planned">
        ${tests.map((test) => `<label class="field"><span>${escapeHtml(test.shortLabel || test.label)}</span><input class="text-input" type="number" min="0" inputmode="numeric" data-workspace-planned="${escapeHtml(test.id)}" value="${escapeHtml(workspaceUi.planned[test.id] || '')}" /></label>`).join('')}
      </div>
      <div class="button-row"><button class="button button-secondary" data-action="workspace-save-planned">Save planned quantities</button></div>
    </details>`;
}

function renderAssemblePanel(workspace) {
  const assembly = workspace.assembly;
  if (!workspaceUi.confirmAssemble) {
    return '';
  }
  return `
    <div class="workspace-confirm" role="alertdialog" aria-label="Assemble final report">
      <strong>Assemble the final report from ${assembly.totalDailyReports} daily report${assembly.totalDailyReports === 1 ? '' : 's'}</strong>
      <span>Included dates: ${assembly.includedDates.map((date) => escapeHtml(formatDisplayDate(date) || date)).join(', ') || '-'}</span>
      ${assembly.hasUnapproved ? `<span><strong>Not yet approved:</strong> ${assembly.flaggedDates.map((date) => escapeHtml(formatDisplayDate(date) || date)).join(', ')}. These days are still included, and the final report will say they were not approved when it was assembled.</span>` : '<span>Every day is approved.</span>'}
      ${workspace.finals.length ? '<span>An earlier final report exists. It is kept as it is; this makes a new draft.</span>' : ''}
      <div class="button-row">
        <button class="button button-primary" data-action="workspace-assemble-confirm" ${workspaceUi.busy ? 'disabled' : ''}>${workspaceUi.busy ? 'Assembling...' : 'Assemble final report'}</button>
        <button class="button button-secondary" data-action="workspace-assemble-cancel">Cancel</button>
      </div>
    </div>`;
}

function workspaceView() {
  const workspace = workspaceUi.workspace;
  if (!workspace) {
    return '<section class="surface empty-state"><p>Loading project...</p></section>';
  }
  const project = workspace.project;
  const existing = workspace.days.find((day) => day.date === workspaceUi.workDate);
  return `
    <section class="surface detail-shell workspace-shell">
      <div class="detail-hero">
        <div class="detail-hero-copy">
          <p class="section-kicker">Project workspace</p>
          <h2>${escapeHtml(project.clientName)}</h2>
          <p>${escapeHtml(project.projectNo)}${project.siteLocation ? ` · ${escapeHtml(project.siteLocation)}` : ''}${project.unit ? ` · ${escapeHtml(project.unit)}` : ''}</p>
          <div class="detail-hero-tags">${pill(workspace.state === 'completed' ? 'healthy' : 'warning', workspace.state === 'completed' ? 'Completed' : 'Active')}</div>
        </div>
        <div class="button-row detail-hero-actions">
          <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
          <button class="button button-secondary" data-action="workspace-print">Print sheets for site</button>
        </div>
      </div>

      <div class="workspace-grid">
        <section class="surface section-card">
          <p class="section-kicker">Site days</p>
          <h3>Daily reports</h3>
          <div class="button-row">
            <label class="field"><span>Work date</span><input class="text-input" type="date" data-workspace-field="workDate" value="${escapeHtml(workspaceUi.workDate)}" /></label>
            <button class="button button-primary" data-action="workspace-start-day">${existing ? `Reopen ${escapeHtml(formatDisplayDate(workspaceUi.workDate) || workspaceUi.workDate)}` : "Start today's report"}</button>
          </div>
          <p class="muted">${existing ? 'This date already has a daily report: it is reopened, not duplicated.' : 'Project details carry over from the previous day. The plan for the next day does not.'}</p>
          ${
            workspace.days.length
              ? `<table class="data-table workspace-days"><thead><tr><th>Date</th><th>Status</th><th></th></tr></thead><tbody>
              ${workspace.days
                .map(
                  (day) => `<tr>
                <td>${escapeHtml(formatDisplayDate(day.date) || day.date)}</td>
                <td>${statusPill(day.status)}</td>
                <td><div class="day-actions">
                  ${day.status === 'draft' ? `<button class="button button-secondary" data-action="workspace-edit-day" data-id="${day.reportId}">Edit</button>` : ''}
                  <button class="button button-secondary" data-action="open-review" data-id="${day.reportId}">Review</button>
                  <button class="button button-ghost" data-action="export-pdf" data-id="${day.reportId}">PDF</button>
                </div></td></tr>`
                )
                .join('')}
            </tbody></table>`
              : '<p class="muted">No site day recorded yet.</p>'
          }
        </section>

        <section class="surface section-card">
          <p class="section-kicker">Progress</p>
          <h3>Against the plan</h3>
          ${renderProgress(workspace)}
          ${renderPlannedForm(workspace)}
        </section>
      </div>

      <section class="surface section-card">
        <p class="section-kicker">Final report</p>
        <h3>${workspace.finals.length ? 'Final reports' : 'Not assembled yet'}</h3>
        ${
          workspace.finals.length
            ? `<table class="data-table"><thead><tr><th>Revision</th><th>Status</th><th>Days included</th><th></th></tr></thead><tbody>
            ${workspace.finals
              .map(
                (final) => `<tr><td>R${final.revision}</td><td>${statusPill(final.status)}</td>
                <td>${final.includedDates.length}${final.flaggedDates.length ? ` (${final.flaggedDates.length} not approved then)` : ''}</td>
                <td><button class="button button-secondary" data-action="open-review" data-id="${final.id}">Review</button></td></tr>`
              )
              .join('')}
          </tbody></table>`
            : '<p class="muted">Assemble the final report once the site days are recorded. Every day is stitched in; days not yet approved are listed and flagged.</p>'
        }
        ${renderAssemblePanel(workspace)}
        <div class="button-row">
          <button class="button button-primary" data-action="workspace-assemble" ${workspace.days.length ? '' : 'disabled'}>Assemble final report</button>
        </div>
      </section>
    </section>`;
}

// ---- Entering a site day in the report builder -------------------------------------------------

function startNewDay(workDate) {
  const workspace = workspaceUi.workspace;
  const draft = createDraft();
  draft.project = { ...draft.project, ...cloneSerializable(workspace.dayTemplate.project), reportDate: workDate, nextDayPlan: '' };
  if (Object.keys(workspace.dayTemplate.tests || {}).length) {
    draft.tests = { ...draft.tests, ...workspace.dayTemplate.tests };
  }
  closeObservationEditor();
  state.draft = draft;
  state.editing = { projectId: workspace.project.id, workDate, reportId: '' };
  state.stepIndex = 0;
  state.view = 'builder';
  resetAllOcrStates(false);
  clearDraftSnapshot();
  render();
}

async function editDay(reportId) {
  try {
    const report = await api(`/api/reports/${encodeURIComponent(reportId)}`);
    closeObservationEditor();
    state.draft = restoreDraftShape(report);
    state.editing = { projectId: report.projectId, workDate: report.collectionDate, reportId: report.id };
    state.stepIndex = 0;
    state.view = 'builder';
    resetAllOcrStates(false);
    clearDraftSnapshot();
    render();
  } catch (error) {
    showToast(error.message, 'critical');
  }
}

async function saveWorkspaceDay() {
  const editing = state.editing;
  state.saving = true;
  closeObservationEditor();
  resetAllOcrStates(false);
  render();
  try {
    const result = editing.reportId
      ? await api(`/api/reports/${encodeURIComponent(editing.reportId)}`, { method: 'PUT', body: JSON.stringify(state.draft) })
      : await api(`/api/projects/${encodeURIComponent(editing.projectId)}/days`, { method: 'POST', body: JSON.stringify({ workDate: editing.workDate, report: state.draft }) });
    const warnings = ((result.assessment && result.assessment.validation && result.assessment.validation.warnings) || []).length;
    state.editing = null;
    state.draft = createDraft();
    state.stepIndex = 0;
    clearDraftSnapshot();
    await loadReports();
    await loadProjects();
    await openWorkspace(editing.projectId);
    showToast(warnings ? `Site day saved with ${warnings} warning${warnings === 1 ? '' : 's'} to check.` : 'Site day saved.', warnings ? 'warning' : 'healthy');
  } catch (error) {
    if (error.code === 'DAY_EXISTS') {
      showToast(error.message, 'warning');
    } else {
      const extra = Array.isArray(error.errors) && error.errors.length > 1 ? ` (${error.errors.length} problems; see the sheets)` : '';
      showToast(`${error.message}${extra}`, 'critical');
    }
  } finally {
    state.saving = false;
    render();
  }
}

// ---- Clicks and typing -------------------------------------------------------------------------

function workspacePrintProject() {
  const workspace = workspaceUi.workspace;
  return { ...workspace.dayTemplate.project, reportDate: workspaceUi.workDate || todayIso(), tests: workspace.dayTemplate.tests || {} };
}

// Returns true when the click was one of ours.
function handleWorkspaceAction(action, button) {
  if (action === 'open-workspace') {
    openWorkspace(button.dataset.id);
    return true;
  }
  if (action === 'open-workspace-zoho') {
    openWorkspaceFromZoho(button.dataset.projectId);
    return true;
  }
  if (action === 'workspace-back') {
    const projectId = state.editing && state.editing.projectId;
    state.editing = null;
    clearDraftSnapshot();
    if (projectId) {
      openWorkspace(projectId);
    } else {
      state.view = 'dashboard';
      render();
    }
    return true;
  }
  if (action === 'workspace-start-day') {
    const date = workspaceUi.workDate;
    if (!date) {
      showToast('Choose the work date first.', 'warning');
      return true;
    }
    const existing = workspaceUi.workspace.days.find((day) => day.date === date);
    if (!existing) {
      startNewDay(date);
    } else if (existing.status === 'draft') {
      editDay(existing.reportId);
    } else {
      showToast(`${formatDisplayDate(date) || date} is ${DAY_STATUS_LABEL[existing.status].toLowerCase()}: opening it for review.`, 'warning');
      openReview(existing.reportId);
    }
    return true;
  }
  if (action === 'workspace-edit-day') {
    editDay(button.dataset.id);
    return true;
  }
  if (action === 'workspace-print') {
    openPrintSheets('', workspacePrintProject());
    return true;
  }
  if (action === 'workspace-assemble') {
    workspaceUi.confirmAssemble = true;
    render();
    return true;
  }
  if (action === 'workspace-assemble-cancel') {
    workspaceUi.confirmAssemble = false;
    render();
    return true;
  }
  if (action === 'workspace-assemble-confirm') {
    workspaceUi.busy = true;
    render();
    api(`/api/projects/${encodeURIComponent(workspaceUi.workspace.project.id)}/final/assemble`, { method: 'POST', body: JSON.stringify({ actor: (reviewUi && reviewUi.actor) || '' }) })
      .then(async (result) => {
        workspaceUi.confirmAssemble = false;
        showToast(result.hasUnapproved ? `Final report assembled. Not yet approved: ${result.flaggedDates.join(', ')}.` : 'Final report assembled.', result.hasUnapproved ? 'warning' : 'healthy');
        await loadProjects();
        await openReview(result.report.id);
      })
      .catch((error) => showToast(error.message, 'critical'))
      .finally(() => {
        workspaceUi.busy = false;
        render();
      });
    return true;
  }
  if (action === 'workspace-save-planned') {
    api(`/api/projects/${encodeURIComponent(workspaceUi.workspace.project.id)}`, { method: 'PUT', body: JSON.stringify({ project: { plannedQuantities: workspaceUi.planned } }) })
      .then((workspace) => {
        workspaceUi.workspace = workspace;
        showToast('Planned quantities saved.', 'healthy');
        render();
      })
      .catch((error) => showToast(error.message, 'critical'));
    return true;
  }
  return false;
}

// Typing in the workspace: kept in state without re-rendering. Returns true when handled.
function handleWorkspaceInput(target) {
  if (target.dataset && target.dataset.workspaceField === 'workDate') {
    workspaceUi.workDate = target.value;
    render();
    return true;
  }
  if (target.dataset && target.dataset.workspacePlanned) {
    workspaceUi.planned[target.dataset.workspacePlanned] = target.value;
    return true;
  }
  return false;
}
