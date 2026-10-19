// Step 1: the project's tasklists and tasks from Zoho (6D). The server decides which tasks are
// measurement tasks, which ESHA test each is and what the planned quantity is (src/zohoTasks.js);
// this file shows that, applies the selection to the draft (tests, planned quantities, engineer) and
// marks pre-filled values "from Zoho". Everything pre-filled stays editable.

const planUi = {
  projectId: '',
  plan: null,
  loading: false,
  error: '',
  selected: {}, // task id -> boolean
  plannedFrom: {}, // test id -> true while the planned quantity is still the Zoho value
  autoEngineer: ''
};

function resetPlanUi() {
  Object.assign(planUi, { projectId: '', plan: null, loading: false, error: '', selected: {}, plannedFrom: {}, autoEngineer: '' });
}

async function loadZohoPlan(projectId) {
  resetPlanUi();
  if (!projectId) {
    return;
  }
  planUi.projectId = projectId;
  planUi.loading = true;
  render();
  try {
    const result = await api(`/api/zoho/projects/${encodeURIComponent(projectId)}/plan`);
    if (planUi.projectId !== projectId) {
      return;
    }
    planUi.plan = result.plan;
    result.plan.tasks.forEach((task) => {
      planUi.selected[task.id] = task.selected;
    });
    applyPlanToDraft(true);
    if (!result.plan.d03 && !safeText(state.draft.project.engineerName, '')) {
      const zohoProject = state.zoho.projects.find((item) => String(item.id) === String(projectId));
      state.draft.project.engineerName = safeText(zohoProject && zohoProject.ownerName, '');
    }
  } catch (error) {
    planUi.error = error.message;
  } finally {
    planUi.loading = false;
    render();
  }
}

function selectedPlanTasks() {
  return planUi.plan ? planUi.plan.tasks.filter((task) => planUi.selected[task.id]) : [];
}

// Puts the selection into the draft. `initial` also sets the tests that have no measurement task to
// off, so a D03 project starts with exactly the tests its Onsite tasks name.
function applyPlanToDraft(initial = false) {
  const plan = planUi.plan;
  if (!plan || !plan.d03) {
    return;
  }
  const project = state.draft.project;
  const chosen = selectedPlanTasks().filter((task) => task.measurement);
  const chosenTests = new Set(chosen.map((task) => task.testId));
  const planTests = Object.keys(plan.tests);
  if (initial && planTests.length) {
    Object.keys(state.draft.tests).forEach((id) => {
      state.draft.tests[id] = chosenTests.has(id);
    });
  }
  planTests.forEach((id) => {
    state.draft.tests[id] = chosenTests.has(id);
  });
  project.plannedQuantities = project.plannedQuantities || {};
  planTests.forEach((id) => {
    const test = plan.tests[id];
    const current = String(project.plannedQuantities[id] === undefined ? '' : project.plannedQuantities[id]).trim();
    if (chosenTests.has(id) && test.planned !== null && (!current || planUi.plannedFrom[id])) {
      project.plannedQuantities[id] = String(test.planned);
      planUi.plannedFrom[id] = true;
    } else if (!chosenTests.has(id) && planUi.plannedFrom[id]) {
      project.plannedQuantities[id] = '';
      planUi.plannedFrom[id] = false;
    }
  });
  const engineers = [...new Set(chosen.flatMap((task) => task.owners.map((owner) => owner.fullName || owner.name)).filter(Boolean))];
  if (engineers.length) {
    const name = engineers.join(', ');
    if (!safeText(project.engineerName, '') || project.engineerName === planUi.autoEngineer) {
      project.engineerName = name;
      planUi.autoEngineer = name;
    }
  }
}

// Under a planned-quantity field: "from Zoho", the hint, what Zoho says, the warnings.
function plannedNote(testId) {
  const test = planUi.plan && planUi.plan.tests ? planUi.plan.tests[testId] : null;
  if (!test) {
    return '';
  }
  return `
    <div class="planned-note">
      ${planUi.plannedFrom[testId] ? `<span class="pill pill-healthy" data-from-zoho="${escapeHtml(testId)}">from Zoho</span>` : ''}
      ${test.hint ? `<span class="muted">${escapeHtml(test.hint)} (not filled in)</span>` : ''}
      ${test.zohoSays ? `<span class="muted">Zoho says: ${escapeHtml(test.zohoSays)}</span>` : ''}
      ${test.warnings.map((message) => `<span class="planned-warning">${escapeHtml(message)}</span>`).join('')}
    </div>`;
}

function renderZohoTaskPanel() {
  if (state.editing || !state.draft.project.zohoProjectId) {
    return '';
  }
  if (planUi.loading) {
    return '<div class="zoho-tasks"><p class="muted">Loading the project\'s tasks from Zoho...</p></div>';
  }
  if (planUi.error) {
    return `<div class="zoho-tasks"><p class="form-error" role="alert">Could not read the Zoho tasks: ${escapeHtml(planUi.error)}</p></div>`;
  }
  const plan = planUi.plan;
  if (!plan) {
    return '';
  }
  const picked = selectedPlanTasks().length;
  return `
    <div class="zoho-tasks">
      <p class="section-kicker">Zoho tasks</p>
      <h4>${plan.d03 ? 'Measurement tasks are pre-selected' : 'Tasks of this project'}</h4>
      <p class="muted">${plan.d03 ? `${picked} task${picked === 1 ? '' : 's'} selected. The selection sets the tests, the planned quantities and the engineer; change it if the site visit differs.` : 'This is not a D03 project: the tasks are listed, nothing is selected for you.'}</p>
      ${plan.warnings.length ? `<ul class="planned-warnings">${plan.warnings.map((message) => `<li>${escapeHtml(message)}</li>`).join('')}</ul>` : ''}
      ${plan.taskLists
        .map(
          (list) => `
        <details class="zoho-tasklist" ${list.taskIds.some((id) => planUi.selected[id]) ? 'open' : ''}>
          <summary>${escapeHtml(list.name)} <span class="muted">(${list.taskIds.length})</span></summary>
          <div class="zoho-task-rows">
            ${list.taskIds
              .map((id) => plan.tasks.find((task) => task.id === id))
              .map(
                (task) => `
              <label class="zoho-task-row ${task.irt ? 'is-irt' : ''}">
                <input type="checkbox" data-zoho-task="${escapeHtml(task.id)}" ${planUi.selected[task.id] ? 'checked' : ''} ${plan.d03 ? '' : 'disabled'} />
                <span class="zoho-task-main"><strong>${escapeHtml(task.name)}</strong>
                  <span class="muted">${escapeHtml(task.ownerName || 'No owner')} · ${escapeHtml(task.startDate || '-')}${task.dueDate && task.dueDate !== task.startDate ? ` to ${escapeHtml(task.dueDate)}` : ''} · ${escapeHtml(task.status || '-')}${task.category ? ` · ${escapeHtml(task.category)}` : ''}${task.eshaTest ? ` · ${escapeHtml(task.eshaTest)}` : ''}</span>
                  ${task.warnings.map((message) => `<span class="planned-warning">${escapeHtml(message)}</span>`).join('')}
                </span>
              </label>`
              )
              .join('')}
          </div>
        </details>`
        )
        .join('')}
    </div>`;
}

// Returns true when handled.
function handlePlanChange(target) {
  if (target.dataset && target.dataset.zohoTask) {
    planUi.selected[target.dataset.zohoTask] = target.checked;
    applyPlanToDraft(false);
    render();
    return true;
  }
  return false;
}

// Typing in a planned-quantity field ends "from Zoho" (the value is the engineer's now).
function planOnPlannedEdit(bind) {
  const match = /^project\.plannedQuantities\.(.+)$/.exec(bind || '');
  if (match && planUi.plannedFrom[match[1]]) {
    planUi.plannedFrom[match[1]] = false;
    const badge = document.querySelector(`[data-from-zoho="${match[1]}"]`);
    if (badge) {
      badge.remove();
    }
  }
}
