// "Print sheets for site": choose the work date and tests, pick the Zoho task (its owner is
// printed as the engineer) or type the engineer, and download the printable measurement sheets.
// Rendering and state only; the sheets themselves are generated on the server.

const printState = {
  open: false,
  workDate: '',
  selected: {},
  taskId: '',
  engineerName: '',
  project: null, // the workspace's project when opened from a project workspace; else the wizard draft
  tasks: [],
  loadingTasks: false,
  busy: false,
  error: ''
};

// The project the sheets are printed for: the workspace's, or the report being built.
function printProject() {
  return printState.project || state.draft.project;
}

function printableSheets() {
  return allSheets().filter((sheet) => !sheet.retired);
}

function renderPrintButton(extraClass = '') {
  return `<button class="button button-secondary print-sheets-button ${extraClass}" type="button" data-action="open-print-sheets">Print sheets for site</button>`;
}

function renderPrintDialog() {
  if (!printState.open) {
    return '';
  }
  const project = printProject();
  const sheets = printableSheets();
  const taskOptions = printState.tasks
    .map((task) => `<option value="${escapeHtml(task.id)}" ${String(task.id) === printState.taskId ? 'selected' : ''}>${escapeHtml(task.name || task.id)}${task.ownerName ? ` · ${escapeHtml(task.ownerName)}` : ''}</option>`)
    .join('');
  return `
    <div class="print-dialog-layer">
      <button class="print-dialog-backdrop" type="button" data-action="close-print-sheets" aria-label="Close"></button>
      <section class="print-dialog" role="dialog" aria-modal="true" aria-label="Print sheets for site">
        <div class="observation-drawer-head">
          <div>
            <p class="section-kicker">Site visit</p>
            <h3 aria-level="2">Print sheets for site</h3>
            <p class="muted">${escapeHtml(project.projectNo || 'No project number')} · ${escapeHtml(project.clientName || 'No client')}. Instrument, serial number and readings stay blank; pit and point tags from earlier reports of this project are pre-printed.</p>
          </div>
          <button class="icon-button" type="button" data-action="close-print-sheets" aria-label="Close"><span class="icon-button-symbol">&times;</span></button>
        </div>
        <div class="print-dialog-body">
          <label class="field"><span>Work date</span><input class="text-input" type="date" data-print-field="workDate" value="${escapeHtml(printState.workDate)}" /></label>
          <fieldset class="print-sheet-list">
            <legend>Tests to print</legend>
            ${sheets
              .map((sheet) => `<label class="print-sheet-option"><input type="checkbox" data-print-sheet="${escapeHtml(sheet.id)}" ${printState.selected[sheet.id] ? 'checked' : ''} /><span><b>${escapeHtml(sheet.code)}</b> ${escapeHtml(sheet.title)}</span></label>`)
              .join('')}
          </fieldset>
          <label class="field">
            <span>Zoho task (its owner is printed as the engineer)</span>
            <select class="text-input" data-print-field="taskId" ${project.zohoProjectId ? '' : 'disabled'}>
              <option value="">${project.zohoProjectId ? (printState.loadingTasks ? 'Loading tasks...' : 'No task - use the name below') : 'Link a Zoho project in step 1 to pick a task'}</option>
              ${taskOptions}
            </select>
          </label>
          <label class="field"><span>Engineer name (used when no task is chosen)</span><input class="text-input" type="text" data-print-field="engineerName" value="${escapeHtml(printState.engineerName)}" placeholder="Engineer name" /></label>
          ${printState.error ? `<p class="form-error" role="alert">${escapeHtml(printState.error)}</p>` : ''}
          <div class="button-row">
            <button class="button button-secondary" type="button" data-action="close-print-sheets">Cancel</button>
            <button class="button button-primary" type="button" data-action="download-print-sheets" ${printState.busy ? 'disabled' : ''}>${printState.busy ? 'Preparing...' : 'Download PDF'}</button>
          </div>
        </div>
      </section>
    </div>`;
}

async function openPrintSheets(currentSheetId = '', workspaceProject = null) {
  printState.project = workspaceProject;
  const project = printProject();
  const selected = {};
  const tests = workspaceProject ? workspaceProject.tests || {} : state.draft.tests;
  printableSheets().forEach((sheet) => {
    if (tests[sheet.id] || sheet.id === currentSheetId) {
      selected[sheet.id] = true;
    }
  });
  Object.assign(printState, {
    open: true,
    workDate: project.reportDate || new Date().toISOString().slice(0, 10),
    selected,
    taskId: '',
    engineerName: project.engineerName || '',
    tasks: [],
    loadingTasks: Boolean(project.zohoProjectId),
    busy: false,
    error: ''
  });
  render();
  if (project.zohoProjectId) {
    try {
      printState.tasks = await api(`/api/zoho/projects/${encodeURIComponent(project.zohoProjectId)}/tasks`);
    } catch (_error) {
      printState.tasks = [];
    }
    printState.loadingTasks = false;
    if (printState.open) {
      render();
    }
  }
}

async function downloadPrintSheets() {
  const project = printProject();
  const sheets = Object.keys(printState.selected).filter((id) => printState.selected[id]);
  printState.error = '';
  if (!project.projectNo) {
    printState.error = 'Enter the project number in step 1 first.';
  } else if (!printState.workDate) {
    printState.error = 'Choose the work date.';
  } else if (!sheets.length) {
    printState.error = 'Tick at least one test to print.';
  }
  if (printState.error) {
    render();
    return;
  }
  printState.busy = true;
  render();
  try {
    const response = await fetch('/api/sheets/print', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        project: {
          projectNo: project.projectNo,
          clientName: project.clientName,
          siteLocation: project.siteLocation,
          workOrder: project.workOrder,
          zohoProjectId: project.zohoProjectId
        },
        workDate: printState.workDate,
        sheets,
        taskId: printState.taskId,
        engineerName: printState.engineerName
      })
    });
    if (!response.ok) {
      const data = await response.json().catch(() => ({ message: 'Could not generate the sheets.' }));
      throw new Error(data.message || 'Could not generate the sheets.');
    }
    const pageCount = (response.headers.get('X-Sheet-Ids') || '').split(',').filter(Boolean).length;
    const warnings = decodeURIComponent(response.headers.get('X-Print-Warnings') || '');
    const url = URL.createObjectURL(await response.blob());
    const link = document.createElement('a');
    link.href = url;
    link.download = `ESHA_Sheets_${project.projectNo.replace(/[^A-Za-z0-9_-]+/g, '_')}_${printState.workDate}.pdf`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 60000);
    printState.open = false;
    printState.busy = false;
    showToast(`${pageCount} sheet page${pageCount === 1 ? '' : 's'} ready.${warnings ? ` ${warnings}` : ''}`, warnings ? 'warning' : 'healthy');
  } catch (error) {
    printState.busy = false;
    printState.error = error.message;
    render();
  }
}

// Returns true when the click was one of ours.
function handlePrintAction(action, button) {
  if (action === 'open-print-sheets') {
    openPrintSheets(button.dataset.sheet || '');
    return true;
  }
  if (action === 'close-print-sheets') {
    printState.open = false;
    render();
    return true;
  }
  if (action === 'download-print-sheets') {
    downloadPrintSheets();
    return true;
  }
  return false;
}

// Field changes in the dialog are kept in state without re-rendering (typing stays smooth).
function handlePrintInput(target) {
  if (!target.dataset) {
    return false;
  }
  if (target.dataset.printSheet) {
    printState.selected[target.dataset.printSheet] = target.checked;
    return true;
  }
  if (target.dataset.printField) {
    printState[target.dataset.printField] = target.value;
    return true;
  }
  return false;
}
