// Scanned sheets (6I): upload printed, hand-filled sheets (a PDF or photos), see what was read next to the scan, correct it,
// and Accept it into the project's daily report. Also the "Suggest wording" helper on flagged entry-sheet rows. Everything
// that decides anything (what is flagged, what is written, who may) is on the server; this file shows it and sends choices.

const scanUi = { files: [], scans: [], loading: false, uploading: false, message: '', detail: null, zoom: 1, assignFor: null, assign: { project: '', workDate: '', sheetKey: '', start: '1' }, confirm: null, errors: null, busy: false, timer: null, purge: { from: '', to: '' }, usage: null, purgeMessage: '' };
const suggestUi = {}; // rowId -> { busy, text, source, note }

const SCAN_STATUS = { queued: ['neutral', 'Waiting'], reading: ['warning', 'Reading'], read: ['healthy', 'Ready to review'], needs_sheet: ['warning', 'Which sheet is this?'], failed: ['critical', 'Could not be read'], accepted: ['healthy', 'Accepted'], discarded: ['neutral', 'Discarded'] };
const FLAG_CLASS = 'scan-flag';

async function scanFetch(path, options = {}) {
  const response = await fetch(path, { headers: options.body && !(options.body instanceof FormData) ? { 'Content-Type': 'application/json' } : {}, ...options });
  if (response.status === 401) {
    authUi.user = null;
    state.view = 'signin';
    render();
  }
  const data = await response.json().catch(() => ({}));
  return { status: response.status, ok: response.ok, data };
}

// ---- the list ---------------------------------------------------------------------------------------
async function loadScans() {
  const result = await scanFetch('/api/scans');
  if (result.ok) scanUi.scans = result.data.scans;
  const busy = scanUi.scans.some((scan) => ['queued', 'reading'].includes(scan.status));
  window.clearTimeout(scanUi.timer);
  if (busy && ['scans', 'scan-review'].includes(state.view)) scanUi.timer = window.setTimeout(() => loadScans().then(() => (state.view === 'scans' ? render() : null)), 1500);
  return result;
}

function openScans() {
  state.view = 'scans';
  scanUi.message = '';
  render();
  loadScans().then(render);
}

function printableSheets() {
  return allSheets().filter((sheet) => !sheet.retired && sheet.code);
}

function scanRow(scan) {
  const [tone, label] = SCAN_STATUS[scan.status] || ['neutral', scan.status];
  const assign = scanUi.assignFor === scan.id;
  const projects = (workspaceUi && workspaceUi.projects) || [];
  return `
    <li class="kb-hit scan-item" data-scan="${escapeHtml(scan.id)}" data-scan-status="${escapeHtml(scan.status)}">
      <div>${pill(tone, label)} <strong>${escapeHtml(scan.fileName || 'scan')}</strong> <span class="muted">· page ${scan.pageIndex + 1}${scan.sheetCode ? ` · ${escapeHtml(scan.sheetCode)}` : ''}${scan.projectNo ? ` · ${escapeHtml(scan.projectNo)}` : ''}${scan.workDate ? ` · ${escapeHtml(scan.workDate)}` : ''}</span>${scan.status === 'read' && scan.flagged ? ` <span class="soft-chip">${scan.flagged} to check</span>` : ''}</div>
      ${scan.note ? `<div class="muted">${escapeHtml(scan.note)}</div>` : ''}
      ${scan.error ? `<div class="muted" data-scan-error>${escapeHtml(scan.error.message)}</div>` : ''}
      <div class="button-row">
        ${scan.status === 'read' || scan.status === 'accepted' ? `<button class="button button-secondary" data-action="scan-open" data-id="${escapeHtml(scan.id)}">${scan.status === 'read' ? 'Review' : 'View'}</button>` : ''}
        ${scan.status === 'failed' ? `<button class="button button-secondary" data-action="scan-retry" data-id="${escapeHtml(scan.id)}">Try again</button>` : ''}
        ${['needs_sheet'].includes(scan.status) ? `<button class="button button-secondary" data-action="scan-assign-open" data-id="${escapeHtml(scan.id)}">Choose the sheet</button>` : ''}
        ${['read', 'failed', 'needs_sheet', 'queued'].includes(scan.status) ? `<button class="button button-ghost" data-action="scan-discard" data-id="${escapeHtml(scan.id)}">Discard</button>` : ''}
      </div>
      ${assign ? `
        <div class="field-grid" data-scan-assign>
          <label class="field"><span>Project</span><select class="text-input" data-scan-assign-field="project"><option value="">Choose...</option>${projects.map((project) => `<option value="${escapeHtml(project.id)}" ${scanUi.assign.project === project.id ? 'selected' : ''}>${escapeHtml(project.projectNo)} · ${escapeHtml(project.clientName)}</option>`).join('')}</select></label>
          <label class="field"><span>Work date</span><input class="text-input" type="date" data-scan-assign-field="workDate" value="${escapeHtml(scanUi.assign.workDate)}" /></label>
          <label class="field"><span>Sheet</span><select class="text-input" data-scan-assign-field="sheetKey"><option value="">Choose...</option>${printableSheets().map((sheet) => `<option value="${escapeHtml(sheet.id)}" ${scanUi.assign.sheetKey === sheet.id ? 'selected' : ''}>${escapeHtml(sheet.code)} ${escapeHtml(sheet.title)}</option>`).join('')}</select></label>
          <label class="field"><span>${scanUi.assign.sheetKey === 'soilResistivity' ? 'Page number of this location sheet' : 'Sr. number printed on the first row of this page'}</span><input class="text-input" type="number" min="1" data-scan-assign-field="start" value="${escapeHtml(scanUi.assign.start)}" /></label>
        </div>
        <div class="button-row"><button class="button button-primary" data-action="scan-assign" data-id="${escapeHtml(scan.id)}" ${scanUi.assign.project && scanUi.assign.workDate && scanUi.assign.sheetKey ? '' : 'disabled'}>Read it as this sheet</button><button class="button button-ghost" data-action="scan-assign-close">Cancel</button></div>` : ''}
    </li>`;
}

function scansView() {
  return `
    <section class="surface section-card kb-card" data-scans>
      <p class="section-kicker">Scanned sheets</p>
      <h1>Upload filled sheets</h1>
      <p class="muted">Upload a PDF or photos (PNG, JPEG) of the printed measurement sheets. Each page is matched to its issued sheet by the QR code and read; you check it next to the scan, then accept it into that project's daily report. Nothing reaches a report until you accept.</p>
      <div class="field-grid"><label class="field"><span>PDF or photos</span><input class="text-input" type="file" multiple accept="application/pdf,image/png,image/jpeg" data-scan-files /></label></div>
      ${scanUi.files.length ? `<p class="muted" data-scan-chosen>Chosen: ${scanUi.files.map((file) => escapeHtml(file.name)).join(', ')}</p>` : ''}
      <div class="button-row">
        <button class="button button-primary" data-action="scan-upload" ${scanUi.uploading ? 'disabled' : ''}>${scanUi.uploading ? 'Uploading...' : 'Upload and read'}</button>
        <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
      </div>
      ${scanUi.message ? `<p class="review-message" role="status" data-scan-message>${escapeHtml(scanUi.message)}</p>` : ''}
      ${scanUi.scans.length ? `<ul class="kb-hits" data-scan-list>${scanUi.scans.filter((scan) => scan.status !== 'discarded').map(scanRow).join('')}</ul>` : '<p class="muted" data-scan-list>No scans yet.</p>'}
    </section>`;
}

// ---- the review screen ---------------------------------------------------------------------------------
function cellEditor(def, cell) {
  if (def.preprinted) return `<span class="scan-pre">${escapeHtml(def.value)}</span>`;
  const path = escapeHtml(def.path);
  if (def.kind === 'ticks') {
    return `<div class="scan-ticks">${def.options.map((option) => `<label><input type="checkbox" data-scan-tick="${path}" data-value="${escapeHtml(option.value)}" ${cell.selected.includes(String(option.value)) ? 'checked' : ''} /> ${escapeHtml(option.label)}</label>`).join('')}</div>`;
  }
  if (def.kind === 'tick') return `<label><input type="checkbox" data-scan-tick="${path}" data-value="true" ${cell.value === 'true' ? 'checked' : ''} /> yes</label>`;
  return `<input class="table-input" type="text" ${def.kind === 'number' ? 'inputmode="decimal"' : ''} autocomplete="off" data-scan-cell="${path}" value="${escapeHtml(cell.value)}" aria-label="${escapeHtml(def.label)}" />`;
}

function cellBox(def, cells, { tag = 'td', attrs = '' } = {}) {
  const cell = cells[def.path] || { value: '', selected: [], flags: [] };
  const flags = (cell.flags || []).map((flag) => flag.text);
  const classes = [flags.length ? FLAG_CLASS : '', cell.edited ? 'scan-edited' : ''].filter(Boolean).join(' ');
  const hint = [...flags, cell.alt ? `The other read had: ${cell.alt.value !== undefined ? cell.alt.value || '(blank)' : (cell.alt.selected || []).join(', ') || '(none)'}` : '', cell.was ? `Read as: ${cell.was.value || (cell.was.selected || []).join(', ') || '(blank)'}` : ''].filter(Boolean).join(' | ');
  return `<${tag} class="${classes}" ${attrs} data-cell-box="${escapeHtml(def.path)}" ${hint ? `title="${escapeHtml(hint)}"` : ''}>${cellEditor(def, cell)}${flags.length ? `<div class="scan-flag-text">${escapeHtml(flags.join(' '))}</div>` : ''}</${tag}>`;
}

function scanTable(detail) {
  const form = detail.form;
  const byPath = Object.fromEntries(form.cells.map((def) => [def.path, def]));
  const itemOffset = Math.floor(form.rowOffset / form.group);
  const columns = form.columns.filter((col) => col.k !== 'sr' || true);
  const head = `<tr>${columns.map((col) => `<th scope="col">${escapeHtml(col.label)}</th>`).join('')}</tr>`;
  const rows = [];
  for (let r = 0; r < form.rowCount; r += 1) {
    const first = r % form.group === 0;
    const item = Math.floor(r / form.group);
    const cellsHtml = columns.map((col) => {
      if (col.k === 'sr') return col.merge ? (first ? `<td rowspan="${form.group}" class="scan-fixed">${itemOffset + item + 1}</td>` : '') : `<td class="scan-fixed">${form.rowOffset + r + 1}</td>`;
      if (col.k === 'fixed') return `<td class="scan-fixed">${escapeHtml(form.layout === 'soilMatrix' ? form.spacings[r] : (col.values || [])[r % Math.max(1, (col.values || []).length)] || '')}</td>`;
      if (col.merge && !first) return '';
      const def = byPath[`r${r}.${col.key}`];
      if (!def) return '<td></td>';
      return cellBox(def, detail.cells, { attrs: col.merge ? `rowspan="${form.group}"` : '' });
    });
    rows.push(`<tr class="sheet-row">${cellsHtml.join('')}</tr>`);
  }
  return `<div class="scan-table-wrap"><table class="data-table scan-table" data-scan-table><thead>${head}</thead><tbody>${rows.join('')}</tbody></table></div>`;
}

function scanReviewView() {
  const detail = scanUi.detail;
  if (!detail || !detail.form) return '<section class="surface empty-state"><p>Loading...</p></section>';
  const flagged = Object.values(detail.cells).filter((cell) => cell.flags && cell.flags.length).length;
  const headerCells = detail.form.cells.filter((def) => def.section === 'header');
  const accepted = detail.status === 'accepted';
  const editable = detail.status === 'read' && detail.canChange;
  return `
    <section class="surface detail-shell review-shell scan-review" data-scan-review="${escapeHtml(detail.id)}">
      <div class="detail-hero">
        <div class="detail-hero-copy">
          <p class="section-kicker">${escapeHtml(detail.form.code)} · ${escapeHtml(detail.form.title)}</p>
          <h1>${escapeHtml(detail.clientName || '')}</h1>
          <p>${escapeHtml(detail.projectNo)} · ${escapeHtml(detail.workDate)} · rows from ${detail.form.rowOffset + 1}${detail.identified ? ` · found by ${escapeHtml(detail.identified)}` : ''}</p>
          <div class="detail-hero-tags">${pill(accepted ? 'healthy' : flagged ? 'warning' : 'healthy', accepted ? 'Accepted' : `${flagged} cell${flagged === 1 ? '' : 's'} to check`)}${detail.escalated ? '<span class="soft-chip">second read used</span>' : ''}</div>
        </div>
        <div class="button-row detail-hero-actions">
          <button class="button button-secondary" data-action="scan-back">Back to scans</button>
          ${editable ? '<button class="button button-primary" data-action="scan-accept">Accept into the daily report</button>' : ''}
          ${editable ? '<button class="button button-ghost" data-action="scan-discard-current">Discard</button>' : ''}
        </div>
      </div>
      <p class="muted">Highlighted cells are ones the reader was unsure of, that are not a number where one is needed, are out of range, have two boxes ticked, or where two reads disagreed (hover for the reason). Everything can be edited. Blank cells do not change the report. Confidence below ${Math.round(detail.threshold * 100)}% is flagged.</p>
      ${scanUi.errors ? `<div class="review-alert" role="alert" data-scan-errors><strong>${escapeHtml(scanUi.errors.message)}</strong>${scanUi.errors.list && scanUi.errors.list.length ? `<ul>${scanUi.errors.list.map((message) => `<li>${escapeHtml(message)}</li>`).join('')}</ul>` : ''}</div>` : ''}
      ${scanUi.confirm ? `<div class="review-warnings" data-scan-confirm><strong>${scanUi.confirm.overwrites.length} existing value${scanUi.confirm.overwrites.length === 1 ? '' : 's'} in the daily report will be overwritten.</strong> The previous values stay in this scan's record.<ul>${scanUi.confirm.overwrites.map((entry) => `<li>${escapeHtml(entry.where)} · ${escapeHtml(entry.label)}: ${escapeHtml(Array.isArray(entry.before) ? entry.before.join(', ') : String(entry.before))} → ${escapeHtml(Array.isArray(entry.after) ? entry.after.join(', ') : String(entry.after))}</li>`).join('')}</ul>
        <div class="button-row"><button class="button button-primary" data-action="scan-accept-confirm" ${scanUi.busy ? 'disabled' : ''}>Overwrite and accept</button><button class="button button-ghost" data-action="scan-accept-cancel">Cancel</button></div></div>` : ''}
      <div class="scan-split">
        <div class="scan-image-pane">
          <div class="button-row"><button class="button button-ghost" data-action="scan-zoom" data-step="-0.25">−</button><span class="muted">${Math.round(scanUi.zoom * 100)}%</span><button class="button button-ghost" data-action="scan-zoom" data-step="0.25">+</button><button class="button button-ghost" data-action="scan-zoom" data-step="0">Fit</button></div>
          <div class="scan-image-scroll"><img data-scan-image src="/api/scans/${encodeURIComponent(detail.id)}/image" alt="The scanned sheet" style="width:${Math.round(scanUi.zoom * 100)}%" /></div>
        </div>
        <div class="scan-data-pane">
          <fieldset class="scan-header" ${editable ? '' : 'disabled'}><legend>Header</legend><div class="scan-header-grid">${headerCells.map((def) => `<div class="scan-header-item"><span class="muted">${escapeHtml(def.label)}</span>${cellBox(def, detail.cells, { tag: 'div' })}</div>`).join('')}</div></fieldset>
          <fieldset ${editable ? '' : 'disabled'}>${scanTable(detail)}</fieldset>
        </div>
      </div>
    </section>`;
}

async function openScanReview(id) {
  const result = await scanFetch(`/api/scans/${encodeURIComponent(id)}`);
  if (!result.ok) {
    showToast(result.data.message || 'Could not open the scan.', 'critical');
    return;
  }
  scanUi.detail = result.data;
  scanUi.zoom = 1;
  scanUi.confirm = null;
  scanUi.errors = null;
  state.view = 'scan-review';
  render();
}

// Edits go to the server one at a time, in order; the answer replaces the cells (the checks are run there).
let scanEditChain = Promise.resolve();
function saveScanEdit(path, edit) {
  const id = scanUi.detail.id;
  scanEditChain = scanEditChain.then(async () => {
    const result = await scanFetch(`/api/scans/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify({ edits: { [path]: edit } }) });
    if (result.ok && scanUi.detail && scanUi.detail.id === id) {
      scanUi.detail = result.data;
      scanUi.confirm = null;
      render();
    } else if (!result.ok) {
      showToast(result.data.message || 'The edit was not saved.', 'critical');
    }
  });
  return scanEditChain;
}

async function acceptScan(confirm) {
  scanUi.busy = true;
  scanUi.errors = null;
  render();
  const result = await scanFetch(`/api/scans/${encodeURIComponent(scanUi.detail.id)}/accept`, { method: 'POST', body: JSON.stringify({ confirm }) });
  scanUi.busy = false;
  if (result.ok) {
    scanUi.confirm = null;
    showToast(`Accepted: ${result.data.written} item${result.data.written === 1 ? '' : 's'} written${result.data.createdDay ? ' (a new daily report was started)' : ''}.`, 'healthy');
    await loadProjects();
    await openWorkspace(result.data.projectId);
    return;
  }
  if (result.data.code === 'NEEDS_CONFIRM') scanUi.confirm = { overwrites: result.data.overwrites };
  else scanUi.errors = { message: result.data.message || 'Could not accept the scan.', list: (result.data.errors || []).map((entry) => entry.message || String(entry)) };
  render();
}

// ---- "Suggest wording" on flagged entry-sheet rows ------------------------------------------------------
// Wording for a flagged point, from its condition codes only; the engineer clicks Insert. The text is never a number or a verdict.
function suggestionHtml(rowId) {
  const entry = suggestUi[rowId];
  if (!entry) return '';
  if (entry.busy) return '<div class="muted" data-suggestion>Drafting wording...</div>';
  return `<div class="auto-eval-text" data-suggestion="${escapeHtml(rowId)}"><span data-suggestion-text>${escapeHtml(entry.text)}</span>
    <div class="muted">${entry.source === 'ai' ? 'Drafted by the assistant.' : 'Standard wording.'}${entry.note ? ` ${escapeHtml(entry.note)}` : ''}</div>
    <button class="button button-secondary" type="button" data-action="suggest-insert" data-row-id="${escapeHtml(rowId)}">Insert</button> <button class="button button-ghost" type="button" data-action="suggest-dismiss" data-row-id="${escapeHtml(rowId)}">Dismiss</button></div>`;
}

function suggestButton(sheetId, row, assessment) {
  const flagged = assessment && ((assessment.status && ['warning', 'critical'].includes(assessment.status.tone)) || ['P1', 'P2'].includes(assessment.priority));
  if (!flagged || !row || !row.rowId) return '';
  return `<button class="suggest-btn" type="button" data-action="suggest-wording" data-section="${escapeHtml(sheetId)}" data-row-id="${escapeHtml(row.rowId)}" title="Suggest remark wording" aria-label="Suggest remark wording">✎</button>`;
}

function findSheetRow(sheetId, rowId) {
  const rows = state.draft && state.draft[sheetId];
  return Array.isArray(rows) ? rows.find((row) => row.rowId === rowId) : null;
}

const REMARK_FIELD_KEYS = ['remark', 'remarks', 'comment', 'observation'];

// ---- events ---------------------------------------------------------------------------------------------
function handleScanAction(action, button) {
  const id = button && button.dataset && button.dataset.id;
  if (action === 'open-scans') { openScans(); return true; }
  if (action === 'scan-upload') {
    const files = scanUi.files;
    if (!files.length) { scanUi.message = 'Choose a PDF or photos first.'; render(); return true; }
    const form = new FormData();
    files.forEach((file) => form.append('files', file));
    scanUi.uploading = true;
    scanUi.message = '';
    render();
    scanFetch('/api/scans', { method: 'POST', body: form }).then(async (result) => {
      scanUi.uploading = false;
      if (result.ok) scanUi.files = [];
      scanUi.message = result.ok ? `${result.data.scans.length} page${result.data.scans.length === 1 ? '' : 's'} uploaded. Reading...` : result.data.message || 'The upload failed.';
      await loadScans();
      render();
    });
    return true;
  }
  if (action === 'scan-open') { openScanReview(id); return true; }
  if (action === 'scan-back') { window.clearTimeout(scanUi.timer); openScans(); return true; }
  if (action === 'scan-retry') { scanFetch(`/api/scans/${encodeURIComponent(id)}/retry`, { method: 'POST', body: '{}' }).then(() => loadScans()).then(render); return true; }
  if (action === 'scan-discard' || action === 'scan-discard-current') {
    const target = action === 'scan-discard' ? id : scanUi.detail.id;
    scanFetch(`/api/scans/${encodeURIComponent(target)}/discard`, { method: 'POST', body: '{}' }).then(() => { if (action === 'scan-discard-current') state.view = 'scans'; return loadScans(); }).then(render);
    return true;
  }
  if (action === 'scan-assign-open') { scanUi.assignFor = id; scanUi.assign = { project: '', workDate: '', sheetKey: '', start: '1' }; render(); return true; }
  if (action === 'scan-assign-close') { scanUi.assignFor = null; render(); return true; }
  if (action === 'scan-assign') {
    const { project, workDate, sheetKey, start } = scanUi.assign;
    const rowOffset = Math.max(0, (Number.parseInt(start, 10) || 1) - 1);
    scanFetch(`/api/scans/${encodeURIComponent(id)}/assign`, { method: 'POST', body: JSON.stringify({ project, workDate, sheetKey, rowOffset }) }).then(async (result) => {
      if (!result.ok) showToast(result.data.message || 'Could not use that choice.', 'critical');
      scanUi.assignFor = null;
      await loadScans();
      render();
    });
    return true;
  }
  if (action === 'scan-zoom') { scanUi.zoom = Number(button.dataset.step) === 0 ? 1 : Math.min(4, Math.max(0.5, scanUi.zoom + Number(button.dataset.step))); render(); return true; }
  if (action === 'scan-accept') { acceptScan(false); return true; }
  if (action === 'scan-accept-confirm') { acceptScan(true); return true; }
  if (action === 'scan-accept-cancel') { scanUi.confirm = null; render(); return true; }
  if (action === 'suggest-wording') {
    const rowId = button.dataset.rowId;
    const sheet = sheetById(button.dataset.section);
    const row = findSheetRow(button.dataset.section, rowId);
    if (!sheet || !row) return true;
    suggestUi[rowId] = { busy: true };
    render();
    scanFetch('/api/observations/suggest', { method: 'POST', body: JSON.stringify({ sheetKey: sheet.id, pointKind: sheet.title, conditionCodes: Array.isArray(row.conditionCodes) ? row.conditionCodes : [] }) }).then((result) => {
      suggestUi[rowId] = result.ok ? { text: result.data.text, source: result.data.source, note: result.data.status && !result.data.status.ok ? 'The assistant is unavailable right now.' : result.data.status && result.data.status.note || '' } : undefined;
      if (!result.ok) showToast(result.data.message || 'No wording could be drafted.', 'warning');
      render();
    });
    return true;
  }
  if (action === 'suggest-insert') {
    const rowId = button.dataset.rowId;
    const entry = suggestUi[rowId];
    const sheetId = Object.keys(state.draft || {}).find((key) => Array.isArray(state.draft[key]) && state.draft[key].some((row) => row && row.rowId === rowId));
    const row = sheetId && findSheetRow(sheetId, rowId);
    const sheet = sheetId && sheetById(sheetId);
    const field = sheet && sheet.rowFields.map((entry_) => entry_.key).find((key) => REMARK_FIELD_KEYS.includes(key));
    if (entry && row && field) {
      row[field] = String(row[field] || '').trim() ? `${String(row[field]).trim()} ${entry.text}` : entry.text;
    }
    delete suggestUi[rowId];
    render();
    return true;
  }
  if (action === 'suggest-dismiss') { delete suggestUi[button.dataset.rowId]; render(); return true; }
  return false;
}

function handleScanInput(target) {
  const data = target.dataset || {};
  if (data.scanAssignField) { scanUi.assign[data.scanAssignField] = target.value; return true; }
  return false;
}

function handleScanChange(target) {
  const data = target.dataset || {};
  if (data.scanFiles !== undefined) { scanUi.files = [...(target.files || [])]; render(); return true; }
  if (data.scanAssignField) {
    scanUi.assign[data.scanAssignField] = target.value;
    // a select changes the page's wording; typing in a field must not re-draw it (a re-draw on blur eats the click on the button)
    if (target.tagName === 'SELECT') render();
    else {
      const go = document.querySelector('[data-action="scan-assign"]');
      if (go) go.disabled = !(scanUi.assign.project && scanUi.assign.workDate && scanUi.assign.sheetKey);
    }
    return true;
  }
  if (data.scanCell) { saveScanEdit(data.scanCell, { value: target.value }); return true; }
  if (data.scanTick) {
    const path = data.scanTick;
    const def = scanUi.detail.form.cells.find((cell) => cell.path === path);
    const boxes = [...document.querySelectorAll(`[data-scan-tick="${CSS.escape(path)}"]`)];
    const selected = boxes.filter((box) => box.checked).map((box) => box.dataset.value);
    saveScanEdit(path, def.kind === 'tick' ? { value: selected.length ? 'true' : '' } : { selected });
    return true;
  }
  return false;
}

// ---- the admin panel: purge old scans, AI usage ---------------------------------------------------------
function scanAdminPanel() {
  return `
    <p class="muted">Scans are kept indefinitely (the image in data/scans, the reading and what an accept overwrote in the database). An admin can delete a date range here; it is recorded in the audit log.</p>
    <div class="field-grid">
      <label class="field"><span>Delete scans uploaded from</span><input class="text-input" type="date" data-scan-purge-field="from" value="${escapeHtml(scanUi.purge.from)}" /></label>
      <label class="field"><span>to</span><input class="text-input" type="date" data-scan-purge-field="to" value="${escapeHtml(scanUi.purge.to)}" /></label>
    </div>
    <div class="button-row"><button class="button button-secondary" data-action="scan-purge" ${scanUi.purge.from && scanUi.purge.to ? '' : 'disabled'}>Delete these scans</button></div>
    ${scanUi.purgeMessage ? `<p class="review-message" role="status" data-scan-purge-message>${escapeHtml(scanUi.purgeMessage)}</p>` : ''}
    <h3 aria-level="2">Reader and AI-check calls</h3>
    ${scanUi.usage && scanUi.usage.length ? `<table class="data-table" data-ai-usage><thead><tr><th>Provider</th><th>Kind</th><th>Calls</th><th>Succeeded</th></tr></thead><tbody>${scanUi.usage.map((row) => `<tr><td>${escapeHtml(row.provider)}</td><td>${escapeHtml(row.kind)}</td><td>${row.calls}</td><td>${row.ok}</td></tr>`).join('')}</tbody></table>` : '<p class="muted" data-ai-usage>No provider calls yet.</p>'}`;
}

async function loadScanAdmin() {
  const result = await scanFetch('/api/admin/ai-usage');
  scanUi.usage = result.ok ? result.data.usage : [];
}

function handleScanAdminAction(action) {
  if (action === 'scan-purge') {
    scanUi.purgeMessage = '';
    scanFetch('/api/admin/scans/purge', { method: 'POST', body: JSON.stringify(scanUi.purge) }).then((result) => {
      scanUi.purgeMessage = result.ok ? `${result.data.removed} scan${result.data.removed === 1 ? '' : 's'} deleted.` : result.data.message || 'Nothing was deleted.';
      render();
    });
    return true;
  }
  return false;
}

document.addEventListener('input', (event) => {
  const data = event.target.dataset || {};
  if (data.scanPurgeField) {
    scanUi.purge[data.scanPurgeField] = event.target.value;
    const button = document.querySelector('[data-action="scan-purge"]');
    if (button) button.disabled = !(scanUi.purge.from && scanUi.purge.to);
  }
});
