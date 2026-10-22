// Microsoft Graph screens (6H): the knowledge-base search, the references a reviewer picks for a final report,
// the admin panels for the Outlook calendar sync and the SharePoint index, and the /?project=<id> deep link
// (the link the calendar events carry). Everything about who may see what is decided on the server.

const kbUi = { query: '', hits: null, searched: '', busy: false, ref: { query: '', hits: null, busy: false }, calendar: null, kb: null, preview: null, message: '' };
const DEEP_LINK_KEY = 'er_deeplink';

// ---- deep link ------------------------------------------------------------------------------------
// The link may be opened before sign-in: the id is kept for the tab and used once the app is ready.
(function captureDeepLink() {
  try {
    const id = new URLSearchParams(window.location.search).get('project');
    if (id) {
      sessionStorage.setItem(DEEP_LINK_KEY, id);
      window.history.replaceState(null, '', window.location.pathname);
    }
  } catch (_error) { /* storage unavailable: the link just opens the dashboard */ }
})();

async function consumeDeepLink() {
  let id = '';
  try {
    id = sessionStorage.getItem(DEEP_LINK_KEY) || '';
    sessionStorage.removeItem(DEEP_LINK_KEY);
  } catch (_error) { id = ''; }
  if (!id) return;
  try {
    const found = await api(`/api/projects/lookup/${encodeURIComponent(id)}`);
    if (found.workspaceId) {
      await openWorkspace(found.workspaceId);
    } else if (found.zohoProjectId && state.zoho && Array.isArray(state.zoho.projects) && state.zoho.projects.some((entry) => String(entry.id) === String(found.zohoProjectId))) {
      await openWorkspaceFromZoho(found.zohoProjectId);
    } else {
      showToast('That project could not be found.', 'warning');
    }
  } catch (error) {
    showToast(error.message, 'critical');
  }
}

// ---- search ---------------------------------------------------------------------------------------
const dateOnly = (value) => String(value || '').slice(0, 10);

function hitHtml(hit, action = '') {
  return `<li class="kb-hit">
    <a href="${escapeHtml(hit.webUrl)}" target="_blank" rel="noopener noreferrer"><strong>${escapeHtml(hit.title)}</strong></a>
    <span class="muted"> · ${escapeHtml(hit.source)} · ${escapeHtml(dateOnly(hit.modified))}</span>
    <div class="muted">${escapeHtml(hit.path)}</div>
    ${hit.snippet ? `<div>${escapeHtml(hit.snippet)}</div>` : ''}
    ${action}
  </li>`;
}

function kbView() {
  return `
    <section class="surface section-card kb-card">
      <p class="section-kicker">Knowledge base</p>
      <h2>Search documents</h2>
      <p class="muted">Reads an index of the Submitted, Reference and Input documents of the project folders and of the technical knowledge hub. You see the general documents and your own projects' documents. Files open in SharePoint; nothing here changes them.</p>
      <div class="field-grid">
        <label class="field"><span>Search</span><input class="text-input" type="search" data-kb-field="query" value="${escapeHtml(kbUi.query)}" placeholder="e.g. pit resistance limit" /></label>
      </div>
      <div class="button-row">
        <button class="button button-primary" data-action="kb-search" ${kbUi.busy ? 'disabled' : ''}>Search</button>
        <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
      </div>
      ${kbUi.hits === null ? '' : kbUi.hits.length ? `<ul class="kb-hits" data-kb-hits>${kbUi.hits.map((hit) => hitHtml(hit)).join('')}</ul>` : `<p class="muted" data-kb-hits>No documents found for "${escapeHtml(kbUi.searched)}".</p>`}
    </section>`;
}

async function runKbSearch() {
  kbUi.busy = true;
  render();
  try {
    kbUi.searched = kbUi.query;
    kbUi.hits = (await api(`/api/kb/search?q=${encodeURIComponent(kbUi.query)}`)).hits;
  } catch (error) {
    showToast(error.message, 'critical');
  }
  kbUi.busy = false;
  render();
}

// ---- references on a final report (reviewer picks; the model never touches them) ---------------------
function reviewReferencesPanel(report, review) {
  if (review.kind !== 'final') return '';
  const references = report.references || [];
  const may = hasRoleClient('reviewer') && review.status !== 'approved';
  return `
    <section class="surface section-card review-block" data-references>
      <p class="section-kicker">References</p>
      <h3>Further reading</h3>
      <p class="muted">Documents listed on the report as references: title, location and date only. They are not used in any figure, result or finding.${may ? ' Only this project\'s own documents and the general knowledge base can be added.' : ''}</p>
      ${references.length ? `<ul class="kb-hits" data-reference-list>${references.map((reference) => `<li class="kb-hit"><strong>${escapeHtml(reference.title)}</strong> <span class="muted">· ${escapeHtml(reference.source)} · ${escapeHtml(dateOnly(reference.modified))}</span><div class="muted">${escapeHtml(reference.path)}</div>${may ? `<button class="button button-ghost" data-action="ref-remove" data-file-id="${reference.fileId}">Remove</button>` : ''}</li>`).join('')}</ul>` : '<p class="muted" data-reference-list>No references chosen.</p>'}
      ${may ? `
        <div class="field-grid"><label class="field"><span>Find a document</span><input class="text-input" type="search" data-ref-field="query" value="${escapeHtml(kbUi.ref.query)}" placeholder="search the knowledge base" /></label></div>
        <div class="button-row"><button class="button button-secondary" data-action="ref-search" ${kbUi.ref.busy ? 'disabled' : ''}>Search</button></div>
        ${kbUi.ref.hits === null ? '' : kbUi.ref.hits.length ? `<ul class="kb-hits" data-ref-hits>${kbUi.ref.hits.map((hit) => hitHtml(hit, references.some((entry) => entry.fileId === hit.id) ? '<span class="muted">Already listed</span>' : `<button class="button button-secondary" data-action="ref-add" data-file-id="${hit.id}">Add as reference</button>`)).join('')}</ul>` : '<p class="muted" data-ref-hits>No documents found.</p>'}` : ''}
    </section>`;
}

// The references screen of a final report (reached from the review screen's header, so the review page keeps its layout).
function referencesView() {
  const report = reviewUi.report;
  const review = reviewUi.review;
  if (!report || !review) return '<section class="surface empty-state"><p>Loading...</p></section>';
  return `
    <section class="surface detail-shell review-shell">
      <div class="detail-hero">
        <div class="detail-hero-copy">
          <p class="section-kicker">Final report · References</p>
          <h2>${escapeHtml(report.project.clientName || '')}</h2>
          <p>${escapeHtml(report.project.projectNo || '')}</p>
        </div>
        <div class="button-row detail-hero-actions"><button class="button button-secondary" data-action="ref-back">Back to review</button></div>
      </div>
      ${reviewReferencesPanel(report, review)}
    </section>`;
}

async function saveReferences(fileIds) {
  try {
    await api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/references`, { method: 'PUT', body: JSON.stringify({ fileIds }) });
    await refreshReview();
  } catch (error) {
    showToast(error.message, 'critical');
  }
  render();
}

// ---- admin panels ----------------------------------------------------------------------------------
const ACTION_LABEL = { create: 'Create', update: 'Update', delete: 'Delete', skip: 'Skipped', forget: 'Forget' };

function calendarSyncPanel() {
  const info = kbUi.calendar;
  if (!info) return '<p class="muted">Loading...</p>';
  const summary = kbUi.preview || info.last;
  const counts = summary ? summary.counts : null;
  return `
    <p class="muted">Each Onsite ESHA task with dates and an owner becomes an all-day event in the owner's Outlook calendar (shown as "working elsewhere"). Only events this app created are ever changed or deleted. ${info.appBaseUrl ? '' : '<strong>APP_BASE_URL is not set, so the events carry no link yet.</strong>'}</p>
    <div class="field-grid">
      <label class="role-box"><input type="checkbox" data-calendar-setting="enabled" ${info.settings.enabled ? 'checked' : ''} /> Calendar sync is on</label>
      <label class="role-box"><input type="checkbox" data-calendar-setting="dryRun" ${info.settings.dryRun ? 'checked' : ''} /> Dry run (plan only, change nothing)</label>
    </div>
    <div class="button-row">
      <button class="button button-secondary" data-action="calendar-preview">Preview changes</button>
      <button class="button button-primary" data-action="calendar-run" ${info.settings.enabled ? '' : 'disabled'}>Run now</button>
    </div>
    ${kbUi.message ? `<p class="review-message" role="status" data-calendar-message>${escapeHtml(kbUi.message)}</p>` : ''}
    ${summary ? `<p class="muted" data-calendar-summary>${kbUi.preview ? 'Preview' : summary.dryRun ? 'Last run (dry run)' : 'Last run'} ${escapeHtml(formatStamp(summary.at))}: ${counts.create} to create, ${counts.update} to update, ${counts.delete} to delete, ${counts.unchanged} unchanged${counts.skip ? `, ${counts.skip} skipped` : ''}${counts.failed ? `, ${counts.failed} failed` : ''}.</p>` : '<p class="muted">Nothing has run yet.</p>'}
    ${summary && summary.warnings && summary.warnings.length ? `<ul class="muted" data-calendar-warnings>${summary.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join('')}</ul>` : ''}
    ${summary && summary.items && summary.items.length ? `<table class="data-table" data-calendar-items><thead><tr><th>Change</th><th>Owner</th><th>Project</th><th>Test</th><th>Dates</th></tr></thead><tbody>
      ${summary.items.map((item) => `<tr><td>${escapeHtml(ACTION_LABEL[item.action] || item.action)}${item.reason ? ` <span class="muted">(${escapeHtml(item.reason)})</span>` : ''}</td><td>${escapeHtml(item.ownerName || item.owner)}</td><td>${escapeHtml(item.project || '')}</td><td>${escapeHtml(item.test || '')}</td><td>${escapeHtml(item.from ? (item.from === item.to ? item.from : `${item.from} to ${item.to}`) : '')}</td></tr>`).join('')}
    </tbody></table>` : ''}`;
}

function kbAdminPanel() {
  const info = kbUi.kb;
  if (!info) return '<p class="muted">Loading...</p>';
  return `
    <p class="muted">A local index of the Submitted, Reference and Input documents (PDF, DOCX, XLSX) of the project folders and of the technical knowledge hub. SharePoint is only ever read.</p>
    <div class="button-row"><button class="button button-primary" data-action="kb-refresh" ${info.running ? 'disabled' : ''}>Refresh now</button></div>
    ${kbUi.message ? `<p class="review-message" role="status" data-kb-message>${escapeHtml(kbUi.message)}</p>` : ''}
    <table class="data-table" data-kb-sites><thead><tr><th>Site</th><th>Scope</th><th>Folders</th><th>Indexed</th><th>No text</th><th>Problems</th><th>Refreshed</th><th>Error</th></tr></thead><tbody>
      ${info.sites.map((site) => `<tr><td>${escapeHtml(site.name)}</td><td>${escapeHtml(site.scope)}</td><td>${escapeHtml(site.folders.join(', ') || 'whole library')}</td><td>${site.files}</td><td>${site.noText}</td><td>${site.problems}</td><td>${escapeHtml(site.refreshedAt ? formatStamp(site.refreshedAt) : 'never')}</td><td>${escapeHtml(site.error)}</td></tr>`).join('')}
    </tbody></table>`;
}

async function loadGraphAdmin(tab) {
  if (tab === 'calendar') kbUi.calendar = await api('/api/admin/calendar-sync');
  if (tab === 'kb') kbUi.kb = await api('/api/admin/kb');
}

// ---- events ------------------------------------------------------------------------------------------
function handleKbAction(action, button) {
  if (action === 'open-kb') {
    state.view = 'kb';
    render();
    return true;
  }
  if (action === 'open-references') {
    kbUi.ref = { query: '', hits: null, busy: false };
    state.view = 'references';
    render();
    return true;
  }
  if (action === 'ref-back') {
    openReview(reviewUi.report.id);
    return true;
  }
  if (action === 'kb-search') {
    runKbSearch();
    return true;
  }
  if (action === 'ref-search') {
    kbUi.ref.busy = true;
    render();
    api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/kb-search?q=${encodeURIComponent(kbUi.ref.query)}`)
      .then((result) => { kbUi.ref.hits = result.hits; })
      .catch((error) => showToast(error.message, 'critical'))
      .finally(() => { kbUi.ref.busy = false; render(); });
    return true;
  }
  if (action === 'ref-add' || action === 'ref-remove') {
    const current = (reviewUi.report.references || []).map((entry) => entry.fileId);
    const id = Number(button.dataset.fileId);
    saveReferences(action === 'ref-add' ? [...new Set([...current, id])] : current.filter((entry) => entry !== id));
    return true;
  }
  if (action === 'calendar-preview' || action === 'calendar-run') {
    const preview = action === 'calendar-preview';
    kbUi.message = preview ? 'Planning...' : 'Running...';
    render();
    api(`/api/admin/calendar-sync/${preview ? 'preview' : 'run'}`, { method: 'POST', body: '{}' })
      .then(async (summary) => {
        kbUi.preview = preview ? summary : null;
        kbUi.message = '';
        kbUi.calendar = await api('/api/admin/calendar-sync');
      })
      .catch((error) => { kbUi.message = ''; showToast(error.message, 'critical'); })
      .finally(() => render());
    return true;
  }
  if (action === 'kb-refresh') {
    kbUi.message = 'Refreshing... this can take a while on the first run.';
    render();
    api('/api/admin/kb/refresh', { method: 'POST', body: '{}' })
      .then(async () => { kbUi.message = 'Refresh done.'; kbUi.kb = await api('/api/admin/kb'); })
      .catch((error) => { kbUi.message = ''; showToast(error.message, 'critical'); })
      .finally(() => render());
    return true;
  }
  return false;
}

// Typing in the search boxes (no re-render per keystroke). Returns true when handled.
function handleKbInput(target) {
  const data = target.dataset || {};
  if (data.kbField) { kbUi[data.kbField] = target.value; return true; }
  if (data.refField) { kbUi.ref[data.refField] = target.value; return true; }
  return false;
}

function handleKbChange(target) {
  const data = target.dataset || {};
  if (data.calendarSetting) {
    api('/api/admin/calendar-sync', { method: 'PUT', body: JSON.stringify({ [data.calendarSetting]: target.checked }) })
      .then((info) => { kbUi.calendar = info; })
      .catch((error) => showToast(error.message, 'critical'))
      .finally(() => render());
    return true;
  }
  return false;
}

document.addEventListener('keydown', (event) => {
  if (event.key !== 'Enter' || !event.target.dataset) return;
  if (event.target.dataset.kbField) { event.preventDefault(); runKbSearch(); }
  if (event.target.dataset.refField) { event.preventDefault(); handleKbAction('ref-search', event.target); }
});
