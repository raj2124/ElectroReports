// Review & approval screen for daily and final reports (standard 6.4):
// draft -> in review -> marked as checked -> approved, with the reviewer checklist, the
// acknowledgement notes, the transitions log and re-issue. All rules live on the server
// (src/review.js, reportStore); this file only renders the state it returns and sends steps.

const reviewUi = {
  review: null,
  report: null,
  loading: false,
  busy: false,
  actor: '',
  note: '',
  acknowledgements: {},
  reissueReason: '',
  message: '',
  logoMessage: ''
};

try {
  reviewUi.actor = window.localStorage.getItem('electroreports.reviewer') || '';
} catch (_error) {
  reviewUi.actor = '';
}

const REVIEW_STATUS_LABEL = { draft: 'Draft', in_review: 'In review', approved: 'Approved' };
const REVIEW_STATUS_TONE = { draft: 'neutral', in_review: 'warning', approved: 'healthy' };
const AUTO_LABEL = { pass: ['healthy', 'Checked by the app'], warn: ['warning', 'Needs attention'], fail: ['critical', 'Failed'] };
const STEP_LABEL = { submit: 'Submitted for review (prepared)', check: 'Marked as checked', approve: 'Approved', return: 'Returned to draft' };

function isReviewable(report) {
  return Boolean(report && report.kind && report.kind !== 'standalone');
}

function formatStamp(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? String(value || '') : date.toLocaleString('en-GB', { dateStyle: 'medium', timeStyle: 'short' });
}

async function openReview(id) {
  reviewUi.loading = true;
  reviewUi.message = '';
  reviewUi.note = '';
  reviewUi.acknowledgements = {};
  reviewUi.reissueReason = '';
  try {
    reviewUi.report = await api(`/api/reports/${encodeURIComponent(id)}`);
    reviewUi.review = await api(`/api/reports/${encodeURIComponent(id)}/review`);
    state.view = 'review';
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    reviewUi.loading = false;
    render();
  }
}

async function refreshReview() {
  const id = reviewUi.report.id;
  reviewUi.report = await api(`/api/reports/${encodeURIComponent(id)}`);
  reviewUi.review = await api(`/api/reports/${encodeURIComponent(id)}/review`);
}

function renderChecklistItem(item, locked) {
  const auto = item.auto ? AUTO_LABEL[item.auto] : null;
  return `
    <li class="review-item">
      <label class="review-tick">
        <input type="checkbox" data-review-tick="${escapeHtml(item.id)}" ${item.ticked ? 'checked' : ''} ${locked ? 'disabled' : ''} />
        <span>${escapeHtml(item.text)}</span>
      </label>
      <div class="review-item-meta">
        ${auto ? pill(auto[0], auto[1]) : '<span class="muted">Reviewer to confirm</span>'}
        ${item.detail ? `<span class="muted">${escapeHtml(item.detail)}</span>` : ''}
      </div>
    </li>`;
}

function renderReviewLogo(report) {
  return `
    <section class="surface section-card review-block">
      <p class="section-kicker">Cover</p>
      <h3>Client logo</h3>
      <p class="muted">Stored once per client and used by every report for ${escapeHtml(report.project.clientName || 'this client')}. PNG or JPG, up to 1 MB.</p>
      <div class="button-row">
        <label class="button button-secondary">Upload for this client<input type="file" accept="image/png,image/jpeg" data-review-logo="client" hidden /></label>
        <label class="button button-secondary">Use a different logo for this report<input type="file" accept="image/png,image/jpeg" data-review-logo="report" hidden /></label>
        ${report.clientLogoOverride ? '<button class="button button-ghost" data-action="review-clear-report-logo">Remove this report\'s logo</button>' : ''}
      </div>
      ${reviewUi.logoMessage ? `<p class="muted" role="status">${escapeHtml(reviewUi.logoMessage)}</p>` : ''}
    </section>`;
}

function reviewView() {
  const review = reviewUi.review;
  const report = reviewUi.report;
  if (!review || !report) {
    return '<section class="surface empty-state"><p>Loading review...</p></section>';
  }
  const locked = review.status === 'approved';
  const kindLabel = review.kind === 'final' ? 'Final report' : 'Daily report';
  const date = report.collectionDate || (report.project && report.project.reportDate) || '';
  const available = review.available;
  const busy = reviewUi.busy ? 'disabled' : '';
  const blocking = review.checklist.some((item) => item.blocking && item.auto === 'fail');
  return `
    <section class="surface detail-shell review-shell">
      <div class="detail-hero">
        <div class="detail-hero-copy">
          <p class="section-kicker">${escapeHtml(kindLabel)}${date ? ` · ${escapeHtml(date)}` : ''}${review.revision ? ` · Revision R${review.revision}` : ''}</p>
          <h2>${escapeHtml(report.project.clientName || '')}</h2>
          <p>${escapeHtml(report.project.projectNo || '')} ${review.documentNo ? `· ${escapeHtml(review.documentNo)}` : ''}</p>
          <div class="detail-hero-tags">${pill(REVIEW_STATUS_TONE[review.status], REVIEW_STATUS_LABEL[review.status])}</div>
        </div>
        <div class="button-row detail-hero-actions">
          <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
          <button class="button button-secondary" data-action="open-report" data-id="${report.id}">View Details</button>
          <button class="button button-success" data-action="export-pdf" data-id="${report.id}" ${state.exporting ? 'disabled' : ''}>${state.exporting ? 'Generating PDF...' : 'Preview / download PDF'}</button>
        </div>
      </div>

      <section class="surface section-card review-block">
        <p class="section-kicker">Who is acting</p>
        <label class="field"><span>Your name (recorded with each step; gives the initials on the cover)</span>
          <input type="text" class="text-input" data-review-field="actor" value="${escapeHtml(reviewUi.actor)}" placeholder="e.g. A. S. Mahto" /></label>
        <p class="muted">Cover initials: Prepared <strong>${escapeHtml(review.initials.prepared || '-')}</strong> · Checked <strong>${escapeHtml(review.initials.checked || '-')}</strong> · Approved <strong>${escapeHtml(review.initials.approved || '-')}</strong></p>
      </section>

      <section class="surface section-card review-block">
        <p class="section-kicker">Standard 6.4</p>
        <h3>Reviewer checklist</h3>
        <ul class="review-list">${review.checklist.map((item) => renderChecklistItem(item, locked)).join('')}</ul>
        <p class="muted">The app compares every count itself (summary, results tables, electrode priorities, scope and the corrective action plan). ${review.consistency.ok ? 'All counts agree.' : ''}</p>
        ${review.consistency.ok ? '' : `<div class="review-alert" role="alert"><strong>Counts differ: this report cannot be released.</strong><ul>${review.consistency.problems.map((problem) => `<li>${escapeHtml(problem)}</li>`).join('')}</ul></div>`}
        ${review.warnings.length ? `<div class="review-warnings"><strong>Warnings</strong><ul>${review.warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join('')}</ul></div>` : ''}
        <p class="muted">Narrative: ${review.narrative ? (review.narrative.aiWritten ? `written by Claude (${escapeHtml(review.narrative.model)}) from the computed results` : escapeHtml(review.narrative.note || 'not AI-written')) : 'not generated yet'} · Sanity check: <strong>${escapeHtml(review.sanityCheck)}</strong></p>
      </section>

      ${review.status === 'in_review' && review.acknowledgements.length ? `
      <section class="surface section-card review-block">
        <p class="section-kicker">Before approval</p>
        <h3>Acknowledgements</h3>
        ${review.acknowledgements.map((item) => `
          <label class="field"><span>${escapeHtml(item.text)}</span>
            <textarea class="text-input" rows="2" data-review-ack="${escapeHtml(item.key)}" placeholder="Reviewer note (required to approve)">${escapeHtml(reviewUi.acknowledgements[item.key] || '')}</textarea></label>`).join('')}
      </section>` : ''}

      <section class="surface section-card review-block">
        <p class="section-kicker">Workflow</p>
        <h3>${escapeHtml(REVIEW_STATUS_LABEL[review.status])}${review.checked ? ` · checked by ${escapeHtml(review.checked.by)}` : ''}</h3>
        ${locked ? '' : `<label class="field"><span>Note (optional)</span><textarea class="text-input" rows="2" data-review-field="note">${escapeHtml(reviewUi.note)}</textarea></label>`}
        <div class="button-row">
          ${available.submit ? `<button class="button button-primary" data-action="review-step" data-step="submit" ${busy}>Submit for review</button>` : ''}
          ${available.check ? `<button class="button button-primary" data-action="review-step" data-step="check" ${busy} ${blocking ? 'disabled' : ''}>Mark as checked</button>` : ''}
          ${available.approve ? `<button class="button button-success" data-action="review-step" data-step="approve" ${busy}>Approve</button>` : ''}
          ${available.return ? `<button class="button button-ghost" data-action="review-step" data-step="return" ${busy}>Return to draft</button>` : ''}
        </div>
        ${review.status === 'in_review' && !review.checked ? '<p class="muted">Approval is blocked until a checker marks the report as checked. Tick every checklist item first.</p>' : ''}
        ${reviewUi.message ? `<p class="review-message" role="status">${escapeHtml(reviewUi.message)}</p>` : ''}
        ${available.reissue ? `
          <div class="review-reissue">
            <p class="muted">An approved report is never edited. Re-issue creates the next revision (R${review.revision + 1}) as a new draft; the revision table on the cover lists every issue.</p>
            <label class="field"><span>Reason for re-issue</span><input type="text" class="text-input" data-review-field="reissueReason" value="${escapeHtml(reviewUi.reissueReason)}" placeholder="e.g. Corrected P1 pit EP-07 location" /></label>
            <button class="button button-secondary" data-action="review-reissue" ${busy}>Re-issue as R${review.revision + 1}</button>
          </div>` : ''}
      </section>

      ${review.kind === 'final' ? renderReviewLogo(report) : ''}

      <section class="surface section-card review-block">
        <p class="section-kicker">Record</p>
        <h3>Transitions</h3>
        ${review.transitions.length ? `
        <table class="data-table"><thead><tr><th>When</th><th>Step</th><th>By</th><th>Note</th></tr></thead><tbody>
          ${review.transitions.map((row) => `<tr><td>${escapeHtml(formatStamp(row.at))}</td><td>${escapeHtml(STEP_LABEL[row.action] || row.action)}</td><td>${escapeHtml(row.actor || '-')}</td><td>${escapeHtml(row.note || '')}</td></tr>`).join('')}
        </tbody></table>` : '<p class="muted">No steps recorded yet.</p>'}
      </section>
    </section>`;
}

async function runReviewStep(step) {
  if (!reviewUi.actor.trim()) {
    showToast('Enter your name first: it is recorded with this step.', 'warning');
    return;
  }
  reviewUi.busy = true;
  reviewUi.message = '';
  render();
  try {
    const result = await api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/review/${step}`, {
      method: 'POST',
      body: JSON.stringify({ actor: reviewUi.actor.trim(), note: reviewUi.note, acknowledgements: reviewUi.acknowledgements })
    });
    reviewUi.review = result;
    reviewUi.report = await api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}`);
    reviewUi.note = '';
    if (result.stepWarnings && result.stepWarnings.length) {
      reviewUi.message = result.stepWarnings.join(' ');
    }
    showToast(`${STEP_LABEL[step === 'submit' ? 'submit' : step === 'return' ? 'return' : step]}.`, 'healthy');
  } catch (error) {
    reviewUi.message = error.message;
    showToast(error.message, 'critical');
  } finally {
    reviewUi.busy = false;
    render();
  }
}

async function saveReviewTick(id, ticked) {
  const ticks = Object.fromEntries(reviewUi.review.checklist.map((item) => [item.id, item.id === id ? ticked : item.ticked]));
  try {
    reviewUi.review = await api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/review/ticks`, { method: 'PUT', body: JSON.stringify({ ticks }) });
  } catch (error) {
    showToast(error.message, 'critical');
  }
  render();
}

function readLogoFile(file) {
  return new Promise((resolve, reject) => {
    if (!/^image\/(png|jpeg)$/.test(file.type)) {
      reject(new Error('The logo must be a PNG or JPEG image.'));
      return;
    }
    if (file.size > 1024 * 1024) {
      reject(new Error(`The logo is ${(file.size / 1024 / 1024).toFixed(1)} MB; the limit is 1 MB.`));
      return;
    }
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('The logo could not be read.'));
    reader.readAsDataURL(file);
  });
}

async function uploadReviewLogo(scope, file) {
  try {
    const dataUrl = await readLogoFile(file);
    const project = reviewUi.report.project;
    if (scope === 'client') {
      await api('/api/client-logo', { method: 'PUT', body: JSON.stringify({ clientId: project.clientId || '', clientName: project.clientName || '', dataUrl }) });
      reviewUi.logoMessage = `Logo saved for ${project.clientName || 'this client'}.`;
    } else {
      await api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/client-logo`, { method: 'PUT', body: JSON.stringify({ dataUrl }) });
      reviewUi.logoMessage = 'Logo set for this report only.';
      reviewUi.report = await api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}`);
    }
  } catch (error) {
    reviewUi.logoMessage = error.message;
  }
  render();
}

// Click actions of this screen. Returns true when the action was handled here.
function handleReviewAction(action, button) {
  if (action === 'open-review') {
    openReview(button.dataset.id);
    return true;
  }
  if (action === 'review-step') {
    runReviewStep(button.dataset.step);
    return true;
  }
  if (action === 'review-reissue') {
    if (!reviewUi.actor.trim()) {
      showToast('Enter your name first.', 'warning');
      return true;
    }
    reviewUi.busy = true;
    render();
    api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/reissue`, { method: 'POST', body: JSON.stringify({ actor: reviewUi.actor.trim(), reason: reviewUi.reissueReason }) })
      .then((copy) => {
        showToast(`Re-issued as R${copy.revision}. Opening the new draft.`, 'healthy');
        return openReview(copy.id);
      })
      .catch((error) => showToast(error.message, 'critical'))
      .finally(() => {
        reviewUi.busy = false;
        render();
      });
    return true;
  }
  if (action === 'review-clear-report-logo') {
    api(`/api/reports/${encodeURIComponent(reviewUi.report.id)}/client-logo`, { method: 'PUT', body: JSON.stringify({ dataUrl: null }) })
      .then(() => refreshReview())
      .then(() => {
        reviewUi.logoMessage = 'This report now uses the client logo.';
        render();
      })
      .catch((error) => showToast(error.message, 'critical'));
    return true;
  }
  return false;
}

// Typing and ticking on this screen (no re-render per keystroke). Returns true when handled.
function handleReviewInput(target) {
  if (target.dataset && target.dataset.reviewField) {
    const field = target.dataset.reviewField;
    reviewUi[field] = target.value;
    if (field === 'actor') {
      try {
        window.localStorage.setItem('electroreports.reviewer', target.value);
      } catch (_error) {
        // remembering the name is a convenience only
      }
    }
    return true;
  }
  if (target.dataset && target.dataset.reviewAck) {
    reviewUi.acknowledgements[target.dataset.reviewAck] = target.value;
    return true;
  }
  return false;
}

function handleReviewChange(target) {
  if (target.dataset && target.dataset.reviewTick) {
    saveReviewTick(target.dataset.reviewTick, target.checked);
    return true;
  }
  if (target.dataset && target.dataset.reviewLogo) {
    const file = target.files && target.files[0];
    if (file) {
      uploadReviewLogo(target.dataset.reviewLogo, file);
    }
    target.value = '';
    return true;
  }
  return false;
}
