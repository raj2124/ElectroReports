// In-app notices on Home (6D): "a new D03 project needs two answers". Notices are addressed to the
// owners of the project's ESHA Onsite tasks (or the project owner), and since 6E they go to the
// signed-in user: the server matches them to the session, nothing is chosen here.

const noticeUi = { notices: [], answers: {}, busy: false };

async function loadNotices() {
  try {
    const list = authUi.user ? await api('/api/notices') : [];
    noticeUi.notices = Array.isArray(list) ? list : [];
  } catch (_error) {
    noticeUi.notices = [];
  }
  if (state.view === 'dashboard') {
    render();
  }
}

function noticeCard(notice) {
  const info = notice.payload;
  const answers = noticeUi.answers[notice.id] || {};
  const yesNo = (key, label) => `
    <fieldset class="notice-question"><legend>${escapeHtml(label)}</legend>
      <label><input type="radio" name="${key}-${notice.id}" data-notice-answer="${notice.id}:${key}" value="yes" ${answers[key] === true ? 'checked' : ''} /> Yes</label>
      <label><input type="radio" name="${key}-${notice.id}" data-notice-answer="${notice.id}:${key}" value="no" ${answers[key] === false ? 'checked' : ''} /> No</label>
    </fieldset>`;
  return `
    <article class="notice-card" data-notice="${notice.id}">
      <div><span class="chip">New project</span> <strong>${escapeHtml(info.projectName)}</strong>${info.clientName ? ` · ${escapeHtml(info.clientName)}` : ''}</div>
      <p class="muted">${info.reason === 'onsite_owner' ? 'You own an ESHA Onsite task on this project.' : 'You are the owner of this project.'} It has no ElectroReports workspace yet. Two questions about the site set how the data is collected.</p>
      <div class="notice-questions">
        ${yesNo('phonesAllowed', 'Phones/tablets allowed on site?')}
        ${yesNo('mobileData', 'Mobile data available?')}
      </div>
      <div class="button-row">
        <button class="button button-primary" data-action="notice-answer" data-id="${notice.id}" ${noticeUi.busy ? 'disabled' : ''}>Save answers and open the workspace</button>
        <button class="button button-ghost" data-action="notice-dismiss" data-id="${notice.id}" ${noticeUi.busy ? 'disabled' : ''}>Dismiss</button>
      </div>
    </article>`;
}

function renderHomeNotices() {
  if (!noticeUi.notices.length) {
    return '';
  }
  return `
    <section class="surface zoho-overview" aria-label="Notices">
      <div class="dashboard-section-copy">
        <p class="section-kicker">Notices</p>
        <h3 aria-level="2">${noticeUi.notices.length} notice${noticeUi.notices.length === 1 ? '' : 's'} for you</h3>
      </div>
      ${noticeUi.notices.map(noticeCard).join('')}
    </section>`;
}

// Returns true when the click was one of ours.
function handleNoticeAction(action, button) {
  if (action === 'notice-dismiss') {
    api(`/api/notices/${encodeURIComponent(button.dataset.id)}/dismiss`, { method: 'POST', body: '{}' })
      .then(() => loadNotices())
      .catch((error) => showToast(error.message, 'critical'));
    return true;
  }
  if (action === 'notice-answer') {
    const id = button.dataset.id;
    const answers = noticeUi.answers[id] || {};
    if (typeof answers.phonesAllowed !== 'boolean' || typeof answers.mobileData !== 'boolean') {
      showToast('Answer both questions first.', 'warning');
      return true;
    }
    noticeUi.busy = true;
    render();
    api(`/api/notices/${encodeURIComponent(id)}/answer`, { method: 'POST', body: JSON.stringify(answers) })
      .then(async (result) => {
        showToast(`Saved. Default data collection: ${result.dataCollectionMode}.`, 'healthy');
        await loadNotices();
        await loadProjects();
        await openWorkspace(result.projectId);
      })
      .catch((error) => showToast(error.message, 'critical'))
      .finally(() => {
        noticeUi.busy = false;
        render();
      });
    return true;
  }
  return false;
}

// Returns true when handled.
function handleNoticeChange(target) {
  if (target.dataset && target.dataset.noticeAnswer) {
    const [id, key] = target.dataset.noticeAnswer.split(':');
    noticeUi.answers[id] = { ...(noticeUi.answers[id] || {}), [key]: target.value === 'yes' };
    return true;
  }
  return false;
}
