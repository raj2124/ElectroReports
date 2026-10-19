// In-app notices on Home (6D): "a new D03 project needs two answers". Notices are addressed to the
// owners of the project's ESHA Onsite tasks (or the project owner). There are no logins yet, so who
// the person is comes from an "I am ..." choice from the Zoho users, remembered in this browser.
// currentUser() / setCurrentUser() are the ONLY identity code on the client: per-engineer logins
// replace them (and noticeRecipientFrom in src/server.js on the server).

const noticeUi = { notices: [], users: [], answers: {}, busy: false };
const IAM_KEY = 'electroreports.iam';

function currentUser() {
  try {
    const stored = JSON.parse(window.localStorage.getItem(IAM_KEY) || 'null');
    return stored && (stored.email || stored.name) ? stored : null;
  } catch (_error) {
    return null;
  }
}

function setCurrentUser(user) {
  try {
    if (user) {
      window.localStorage.setItem(IAM_KEY, JSON.stringify(user));
    } else {
      window.localStorage.removeItem(IAM_KEY);
    }
  } catch (_error) {
    // remembering who you are is a convenience only
  }
}

function noticeQuery() {
  const me = currentUser();
  return me ? `email=${encodeURIComponent(me.email || '')}&name=${encodeURIComponent(me.name || '')}` : '';
}

async function loadNotices() {
  const query = noticeQuery();
  try {
    const list = query ? await api(`/api/notices?${query}`) : [];
    noticeUi.notices = Array.isArray(list) ? list : [];
  } catch (_error) {
    noticeUi.notices = [];
  }
  if (state.view === 'dashboard') {
    render();
  }
}

async function loadNoticeUsers() {
  try {
    const users = await api('/api/zoho/users');
    noticeUi.users = Array.isArray(users) ? users : [];
  } catch (_error) {
    noticeUi.users = [];
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
  const me = currentUser();
  if (!me) {
    return `
      <section class="surface zoho-overview" aria-label="Who are you">
        <div class="dashboard-section-copy">
          <p class="section-kicker">Notices</p>
          <h3>Who are you?</h3>
          <p class="muted">Pick your name to see the notices addressed to you, for example a new project that needs two answers about the site.</p>
        </div>
        <select class="text-input" data-notice-user>
          <option value="">I am...</option>
          ${noticeUi.users.map((user) => `<option value="${escapeHtml(user.email || user.displayName)}" data-name="${escapeHtml(user.displayName || '')}">${escapeHtml(user.displayName || user.email)}</option>`).join('')}
        </select>
      </section>`;
  }
  return `
    <section class="surface zoho-overview" aria-label="Notices">
      <div class="dashboard-section-copy">
        <p class="section-kicker">Notices</p>
        <h3>${noticeUi.notices.length ? `${noticeUi.notices.length} notice${noticeUi.notices.length === 1 ? '' : 's'} for you` : 'No notices'}</h3>
        <p class="muted">You are ${escapeHtml(me.name || me.email)}. <button class="button button-ghost" data-action="notice-change-user">Not you?</button></p>
      </div>
      ${noticeUi.notices.map(noticeCard).join('')}
    </section>`;
}

// Returns true when the click was one of ours.
function handleNoticeAction(action, button) {
  if (action === 'notice-change-user') {
    setCurrentUser(null);
    noticeUi.notices = [];
    render();
    return true;
  }
  if (action === 'notice-dismiss') {
    api(`/api/notices/${encodeURIComponent(button.dataset.id)}/dismiss?${noticeQuery()}`, { method: 'POST', body: '{}' })
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
    api(`/api/notices/${encodeURIComponent(id)}/answer?${noticeQuery()}`, { method: 'POST', body: JSON.stringify(answers) })
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
  if (target.dataset && target.dataset.noticeUser !== undefined) {
    if (target.value) {
      const option = target.selectedOptions[0];
      setCurrentUser({ email: target.value.includes('@') ? target.value : '', name: option.dataset.name || target.value });
      loadNotices().then(() => render());
    }
    return true;
  }
  if (target.dataset && target.dataset.noticeAnswer) {
    const [id, key] = target.dataset.noticeAnswer.split(':');
    noticeUi.answers[id] = { ...(noticeUi.answers[id] || {}), [key]: target.value === 'yes' };
    return true;
  }
  return false;
}
