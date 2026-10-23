// Sign-in, My Profile and the admin screens (6E). Who you are comes from the server session; roles
// decide which buttons are shown, but the server enforces every one of them (src/permissions.js).

const authUi = { user: null, mode: 'none', ready: false, errors: {}, devUsers: [], loaded: false, adminTab: 'users', people: null, totals: null, backfill: { from: '', to: '' }, users: [], audit: { total: 0, entries: [] }, auditFilter: { user: '', action: '', offset: 0 }, message: '' };

const ROLE_LABEL = { engineer: 'Engineer', reviewer: 'Reviewer', approver: 'Approver', admin: 'Admin' };

function hasRoleClient(...roles) {
  return Boolean(authUi.user) && (authUi.user.roles.includes('admin') || roles.some((role) => authUi.user.roles.includes(role)));
}

async function loadAuth() {
  try {
    const me = await fetch('/api/auth/me').then((response) => response.json());
    Object.assign(authUi, { user: me.user, mode: me.mode, ready: me.ready, errors: me.errors || {}, loaded: true });
    if (!me.user && me.mode === 'dev') {
      authUi.devUsers = await fetch('/api/auth/dev-users').then((response) => (response.ok ? response.json() : []));
    }
  } catch (_error) {
    authUi.loaded = true;
  }
}

function signInView() {
  const code = new URLSearchParams(window.location.search).get('auth_error');
  const error = code ? authUi.errors[code] || 'Sign-in failed. Try again.' : '';
  return `
    <section class="surface empty-state signin-card" aria-label="Sign in">
      <div class="empty-state-copy">
        <p class="section-kicker">ElectroReports</p>
        <h1>Sign in</h1>
        <p class="muted">Use your Elegrow account. Every action is recorded against you.</p>
        ${error ? `<p class="form-error" role="alert">${escapeHtml(error)}</p>` : ''}
        ${
          authUi.mode === 'zoho'
            ? authUi.ready
              ? '<a class="button button-primary" id="signInZoho" href="/auth/login">Sign in with Zoho</a>'
              : '<p class="form-error">Sign-in is not configured on this server yet (Zoho client ID and secret).</p>'
            : authUi.mode === 'dev'
              ? `<p class="muted">Development sign-in (not for production).</p>
                 <select class="text-input" data-dev-user aria-label="Sign in as"><option value="">Choose a user...</option>${authUi.devUsers.map((user) => `<option value="${escapeHtml(user.email)}">${escapeHtml(user.name)} (${escapeHtml(user.roles.join(', '))})</option>`).join('')}</select>`
              : '<p class="form-error">This server has no sign-in mode configured (AUTH_MODE).</p>'
        }
      </div>
    </section>`;
}

// Name, roles and sign out, in the header.
function renderUserBar() {
  if (!authUi.user) {
    return '';
  }
  return `
    <div class="user-bar">
      <strong>${escapeHtml(authUi.user.name)}</strong>
      <span class="muted">${authUi.user.roles.map((role) => ROLE_LABEL[role] || role).join(', ')}</span>
      <button class="button button-ghost" data-action="open-schedule">Schedule</button>
      <button class="button button-ghost" data-action="open-kb">Search documents</button>
      ${hasRoleClient('engineer') ? '<button class="button button-ghost" data-action="open-scans">Scan sheets</button>' : ''}
      <button class="button button-ghost" data-action="open-profile">My profile</button>
      ${hasRoleClient('admin') ? '<button class="button button-ghost" data-action="open-admin">Admin</button>' : ''}
      <button class="button button-ghost" data-action="sign-out">Sign out</button>
    </div>`;
}

function profileView() {
  const user = authUi.user;
  return `
    <section class="surface section-card profile-card">
      <p class="section-kicker">My profile</p>
      <h1>${escapeHtml(user.name)}</h1>
      <p class="muted">${escapeHtml(user.email)} · ${user.roles.map((role) => ROLE_LABEL[role] || role).join(', ')}</p>
      <p class="muted">Your name and designation are printed as Prepared, Checked and Approved by on the reports you act on. A signed report keeps what it said when you signed it.</p>
      <label class="field"><span>Name</span><input class="text-input" type="text" data-profile-field="name" value="${escapeHtml(user.name)}" /></label>
      <label class="field"><span>Designation</span><input class="text-input" type="text" data-profile-field="designation" value="${escapeHtml(user.designation)}" placeholder="e.g. Senior Engineer - Earthing" /></label>
      ${authUi.message ? `<p class="review-message" role="status">${escapeHtml(authUi.message)}</p>` : ''}
      <div class="button-row">
        <button class="button button-primary" data-action="save-profile">Save profile</button>
        <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
      </div>
    </section>`;
}

function adminView() {
  const tab = authUi.adminTab;
  const users = authUi.users;
  return `
    <section class="surface section-card admin-card">
      <p class="section-kicker">Admin</p>
      <h1>Administration</h1>
      <div class="button-row admin-tabs">
        <button class="button ${tab === 'users' ? 'button-primary' : 'button-secondary'}" data-action="admin-tab" data-tab="users">Users and roles</button>
        <button class="button ${tab === 'audit' ? 'button-primary' : 'button-secondary'}" data-action="admin-tab" data-tab="audit">Audit log</button>
        <button class="button ${tab === 'people' ? 'button-primary' : 'button-secondary'}" data-action="admin-tab" data-tab="people">Site Work Log</button>
        <button class="button ${tab === 'calendar' ? 'button-primary' : 'button-secondary'}" data-action="admin-tab" data-tab="calendar">Calendar sync</button>
        <button class="button ${tab === 'kb' ? 'button-primary' : 'button-secondary'}" data-action="admin-tab" data-tab="kb">Knowledge base</button>
        <button class="button ${tab === 'scans' ? 'button-primary' : 'button-secondary'}" data-action="admin-tab" data-tab="scans">Scans and AI</button>
        <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
      </div>
      ${
        tab === 'people'
          ? peopleSyncPanel()
          : tab === 'calendar'
          ? calendarSyncPanel()
          : tab === 'kb'
          ? kbAdminPanel()
          : tab === 'scans'
          ? scanAdminPanel()
          : tab === 'users'
          ? `<table class="data-table admin-users"><thead><tr><th>Name</th><th>E-mail</th><th>Roles</th><th>Active</th><th>Last sign-in</th></tr></thead><tbody>
          ${users
            .map(
              (user) => `<tr data-user-row="${user.id}">
            <td>${escapeHtml(user.name)}<br /><span class="muted">${escapeHtml(user.designation || '')}</span></td>
            <td>${escapeHtml(user.email)}</td>
            <td>${['engineer', 'reviewer', 'approver', 'admin'].map((role) => `<label class="role-box"><input type="checkbox" data-admin-role="${user.id}:${role}" ${user.roles.includes(role) ? 'checked' : ''} /> ${ROLE_LABEL[role]}</label>`).join('')}</td>
            <td><label class="role-box"><input type="checkbox" aria-label="Active: ${escapeHtml(user.name)}" data-admin-active="${user.id}" ${user.active ? 'checked' : ''} /></label></td>
            <td>${escapeHtml(user.lastLoginAt ? formatStamp(user.lastLoginAt) : 'never')}</td></tr>`
            )
            .join('')}
          </tbody></table>
          ${authUi.message ? `<p class="review-message" role="status">${escapeHtml(authUi.message)}</p>` : ''}`
          : `<div class="field-grid">
            <label class="field"><span>Action starts with</span><input class="text-input" type="text" data-audit-field="action" value="${escapeHtml(authUi.auditFilter.action)}" placeholder="e.g. review." /></label>
            <label class="field"><span>User</span><select class="text-input" data-audit-field="user"><option value="">Everyone</option>${users.map((user) => `<option value="${user.id}" ${authUi.auditFilter.user === user.id ? 'selected' : ''}>${escapeHtml(user.name)}</option>`).join('')}</select></label>
          </div>
          <p class="muted">${authUi.audit.total} entries. Read only: nothing here can be edited or deleted.</p>
          <table class="data-table"><thead><tr><th>When</th><th>User</th><th>Action</th><th>On</th><th>Detail</th></tr></thead><tbody>
          ${authUi.audit.entries.map((entry) => `<tr><td>${escapeHtml(formatStamp(entry.at))}</td><td>${escapeHtml(entry.userName || entry.userEmail || '-')}</td><td>${escapeHtml(entry.action)}</td><td>${escapeHtml([entry.entityType, entry.entityId].filter(Boolean).join(' '))}</td><td>${escapeHtml(entry.detail ? JSON.stringify(entry.detail).slice(0, 140) : '')}</td></tr>`).join('')}
          </tbody></table>
          <div class="button-row">
            <button class="button button-secondary" data-action="audit-page" data-step="-1" ${authUi.auditFilter.offset ? '' : 'disabled'}>Newer</button>
            <button class="button button-secondary" data-action="audit-page" data-step="1" ${authUi.auditFilter.offset + 100 < authUi.audit.total ? '' : 'disabled'}>Older</button>
          </div>`
      }
    </section>`;
}

// Totals per engineer: completed against planned per test, then on-time, first-time approved, returns, data gaps.
function peopleTotalsTable(totals) {
  if (!totals.engineers.length) return `<p class="muted" data-people-totals>No approved work between ${escapeHtml(totals.from)} and ${escapeHtml(totals.to)}.</p>`;
  const pct = (value) => (value === null ? '-' : `${value}%`);
  return `<table class="data-table" data-people-totals><thead><tr><th>Engineer</th><th>Test</th><th>Completed</th><th>Planned</th><th>On time</th><th>First-time approved</th><th>Returns</th><th>Data gaps</th></tr></thead><tbody>
    ${totals.engineers.map((who) => who.tests.map((test, index) => `<tr data-totals-row="${escapeHtml(who.email)}">
      <td>${index === 0 ? `${escapeHtml(who.name || who.email)}<br /><span class="muted">${who.engineerDays} day${who.engineerDays === 1 ? '' : 's'}</span>` : ''}</td>
      <td>${escapeHtml(test.label)}</td><td>${test.completed} ${escapeHtml(test.unit)}</td><td>${test.planned || '-'}</td>
      ${index === 0 ? `<td rowspan="${who.tests.length}">${pct(who.onTimePct)}</td><td rowspan="${who.tests.length}">${pct(who.firstTimeApprovedPct)}</td><td rowspan="${who.tests.length}">${who.returns}</td><td rowspan="${who.tests.length}">${who.dataGaps}</td>` : ''}
    </tr>`).join('')).join('')}
  </tbody></table>`;
}

const SYNC_STATE_LABEL = { done: 'Sent', pending: 'Waiting', failed: 'Failed' };

// Site Work Log in Zoho People (6G): whether the form exists, what is waiting or failed, and the
// Backfill of approved days. Written when a daily report is approved; finals, sample and test reports are never sent.
function peopleSyncPanel() {
  const info = authUi.people;
  if (!info) return '<p class="muted">Loading...</p>';
  const counts = info.counts || {};
  const last = info.lastBackfill;
  return `
    <p class="muted" data-people-form>${info.formExists === false ? `The People form "${escapeHtml(info.form)}" does not exist yet. Create it in Zoho People; until then entries wait here.` : info.formError ? `Could not check the People form: ${escapeHtml(info.formError)}` : `Writing to the People form "${escapeHtml(info.form)}".`}</p>
    <p class="muted" data-people-counts>Sent ${counts.done || 0} · waiting ${counts.pending || 0} · failed ${counts.failed || 0}</p>
    <div class="field-grid">
      <label class="field"><span>Backfill from (work date)</span><input class="text-input" type="date" data-backfill-field="from" value="${escapeHtml(authUi.backfill.from)}" /></label>
      <label class="field"><span>to</span><input class="text-input" type="date" data-backfill-field="to" value="${escapeHtml(authUi.backfill.to)}" /></label>
    </div>
    <div class="button-row">
      <button class="button button-primary" data-action="people-backfill" ${authUi.backfill.from && authUi.backfill.to ? '' : 'disabled'}>Backfill approved days</button>
      <button class="button button-secondary" data-action="people-retry" ${(counts.pending || 0) + (counts.failed || 0) ? '' : 'disabled'}>Retry waiting and failed</button>
    </div>
    <div class="button-row">
      <button class="button button-secondary" data-action="people-totals" ${authUi.backfill.from && authUi.backfill.to ? '' : 'disabled'}>Show totals for these dates</button>
    </div>
    ${authUi.totals ? peopleTotalsTable(authUi.totals) : ''}
    ${authUi.message ? `<p class="review-message" role="status" data-people-message>${escapeHtml(authUi.message)}</p>` : ''}
    ${last ? `<p class="muted">Last backfill ${escapeHtml(last.from)} to ${escapeHtml(last.to)}: ${last.reports} days, ${last.written} written, ${last.unchanged} unchanged, ${last.failed} failed, ${last.blocked} waiting, ${last.skippedSample} sample or test reports skipped.</p>` : ''}
    ${info.problems.length ? `<table class="data-table"><thead><tr><th>Entry</th><th>Work date</th><th>Test</th><th>State</th><th>Reason</th></tr></thead><tbody>
      ${info.problems.map((row) => `<tr><td>${escapeHtml(row.key)}</td><td>${escapeHtml(row.workDate)}</td><td>${escapeHtml(row.testId)}</td><td>${escapeHtml(SYNC_STATE_LABEL[row.state] || row.state)}</td><td>${escapeHtml(row.error)}</td></tr>`).join('')}
    </tbody></table>` : ''}`;
}

async function loadAdminData() {
  try {
    authUi.users = await api('/api/admin/users');
    if (authUi.adminTab === 'people') {
      authUi.people = await api('/api/admin/people-sync');
    }
    await loadGraphAdmin(authUi.adminTab);
    if (authUi.adminTab === 'scans' && typeof loadScanAdmin === 'function') {
      await loadScanAdmin();
    }
    if (authUi.adminTab === 'audit') {
      const filter = authUi.auditFilter;
      authUi.audit = await api(`/api/admin/audit?limit=100&offset=${filter.offset}&user=${encodeURIComponent(filter.user)}&action=${encodeURIComponent(filter.action)}`);
    }
  } catch (error) {
    showToast(error.message, 'critical');
  }
  render();
}

// Returns true when the click was one of ours.
function handleAuthAction(action, button) {
  if (action === 'sign-out') {
    fetch('/auth/logout', { method: 'POST' }).finally(() => window.location.assign('/'));
    return true;
  }
  if (action === 'open-profile') {
    authUi.message = '';
    state.view = 'profile';
    render();
    return true;
  }
  if (action === 'save-profile') {
    const draft = authUi.profileDraft || { name: authUi.user.name, designation: authUi.user.designation };
    api('/api/profile', { method: 'PUT', body: JSON.stringify({ name: draft.name, designation: draft.designation }) })
      .then((result) => {
        authUi.user = result.user;
        authUi.message = 'Profile saved.';
        render();
      })
      .catch((error) => showToast(error.message, 'critical'));
    return true;
  }
  if (action === 'open-admin') {
    authUi.message = '';
    state.view = 'admin';
    loadAdminData();
    render();
    return true;
  }
  if (action === 'admin-tab') {
    authUi.adminTab = button.dataset.tab;
    loadAdminData();
    return true;
  }
  if (action === 'people-totals') {
    const { from, to } = authUi.backfill;
    api(`/api/admin/people-sync/totals?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`)
      .then((totals) => { authUi.totals = totals; render(); })
      .catch((error) => showToast(error.message, 'critical'));
    return true;
  }
  if (action === 'people-backfill' || action === 'people-retry') {
    const backfill = action === 'people-backfill';
    authUi.message = backfill ? 'Backfilling...' : 'Retrying...';
    render();
    api(backfill ? '/api/admin/people-sync/backfill' : '/api/admin/people-sync/retry', { method: 'POST', body: JSON.stringify(backfill ? authUi.backfill : {}) })
      .then((result) => {
        authUi.message = backfill ? `Backfill done: ${result.written} written, ${result.unchanged} unchanged, ${result.failed} failed, ${result.blocked} waiting.` : `Retry done: ${result.written} written, ${result.failed} failed, ${result.blocked} waiting.`;
        return loadAdminData();
      })
      .catch((error) => {
        authUi.message = '';
        showToast(error.message, 'critical');
        render();
      });
    return true;
  }
  if (action === 'audit-page') {
    authUi.auditFilter.offset = Math.max(0, authUi.auditFilter.offset + Number(button.dataset.step) * 100);
    loadAdminData();
    return true;
  }
  return false;
}

// Returns true when handled.
function handleAuthChange(target) {
  const data = target.dataset || {};
  if (data.devUser !== undefined) {
    if (target.value) {
      fetch('/auth/dev-login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: target.value }) }).then(() => window.location.assign('/'));
    }
    return true;
  }
  if (data.profileField) {
    authUi.profileDraft = { ...(authUi.profileDraft || { name: authUi.user.name, designation: authUi.user.designation }), [data.profileField]: target.value };
    return true;
  }
  if (data.adminRole) {
    const [id, role] = data.adminRole.split(':');
    const user = authUi.users.find((entry) => entry.id === id);
    const roles = target.checked ? [...new Set([...user.roles, role])] : user.roles.filter((entry) => entry !== role);
    saveAdminUser(id, { roles });
    return true;
  }
  if (data.adminActive) {
    saveAdminUser(data.adminActive, { active: target.checked });
    return true;
  }
  if (data.backfillField) {
    authUi.backfill = { ...authUi.backfill, [data.backfillField]: target.value };
    render();
    return true;
  }
  if (data.auditField) {
    authUi.auditFilter = { ...authUi.auditFilter, [data.auditField]: target.value, offset: 0 };
    loadAdminData();
    return true;
  }
  return false;
}

function saveAdminUser(id, patch) {
  api(`/api/admin/users/${encodeURIComponent(id)}`, { method: 'PUT', body: JSON.stringify(patch) })
    .then((user) => {
      authUi.users = authUi.users.map((entry) => (entry.id === id ? user : entry));
      authUi.message = `Saved ${user.name}.`;
      if (authUi.user && authUi.user.id === id) {
        authUi.user = user;
      }
      render();
    })
    .catch((error) => {
      showToast(error.message, 'critical');
      loadAdminData();
    });
}
