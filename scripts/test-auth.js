// Per-engineer sign-in (Session 6E): the route rules on every route, the review gate (a preparer can
// never check or approve their own report; an admin can override with a typed reason; a checker who
// is also the approver is allowed with a warning), ownership, dev mode, session expiry, the Zoho
// OpenID Connect flow against a fake accounts server, rejected accounts, roles and the audit log.
const assert = require('assert');
const fs = require('fs');
const http = require('http');
const net = require('net');
const os = require('os');
const path = require('path');
const Database = require('better-sqlite3');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const permissions = require('../src/permissions');
const { createAuth, readIdToken, parseDevUsers } = require('../src/auth');
const { buildPayload } = require('./fixtures/golden-esha');
const { devEnv, createCaller, login } = require('./lib/dev-auth');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}
const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'auth-'));

async function startServer(env, name = 'app') {
  const port = env.ELECTROREPORTS_PORT ? Number(env.ELECTROREPORTS_PORT) : await freePort();
  const dataFile = path.join(dir, `${name}.json`);
  const server = spawn(process.execPath, [path.join(__dirname, '..', 'src', 'server.js')], {
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: dataFile, ELECTROREPORTS_GENERATED_DIR: path.join(dir, `${name}-pdfs`), ZOHO_USE_MOCK: 'true', ANTHROPIC_API_KEY: '', OPENAI_API_KEY: '', GEMINI_API_KEY: '', ...env },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    server.on('exit', (code) => reject(new Error(`server exited early (${code})`)));
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  return { port, base: `http://127.0.0.1:${port}`, dataFile, stop: () => server.kill() };
}

const draftPayload = (extra = {}) => {
  const payload = buildPayload();
  return { ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', nextDayPlan: '', ...extra } };
};

(async () => {
  // ================= rules, without a server =================================================
  await check('every write route in server.js has an explicit rule (a new route cannot be left open by accident)', () => {
    const source = fs.readFileSync(path.join(__dirname, '..', 'src', 'server.js'), 'utf8');
    const routes = [...source.matchAll(/app\.(post|put|delete)\('([^']+)'/g)].map(([, method, route]) => [method.toUpperCase(), route.replace(/:[A-Za-z]+/g, 'x')]);
    assert.ok(routes.length >= 25, `found ${routes.length} write routes`);
    const uncovered = routes.filter(([method, route]) => !permissions.ROUTE_RULES.some(([ruleMethod, pattern]) => (ruleMethod === '*' || ruleMethod === method) && pattern.test(route)));
    assert.deepStrictEqual(uncovered, [], `routes with no explicit rule: ${JSON.stringify(uncovered)}`);
  });

  await check('route rules: roles, admin implies all, unknown writes are admin-only, reads are open to any signed-in user', () => {
    const user = (...roles) => ({ id: 'U', roles });
    const can = (roles, method, route) => permissions.mayCall(user(...roles), method, route);
    assert.ok(can(['engineer'], 'POST', '/api/reports') && !can(['reviewer'], 'POST', '/api/reports'));
    assert.ok(can(['reviewer'], 'POST', '/api/reports/x/review/check') && !can(['engineer', 'approver'], 'POST', '/api/reports/x/review/check'));
    assert.ok(can(['approver'], 'POST', '/api/reports/x/review/approve') && !can(['reviewer'], 'POST', '/api/reports/x/review/approve'));
    assert.ok(can(['reviewer'], 'POST', '/api/reports/x/review/return') && can(['approver'], 'POST', '/api/reports/x/review/return') && !can(['engineer'], 'POST', '/api/reports/x/review/return'));
    assert.ok(can(['approver'], 'POST', '/api/reports/x/reissue') && !can(['reviewer'], 'POST', '/api/reports/x/reissue'));
    ['POST /api/reports', 'POST /api/reports/x/review/check', 'POST /api/reports/x/review/approve', 'POST /api/reports/x/reissue', 'PUT /api/client-logo', 'POST /api/zoho/poll', 'PUT /api/admin/users/x', 'POST /api/anything-new'].forEach((entry) => {
      const [method, route] = entry.split(' ');
      assert.ok(can(['admin'], method, route), `admin: ${entry}`);
    });
    assert.ok(!can(['engineer', 'reviewer', 'approver'], 'POST', '/api/anything-new'), 'an unknown write is admin-only');
    assert.ok(can([], 'GET', '/api/reports') && can([], 'GET', '/api/projects/x') && !can(['approver'], 'GET', '/api/admin/users'));
  });

  await check('id_token: audience and expiry are checked; the claims are read', () => {
    const token = (claims) => `h.${Buffer.from(JSON.stringify(claims)).toString('base64url')}.s`;
    const future = Math.floor(Date.now() / 1000) + 600;
    assert.strictEqual(readIdToken(token({ aud: 'cid', exp: future, email: 'a@elegrow.com' }), 'cid').email, 'a@elegrow.com');
    assert.strictEqual(readIdToken(token({ aud: ['x', 'cid'], exp: future, email: 'a@elegrow.com' }), 'cid').email, 'a@elegrow.com');
    assert.strictEqual(readIdToken(token({ aud: 'other', exp: future, email: 'a@elegrow.com' }), 'cid'), null);
    assert.strictEqual(readIdToken(token({ aud: 'cid', exp: 1, email: 'a@elegrow.com' }), 'cid'), null);
    assert.strictEqual(readIdToken('not-a-jwt', 'cid'), null);
    assert.deepStrictEqual(parseDevUsers(['A@x.com:engineer+reviewer:A Name:Boss'])[0], { email: 'a@x.com', roles: ['engineer', 'reviewer'], name: 'A Name', designation: 'Boss' });
  });

  await check('the hourly re-check: a deactivated portal user loses the session; if Zoho is down the session is kept', async () => {
    const store = createReportStore({ filePath: path.join(dir, 'recheck.db') });
    const user = store.addUser({ email: 'x@elegrow.com', name: 'X', roles: ['engineer'] });
    let answer = { found: true, active: true };
    const zoho = { lookupPortalUser: async () => { if (answer === 'down') throw new Error('Zoho down'); return answer; } };
    const config = { auth: { mode: 'zoho', zoho: { redirectUri: 'http://127.0.0.1:1/cb', clientId: 'c', clientSecret: 's', accountsUrl: 'http://x' }, allowedDomain: 'elegrow.com', idleHours: 12, recheckMinutes: 60, sessionSecret: 'secret', seedAdmins: [], devUsers: [] } };
    const auth = createAuth({ store, config, zoho, log: { error() {} } });
    const raw = new Database(path.join(dir, 'recheck.db'));
    const attach = async (token) => { const req = { headers: { cookie: `er_session=${token}` } }; await auth.attachUser(req, {}, () => {}); return req.user; };
    const open = (token, who = user) => { store.createSession(auth.hashOf(token), who.id); raw.prepare('UPDATE sessions SET checked_at = ? WHERE id_hash = ?').run(new Date(Date.now() - 2 * 3600 * 1000).toISOString(), auth.hashOf(token)); };
    open('tok-ok');
    assert.strictEqual((await attach('tok-ok')).email, 'x@elegrow.com', 'still active: kept, and re-checked');
    assert.ok(Date.now() - Date.parse(store.getSession(auth.hashOf('tok-ok')).checkedAt) < 5000);
    open('tok-down');
    answer = 'down';
    assert.strictEqual((await attach('tok-down')).email, 'x@elegrow.com', 'Zoho unreachable: the session is kept');
    answer = { found: true, active: false };
    const other = store.addUser({ email: 'y@elegrow.com', name: 'Y', roles: ['engineer'] }); // (a failed check backs off per user for a few minutes)
    open('tok-gone', other);
    assert.strictEqual(await attach('tok-gone'), null, 'no longer active in the portal: signed out');
    assert.strictEqual(store.getSession(auth.hashOf('tok-gone')), null);
    assert.ok(store.listAudit({ action: 'auth.revoked' }).total >= 1);
    raw.close();
  });

  // ================= dev mode server ==========================================================
  const dev = await startServer(devEnv(), 'dev');
  const call = createCaller(dev.base);
  const as = (email) => (method, url, body) => call(method, url, body, email);
  try {
    await check('signed out: the API and the generated PDFs need a session; health and "me" are open', async () => {
      assert.strictEqual((await call('GET', '/api/health', undefined, null)).status, 200);
      const me = await call('GET', '/api/auth/me', undefined, null);
      assert.deepStrictEqual([me.status, me.body.user, me.body.mode], [200, null, 'dev']);
      for (const route of ['/api/reports', '/api/projects', '/api/sheets', '/api/notices', '/api/catalog']) {
        const response = await call('GET', route, undefined, null);
        assert.deepStrictEqual([response.status, response.body.code], [401, 'SIGN_IN_REQUIRED'], route);
      }
      assert.strictEqual((await call('POST', '/api/reports', {}, null)).status, 401);
      assert.strictEqual((await call('GET', '/generated-pdfs/x.pdf', undefined, null)).status, 401);
    });

    await check('dev sign-in: only the configured users, a httpOnly SameSite=Lax cookie, sign-out ends the session', async () => {
      const refused = await fetch(`${dev.base}/auth/dev-login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: 'stranger@elegrow.com' }) });
      assert.strictEqual(refused.status, 403);
      const ok = await fetch(`${dev.base}/auth/dev-login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: 'engineer@elegrow.com' }) });
      assert.strictEqual(ok.status, 200);
      const cookie = ok.headers.getSetCookie().find((entry) => entry.startsWith('er_session='));
      assert.ok(/HttpOnly/i.test(cookie) && /SameSite=Lax/i.test(cookie) && /Path=\//.test(cookie) && !/Secure/i.test(cookie), cookie);
      const header = cookie.split(';')[0];
      const me = await (await fetch(`${dev.base}/api/auth/me`, { headers: { Cookie: header } })).json();
      assert.deepStrictEqual([me.user.email, me.user.roles, me.user.name], ['engineer@elegrow.com', ['engineer'], 'E. N. Gineer']);
      assert.strictEqual((await fetch(`${dev.base}/auth/logout`, { method: 'POST', headers: { Cookie: header } })).status, 200);
      assert.strictEqual((await fetch(`${dev.base}/api/reports`, { headers: { Cookie: header } })).status, 401, 'the session is gone');
      assert.strictEqual((await fetch(`${dev.base}/api/auth/dev-users`)).status, 200);
    });

    await check('a request from another site is refused (Origin must be this host)', async () => {
      const cookie = await login(dev.base, 'engineer@elegrow.com');
      const evil = await fetch(`${dev.base}/api/assessments`, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie, Origin: 'http://evil.example' }, body: JSON.stringify(draftPayload()) });
      assert.deepStrictEqual([evil.status, (await evil.json()).code], [403, 'CROSS_SITE']);
      const own = await fetch(`${dev.base}/api/assessments`, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie, Origin: dev.base }, body: JSON.stringify(draftPayload()) });
      assert.strictEqual(own.status, 200);
    });

    await check('role enforcement on every kind of route: the wrong role gets 403 FORBIDDEN, the right one gets through', async () => {
      const cases = [
        // [method, url, body, forbidden for, allowed for]
        ['POST', '/api/reports', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['PUT', '/api/reports/x', {}, 'approver@elegrow.com', 'engineer@elegrow.com'],
        ['DELETE', '/api/reports/x', undefined, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/reports/x/ai/generate', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/projects', {}, 'approver@elegrow.com', 'engineer@elegrow.com'],
        ['PUT', '/api/projects/x', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/projects/x/days', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/projects/x/final/assemble', {}, 'approver@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/projects/x/zoho-refresh', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/sheets/print', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/ocr/soil/preview', {}, 'approver@elegrow.com', 'engineer@elegrow.com'],
        ['POST', '/api/reports/x/review/submit', {}, 'reviewer@elegrow.com', 'engineer@elegrow.com'],
        ['PUT', '/api/reports/x/review/ticks', {}, 'engineer@elegrow.com', 'reviewer@elegrow.com'],
        ['POST', '/api/reports/x/review/check', {}, 'engineer@elegrow.com', 'reviewer@elegrow.com'],
        ['POST', '/api/reports/x/review/check', {}, 'approver@elegrow.com', 'reviewer@elegrow.com'],
        ['POST', '/api/reports/x/review/approve', {}, 'reviewer@elegrow.com', 'approver@elegrow.com'],
        ['POST', '/api/reports/x/review/approve', {}, 'engineer@elegrow.com', 'approver@elegrow.com'],
        ['POST', '/api/reports/x/review/return', {}, 'engineer@elegrow.com', 'reviewer@elegrow.com'],
        ['POST', '/api/reports/x/review/return', {}, 'engineer@elegrow.com', 'approver@elegrow.com'],
        ['POST', '/api/reports/x/reissue', {}, 'reviewer@elegrow.com', 'approver@elegrow.com'],
        ['PUT', '/api/client-logo', {}, 'reviewer@elegrow.com', 'approver@elegrow.com'],
        ['DELETE', '/api/client-logo', undefined, 'engineer@elegrow.com', 'approver@elegrow.com'],
        ['PUT', '/api/reports/x/client-logo', {}, 'engineer@elegrow.com', 'approver@elegrow.com'],
        ['POST', '/api/zoho/poll', undefined, 'approver@elegrow.com', 'admin@elegrow.com'],
        ['GET', '/api/admin/users', undefined, 'approver@elegrow.com', 'admin@elegrow.com'],
        ['GET', '/api/admin/audit', undefined, 'engineer@elegrow.com', 'admin@elegrow.com'],
        ['PUT', '/api/admin/users/x', {}, 'reviewer@elegrow.com', 'admin@elegrow.com'],
        ['POST', '/api/not-a-route-yet', {}, 'approver@elegrow.com', 'admin@elegrow.com']
      ];
      for (const [method, url, body, forbidden, allowed] of cases) {
        const no = await call(method, url, body, forbidden);
        assert.deepStrictEqual([no.status, no.body && no.body.code], [403, 'FORBIDDEN'], `${method} ${url} as ${forbidden}`);
        const yes = await call(method, url, body, allowed);
        assert.notStrictEqual(yes.body && yes.body.code, 'FORBIDDEN', `${method} ${url} as ${allowed}`);
      }
      for (const email of ['engineer@elegrow.com', 'reviewer@elegrow.com', 'approver@elegrow.com']) {
        assert.strictEqual((await call('GET', '/api/reports', undefined, email)).status, 200);
        assert.strictEqual((await call('POST', '/api/assessments', draftPayload(), email)).status, 200);
      }
    });

    // ---- the review rules, on days created through the routes -------------------------------
    const project = (await call('POST', '/api/projects', { project: { projectNo: '077-D03_TS-26-AUT', clientName: 'Auth Test Ltd', zohoProjectId: 'Z-AUTH' } }, 'engineer@elegrow.com')).body;
    let dayNo = 0;
    const newDay = async (email) => {
      dayNo += 1;
      const created = await call('POST', `/api/projects/${project.projectId}/days`, { workDate: `2026-09-${String(10 + dayNo).padStart(2, '0')}`, report: draftPayload() }, email);
      assert.strictEqual(created.status, 201, JSON.stringify(created.body));
      return created.body;
    };
    const review = (id, step, body, email) => call('POST', `/api/reports/${id}/review/${step}`, body || {}, email);
    const ticksFor = async (id, email, extra = {}) => {
      const ticks = Object.fromEntries((await call('GET', `/api/reports/${id}/review`, undefined, email)).body.checklist.map((item) => [item.id, true]));
      return call('PUT', `/api/reports/${id}/review/ticks`, { ticks, ...extra }, email);
    };

    await check('ownership: a daily is edited and submitted by the engineer who created it (or an admin); other engineers are refused', async () => {
      const day = await newDay('engineer@elegrow.com');
      assert.strictEqual(day.createdBy.startsWith('U-'), true);
      const body = draftPayload({ reportDate: day.collectionDate });
      assert.strictEqual((await call('PUT', `/api/reports/${day.id}`, body, 'engineer2@elegrow.com')).body.code, 'NOT_OWNER');
      assert.strictEqual((await call('PUT', `/api/reports/${day.id}`, body, 'engineer@elegrow.com')).status, 200);
      assert.strictEqual((await call('PUT', `/api/reports/${day.id}`, body, 'admin@elegrow.com')).status, 200);
      assert.strictEqual((await review(day.id, 'submit', {}, 'engineer2@elegrow.com')).body.code, 'NOT_OWNER');
      assert.strictEqual((await call('DELETE', `/api/reports/${day.id}`, undefined, 'engineer2@elegrow.com')).body.code, 'NOT_OWNER');
      assert.strictEqual((await call('DELETE', `/api/reports/${day.id}`, undefined, 'engineer@elegrow.com')).status, 204);
    });

    await check('a report with no creator is admin-only to edit, submit or delete', async () => {
      const store = createReportStore({ filePath: dev.dataFile });
      const p = store.getProject(project.projectId);
      const old = store.addReport({ ...normalizeReportInput(draftPayload()), kind: 'daily', projectId: p.id, collectionDate: '2026-01-05' });
      assert.strictEqual(old.createdBy, undefined);
      assert.strictEqual((await call('PUT', `/api/reports/${old.id}`, { ...draftPayload(), project: { ...draftPayload().project, reportDate: '2026-01-05' } }, 'engineer@elegrow.com')).body.code, 'NOT_OWNER');
      assert.strictEqual((await review(old.id, 'submit', {}, 'engineer@elegrow.com')).body.code, 'NOT_OWNER');
      assert.strictEqual((await call('DELETE', `/api/reports/${old.id}`, undefined, 'engineer@elegrow.com')).status, 403);
      assert.strictEqual((await call('PUT', `/api/reports/${old.id}`, { ...draftPayload(), project: { ...draftPayload().project, reportDate: '2026-01-05' } }, 'admin@elegrow.com')).status, 200);
      assert.strictEqual((await review(old.id, 'submit', {}, 'admin@elegrow.com')).status, 200);
    });

    await check('the full path by three people: submit, tick and check, approve; each step is recorded against the signed-in user', async () => {
      const day = await newDay('engineer@elegrow.com');
      assert.strictEqual((await review(day.id, 'submit', { actor: 'Somebody Else', designation: 'Imposter' }, 'engineer@elegrow.com')).status, 200);
      assert.strictEqual((await ticksFor(day.id, 'reviewer@elegrow.com')).status, 200);
      assert.strictEqual((await review(day.id, 'check', { note: 'ok' }, 'reviewer@elegrow.com')).status, 200);
      const approved = await review(day.id, 'approve', {}, 'approver@elegrow.com');
      assert.strictEqual(approved.status, 200);
      assert.deepStrictEqual(approved.body.initials, { prepared: 'ENG', checked: 'AP', approved: 'JCS' }, 'the names come from the profiles, not from the request');
      const rows = approved.body.transitions;
      assert.deepStrictEqual(rows.map((row) => row.actor), ['E. N. Gineer', 'A. P.', 'J. C. Shah']);
      assert.deepStrictEqual(rows.map((row) => row.designation), ['Site Engineer', 'Reviewer', 'Director']);
      assert.ok(rows.every((row) => row.userId.startsWith('U-')), 'every transition carries a user id');
    });

    await check('self-review block: the preparer can never check or approve their own report; an engineer+reviewer cannot override', async () => {
      const day = await newDay('engrev@elegrow.com');
      await review(day.id, 'submit', {}, 'engrev@elegrow.com');
      const ticked = await ticksFor(day.id, 'engrev@elegrow.com');
      assert.deepStrictEqual([ticked.status, ticked.body.code, ticked.body.canOverride], [403, 'SELF_REVIEW_BLOCKED', false]);
      const own = await review(day.id, 'check', { overrideReason: 'I really want to check my own report' }, 'engrev@elegrow.com');
      assert.deepStrictEqual([own.status, own.body.code, own.body.canOverride], [403, 'SELF_REVIEW_BLOCKED', false], 'a typed reason does not help a non-admin');
      assert.strictEqual((await call('GET', `/api/reports/${day.id}/review`, undefined, 'engrev@elegrow.com')).body.access.check, false);
      assert.strictEqual((await ticksFor(day.id, 'reviewer@elegrow.com')).status, 200);
      assert.strictEqual((await review(day.id, 'check', {}, 'reviewer@elegrow.com')).status, 200);
      // an engineer+approver who prepared cannot approve
      const mine = await newDay('engapp@elegrow.com');
      await review(mine.id, 'submit', {}, 'engapp@elegrow.com');
      await ticksFor(mine.id, 'reviewer@elegrow.com');
      await review(mine.id, 'check', {}, 'reviewer@elegrow.com');
      const approve = await review(mine.id, 'approve', {}, 'engapp@elegrow.com');
      assert.deepStrictEqual([approve.status, approve.body.code], [403, 'SELF_REVIEW_BLOCKED']);
      assert.strictEqual((await review(mine.id, 'approve', {}, 'approver@elegrow.com')).status, 200);
    });

    await check('admin override: only with a typed reason (10+ characters); recorded on the transition and in the audit log', async () => {
      const day = await newDay('admin@elegrow.com');
      await review(day.id, 'submit', {}, 'admin@elegrow.com');
      const blocked = await ticksFor(day.id, 'admin@elegrow.com');
      assert.deepStrictEqual([blocked.status, blocked.body.code, blocked.body.canOverride], [403, 'SELF_REVIEW_BLOCKED', true]);
      assert.strictEqual((await call('GET', `/api/reports/${day.id}/review`, undefined, 'admin@elegrow.com')).body.access.needsOverride, true);
      assert.strictEqual((await review(day.id, 'check', { overrideReason: 'too short' }, 'admin@elegrow.com')).status, 403);
      assert.strictEqual((await ticksFor(day.id, 'admin@elegrow.com', { overrideReason: 'Director check: reviewers are on site' })).status, 200);
      const checked = await review(day.id, 'check', { overrideReason: 'Director check: reviewers are on site' }, 'admin@elegrow.com');
      assert.strictEqual(checked.status, 200);
      const row = checked.body.transitions.find((entry) => entry.action === 'check');
      assert.ok(row.note.includes('Admin override by H. S. Patil: Director check: reviewers are on site'), row.note);
      const audit = (await call('GET', '/api/admin/audit?action=review.override', undefined, 'admin@elegrow.com')).body;
      assert.ok(audit.entries.some((entry) => entry.entityId === day.id && entry.detail.reason.includes('Director check')));
    });

    await check('a checker who is also the approver is allowed, with a warning (small team)', async () => {
      const day = await newDay('engineer@elegrow.com');
      await review(day.id, 'submit', {}, 'engineer@elegrow.com');
      await ticksFor(day.id, 'both@elegrow.com');
      assert.strictEqual((await review(day.id, 'check', {}, 'both@elegrow.com')).status, 200);
      const approved = await review(day.id, 'approve', {}, 'both@elegrow.com');
      assert.strictEqual(approved.status, 200);
      assert.ok(approved.body.stepWarnings.some((warning) => /both checked and approved/.test(warning)), JSON.stringify(approved.body.stepWarnings));
    });

    await check('reviewers may return to draft; an engineer may not', async () => {
      const day = await newDay('engineer@elegrow.com');
      await review(day.id, 'submit', {}, 'engineer@elegrow.com');
      assert.strictEqual((await review(day.id, 'return', {}, 'engineer@elegrow.com')).status, 403);
      const returned = await review(day.id, 'return', {}, 'reviewer@elegrow.com');
      assert.strictEqual(returned.status, 200);
      assert.strictEqual(returned.body.status, 'draft');
    });

    await check('My profile: name and designation come from it; a signed report keeps what it said; the change shows next time', async () => {
      const day = await newDay('engineer@elegrow.com');
      await review(day.id, 'submit', {}, 'engineer@elegrow.com');
      assert.strictEqual((await call('PUT', '/api/profile', { name: '' }, 'engineer@elegrow.com')).status, 400);
      const changed = await call('PUT', '/api/profile', { name: 'Eng Neer-Gineer', designation: 'Principal Engineer' }, 'engineer@elegrow.com');
      assert.deepStrictEqual([changed.body.user.name, changed.body.user.designation], ['Eng Neer-Gineer', 'Principal Engineer']);
      const state = (await call('GET', `/api/reports/${day.id}/review`, undefined, 'engineer@elegrow.com')).body;
      assert.deepStrictEqual([state.transitions[0].actor, state.transitions[0].designation], ['E. N. Gineer', 'Site Engineer'], 'the signed step is not rewritten');
      const later = await newDay('engineer@elegrow.com');
      await review(later.id, 'submit', {}, 'engineer@elegrow.com');
      const second = (await call('GET', `/api/reports/${later.id}/review`, undefined, 'engineer@elegrow.com')).body;
      assert.deepStrictEqual([second.transitions[0].actor, second.transitions[0].designation], ['Eng Neer-Gineer', 'Principal Engineer']);
    });

    await check('admin: roles are editable and take effect at once; the last admin cannot be removed; a deactivated user is signed out and refused', async () => {
      const users = (await call('GET', '/api/admin/users', undefined, 'admin@elegrow.com')).body;
      const byEmail = Object.fromEntries(users.map((user) => [user.email, user]));
      assert.deepStrictEqual(byEmail['engineer2@elegrow.com'].roles, ['engineer'], 'a new sign-in starts as an engineer');
      // give engineer2 the reviewer role: they can now check
      const day = await newDay('engineer@elegrow.com');
      await review(day.id, 'submit', {}, 'engineer@elegrow.com');
      assert.strictEqual((await ticksFor(day.id, 'engineer2@elegrow.com')).status, 403);
      assert.strictEqual((await call('PUT', `/api/admin/users/${byEmail['engineer2@elegrow.com'].id}`, { roles: ['engineer', 'reviewer'] }, 'admin@elegrow.com')).status, 200);
      assert.strictEqual((await ticksFor(day.id, 'engineer2@elegrow.com')).status, 200);
      // last admin
      const admin = byEmail['admin@elegrow.com'];
      const last = await call('PUT', `/api/admin/users/${admin.id}`, { roles: ['engineer'] }, 'admin@elegrow.com');
      assert.deepStrictEqual([last.status, last.body.code], [409, 'LAST_ADMIN']);
      assert.strictEqual((await call('PUT', `/api/admin/users/${admin.id}`, { active: false }, 'admin@elegrow.com')).body.code, 'LAST_ADMIN');
      // deactivate engineer2
      await call('PUT', `/api/admin/users/${byEmail['engineer2@elegrow.com'].id}`, { active: false }, 'admin@elegrow.com');
      assert.strictEqual((await call('GET', '/api/reports', undefined, 'engineer2@elegrow.com')).status, 401, 'the session is revoked');
      const again = await fetch(`${dev.base}/auth/dev-login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: 'engineer2@elegrow.com' }) });
      assert.deepStrictEqual([again.status, (await again.json()).code], [403, 'deactivated']);
      assert.strictEqual((await call('PUT', `/api/admin/users/missing`, { roles: [] }, 'admin@elegrow.com')).status, 404);
    });

    await check('audit log: sign-ins, creates, every review step and role changes are there; admin-only and read-only', async () => {
      const all = (await call('GET', '/api/admin/audit?limit=500', undefined, 'admin@elegrow.com')).body;
      const actions = new Set(all.entries.map((entry) => entry.action));
      ['auth.signin', 'auth.forbidden', 'project.create', 'project.day_save', 'review.submit', 'review.check', 'review.approve', 'review.return', 'review.ticks', 'review.override', 'admin.user_change', 'profile.update', 'report.update', 'report.delete'].forEach((action) => assert.ok(actions.has(action), `audit lacks ${action}`));
      assert.ok(all.entries.every((entry) => entry.at && (entry.userId || entry.action.startsWith('auth.'))));
      const mine = all.entries.find((entry) => entry.action === 'review.approve');
      assert.ok(mine.userName && mine.entityType === 'report' && mine.entityId);
      const filtered = (await call('GET', '/api/admin/audit?action=review.&limit=5', undefined, 'admin@elegrow.com')).body;
      assert.ok(filtered.entries.length <= 5 && filtered.entries.every((entry) => entry.action.startsWith('review.')));
      assert.ok(filtered.total >= 8);
      assert.strictEqual((await call('GET', '/api/admin/audit', undefined, 'engineer@elegrow.com')).status, 403);
      assert.notStrictEqual((await call('DELETE', '/api/admin/audit', undefined, 'admin@elegrow.com')).status, 200, 'no route deletes the log');
      assert.notStrictEqual((await call('PUT', '/api/admin/audit', {}, 'admin@elegrow.com')).status, 200);
    });
  } finally {
    dev.stop();
  }

  // ================= session expiry ============================================================
  await check('a session ends after the idle timeout (and use keeps it alive)', async () => {
    const quick = await startServer(devEnv({ AUTH_SESSION_IDLE_HOURS: String(2.4 / 3600) }), 'idle'); // 2.4 s
    try {
      const cookie = await login(quick.base, 'engineer@elegrow.com');
      const get = () => fetch(`${quick.base}/api/reports`, { headers: { Cookie: cookie } }).then((response) => response.status);
      assert.strictEqual(await get(), 200);
      await sleep(1200);
      assert.strictEqual(await get(), 200, 'used within the window');
      await sleep(1200);
      assert.strictEqual(await get(), 200, 'the window slides with use');
      await sleep(3200);
      assert.strictEqual(await get(), 401, 'idle for longer than the timeout: signed out');
    } finally {
      quick.stop();
    }
  });

  // ================= not configured / zoho mode ================================================
  await check('with no sign-in mode configured the API stays closed and dev login does not exist', async () => {
    const closed = await startServer({ AUTH_MODE: '', ZOHO_POLL_MINUTES: '0' }, 'closed');
    try {
      const me = await (await fetch(`${closed.base}/api/auth/me`)).json();
      assert.deepStrictEqual([me.mode, me.ready, me.user], ['none', false, null]);
      assert.strictEqual((await fetch(`${closed.base}/api/reports`)).status, 401);
      assert.strictEqual((await fetch(`${closed.base}/auth/dev-login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: 'admin@elegrow.com' }) })).status, 404);
      const login302 = await fetch(`${closed.base}/auth/login`, { redirect: 'manual' });
      assert.deepStrictEqual([login302.status, login302.headers.get('location')], [302, '/?auth_error=config']);
    } finally {
      closed.stop();
    }
  });

  // a fake Zoho accounts server: the token endpoint and the profile endpoint
  const jwt = (claims) => `eyJhbGciOiJSUzI1NiJ9.${Buffer.from(JSON.stringify(claims)).toString('base64url')}.sig`;
  const accounts = http.createServer((request, response) => {
    let body = '';
    request.on('data', (chunk) => { body += chunk; });
    request.on('end', () => {
      const reply = (data, status = 200) => { response.writeHead(status, { 'Content-Type': 'application/json' }); response.end(JSON.stringify(data)); };
      if (request.url.startsWith('/oauth/v2/token')) {
        const form = new URLSearchParams(body);
        const exp = Math.floor(Date.now() / 1000) + 600;
        const people = { good: 'engineer.one@elegrow.com', seed: 'shivam@elegrow.com', outsider: 'someone@gmail.com', former: 'former@elegrow.com', stranger: 'unknown@elegrow.com' };
        if (form.get('client_id') !== 'test-client' || form.get('client_secret') !== 'test-secret') return reply({ error: 'invalid_client' }, 400);
        const code = form.get('code');
        if (people[code]) return reply({ access_token: 'at', id_token: jwt({ aud: 'test-client', exp, email: people[code], email_verified: true, name: code === 'seed' ? 'Shivam Agrawal' : 'Engineer One', sub: '600123' }) });
        if (code === 'badaud') return reply({ access_token: 'at', id_token: jwt({ aud: 'someone-else', exp, email: 'engineer.one@elegrow.com' }) });
        if (code === 'noemail') return reply({ access_token: 'at-profile', id_token: jwt({ aud: 'test-client', exp, sub: '9' }) });
        return reply({ error: 'invalid_code' }, 400);
      }
      if (request.url.startsWith('/oauth/user/info')) {
        return request.headers.authorization === 'Zoho-oauthtoken at-profile' ? reply({ Email: 'reviewer.one@elegrow.com', Display_Name: 'Reviewer One', ZUID: '777' }) : reply({}, 401);
      }
      return reply({}, 404);
    });
  });
  await new Promise((resolve) => accounts.listen(0, '127.0.0.1', resolve));
  const accountsUrl = `http://127.0.0.1:${accounts.address().port}`;
  const zohoEnv = (redirectPort, extra = {}) => ({ AUTH_MODE: 'zoho', AUTH_ZOHO_CLIENT_ID: 'test-client', AUTH_ZOHO_CLIENT_SECRET: 'test-secret', AUTH_ZOHO_ACCOUNTS_URL: accountsUrl, AUTH_SESSION_SECRET: 'another-test-secret', AUTH_SEED_ADMINS: 'shivam@elegrow.com', ZOHO_POLL_MINUTES: '0', ...extra });

  try {
    const port = await freePort();
    const app = await startServer(zohoEnv(port, { ELECTROREPORTS_PORT: String(port), AUTH_ZOHO_REDIRECT_URI: `http://127.0.0.1:${port}/auth/callback` }), 'zoho');
    try {
      const flow = async (code, { state, stateCookie } = {}) => {
        const start = await fetch(`${app.base}/auth/login`, { redirect: 'manual' });
        const location = new URL(start.headers.get('location'));
        const cookie = start.headers.getSetCookie().find((entry) => entry.startsWith('er_oauth_state=')).split(';')[0];
        const query = new URLSearchParams({ code, state: state === undefined ? location.searchParams.get('state') : state });
        const callback = await fetch(`${app.base}/auth/callback?${query}`, { redirect: 'manual', headers: stateCookie === null ? {} : { Cookie: stateCookie || cookie } });
        return { start, location, callback, redirect: callback.headers.get('location') };
      };

      await check('Sign in with Zoho: the authorise URL asks for openid email profile with a state; dev login is not available', async () => {
        const { start, location } = await flow('good');
        assert.strictEqual(start.status, 302);
        assert.strictEqual(location.origin + location.pathname, `${accountsUrl}/oauth/v2/auth`);
        assert.deepStrictEqual([location.searchParams.get('scope'), location.searchParams.get('client_id'), location.searchParams.get('response_type'), location.searchParams.get('redirect_uri')], ['openid email profile', 'test-client', 'code', `http://127.0.0.1:${port}/auth/callback`]);
        assert.ok(/^[0-9a-f]{32}$/.test(location.searchParams.get('state')));
        const stateCookie = start.headers.getSetCookie().find((entry) => entry.startsWith('er_oauth_state='));
        assert.ok(/HttpOnly/i.test(stateCookie) && /SameSite=Lax/i.test(stateCookie) && /Max-Age=600/.test(stateCookie));
        assert.strictEqual((await fetch(`${app.base}/auth/dev-login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email: 'admin@elegrow.com' }) })).status, 404);
        assert.strictEqual((await fetch(`${app.base}/api/auth/dev-users`)).status, 404);
      });

      await check('an active @elegrow.com portal user signs in: the e-mail comes from the id_token, new users are engineers, the cookie is httpOnly SameSite=Lax', async () => {
        const result = await flow('good');
        assert.deepStrictEqual([result.callback.status, result.redirect], [302, '/']);
        const cookie = result.callback.headers.getSetCookie().find((entry) => entry.startsWith('er_session='));
        assert.ok(/HttpOnly/i.test(cookie) && /SameSite=Lax/i.test(cookie) && !/Secure/i.test(cookie) && !/Max-Age/i.test(cookie), cookie);
        const me = await (await fetch(`${app.base}/api/auth/me`, { headers: { Cookie: cookie.split(';')[0] } })).json();
        assert.deepStrictEqual([me.user.email, me.user.roles, me.user.name], ['engineer.one@elegrow.com', ['engineer'], 'Engineer One']);
        assert.strictEqual((await fetch(`${app.base}/api/reports`, { headers: { Cookie: cookie.split(';')[0] } })).status, 200);
        assert.strictEqual((await fetch(`${app.base}/api/admin/users`, { headers: { Cookie: cookie.split(';')[0] } })).status, 403);
      });

      await check('the seeded admin signs in as admin (and gets the real name on first sign-in)', async () => {
        const result = await flow('seed');
        const cookie = result.callback.headers.getSetCookie().find((entry) => entry.startsWith('er_session=')).split(';')[0];
        const me = await (await fetch(`${app.base}/api/auth/me`, { headers: { Cookie: cookie } })).json();
        assert.deepStrictEqual([me.user.email, me.user.roles.includes('admin'), me.user.name], ['shivam@elegrow.com', true, 'Shivam Agrawal']);
        assert.strictEqual((await fetch(`${app.base}/api/admin/users`, { headers: { Cookie: cookie } })).status, 200);
      });

      await check('a non-elegrow account, an inactive portal user and a non-member are rejected, with no session', async () => {
        for (const [code, reason] of [['outsider', 'domain'], ['former', 'inactive'], ['stranger', 'not_in_portal']]) {
          const result = await flow(code);
          assert.strictEqual(result.redirect, `/?auth_error=${reason}`, code);
          assert.ok(!result.callback.headers.getSetCookie().some((entry) => entry.startsWith('er_session=')), `${code}: no session cookie`);
        }
      });

      await check('a wrong or missing state, a token for another client, and an error from Zoho are refused', async () => {
        assert.strictEqual((await flow('good', { state: 'forged' })).redirect, '/?auth_error=state');
        assert.strictEqual((await flow('good', { stateCookie: null })).redirect, '/?auth_error=state');
        assert.strictEqual((await flow('good', { stateCookie: 'er_oauth_state=abc.def' })).redirect, '/?auth_error=state');
        assert.strictEqual((await flow('badaud')).redirect, '/?auth_error=failed');
        assert.strictEqual((await flow('nonsense')).redirect, '/?auth_error=failed');
        const denied = await fetch(`${app.base}/auth/callback?error=access_denied&state=x`, { redirect: 'manual' });
        assert.ok(['/?auth_error=state', '/?auth_error=failed'].includes(denied.headers.get('location')));
      });

      await check('without an e-mail in the id_token the profile endpoint is the fallback', async () => {
        const result = await flow('noemail');
        assert.strictEqual(result.redirect, '/');
        const cookie = result.callback.headers.getSetCookie().find((entry) => entry.startsWith('er_session=')).split(';')[0];
        assert.strictEqual((await (await fetch(`${app.base}/api/auth/me`, { headers: { Cookie: cookie } })).json()).user.email, 'reviewer.one@elegrow.com');
      });

      await check('rejected sign-ins are in the audit log with the reason, and accepted ones too', async () => {
        const seed = await flow('seed');
        const cookie = seed.callback.headers.getSetCookie().find((entry) => entry.startsWith('er_session=')).split(';')[0];
        const audit = await (await fetch(`${app.base}/api/admin/audit?limit=200`, { headers: { Cookie: cookie } })).json();
        const rejected = audit.entries.filter((entry) => entry.action === 'auth.rejected');
        assert.deepStrictEqual(rejected.map((entry) => entry.detail.reason).filter((reason) => ['domain', 'inactive', 'not_in_portal', 'state'].includes(reason)).sort().slice(0, 3), ['domain', 'inactive', 'not_in_portal']);
        assert.ok(rejected.some((entry) => entry.entityId === 'someone@gmail.com'));
        assert.ok(audit.entries.some((entry) => entry.action === 'auth.signin' && entry.userEmail === 'engineer.one@elegrow.com'));
      });
    } finally {
      app.stop();
    }

    await check('over HTTPS the session cookie is also Secure', async () => {
      const securePort = await freePort();
      const secure = await startServer(zohoEnv(securePort, { ELECTROREPORTS_PORT: String(securePort), AUTH_ZOHO_REDIRECT_URI: 'https://reports.example.test/auth/callback' }), 'secure');
      try {
        const start = await fetch(`${secure.base}/auth/login`, { redirect: 'manual' });
        const cookie = start.headers.getSetCookie().find((entry) => entry.startsWith('er_oauth_state='));
        assert.ok(/Secure/i.test(cookie));
        const state = new URL(start.headers.get('location')).searchParams.get('state');
        const callback = await fetch(`${secure.base}/auth/callback?code=good&state=${state}`, { redirect: 'manual', headers: { Cookie: cookie.split(';')[0] } });
        assert.ok(/Secure/i.test(callback.headers.getSetCookie().find((entry) => entry.startsWith('er_session='))));
      } finally {
        secure.stop();
      }
    });
  } finally {
    accounts.close();
  }

  console.log(`\n${passed} auth checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
