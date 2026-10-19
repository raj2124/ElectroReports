// Test helper for the sign-in added in Session 6E: servers under test run with AUTH_MODE=dev and a
// fixed set of users, and requests carry the session cookie of the user they act as.
// Users: email:roles:Name:Designation. The names give the initials the report tests expect.
const USERS = [
  'admin@elegrow.com:admin:H. S. Patil:Head - Earthing Services',
  'engineer@elegrow.com:engineer:E. N. Gineer:Site Engineer',
  'engineer2@elegrow.com:engineer:S. Econd:Site Engineer',
  'reviewer@elegrow.com:reviewer:A. P.:Reviewer',
  'approver@elegrow.com:approver:J. C. Shah:Director',
  'both@elegrow.com:reviewer+approver:B. O. Th:Reviewer and Approver',
  'engrev@elegrow.com:engineer+reviewer:E. R.:Engineer and Reviewer',
  'engapp@elegrow.com:engineer+approver:E. A.:Engineer and Approver',
  'anup@elegrow.com:engineer:Anup Parekh:Engineer',
  'nidhi@elegrow.com:engineer:Nidhi Shah:Engineer',
  'karishma@elegrow.com:engineer:Karishma Prajapati:Engineer'
].join(',');

// Environment for a server under test.
function devEnv(extra = {}) {
  return { AUTH_MODE: 'dev', AUTH_DEV_USERS: USERS, AUTH_SESSION_SECRET: 'test-secret-not-used-anywhere-else', AUTH_SEED_ADMINS: '', ZOHO_POLL_MINUTES: '0', ...extra };
}

// Signs in and returns the cookie header value ("er_session=...").
async function login(baseUrl, email) {
  const response = await fetch(`${baseUrl}/auth/dev-login`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email }) });
  if (!response.ok) {
    throw new Error(`dev login of ${email} failed: ${response.status}`);
  }
  const cookie = response.headers.getSetCookie().find((entry) => entry.startsWith('er_session='));
  return cookie.split(';')[0];
}

// call(method, url, body, as) with one session per user, signed in on first use.
function createCaller(baseUrl, defaultUser = 'admin@elegrow.com') {
  const cookies = {};
  return async (method, url, body, as = defaultUser) => {
    if (as && !cookies[as]) {
      cookies[as] = await login(baseUrl, as);
    }
    const response = await fetch(baseUrl + url, { method, headers: { 'Content-Type': 'application/json', ...(as ? { Cookie: cookies[as] } : {}) }, body: body === undefined ? undefined : JSON.stringify(body) });
    const type = response.headers.get('content-type') || '';
    return { status: response.status, body: type.includes('json') ? await response.json() : Buffer.from(await response.arrayBuffer()), type, headers: response.headers };
  };
}

module.exports = { USERS, devEnv, login, createCaller };

// Browser tests: signs the Playwright context in as `email` (the session cookie goes into the
// context's cookie jar). Reload the page afterwards so the screen shows that user.
async function browserLogin(context, baseUrl, email) {
  const response = await context.request.post(`${baseUrl}/auth/dev-login`, { data: { email } });
  if (!response.ok()) {
    throw new Error(`dev login of ${email} failed: ${response.status()}`);
  }
}

module.exports.browserLogin = browserLogin;
