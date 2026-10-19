// Sign-in, My profile and the admin screens in a real browser (Session 6E), in dev auth mode.
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');
const { devEnv, browserLogin } = require('./lib/dev-auth');

const ROOT = path.join(__dirname, '..');
const SHOTS = path.join(ROOT, 'reference', 'screenshots');
let checks = 0;
const failures = [];
function expect(condition, message) {
  checks += 1;
  if (!condition) {
    failures.push(message);
    console.log(`  FAIL - ${message}`);
  }
}
const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'auth-ui-'));
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const browser = await chromium.launch();
  try {
    fs.mkdirSync(SHOTS, { recursive: true });
    const context = await browser.newContext({ viewport: { width: 1288, height: 900 } });
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));

    // signed out: only the sign-in screen
    await page.goto(base);
    await page.waitForSelector('.signin-card');
    expect(await page.locator('.project-card, [data-action="new-report"]').count() === 0, 'signed out: no app screens, only sign-in');
    expect((await page.innerText('.signin-card')).includes('Sign in'), 'the sign-in screen asks to sign in');
    await page.screenshot({ path: path.join(SHOTS, 'Signin_1288.png') });

    // the dev picker signs in
    await page.selectOption('[data-dev-user]', 'engineer@elegrow.com');
    await page.waitForSelector('.user-bar');
    expect((await page.innerText('.user-bar')).includes('E. N. Gineer') && (await page.innerText('.user-bar')).includes('Engineer'), 'the header shows the name and roles');
    expect(await page.locator('[data-action="open-admin"]').count() === 0, 'an engineer has no Admin button');
    expect(await page.locator('[data-action="new-report"]').count() > 0, 'the app is available once signed in');

    // My profile
    await page.click('[data-action="open-profile"]');
    await page.waitForSelector('.profile-card');
    await page.fill('[data-profile-field="name"]', 'Eng Gineer');
    await page.fill('[data-profile-field="designation"]', 'Principal Engineer');
    await page.click('[data-action="save-profile"]');
    await page.waitForFunction(() => /Profile saved/.test(document.querySelector('.profile-card').innerText));
    expect((await page.innerText('.user-bar')).includes('Eng Gineer'), 'the profile change shows in the header');
    expect(await page.evaluate(() => authUi.user.designation) === 'Principal Engineer', 'the designation is saved');
    await page.screenshot({ path: path.join(SHOTS, 'Profile_1288.png') });

    // an engineer is refused admin routes by the server even if the screen were opened by hand
    const refused = await page.evaluate(async () => (await fetch('/api/admin/users')).status);
    expect(refused === 403, 'the server refuses admin routes to an engineer');

    // admin: users, roles, audit
    await browserLogin(context, base, 'admin@elegrow.com');
    await page.goto(base);
    await page.waitForSelector('[data-action="open-admin"]');
    await page.click('[data-action="open-admin"]');
    await page.waitForSelector('.admin-users tbody tr');
    expect(await page.locator('.admin-users tbody tr').count() >= 2, 'the admin sees the users');
    const row = page.locator('.admin-users tbody tr', { hasText: 'Eng Gineer' });
    expect(await row.count() === 1, 'the user who signed in is listed (with the edited name)');
    await row.locator('[data-admin-role$=":reviewer"]').check();
    await page.waitForFunction(() => /Saved Eng Gineer/.test(document.querySelector('.admin-card').innerText));
    expect(await page.evaluate(() => authUi.users.find((user) => user.email === 'engineer@elegrow.com').roles.includes('reviewer')), 'a role is granted from the screen');
    await page.screenshot({ path: path.join(SHOTS, 'Admin_users_1288.png'), fullPage: true });
    await page.click('[data-action="admin-tab"][data-tab="audit"]');
    await page.waitForSelector('.admin-card table:not(.admin-users)');
    const audit = await page.innerText('.admin-card');
    expect(audit.includes('auth.signin') && audit.includes('admin.user_change') && audit.includes('profile.update') && /Read only/.test(audit), 'the audit view lists sign-ins, role and profile changes, read only');
    expect(await page.locator('.admin-card [data-action*="delete"], .admin-card button:has-text("Delete")').count() === 0, 'the audit view has no way to delete');
    await page.screenshot({ path: path.join(SHOTS, 'Admin_audit_1288.png'), fullPage: true });

    // sign out
    await page.click('[data-action="sign-out"]');
    await page.waitForSelector('.signin-card');
    expect(await page.evaluate(async () => (await fetch('/api/reports')).status) === 401, 'after signing out the API is closed');

    // a session that ends while the app is open sends the user back to sign-in
    await page.selectOption('[data-dev-user]', 'engineer@elegrow.com');
    await page.waitForSelector('.user-bar');
    await page.evaluate(() => fetch('/auth/logout', { method: 'POST' })); // ends the session behind the screen's back
    await page.evaluate(() => loadReports());
    await page.waitForSelector('.signin-card');
    expect(true, 'an API 401 shows the sign-in screen');
    expect(errors.length === 0, `no page errors (${errors.join('; ')})`);
  } finally {
    await browser.close();
    server.kill();
  }
  console.log(`\nauth-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} failed`);
    process.exit(1);
  }
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
