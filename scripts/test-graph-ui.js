// Session 6H in a real browser, on the mock Zoho portal and the mock Graph: the /?project=<id> deep link (signed in,
// before sign-in, unknown id), the admin panels for the calendar sync and the knowledge base, the document search,
// and the reviewer's references on a final report. Needs Playwright's Chromium.
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');
const { devEnv, browserLogin } = require('./lib/dev-auth');
const { buildPayload } = require('./fixtures/golden-esha');

const ROOT = path.join(__dirname, '..');
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'graph-ui-'));
  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', GRAPH_USE_MOCK: 'true', GRAPH_KB_SITES: '', GRAPH_KB_REFRESH_MINUTES: '0', APP_BASE_URL: 'https://reports.example.test', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const browser = await chromium.launch();
  try {
    const base = `http://127.0.0.1:${port}`;
    const context = await browser.newContext({ viewport: { width: 1288, height: 900 } });
    await browserLogin(context, base, 'admin@elegrow.com');
    const post = async (url, data = {}) => (await context.request.post(`${base}${url}`, { data })).json();

    // data: a project with one day and a final report, the index and the Zoho tasks
    const { projectId } = await post('/api/projects', { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' });
    const payload = buildPayload();
    await post(`/api/projects/${projectId}/days`, { workDate: '2026-10-12', report: { ...payload, project: { ...payload.project, reportDate: '2026-10-12' } } });
    const { report: final } = await post(`/api/projects/${projectId}/final/assemble`);
    await post('/api/admin/kb/refresh');
    await post('/api/schedule/refresh');

    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    const view = () => page.evaluate(() => state.view);

    // ---- deep link, signed in
    await page.goto(`${base}/?project=${encodeURIComponent(projectId)}`);
    await page.waitForFunction(() => state.view === 'workspace');
    expect((await page.evaluate(() => workspaceUi.workspace.project.projectNo)) === '024-D03_TS-26-GFL', 'the link /?project=<workspace id> opens that workspace');
    expect(new URL(page.url()).search === '', 'the id is taken out of the address bar');

    // ---- deep link with an id nobody has
    await page.goto(`${base}/?project=does-not-exist`);
    await page.waitForSelector('.toast');
    expect((await page.innerText('.toast')).includes('could not be found') && (await view()) === 'dashboard', 'an unknown project gives a notice and the dashboard');

    // ---- deep link before sign-in: kept for the tab, used after sign-in
    const fresh = await browser.newContext({ viewport: { width: 1288, height: 900 } });
    const tab = await fresh.newPage();
    await tab.goto(`${base}/?project=${encodeURIComponent(projectId)}`);
    await tab.waitForFunction(() => state.view === 'signin');
    expect(await tab.evaluate(() => sessionStorage.getItem('er_deeplink')) === projectId, 'the link is kept while signed out');
    await browserLogin(fresh, base, 'engineer@elegrow.com');
    await tab.goto(`${base}/`);
    await tab.waitForFunction(() => state.view === 'workspace');
    expect(await tab.evaluate(() => sessionStorage.getItem('er_deeplink')) === null, 'and used once');
    await fresh.close();

    // ---- admin: calendar sync
    await page.goto(`${base}/`);
    await page.waitForSelector('.user-bar');
    await page.click('[data-action="open-admin"]');
    await page.click('[data-tab="calendar"]');
    await page.waitForSelector('[data-calendar-setting="enabled"]');
    expect(!(await page.isChecked('[data-calendar-setting="enabled"]')) && (await page.isChecked('[data-calendar-setting="dryRun"]')), 'calendar sync is off and a dry run by default');
    expect(await page.isDisabled('[data-action="calendar-run"]'), 'Run now is disabled while it is off');
    await page.click('[data-action="calendar-preview"]');
    await page.waitForSelector('[data-calendar-summary]');
    const summary = await page.innerText('[data-calendar-summary]');
    expect(/^Preview/.test(summary) && /to create/.test(summary), `a preview is shown (${summary.slice(0, 60)})`);
    expect((await context.request.get(`${base}/api/admin/calendar-sync`).then((response) => response.json())).owned === 0, 'a preview writes nothing');
    await page.check('[data-calendar-setting="enabled"]');
    await page.waitForSelector('[data-action="calendar-run"]:not([disabled])');
    expect(!(await page.isDisabled('[data-action="calendar-run"]')), 'switching it on enables Run now');

    // ---- admin: knowledge base
    await page.click('[data-tab="kb"]');
    await page.waitForSelector('[data-kb-sites]');
    const rows = await page.locator('[data-kb-sites] tbody tr').allInnerTexts();
    expect(rows.length === 3 && rows.some((row) => /ETPL_OP_26_27/.test(row)) && rows.some((row) => /FY23-24-StandardsGuideline.*Shared Documents\/ETPL Knowledge Hub\/02 Technical Knowledge/s.test(row)), 'the three configured sites are listed with the knowledge hub folder');
    await page.click('[data-action="kb-refresh"]');
    await page.waitForSelector('[data-kb-message]:has-text("Refresh done")');
    expect(/\t[1-9]/.test((await page.locator('[data-kb-sites] tbody tr').first().innerText())), 'after a refresh the first site has indexed files');

    // ---- search
    await page.click('[data-action="go-dashboard"]');
    await page.click('[data-action="open-kb"]');
    await page.fill('[data-kb-field="query"]', 'resistance');
    await page.press('[data-kb-field="query"]', 'Enter');
    await page.waitForSelector('[data-kb-hits] .kb-hit');
    const hits = await page.locator('.kb-hit').allInnerTexts();
    expect(hits.length >= 3 && hits.some((hit) => /Earthing practice notes/.test(hit)) && hits.some((hit) => /GFL ESHA report/.test(hit)), 'the admin finds general and client documents');
    expect((await page.locator('.kb-hit a').first().getAttribute('href')).startsWith('https://'), 'each hit links to its SharePoint file');
    await page.fill('[data-kb-field="query"]', 'zzzznotthere');
    await page.click('[data-action="kb-search"]');
    await page.waitForSelector('p[data-kb-hits]');
    expect(/No documents found/.test(await page.innerText('p[data-kb-hits]')), 'no hits says so');

    // ---- references on the final report
    await page.evaluate((id) => openReview(id), final.id);
    await page.waitForSelector('[data-action="open-references"]');
    expect((await page.innerText('[data-action="open-references"]')).trim() === 'References', 'the review header has a References button (no count yet)');
    await page.click('[data-action="open-references"]');
    await page.waitForSelector('[data-references]');
    expect((await page.innerText('[data-reference-list]')).includes('No references chosen'), 'a new final report has no references');
    await page.fill('[data-ref-field="query"]', 'resistance');
    await page.click('[data-references] [data-action="ref-search"]');
    await page.waitForSelector('[data-ref-hits] .kb-hit');
    const offered = await page.locator('[data-ref-hits] .kb-hit').allInnerTexts();
    expect(offered.some((hit) => /GFL ESHA report/.test(hit)) && offered.some((hit) => /Earthing practice notes/.test(hit)) && !offered.some((hit) => /APL reference/.test(hit)), "only this project's documents and the general base are offered");
    await page.locator('[data-ref-hits] .kb-hit', { hasText: 'GFL ESHA report' }).locator('[data-action="ref-add"]').click();
    await page.waitForSelector('[data-reference-list] .kb-hit');
    expect((await page.innerText('[data-reference-list]')).includes('GFL ESHA report'), 'the reference is listed');
    await page.click('[data-reference-list] [data-action="ref-remove"]');
    await page.waitForSelector('p[data-reference-list]');
    expect((await page.innerText('[data-references]')).includes('No references chosen'), 'and can be removed');
    await page.locator('[data-ref-hits] .kb-hit', { hasText: 'Earthing practice notes' }).locator('[data-action="ref-add"]').click();
    await page.waitForSelector('[data-reference-list] .kb-hit');
    await page.click('[data-action="ref-back"]');
    await page.waitForSelector('[data-action="open-references"]:has-text("(1)")');
    expect((await page.innerText('[data-action="open-references"]')).includes('(1)'), 'back on the review screen the button shows the count');
    expect(errors.length === 0, `no page errors (${errors.join('; ')})`);
  } finally {
    await browser.close();
    server.kill();
  }
  console.log(`\ngraph-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} FAILED`);
    process.exit(1);
  }
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
