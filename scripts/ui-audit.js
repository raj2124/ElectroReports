// Session 6J: UI audit harness (read-only; not part of npm test). Boots the app on the mocks, seeds a project with a daily and a final
// report and some scanned sheets, then visits every screen at four widths and records
//   - axe-core results (WCAG 2.0/2.1/2.2 A+AA and best practice), and
//   - our own checks: horizontal overflow, small touch targets, tiny text, clickable things the keyboard cannot reach,
//     missing focus indicator, clipped text, unlabeled controls, broken images, duplicate ids, console errors, failed requests.
// Results go to reference/ui-audit/results.json and screenshots to reference/screenshots/audit/ (git-ignored). Needs Playwright's Chromium.
//   node scripts/ui-audit.js [--widths=360,768,1280,1920] [--only=substring]
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');
const AxeBuilder = require('@axe-core/playwright').default;
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { devEnv, browserLogin, login } = require('./lib/dev-auth');
const { fillIssued, renderSheetsPdf } = require('./lib/scan-fixtures');

const ROOT = path.join(__dirname, '..');
const OUT = path.join(ROOT, 'reference', 'ui-audit');
const SHOTS = path.join(ROOT, 'reference', 'screenshots', 'audit');
const arg = (name, fallback) => (process.argv.find((a) => a.startsWith(`--${name}=`)) || '').split('=')[1] || fallback;
const WIDTHS = arg('widths', '360,768,1280,1920').split(',').map(Number);
const ONLY = arg('only', '');
const HEIGHT = { 360: 740, 768: 1024, 1280: 800, 1920: 1080 };

const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});

const PROJECT = { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' };
const ROWS = [
  { tag: 'P-1', location: 'Yard', electrodeType: 'Rod', materialType: 'GI', length: '3', diameter: '20', resistanceWithGrid: '0.52', method: '61.8%', conditionCodes: ['C1', 'C6'] },
  { tag: 'P-2', location: 'Gate', electrodeType: 'Plate', materialType: 'Cu', length: '0.6', diameter: '3', resistanceWithGrid: '1.20', method: '61.8%' }
];

// ---- in-page checks (serialised into the page, so no closures) ----
function pageChecks({ narrow }) {
  const describe = (el) => {
    const parts = [el.tagName.toLowerCase()];
    if (el.id) parts.push(`#${el.id}`);
    const cls = [...el.classList].slice(0, 2).join('.');
    if (cls) parts.push(`.${cls}`);
    const da = el.getAttribute('data-action');
    if (da) parts.push(`[data-action=${da}]`);
    const text = (el.getAttribute('aria-label') || el.innerText || el.value || '').trim().replace(/\s+/g, ' ').slice(0, 30);
    const where = el.matches('input') ? ` in <${el.parentElement.tagName.toLowerCase()}${el.parentElement.className ? '.' + el.parentElement.className.toString().split(' ')[0] : ''}>` : '';
    return text ? `${parts.join('')} "${text}"${where}` : parts.join('') + where;
  };
  const visible = (el) => {
    const r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) return false;
    if (el.closest('details:not([open])') && !el.closest('summary')) return false; // content of a closed <details> is not shown
    const s = getComputedStyle(el);
    return s.visibility !== 'hidden' && s.display !== 'none' && Number(s.opacity) > 0;
  };
  const inScroller = (el) => {
    for (let p = el.parentElement; p && p !== document.body; p = p.parentElement) {
      const s = getComputedStyle(p);
      if (/(auto|scroll)/.test(s.overflowX) && p.scrollWidth > p.clientWidth) return true;
      if (/(hidden|clip)/.test(s.overflowX) && p.getBoundingClientRect().right <= innerWidth + 1) return true;
    }
    return false;
  };
  const out = {};
  const vw = document.documentElement.clientWidth;
  out.pageOverflowPx = document.documentElement.scrollWidth - vw;
  const over = [...document.querySelectorAll('body *')].filter((el) => visible(el) && !inScroller(el) && el.getBoundingClientRect().right > vw + 1 && getComputedStyle(el).position !== 'fixed');
  out.overflowing = over.filter((el) => !over.some((other) => other !== el && el.contains(other))).slice(0, 6).map((el) => `${describe(el)} right=${Math.round(el.getBoundingClientRect().right)}`);

  const interactive = [...document.querySelectorAll('button, a[href], input:not([type=hidden]), select, textarea, summary, [role=button], [role=tab], [role=checkbox], [tabindex]:not([tabindex="-1"])')].filter(visible);
  const sized = (limit) => interactive.filter((el) => {
    const target = el.matches('input[type=checkbox], input[type=radio]') && el.closest('label') ? el.closest('label') : el;
    const r = target.getBoundingClientRect();
    return r.width < limit || r.height < limit;
  });
  out.targetsUnder24 = sized(24).slice(0, 12).map((el) => `${describe(el)} ${Math.round(el.getBoundingClientRect().width)}x${Math.round(el.getBoundingClientRect().height)}`);
  out.targetsUnder24Count = sized(24).length;
  if (narrow) {
    out.targetsUnder44Count = sized(44).length;
    out.targetsUnder44 = sized(44).slice(0, 8).map((el) => `${describe(el)} ${Math.round(el.getBoundingClientRect().width)}x${Math.round(el.getBoundingClientRect().height)}`);
  }
  out.interactiveCount = interactive.length;

  const tiny = [];
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    if (!n.textContent.trim()) continue;
    const el = n.parentElement;
    if (!el || !visible(el) || el.closest('script, style, noscript')) continue;
    const size = parseFloat(getComputedStyle(el).fontSize);
    if (size < 11) tiny.push(`${describe(el)} ${size}px`);
  }
  out.tinyTextCount = tiny.length;
  out.tinyText = [...new Set(tiny)].slice(0, 8);
  const small = [];
  const w2 = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let n = w2.nextNode(); n; n = w2.nextNode()) {
    if (!n.textContent.trim()) continue;
    const el = n.parentElement;
    if (!el || !visible(el) || el.closest('script, style, noscript')) continue;
    if (parseFloat(getComputedStyle(el).fontSize) < 12) small.push(el);
  }
  out.under12Count = small.length;

  out.unreachable = [...document.querySelectorAll('[data-action], [onclick]')].filter((el) => visible(el) && !el.matches('button, a[href], input, select, textarea, summary, [tabindex], label') && !el.closest('button, a[href], label')).slice(0, 8).map(describe);
  const labelled = (el) => el.getAttribute('aria-label') || el.getAttribute('aria-labelledby') || el.getAttribute('title') || (el.labels && el.labels.length) || el.closest('label');
  out.unlabeledControls = [...document.querySelectorAll('input:not([type=hidden]):not([type=checkbox]):not([type=radio]), select, textarea')].filter((el) => visible(el) && !labelled(el)).slice(0, 10).map((el) => `${describe(el)} placeholder="${el.placeholder || ''}"`);
  out.unlabeledCount = [...document.querySelectorAll('input:not([type=hidden]), select, textarea')].filter((el) => visible(el) && !labelled(el)).length;
  out.nameless = [...document.querySelectorAll('button, a[href]')].filter((el) => visible(el) && !(el.innerText || '').trim() && !el.getAttribute('aria-label') && !el.getAttribute('title') && !el.querySelector('img[alt]:not([alt=""])')).slice(0, 8).map(describe);
  out.brokenImages = [...document.images].filter((img) => visible(img) && img.complete && img.naturalWidth === 0).map((img) => img.src.slice(0, 80));
  out.noAlt = [...document.images].filter((img) => visible(img) && !img.hasAttribute('alt')).length;
  const ids = {};
  document.querySelectorAll('[id]').forEach((el) => { ids[el.id] = (ids[el.id] || 0) + 1; });
  out.duplicateIds = Object.entries(ids).filter(([, n]) => n > 1).map(([id, n]) => `${id} x${n}`).slice(0, 8);
  out.clipped = [...document.querySelectorAll('body *')].filter((el) => {
    if (!visible(el) || el.children.length > 0 || !(el.textContent || '').trim()) return false;
    const s = getComputedStyle(el);
    return /(hidden|clip)/.test(s.overflowX) && el.scrollWidth > el.clientWidth + 2 && s.textOverflow !== 'ellipsis' && !el.matches('input, textarea, select');
  }).slice(0, 8).map((el) => `${describe(el)} ${el.scrollWidth}>${el.clientWidth}`);
  out.ellipsisNoTitle = [...document.querySelectorAll('body *')].filter((el) => visible(el) && getComputedStyle(el).textOverflow === 'ellipsis' && el.scrollWidth > el.clientWidth + 1 && !el.title && !el.closest('[title]')).slice(0, 6).map(describe);
  out.h1Count = [...document.querySelectorAll('h1')].filter(visible).length;
  out.headings = [...document.querySelectorAll('h1, h2, h3, h4, h5, h6')].filter(visible).map((h) => Number(h.tagName[1]));
  out.hasLang = Boolean(document.documentElement.lang);
  out.hasViewport = Boolean(document.querySelector('meta[name=viewport]'));
  out.hasMain = Boolean(document.querySelector('main, [role=main]'));
  out.title = document.title;
  return out;
}

// Tabs through up to `limit` stops and reports the ones with no visible change on focus. Desktop width only.
async function focusAudit(page, limit = 40) {
  await page.evaluate(() => { document.activeElement && document.activeElement.blur(); window.scrollTo(0, 0); });
  const stops = [];
  const seen = new Set();
  for (let i = 0; i < limit; i += 1) {
    await page.keyboard.press('Tab');
    const info = await page.evaluate(() => {
      const el = document.activeElement;
      if (!el || el === document.body) return null;
      const sig = () => { const s = getComputedStyle(el); return [s.outlineStyle, s.outlineWidth, s.outlineColor, s.boxShadow, s.borderColor, s.backgroundColor, s.textDecorationLine, s.color].join('|'); };
      const withFocus = sig();
      const r = el.getBoundingClientRect();
      const key = `${el.tagName}|${el.getAttribute('data-action') || ''}|${el.className}|${el.getAttribute('aria-label') || (el.innerText || el.value || '').slice(0, 20)}`;
      el.blur();
      const without = sig();
      el.focus({ focusVisible: true });
      const s = getComputedStyle(el);
      return { key, changed: withFocus !== without, outline: `${s.outlineStyle} ${s.outlineWidth}`, hidden: r.width === 0 || r.height === 0 || getComputedStyle(el).visibility === 'hidden' };
    });
    if (!info) continue;
    if (seen.has(info.key)) continue;
    seen.add(info.key);
    stops.push(info);
  }
  return { stops: stops.length, noIndicator: stops.filter((s) => !s.changed).map((s) => s.key.slice(0, 80)).slice(0, 10), noIndicatorCount: stops.filter((s) => !s.changed).length, hiddenStops: stops.filter((s) => s.hidden).map((s) => s.key.slice(0, 60)).slice(0, 5) };
}

(async () => {
  fs.mkdirSync(OUT, { recursive: true });
  fs.mkdirSync(SHOTS, { recursive: true });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ui-audit-'));
  const dataFile = path.join(dir, 'reports.json');
  const store = createReportStore({ filePath: dataFile });
  const project = store.addProject(PROJECT);
  const payload = buildPayload();
  const normalized = normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: '2026-09-14' } });
  store.addReport({ ...normalized, kind: 'daily', projectId: project.id, collectionDate: '2026-09-14' });
  store.assembleFinal(project.id);

  const truthFile = path.join(dir, 'truth.json');
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: dataFile, ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', GRAPH_USE_MOCK: 'true', OCR_USE_MOCK: 'true', OCR_MOCK_TRUTH_FILE: truthFile, AI_CHECKS_USE_MOCK: 'true', GEMINI_API_KEY: '', OPENAI_API_KEY: '', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });

  // scanned sheets for the scan screens (uploaded through the API once)
  const cookie = await login(base, 'admin@elegrow.com');
  const api = async (method, url, body) => (await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: body ? JSON.stringify(body) : undefined })).json();
  const scanProject = (await api('POST', '/api/projects', PROJECT)).projectId;
  const issue = async (sheetKey, workDate) => {
    const response = await fetch(`${base}/api/sheets/print`, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: JSON.stringify({ project: PROJECT, workDate, sheets: [sheetKey] }) });
    return api('GET', `/api/sheets/issued/${response.headers.get('x-sheet-ids').split(',')[0]}`);
  };
  let scanReady = false;
  try {
    const electrode = await fillIssued(await issue('electrodeResistance', '2026-10-12'), { instrument: 'mi3290', serialNo: 'SN77', rows: ROWS });
    fs.writeFileSync(truthFile, JSON.stringify({ [electrode.issued.sheetId]: { primary: { ...electrode.truth, 'r0.resistanceWithGrid': { value: '0.5?', confidence: 0.55 } } } }));
    scanReady = { pdf: (await renderSheetsPdf([electrode.ctx])).buffer };
  } catch (error) {
    console.log(`scan fixture failed: ${error.message}`);
  }

  const goHome = async (page) => { await page.goto(`${base}/`); await page.waitForSelector('.user-bar, [data-view-signin], .signin', { timeout: 10000 }).catch(() => {}); await page.waitForTimeout(250); };
  const click = async (page, selector) => { await page.locator(selector).first().click({ timeout: 6000 }); await page.waitForTimeout(350); };
  const projectCardOpen = `[data-action="open-workspace"][data-id="${project.id}"]`;
  const setStep = async (page, index) => page.evaluate((i) => { state.stepIndex = i; render(); }, index);
  const sheetStepCount = async (page) => page.evaluate(() => getSteps().length);

  // screens: [name, role, async setup(page)]
  const ALL = ['admin@elegrow.com', 'engineer@elegrow.com', 'reviewer@elegrow.com', 'approver@elegrow.com'];
  const rolesFor = (name) => (/^(builder|print|scan)/.test(name) ? ['admin@elegrow.com', 'engineer@elegrow.com'] : /^(admin-|schedule-settings)/.test(name) ? ['admin@elegrow.com'] : ALL);
  const screens = [
    ['signin', null, async (page) => { await page.goto(`${base}/`); await page.waitForTimeout(600); }],
    ['dashboard', 'admin@elegrow.com', async (page) => { await goHome(page); await page.waitForSelector('.dashboard, [data-action="new-report"]'); }],
    ['dashboard-search-empty', 'admin@elegrow.com', async (page) => { await goHome(page); const s = page.locator('input[type=search], input[data-bind="search"], [data-search]').first(); if (await s.count()) { await s.fill('zzzz-no-match'); await page.waitForTimeout(300); } }],
    ['builder-project', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="new-report"]'); await page.waitForSelector('[data-bind="project.projectNo"]'); }],
    ['builder-selection', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="new-report"]'); await page.waitForSelector('[data-bind="project.projectNo"]'); await setStep(page, 1); }],
    ['detail', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-report"]'); await page.waitForSelector('.detail-shell, [data-action="export-pdf"], .surface'); }],
    ['workspace', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, projectCardOpen); await page.waitForFunction(() => state.view === 'workspace'); }],
    ['review-final', 'admin@elegrow.com', async (page) => { await goHome(page); await page.waitForSelector('[data-action="open-review"]'); const id = await page.evaluate(() => state.reports.find((r) => r.kind === 'final').id); await click(page, `[data-action="open-review"][data-id="${id}"]`); await page.waitForSelector('.review-shell'); }],
    ['review-daily', 'admin@elegrow.com', async (page) => { await goHome(page); await page.waitForSelector('[data-action="open-review"]'); const id = await page.evaluate(() => state.reports.find((r) => r.kind === 'daily').id); await click(page, `[data-action="open-review"][data-id="${id}"]`); await page.waitForSelector('.review-shell'); }],
    ['references', 'admin@elegrow.com', async (page) => { await goHome(page); const id = await page.evaluate(() => state.reports.find((r) => r.kind === 'final').id); await click(page, `[data-action="open-review"][data-id="${id}"]`); await click(page, '[data-action="open-references"]'); await page.waitForFunction(() => state.view === 'references'); }],
    ['schedule', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-schedule"]'); await page.waitForFunction(() => state.view === 'schedule'); await page.waitForTimeout(600); }],
    ['schedule-settings', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-schedule"]'); await page.waitForTimeout(600); await click(page, '[data-action="schedule-toggle-settings"]'); }],
    ['kb', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-kb"]'); await page.waitForFunction(() => state.view === 'kb'); }],
    ['profile', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-profile"]'); await page.waitForFunction(() => state.view === 'profile'); }],
    ...['users', 'audit', 'people', 'calendar', 'kb'].map((tab) => [`admin-${tab}`, 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-admin"]'); await page.waitForFunction(() => state.view === 'admin'); await click(page, `[data-action="admin-tab"][data-tab="${tab}"]`); await page.waitForTimeout(500); }]),
    ['print-sheets-dialog', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="new-report"]'); await page.locator('[data-action="open-print-sheets"]').click({ force: true }); await page.waitForSelector('.print-dialog-backdrop'); }],
    ['scans-list', 'admin@elegrow.com', async (page) => {
      await goHome(page); await click(page, '[data-action="open-scans"]'); await page.waitForSelector('[data-scans]');
      if (scanReady && !(await page.locator('.scan-item').count())) {
        await page.setInputFiles('[data-scan-files]', { name: 'visit.pdf', mimeType: 'application/pdf', buffer: scanReady.pdf });
        await page.click('[data-action="scan-upload"]');
        await page.waitForSelector('.scan-item[data-scan-status="read"]', { timeout: 20000 });
      }
    }],
    ['scan-review', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, '[data-action="open-scans"]'); await page.waitForSelector('.scan-item[data-scan-status="read"]', { timeout: 15000 }); await click(page, '.scan-item[data-scan-status="read"] [data-action="scan-open"]'); await page.waitForSelector('[data-scan-review]'); await page.waitForTimeout(500); }],
    // builder with real data: edit the seeded site day, then every sheet step
    ['builder-day-project', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, projectCardOpen); await page.waitForFunction(() => state.view === 'workspace'); await click(page, '[data-action="workspace-edit-day"]'); await page.waitForFunction(() => state.view === 'builder'); }]
  ];
  // one screen per sheet step of the seeded day (built dynamically below once the step count is known)
  const dayStepScreens = [];
  const probe = await chromium.launch();
  try {
    const pctx = await probe.newContext({ viewport: { width: 1280, height: 800 } });
    await browserLogin(pctx, base, 'admin@elegrow.com');
    const pp = await pctx.newPage();
    await goHome(pp); await click(pp, projectCardOpen); await click(pp, '[data-action="workspace-edit-day"]');
    await pp.waitForFunction(() => state.view === 'builder');
    const steps = await pp.evaluate(() => getSteps().map((s) => s.id));
    steps.forEach((id, index) => {
      if (index < 2) return;
      dayStepScreens.push([`builder-sheet-${id}`, 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, projectCardOpen); await page.waitForFunction(() => state.view === 'workspace'); await click(page, '[data-action="workspace-edit-day"]'); await page.waitForFunction(() => state.view === 'builder'); await setStep(page, index); await page.waitForTimeout(300); }]);
    });
  } catch (error) {
    console.log(`could not enumerate builder steps: ${error.message}`);
  }
  await probe.close();
  const extraTests = ['continuityTest', 'earthContinuityTest', 'towerFootingResistance'];
  const enableTests = (id) => async (page) => { await goHome(page); await click(page, projectCardOpen); await page.waitForFunction(() => state.view === 'workspace'); await click(page, '[data-action="workspace-edit-day"]'); await page.waitForFunction(() => state.view === 'builder'); await page.evaluate((t) => { state.draft.tests[t] = true; const i = getSteps().findIndex((x) => x.id === t); state.stepIndex = i; render(); }, id); await page.waitForTimeout(400); };
  extraTests.forEach((id) => dayStepScreens.push([`builder-sheet-${id}`, 'admin@elegrow.com', enableTests(id)]));
  screens.push(...dayStepScreens);
  // observation drawer on the first sheet step that has one
  screens.push(['builder-observation-drawer', 'admin@elegrow.com', async (page) => { await goHome(page); await click(page, projectCardOpen); await page.waitForFunction(() => state.view === 'workspace'); await click(page, '[data-action="workspace-edit-day"]'); await page.waitForFunction(() => state.view === 'builder'); const n = await sheetStepCount(page); for (let i = 2; i < n; i += 1) { await setStep(page, i); if (await page.locator('[data-action="open-row-observation"]').count()) break; } await page.locator('[data-action="open-row-observation"]').first().click({ force: true }); await page.waitForTimeout(400); }]);

  const browser = await chromium.launch();
  const results = [];
  const started = Date.now();
  try {
    for (const width of WIDTHS) {
      const height = HEIGHT[width] || 900;
      const roles = {};
      const expanded = screens.flatMap(([name, role, setup]) => (role ? rolesFor(name).map((r) => [name, r, setup]) : [[name, role, setup]]));
      for (const [name, role, setup] of expanded) {
        if (ONLY && !name.includes(ONLY)) continue;
        if (!roles[role]) {
          roles[role] = await browser.newContext({ viewport: { width, height }, deviceScaleFactor: 1, hasTouch: width < 800, isMobile: width < 500 });
          await roles[role].addInitScript(() => { try { localStorage.clear(); sessionStorage.clear(); } catch (e) { /* none */ } });
          if (role) await browserLogin(roles[role], base, role);
        }
        const context = roles[role];
        const page = await context.newPage();
        const consoleErrors = [];
        const failedRequests = [];
        page.on('pageerror', (e) => consoleErrors.push(`pageerror: ${e.message}`.slice(0, 160)));
        page.on('console', (m) => m.type() === 'error' && consoleErrors.push(`console: ${m.text()}`.slice(0, 160)));
        page.on('response', (r) => r.status() >= 400 && failedRequests.push(`${r.status()} ${new URL(r.url()).pathname}`));
        const record = { screen: name, width, role: role ? role.split('@')[0] : 'signed-out' };
        try {
          await setup(page);
          await page.waitForTimeout(250);
          record.url = page.url().replace(base, '');
          record.view = await page.evaluate(() => (typeof state !== 'undefined' ? state.view : null));
          record.own = await page.evaluate(pageChecks, { narrow: width < 800 });
          const axe = await new AxeBuilder({ page }).withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa', 'wcag22aa', 'best-practice']).analyze();
          record.axe = axe.violations.map((v) => ({ id: v.id, impact: v.impact, help: v.help, nodes: v.nodes.length, targets: v.nodes.slice(0, 3).map((n) => n.target.join(' ').slice(0, 100)), tags: v.tags.filter((t) => /^wcag\d+a+$|wcag22aa|best/.test(t)) }));
          record.axeIncomplete = axe.incomplete.map((v) => ({ id: v.id, nodes: v.nodes.length }));
          if (width === 1280) record.focus = await focusAudit(page);
          await page.screenshot({ path: path.join(SHOTS, `${name}_${record.role}_${width}.png`), fullPage: true }).catch(() => {});
        } catch (error) {
          record.error = error.message.split('\n').slice(0, 3).join(' ').slice(0, 300);
        }
        record.consoleErrors = [...new Set(consoleErrors)].slice(0, 5);
        record.failedRequests = [...new Set(failedRequests)].slice(0, 5);
        results.push(record);
        console.log(`${width} ${name}${record.error ? ` ERROR ${record.error}` : ` axe=${(record.axe || []).length} overflow=${record.own && record.own.pageOverflowPx}`}`);
        await page.close();
      }
      for (const ctx of Object.values(roles)) await ctx.close();
    }
  } finally {
    await browser.close();
    server.kill();
    fs.writeFileSync(path.join(OUT, 'results.json'), JSON.stringify({ at: new Date().toISOString(), widths: WIDTHS, seconds: Math.round((Date.now() - started) / 1000), results }, null, 1));
    console.log(`wrote ${path.join(OUT, 'results.json')} (${results.length} screen/width runs)`);
  }
})().catch((error) => { console.error(error); process.exit(1); });
