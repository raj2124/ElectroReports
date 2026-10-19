// Project workspace in a real browser (Session 6C.4): the home list of active projects with their
// workspace links, the workspace (site days with status, progress, finals), "Start today's report"
// (new day carries the project details but not the plan for the next day; an existing date is
// reopened), "Assemble final report" (included dates, unapproved days flagged, opens the review)
// and "Print sheets for site". Needs Playwright's Chromium. Screenshots go to reference/screenshots/.
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { buildReviewState } = require('../src/review');

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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'workspace-ui-'));
  const dataFile = path.join(dir, 'reports.json');
  const store = createReportStore({ filePath: dataFile });
  const project = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej, Bharuch', workOrder: 'W-1', engineerName: 'H. S. Patil', unit: 'FI Plant', unitCode: 'FI', zohoProjectId: 'Z-1001', plannedQuantities: { electrodeResistance: 10, soilResistivity: 2 } });
  const payload = buildPayload();
  const day = (date, pits, extra = {}) => {
    const normalized = normalizeReportInput({
      ...payload,
      project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: date, nextDayPlan: 'Earth pits EP-07 to EP-10.', dataCollectionMode: 'Printed sheets', plannedQuantities: { electrodeResistance: 10, soilResistivity: 2 } },
      tests: { soilResistivity: true, electrodeResistance: true },
      electrodeResistance: payload.electrodeResistance.slice(...pits),
      ...extra
    });
    delete normalized.loopImpedanceTest;
    return store.addReport({ ...normalized, kind: 'daily', projectId: project.id, collectionDate: date });
  };
  const first = day('2026-09-14', [0, 6]);
  const second = day('2026-09-15', [6, 10]);
  store.transition(first.id, 'in_review', { actor: 'H. S. Patil' });
  store.setReviewTicks(first.id, Object.fromEntries(buildReviewState(store, store.getReport(first.id)).checklist.filter((item) => item.manual).map((item) => [item.id, true])));
  store.markChecked(first.id, { actor: 'A. P.' });
  store.transition(first.id, 'approved', { actor: 'J. C. Shah' });

  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: dataFile, ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ANTHROPIC_API_KEY: '' },
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
    await page.goto(`http://127.0.0.1:${port}/`);

    // Home: active projects with a workspace link; Zoho projects open a workspace too.
    await page.waitForSelector('.project-card');
    expect(await page.locator('.project-card').count() === 1, 'the home screen lists the active project');
    const card = await page.innerText('.project-card');
    expect(card.includes('024-D03_TS-26-GFL') && card.includes('Gujarat Fluorochemicals Ltd') && /Site days\s*2/.test(card) && /Approved\s*1/.test(card), 'the project card shows number, client, site days and approved days');
    expect(await page.locator('[data-action="open-workspace-zoho"]').count() > 0, 'each Zoho project offers "Open workspace"');
    await page.screenshot({ path: path.join(SHOTS, 'Home_projects_1288.png'), fullPage: true });

    // Workspace
    await page.click('.project-card [data-action="open-workspace"]');
    await page.waitForSelector('.workspace-shell');
    const text = () => page.innerText('.workspace-shell');
    let screen = await text();
    expect(/project workspace/i.test(screen) && /\bactive\b/i.test(screen), 'the workspace opens, active');
    const rows = await page.locator('.workspace-days tbody tr').allInnerTexts();
    expect(rows.length === 2 && /14.*approved/is.test(rows[0]) && /15.*draft/is.test(rows[1]), `site days list with their status (got ${JSON.stringify(rows)})`);
    expect(/Electrode Resistance Test\s*10 of 10 pits · 100%/.test(screen.replace(/\s+/g, ' ')) || /10 of 10 pits · 100%/.test(screen.replace(/\s+/g, ' ')), 'cumulative progress against the plan is shown');
    expect(screen.includes('Not assembled yet'), 'no final report yet');
    await page.screenshot({ path: path.join(SHOTS, 'Workspace_1288.png'), fullPage: true });

    // Start a new day: details carried over, the next-day plan is not
    await page.fill('[data-workspace-field="workDate"]', '2026-09-16');
    expect((await page.innerText('[data-action="workspace-start-day"]')).includes("Start today's report"), 'a new date offers "Start today\'s report"');
    await page.click('[data-action="workspace-start-day"]');
    await page.waitForSelector('[data-bind="project.projectNo"]');
    const carried = await page.evaluate(() => ({ editing: state.editing, project: state.draft.project }));
    expect(carried.editing && carried.editing.workDate === '2026-09-16' && !carried.editing.reportId, 'the builder holds a new site day of the project');
    expect(carried.project.clientName === 'Gujarat Fluorochemicals Ltd' && carried.project.siteLocation === 'Dahej, Bharuch' && carried.project.unitCode === 'FI' && carried.project.zohoProjectId === 'Z-1001', 'project details carry over');
    expect(carried.project.reportDate === '2026-09-16', 'the work date is the chosen date');
    expect(!carried.project.nextDayPlan, 'the plan for the next day does NOT carry over');
    expect(await page.locator('[data-bind="project.nextDayPlan"]').count() === 1, 'the project step has a "Plan for next day" field for a site day');
    expect(await page.locator('[data-bind="project.reportDate"]').count() === 0, 'the work date is fixed (not editable) for a site day');
    expect(await page.locator('.zoho-sync-panel.is-hidden').count() === 1, 'the Zoho picker is hidden while entering a site day');

    // Save it (the entered data is the existing day's readings, loaded into the builder)
    await page.evaluate(async (id) => {
      const report = await api(`/api/reports/${id}`);
      const keep = state.editing;
      state.draft = restoreDraftShape(report);
      state.draft.project.reportDate = '2026-09-16';
      state.draft.project.nextDayPlan = 'Report assembly';
      state.editing = keep;
      await saveReport();
    }, second.id);
    await page.waitForSelector('.workspace-days tbody tr:nth-child(3)');
    expect(await page.locator('.workspace-days tbody tr').count() === 3, 'the saved day appears in the workspace');
    expect(/16.*draft/is.test(await page.locator('.workspace-days tbody tr:nth-child(3)').innerText()), 'a new day is a draft');

    // Reopen an existing date instead of duplicating it
    await page.fill('[data-workspace-field="workDate"]', '2026-09-15');
    expect((await page.innerText('[data-action="workspace-start-day"]')).startsWith('Reopen'), 'an existing date offers "Reopen"');
    await page.click('[data-action="workspace-start-day"]');
    await page.waitForSelector('[data-bind="project.projectNo"]');
    const reopened = await page.evaluate(() => state.editing);
    expect(Boolean(reopened && reopened.reportId), 'the existing day is reopened for editing');
    expect(await page.evaluate(() => state.draft.project.nextDayPlan) === 'Earth pits EP-07 to EP-10.', 'the reopened day keeps its own plan');
    await page.click('[data-action="workspace-back"]');
    await page.waitForSelector('.workspace-shell');
    expect(await page.locator('.workspace-days tbody tr').count() === 3, 'going back to the project does not create a day');

    // An approved date is opened for review, not for editing
    await page.fill('[data-workspace-field="workDate"]', '2026-09-14');
    await page.click('[data-action="workspace-start-day"]');
    await page.waitForSelector('.review-shell');
    expect(/approved/i.test(await page.innerText('.review-shell')), 'an approved day opens in review');
    await page.click('[data-action="open-workspace"]');
    await page.waitForSelector('.workspace-shell');

    // Print sheets for site, from the workspace
    await page.click('[data-action="workspace-print"]');
    await page.waitForSelector('.print-dialog');
    const dialog = await page.innerText('.print-dialog');
    expect(dialog.includes('024-D03_TS-26-GFL') && dialog.includes('Gujarat Fluorochemicals Ltd'), 'the print dialog is for the workspace\'s project');
    await page.click('.print-dialog [data-action="close-print-sheets"]');

    // Assemble the final report
    await page.click('[data-action="workspace-assemble"]');
    await page.waitForSelector('.workspace-confirm');
    const confirm = (await page.innerText('.workspace-confirm')).replace(/\s+/g, ' ');
    expect(confirm.includes('3 daily reports') && confirm.includes('14') && confirm.includes('15') && confirm.includes('16'), 'the confirmation lists the included dates');
    const flagged = confirm.split('Not yet approved:')[1] || '';
    expect(flagged.includes('15') && flagged.includes('16') && !flagged.includes('14 Sept'), `days not yet approved are flagged, the approved day is not (${flagged.slice(0, 60)})`);
    await page.screenshot({ path: path.join(SHOTS, 'Workspace_assemble_1288.png'), fullPage: true });
    await page.click('[data-action="workspace-assemble-confirm"]');
    await page.waitForSelector('.review-shell');
    const review = await page.innerText('.review-shell');
    expect(/final report/i.test(review) && /draft/i.test(review), 'assembling opens the new final report in review, as a draft');
    expect(await page.locator('[data-action="open-workspace"]').count() === 1, 'the review screen links back to the project workspace');
    await page.click('[data-action="open-workspace"]');
    await page.waitForSelector('.workspace-shell');
    expect(await page.locator('.workspace-shell h3', { hasText: 'Final reports' }).count() === 1, 'the final report is listed in the workspace');
    // re-assemble: a new draft, the old one kept
    await page.click('[data-action="workspace-assemble"]');
    expect((await page.innerText('.workspace-confirm')).includes('An earlier final report exists'), 're-assembling says the earlier final is kept');
    await page.click('[data-action="workspace-assemble-confirm"]');
    await page.waitForSelector('.review-shell');
    await page.click('[data-action="open-workspace"]');
    await page.waitForSelector('.workspace-shell');
    expect((await page.locator('.workspace-shell table').last().locator('tbody tr').count()) === 2, 'both final reports are listed');

    // Planned quantities
    await page.locator('.workspace-planned-form summary').click();
    await page.fill('[data-workspace-planned="electrodeResistance"]', '20');
    await page.click('[data-action="workspace-save-planned"]');
    await page.waitForFunction(() => /10 of 20 pits · 50%/.test(document.querySelector('.workspace-shell').innerText.replace(/\s+/g, ' ')));
    expect(true, 'saving the planned quantities updates the progress');

    for (const width of [1288, 390]) {
      await page.setViewportSize({ width, height: 900 });
      await page.waitForTimeout(200);
      const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
      const wide = overflow > 0 ? await page.evaluate(() => [...document.querySelectorAll('.workspace-shell *')].filter((el) => el.getBoundingClientRect().right > window.innerWidth + 1).slice(0, 6).map((el) => `${el.tagName}.${String(el.className).slice(0, 40)}`)) : [];
      expect(overflow <= 0, `the workspace does not scroll sideways at ${width}px (overflow ${overflow}; ${wide.join(', ')})`);
      if (width === 390) {
        await page.screenshot({ path: path.join(SHOTS, 'Workspace_390.png'), fullPage: true });
      }
    }

    // Zoho project -> workspace (find or create)
    await page.setViewportSize({ width: 1288, height: 900 });
    await page.click('[data-action="go-dashboard"]');
    await page.waitForSelector('[data-action="open-workspace-zoho"]');
    await page.locator('[data-action="open-workspace-zoho"]').first().click();
    await page.waitForSelector('.workspace-shell');
    expect((await page.innerText('.workspace-shell')).includes('No site day recorded yet'), 'a Zoho project opens a (new) workspace');
  } finally {
    await browser.close();
    server.kill();
  }

  console.log(`\nworkspace-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} failed`);
    process.exit(1);
  }
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
