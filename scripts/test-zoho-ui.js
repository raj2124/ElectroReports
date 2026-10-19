// Session 6D in a real browser, on the mock Zoho portal: the "I am ..." choice and the new-project
// notice on Home (two questions, default data-collection mode, workspace opens), and step 1 with the
// project's Zoho tasks (measurement tasks pre-selected, planned quantities "from Zoho", the engineer
// from the task owner, no Work Order field). Needs Playwright's Chromium.
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');

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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'zoho-ui-'));
  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ZOHO_POLL_MINUTES: '0', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const browser = await chromium.launch();
  try {
    fs.mkdirSync(SHOTS, { recursive: true });
    const page = await (await browser.newContext({ viewport: { width: 1288, height: 900 } })).newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto(`http://127.0.0.1:${port}/`);

    // ---- Notices ---------------------------------------------------------------------------
    await page.waitForSelector('[data-notice-user]');
    expect(await page.locator('[data-notice-user]').count() === 1, 'Home asks who you are before showing notices');
    await page.request.post(`http://127.0.0.1:${port}/api/zoho/poll`);
    await page.selectOption('[data-notice-user]', { label: 'Anup Parekh' });
    await page.waitForSelector('.notice-card');
    let notice = await page.innerText('.notice-card');
    expect(notice.includes('056-D03_TS-26-CLL') && /Phones\/tablets allowed on site\?/.test(notice) && /Mobile data available\?/.test(notice), 'the notice names the project and asks the two questions');
    expect(/You own an ESHA Onsite task/.test(notice), 'the notice says why it is for this person');
    expect(await page.evaluate(() => JSON.parse(localStorage.getItem('electroreports.iam')).name) === 'Anup Parekh', 'who you are is remembered');
    await page.screenshot({ path: path.join(SHOTS, 'Home_notice_1288.png'), fullPage: true });
    await page.click('[data-action="notice-answer"]');
    await page.waitForTimeout(300);
    expect(await page.locator('.notice-card').count() === 1, 'answering needs both questions');
    await page.check('[data-notice-answer$=":phonesAllowed"][value="yes"]');
    await page.check('[data-notice-answer$=":mobileData"][value="no"]');
    await page.click('[data-action="notice-answer"]');
    await page.waitForSelector('.workspace-shell');
    expect(await page.inputValue('[data-workspace-detail="dataCollectionMode"]') === 'Printed sheets', 'phones yes + data no -> printed sheets by default');
    expect((await page.innerText('.workspace-shell')).includes('Site access answered by'), 'the workspace shows the answers');
    await page.fill('[data-workspace-detail="workOrder"]', '1002-1261201769');
    await page.click('[data-action="workspace-save-details"]');
    await page.waitForTimeout(400);
    expect(await page.evaluate(() => workspaceUi.workspace.project.workOrder) === '1002-1261201769', 'the Work Order is an optional field of the workspace project details');
    await page.click('[data-action="go-dashboard"]');
    await page.waitForSelector('.project-card');
    expect(await page.locator('.notice-card').count() === 0, 'an answered notice does not come back');
    await page.reload();
    await page.waitForSelector('.project-card');
    expect(await page.locator('.notice-card').count() === 0, 'nor after a reload');

    // ---- Step 1 with the Zoho tasks -----------------------------------------------------------
    await page.click('[data-action="use-zoho-project"][data-project-id="27127000009850033"]');
    await page.waitForSelector('.zoho-task-row', { state: 'attached' });
    expect(await page.locator('.zoho-tasklist').count() === 4, 'the tasklists are shown');
    expect(await page.locator('.zoho-task-row').count() === 6, 'all six tasks are listed');
    expect(await page.locator('[data-zoho-task]:checked').count() === 2, 'the two Onsite ESHA measurement tasks are pre-selected');
    const irt = page.locator('.zoho-task-row.is-irt');
    expect(await irt.count() === 1 && !(await irt.locator('input').isChecked()), 'the IRT task is listed and not selected');
    const panel = await page.innerText('.zoho-tasks');
    expect(panel.includes('Anup Parekh') && panel.includes('Onsite') && panel.includes('Planned'), 'tasks show owner, category and status');
    const draft = await page.evaluate(() => ({ tests: state.draft.tests, planned: state.draft.project.plannedQuantities, engineer: state.draft.project.engineerName, workOrder: state.draft.project.workOrder }));
    expect(draft.tests.soilResistivity && draft.tests.electrodeResistance && !draft.tests.continuityTest, 'the tests follow the selected tasks');
    expect(draft.planned.electrodeResistance === '170' && draft.planned.soilResistivity === '4', 'planned quantities are pre-filled from Zoho (170 pits, 4)');
    expect(draft.engineer === 'Anup Parekh', 'the engineer is the task owner, not the project owner');
    expect(await page.locator('[data-from-zoho]').count() === 2, 'pre-filled values are marked "from Zoho"');
    expect(await page.locator('[data-bind="project.workOrder"]').count() === 0, 'step 1 has no Work Order field');
    await page.screenshot({ path: path.join(SHOTS, 'Step1_zoho_tasks_1288.png'), fullPage: true });

    // editing a value ends "from Zoho"; the selection can be changed
    await page.fill('[data-bind="project.plannedQuantities.electrodeResistance"]', '150');
    expect(await page.locator('[data-from-zoho="electrodeResistance"]').count() === 0 && await page.locator('[data-from-zoho="soilResistivity"]').count() === 1, 'an edited value is no longer "from Zoho"; the other stays');
    await page.uncheck('[data-zoho-task="r3"]');
    const after = await page.evaluate(() => ({ soil: state.draft.tests.soilResistivity, planned: state.draft.project.plannedQuantities }));
    expect(!after.soil && after.planned.soilResistivity === '' && after.planned.electrodeResistance === '150', 'deselecting a task turns its test off and clears its Zoho value; the typed value stays');
    await page.check('[data-zoho-task="r3"]');
    expect(await page.evaluate(() => state.draft.project.plannedQuantities.soilResistivity) === '4', 'selecting it again fills the Zoho value again');
    expect(await page.evaluate(() => state.draft.project.engineerName) === 'Anup Parekh', 'the engineer stays');

    // a project that is not D03: tasks listed, nothing chosen for you
    await page.evaluate(() => { state.draft.project.engineerName = ''; });
    await page.click('[data-action="go-dashboard"]');
    await page.click('[data-action="use-zoho-project"][data-project-id="1001"]');
    await page.waitForSelector('.zoho-task-row', { state: 'attached' });
    expect(await page.locator('[data-zoho-task]:checked').count() === 0 && await page.locator('[data-zoho-task]:disabled').count() === 2, 'a non-D03 project lists its tasks with nothing selected');
    expect(await page.evaluate(() => state.draft.project.engineerName) === 'Rakesh Patel', 'a non-D03 project keeps the project owner as the engineer');
    expect(errors.length === 0, `no page errors (${errors.join('; ')})`);
  } finally {
    await browser.close();
    server.kill();
  }
  console.log(`\nzoho-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} failed`);
    process.exit(1);
  }
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
