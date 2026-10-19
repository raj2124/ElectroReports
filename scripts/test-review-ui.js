// Review & approval screen and the new project-step fields in a real browser (Session 6C).
// Seeds a temporary database, opens the review screen of a FINAL report at two widths, and walks
// the workflow: name required, checklist ticks, mark as checked, acknowledgement before approval,
// approve, re-issue. Needs Playwright's Chromium. Screenshots go to reference/screenshots/ (git-ignored).
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'review-ui-'));
  const dataFile = path.join(dir, 'reports.json');
  const store = createReportStore({ filePath: dataFile });
  const project = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' });
  const payload = buildPayload();
  const normalized = normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: '2026-09-14' } });
  store.addReport({ ...normalized, kind: 'daily', projectId: project.id, collectionDate: '2026-09-14' });
  store.assembleFinal(project.id);

  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: dataFile, ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', ANTHROPIC_API_KEY: '' },
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
    const base = `http://127.0.0.1:${port}`;
    await browserLogin(context, base, 'admin@elegrow.com'); // H. S. Patil: prepares
    await page.goto(`${base}/`);

    // Project step: planned quantities and unit code.
    await page.click('[data-action="new-report"]');
    await page.waitForSelector('[data-bind="project.projectNo"]');
    expect(await page.locator('[data-bind="project.unitCode"]').count() === 1, 'project step has a Unit code field');
    expect(await page.locator('[data-bind="project.plannedQuantities.electrodeResistance"]').count() === 1, 'project step has a planned quantity for each selected test');
    await page.fill('[data-bind="project.unitCode"]', 'f-i plant 123456');
    expect(await page.inputValue('[data-bind="project.unitCode"]') === 'FIPLANT1', 'unit code keeps letters/digits, upper case, 8 characters');
    await page.fill('[data-bind="project.plannedQuantities.electrodeResistance"]', '10');
    expect(await page.evaluate(() => state.draft.project.plannedQuantities.electrodeResistance) === '10', 'planned quantity is stored on the draft project');
    await page.click('[data-action="go-dashboard"]');

    // Review screen.
    await page.waitForSelector('[data-action="open-review"]');
    expect(await page.locator('[data-action="open-review"]').count() === 2, 'daily and final reports have a Review button; the standalone path is untouched');
    const finalId = await page.evaluate(() => state.reports.find((report) => report.kind === 'final').id);
    await page.locator(`[data-action="open-review"][data-id="${finalId}"]`).click();
    await page.waitForSelector('.review-shell');
    const text = () => page.innerText('.review-shell');
    // Signs in as another user (a new session cookie) and comes back to this review screen.
    const switchTo = async (email) => {
      await browserLogin(context, base, email);
      await page.goto(`${base}/`);
      await page.waitForSelector(`[data-action="open-review"][data-id="${finalId}"]`);
      await page.locator(`[data-action="open-review"][data-id="${finalId}"]`).click();
      await page.waitForSelector('.review-shell');
    };
    let screen = await text();
    ['Reviewer checklist', 'Counts match across scope', 'Every P1 has a photo', 'Sanity check: not run', 'Narrative:', 'Transitions', 'Draft', 'Client logo'].forEach((expected) => expect(screen.includes(expected) || screen.toLowerCase().includes(expected.toLowerCase()), `review screen shows "${expected}"`));
    expect(await page.locator('[data-action="review-step"][data-step="submit"]').count() === 1, 'a draft offers Submit for review');
    expect(await page.locator('[data-action="review-step"][data-step="approve"]').count() === 0, 'Approve is not offered before the report is checked');

    expect((await text()).includes('Signed in as H. S. Patil'), 'the review screen shows who is signed in (no name to type)');
    expect(await page.locator('[data-review-field="actor"]').count() === 0, 'there is no name or designation field any more');
    await page.click('[data-action="review-step"][data-step="submit"]');
    await page.waitForSelector('[data-action="review-step"][data-step="check"]');
    screen = await text();
    expect(screen.includes('In review'), 'status becomes In review');
    expect(await page.locator('[data-action="review-step"][data-step="approve"]').count() === 0, 'still no Approve: it is blocked until checked');

    // the preparer (an admin) is offered an override, not a plain check
    expect(await page.locator('[data-review-field="overrideReason"]').count() === 1, 'the preparer sees that they cannot check their own report (admin: override with a reason)');
    await switchTo('reviewer@elegrow.com'); // A. P.: checks
    expect(await page.locator('[data-action="review-step"][data-step="submit"]').count() === 0 && await page.locator('[data-action="review-step"][data-step="approve"]').count() === 0, 'a reviewer is offered Check and Return, not Submit or Approve');
    await page.click('[data-action="review-step"][data-step="check"]');
    await page.waitForFunction(() => document.querySelector('.review-message'));
    expect((await text()).includes('Tick every checklist item'), 'marking as checked before ticking the checklist is refused with the reason');
    for (const box of await page.locator('[data-review-tick]').all()) {
      await box.check();
      await page.waitForTimeout(150);
    }
    expect(await page.locator('[data-review-tick]:checked').count() === 5, 'all five checklist items can be ticked and are kept');
    await page.screenshot({ path: path.join(SHOTS, 'Review_1288.png'), fullPage: true });
    await page.click('[data-action="review-step"][data-step="check"]');
    await page.waitForFunction(() => /checked by A\. P\./.test(document.querySelector('.review-shell').innerText));
    expect((await text()).includes('checked by A. P.'), 'the checker is shown');
    expect(await page.locator('[data-action="review-step"][data-step="approve"]').count() === 0, 'a reviewer who is not an approver is not offered Approve');

    await switchTo('approver@elegrow.com'); // J. C. Shah: approves
    await page.waitForSelector('[data-action="review-step"][data-step="approve"]');
    await page.click('[data-action="review-step"][data-step="approve"]');
    await page.waitForFunction(() => /Acknowledge before approving/.test(document.querySelector('.review-message') ? document.querySelector('.review-message').textContent : ''));
    expect(await page.locator('[data-review-ack="sanityCheck"]').count() === 1, 'the sanity-check acknowledgement box is shown');
    expect(await page.locator('[data-review-ack="p1Photos"]').count() === 1, 'the P1-photo acknowledgement box is shown');
    await page.fill('[data-review-ack="sanityCheck"]', 'Provider unavailable; narrative read against the tables.');
    await page.fill('[data-review-ack="p1Photos"]', 'Photos to follow in Annexure B.');
    // The toast must not sit over the action buttons.
    await page.evaluate(() => showToast('Probe toast for the overlap check.', 'warning'));
    await page.waitForSelector('.toast');
    const covered = () => page.evaluate(() => {
      const toast = document.querySelector('.toast').getBoundingClientRect();
      return [...document.querySelectorAll('.review-shell button, .review-shell textarea, .review-shell input')].filter((button) => {
        const box = button.getBoundingClientRect();
        return box.width && box.height && !(box.right < toast.left || box.left > toast.right || box.bottom < toast.top || box.top > toast.bottom);
      }).map((button) => button.textContent.trim() || button.dataset.reviewField || 'field');
    });
    await page.evaluate(() => window.scrollTo(0, 0));
    const overlapTop = await covered();
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    const overlap = [...overlapTop, ...(await covered())];
    expect(overlap.length === 0, `the toast does not cover any button on the review screen (covers: ${overlap.join(', ')})`);
    await page.click('[data-action="review-step"][data-step="approve"]');
    await page.waitForFunction(() => /Approved/.test(document.querySelector('.detail-hero-tags') ? document.querySelector('.detail-hero-tags').textContent : ''));
    screen = await text();
    expect(screen.includes('Prepared HSP') || /Prepared\s+HSP/.test(screen), 'cover initials: prepared HSP');
    expect(/Checked\s+AP/.test(screen) && /Approved\s+JCS/.test(screen), 'cover initials: checked AP, approved JCS');
    expect(await page.locator('.review-shell tbody tr').count() === 3, 'three transitions recorded');
    expect(await page.locator('[data-review-tick]:not([disabled])').count() === 0, 'an approved report cannot be re-ticked');
    expect(await page.locator('[data-action="review-reissue"]').count() === 1, 'an approved report offers Re-issue as R1');

    await page.fill('[data-review-field="reissueReason"]', 'Client comments');
    await page.click('[data-action="review-reissue"]');
    await page.waitForFunction(() => /Revision R1/.test(document.querySelector('.review-shell') ? document.querySelector('.review-shell').textContent : ''));
    expect((await text()).includes('Draft'), 'the re-issued R1 opens as a draft');

    for (const width of [1288, 390]) {
      await page.setViewportSize({ width, height: 900 });
      await page.waitForTimeout(200);
      const overflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
      expect(overflow <= 0, `review screen does not scroll sideways at ${width}px (overflow ${overflow})`);
      if (width === 390) {
        await page.screenshot({ path: path.join(SHOTS, 'Review_390.png'), fullPage: true });
      }
    }
  } finally {
    await browser.close();
    server.kill();
  }

  console.log(`\nreview-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} failed`);
    process.exit(1);
  }
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
