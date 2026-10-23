// Session 6F in a real browser, on the mock Zoho portals (Projects and People): the Schedule screen with
// week and month views, the whole team with your own row first, bars per Onsite task, approved and pending
// leave, holidays and non-working days, warnings, refresh (limited to one a minute), the bar click that
// opens (or offers to create) the project workspace, and the admin settings. Needs Playwright's Chromium.
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
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'schedule-ui-'));
  const port = await freePort();
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
    const base = `http://127.0.0.1:${port}`;
    const context = await browser.newContext({ viewport: { width: 1288, height: 900 } });
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    const goTo = async (anchor, mode = 'week') => {
      await page.evaluate(([a, m]) => { scheduleUi.anchor = a; scheduleUi.mode = m; return loadSchedule(); }, [anchor, mode]);
      await page.waitForSelector('.sched-row:not(.sched-header)');
    };
    const rowNames = () => page.locator('.sched-row:not(.sched-header) .sched-name strong').allInnerTexts();

    // ---- Nidhi: an engineer with no Onsite task herself ---------------------------------------------
    await browserLogin(context, base, 'nidhi@elegrow.com');
    await page.goto(`${base}/`);
    await page.waitForSelector('.user-bar');
    await page.click('[data-action="open-schedule"]');
    await page.waitForSelector('.sched-card');
    await goTo('2026-10-21');
    const names = await rowNames();
    expect(names[0] === 'Nidhi Shah' && (await page.locator('.sched-row.sched-me').count()) === 1, 'your own row is first, even with no task');
    expect(['Anup Parekh', 'Karishma Prajapati'].every((name) => names.includes(name)) && names.length === 3, 'the whole team is on the calendar (people with a task or an engineer role)');
    expect((await page.innerText('.sched-card')).includes('19 Oct to 25 Oct 2026'), 'the week runs Monday to Sunday');
    expect(await page.locator('.sched-day').count() === 7, 'seven day columns');

    // bars, leave and holidays of the week 19-25 Oct
    expect(await page.locator('.sched-bar').count() === 3, 'three Onsite tasks in this week (056 pits and the two 024 tasks)');
    expect(await page.locator('.sched-bar', { hasText: 'Step & Touch' }).count() === 0 && await page.locator('.sched-bar', { hasText: 'Thermography' }).count() === 0 && await page.locator('.sched-bar', { hasText: 'Report Drafting' }).count() === 0, 'closed, IRT and offsite tasks are not shown');
    expect(await page.locator('.sched-leave.sched-leave-approved').count() === 1 && await page.locator('.sched-leave.sched-leave-pending').count() === 1, 'approved and pending leave are both on the calendar');
    const opacity = await page.evaluate(() => [getComputedStyle(document.querySelector('.sched-leave.sched-leave-approved')).opacity, getComputedStyle(document.querySelector('.sched-leave.sched-leave-pending')).opacity].map(Number));
    expect(opacity[0] === 1 && opacity[1] < 1, 'pending leave is dimmed');
    expect((await page.locator('.sched-day').nth(1).innerText()).includes('Holiday') && (await page.locator('.sched-day').nth(1).getAttribute('title')) === 'Dussehra', '20 Oct is marked as a holiday (from Zoho People)');
    expect((await page.locator('.sched-day').nth(6).getAttribute('class')).includes('sched-off'), 'Sunday is a non-working day');
    expect((await page.locator('.sched-day').nth(5).getAttribute('class')).includes('sched-off') === false, 'Saturday is a working day by default');
    expect(await page.locator('.sched-bar-warn').count() === 3, 'every task in this week has something to warn about');
    expect((await page.locator('.sched-bar').first().getAttribute('title')).includes('Warning:'), 'a bar lists its warnings in its tooltip');
    expect(await page.locator('.sched-clash').count() >= 2, 'the days Anup is on two tasks are marked');
    const warnings = await page.innerText('.sched-warnings');
    expect(/Anup Parekh is on two Onsite tasks on 21 Oct to 22 Oct/.test(warnings), 'the clash is described');
    expect(/Karishma Prajapati is on approved leave/.test(warnings) && /Anup Parekh is on pending leave/.test(warnings), 'leave warnings say approved or pending');
    expect(/170 pits at about 35 a day needs 4\.9 working days/.test(warnings), 'the quantity warning shows the arithmetic');
    expect(/read (just now|\d+ min ago)/.test(await page.innerText('.sched-head')) && (await page.innerText('.sched-head')).includes('mock portal'), 'the age of the Zoho data is shown');
    await page.screenshot({ path: path.join(SHOTS, 'Schedule_week_1288.png'), fullPage: true });

    // navigation and the other views
    await page.click('[data-action="schedule-step"][data-step="1"]');
    await page.waitForFunction(() => scheduleUi.data && scheduleUi.data.window.from === '2026-10-26');
    expect((await rowNames()).includes('No owner') && await page.locator('.sched-bar').count() === 2, 'next week: the task nobody owns has its own row');
    expect(/has no owner/.test(await page.innerText('.sched-warnings')) && /has no start or end date/.test(await page.innerText('.sched-warnings')), 'no owner and no dates are warned about');
    await page.click('[data-action="schedule-today"]');
    await page.waitForFunction(() => scheduleUi.data && scheduleUi.data.window.from <= todayIso() && scheduleUi.data.window.to >= todayIso());
    await goTo('2026-10-21', 'month');
    expect(await page.locator('.sched-day').count() === 31 && (await page.innerText('.sched-card h1')).includes('October 2026'), 'the month view has every day of the month');
    expect(await page.locator('.sched-bar').count() === 5, 'five Onsite tasks in October');
    await page.screenshot({ path: path.join(SHOTS, 'Schedule_month_1288.png'), fullPage: true });
    await page.click('[data-action="schedule-step"][data-step="1"]');
    await page.waitForFunction(() => scheduleUi.data && scheduleUi.data.window.from === '2026-11-01');
    expect(await page.locator('.sched-bar').count() === 1 && /Far more days than the planned quantity needs/.test(await page.innerText('.sched-warnings')) && /11 are booked/.test(await page.innerText('.sched-warnings')), 'November: the long riser task, with the "more days than needed" warning');
    expect(await page.locator('.sched-clip-start').count() === 0, 'a task that starts inside the month is not clipped');
    expect((await page.locator('.sched-row', { hasText: 'Mohana Vamsee' }).innerText()).includes('Not signed in yet'), 'someone who never signed in gets a row from their Zoho name');

    // ---- the bar opens the workspace (or offers to create it) -----------------------------------------
    await goTo('2026-10-19');
    await page.click('.sched-bar:has-text("056-D03_TS-26-CLL")');
    expect(await page.locator('.notice-card:has-text("has no ElectroReports workspace yet")').count() === 1, 'a project with no workspace is offered one');
    await page.click('[data-action="schedule-cancel-create"]');
    expect(await page.locator('.notice-card:has-text("no ElectroReports workspace")').count() === 0, 'cancel closes the offer');
    await page.click('.sched-bar:has-text("056-D03_TS-26-CLL")');
    await page.click('[data-action="schedule-create-workspace"]');
    await page.waitForSelector('.workspace-shell');
    expect((await page.innerText('.workspace-shell')).includes('056-D03_TS-26-CLL'), 'the workspace of the project opens');
    await page.click('[data-action="open-schedule"]');
    await page.waitForSelector('.sched-card');
    await goTo('2026-10-19');
    await page.click('.sched-bar:has-text("056-D03_TS-26-CLL")');
    await page.waitForSelector('.workspace-shell');
    expect(await page.locator('.notice-card:has-text("no ElectroReports workspace")').count() === 0, 'once the workspace exists the bar opens it directly');

    // ---- refresh ---------------------------------------------------------------------------------------
    await page.click('[data-action="open-schedule"]');
    await page.waitForSelector('.sched-card');
    await page.click('[data-action="schedule-refresh"]');
    await page.waitForSelector('.review-message');
    expect((await page.innerText('.review-message')).includes('Refreshed from Zoho'), 'refresh reads Zoho again');
    await page.click('[data-action="schedule-refresh"]');
    await page.waitForFunction(() => /once a minute/.test(document.querySelector('.review-message') ? document.querySelector('.review-message').innerText : ''));
    expect(true, 'a second refresh within a minute says so');
    expect(await page.locator('[data-action="schedule-toggle-settings"]').count() === 0, 'only an admin sees the settings');

    // ---- phone width -----------------------------------------------------------------------------------
    await page.setViewportSize({ width: 390, height: 800 });
    await goTo('2026-10-21');
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1), 'at phone width the page does not scroll sideways (the grid scrolls inside its box)');
    await page.screenshot({ path: path.join(SHOTS, 'Schedule_week_390.png'), fullPage: true });
    await page.setViewportSize({ width: 1288, height: 900 });

    // ---- admin: settings -------------------------------------------------------------------------------
    const admin = await browser.newContext({ viewport: { width: 1288, height: 900 } });
    const adminPage = await admin.newPage();
    adminPage.on('pageerror', (error) => errors.push(error.message));
    await browserLogin(admin, base, 'admin@elegrow.com');
    await adminPage.goto(`${base}/`);
    await adminPage.waitForSelector('.user-bar');
    await adminPage.click('[data-action="open-schedule"]');
    await adminPage.waitForSelector('.sched-card');
    await adminPage.evaluate(() => { scheduleUi.anchor = '2026-10-21'; return loadSchedule(); });
    await adminPage.waitForSelector('.sched-row:not(.sched-header)');
    expect((await adminPage.locator('.sched-row:not(.sched-header) .sched-name strong').first().innerText()) === 'H. S. Patil', 'an admin who is not an engineer still sees their own row first');
    await adminPage.click('[data-action="schedule-toggle-settings"]');
    await adminPage.waitForSelector('.sched-settings');
    expect(await adminPage.locator('[data-schedule-day]').count() === 7 && await adminPage.locator('[data-schedule-rate]').count() === 9, 'the working days and nine rates can be edited');
    expect(await adminPage.locator('.sched-settings .chip', { hasText: 'estimate' }).count() === 9, 'every rate is labelled "estimate"');
    expect(await adminPage.locator('[data-schedule-day="1"]').isChecked() && await adminPage.locator('[data-schedule-day="6"]').isChecked() && !(await adminPage.locator('[data-schedule-day="0"]').isChecked()), 'Monday to Saturday by default');
    await adminPage.screenshot({ path: path.join(SHOTS, 'Schedule_settings_1288.png'), fullPage: true });
    await adminPage.fill('[data-schedule-rate="electrodeResistance"]', '50');
    await adminPage.uncheck('[data-schedule-day="6"]');
    await adminPage.click('[data-action="schedule-save-settings"]');
    await adminPage.waitForFunction(() => scheduleUi.data && scheduleUi.data.settings.rates.electrodeResistance.perDay === 50);
    expect(await adminPage.evaluate(() => scheduleUi.data.settings.workingDays.join()) === '1,2,3,4,5', 'the new working days are saved');
    expect(!/170 pits at about/.test(await adminPage.innerText('.sched-warnings')), 'the edited rate is used by the warnings');
    expect((await adminPage.locator('.sched-day').nth(5).getAttribute('class')).includes('sched-off'), 'Saturday is now a non-working day');

    expect(errors.length === 0, `no page errors (${errors.join('; ')})`);
  } finally {
    await browser.close();
    server.kill();
  }
  console.log(`\nschedule-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} failed`);
    process.exit(1);
  }
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
