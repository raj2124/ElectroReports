// Session 6I in a real browser on the mocks: upload a PDF of filled sheets, the review screen (scan beside the parsed cells in the
// sheet's layout, flagged cells, editing, zoom), Accept into the daily report with the overwrite confirmation, a page whose QR will not
// scan (the engineer chooses the sheet), the admin screen, and the "Suggest wording" helper. Needs Playwright's Chromium.
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');
const { devEnv, browserLogin, login } = require('./lib/dev-auth');
const { fillIssued, renderSheetsPdf, withoutQr } = require('./lib/scan-fixtures');
const { splitUpload } = require('../src/scan/split');

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
const PROJECT = { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' };
const ROWS = [
  { tag: 'P-1', location: 'Yard', electrodeType: 'Rod', materialType: 'GI', length: '3', diameter: '20', resistanceWithGrid: '0.52', method: '61.8%', conditionCodes: ['C1', 'C6'] },
  { tag: 'P-2', location: 'Gate', electrodeType: 'Plate', materialType: 'Cu', length: '0.6', diameter: '3', resistanceWithGrid: '1.20', method: '61.8%' }
];

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'scan-ui-'));
  const truthFile = path.join(dir, 'truth.json');
  const setTruth = (truth) => fs.writeFileSync(truthFile, JSON.stringify(truth));
  const port = await freePort();
  const base = `http://127.0.0.1:${port}`;
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', GRAPH_USE_MOCK: 'true', OCR_USE_MOCK: 'true', OCR_MOCK_TRUTH_FILE: truthFile, AI_CHECKS_USE_MOCK: 'true', GEMINI_API_KEY: '', OPENAI_API_KEY: '', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const browser = await chromium.launch();
  try {
    const context = await browser.newContext({ viewport: { width: 1400, height: 1000 } });
    await browserLogin(context, base, 'engineer@elegrow.com');
    const cookie = await login(base, 'engineer@elegrow.com');
    const api = async (method, url, body) => (await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: body ? JSON.stringify(body) : undefined })).json();
    const workspace = (await api('POST', '/api/projects', PROJECT)).projectId;
    const issue = async (sheetKey, workDate) => {
      const response = await fetch(`${base}/api/sheets/print`, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: JSON.stringify({ project: PROJECT, workDate, sheets: [sheetKey] }) });
      return api('GET', `/api/sheets/issued/${response.headers.get('x-sheet-ids').split(',')[0]}`);
    };
    const electrode = await fillIssued(await issue('electrodeResistance', '2026-10-12'), { instrument: 'mi3290', serialNo: 'SN77', rows: ROWS });
    const loop = await fillIssued(await issue('loopImpedanceTest', '2026-10-12'), { rows: [{ location: 'MCC-1', feederTag: 'F-1', deviceType: 'MCCB', deviceRating: '63', loopImpedance: '0.31', voltage: '230' }, { loopImpedance: '0.33' }, { loopImpedance: '0.35' }] });
    const both = (await renderSheetsPdf([electrode.ctx, loop.ctx])).buffer;
    setTruth({
      [electrode.issued.sheetId]: { primary: { ...electrode.truth, 'r0.resistanceWithGrid': { value: '0.5?', confidence: 0.55 }, 'r1.length': { value: '0.6', confidence: 0.7 } }, escalation: { ...electrode.truth, 'r1.length': { value: '0.8', confidence: 0.95 } } },
      [loop.issued.sheetId]: { primary: loop.truth }
    });

    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto(`${base}/`);
    await page.waitForSelector('.user-bar');

    // ---- upload
    await page.click('[data-action="open-scans"]');
    await page.waitForSelector('[data-scans]');
    await page.setInputFiles('[data-scan-files]', { name: 'visit.pdf', mimeType: 'application/pdf', buffer: both });
    await page.click('[data-action="scan-upload"]');
    await page.waitForSelector('.scan-item[data-scan-status="read"] >> nth=1', { timeout: 20000 });
    const items = await page.locator('.scan-item').allInnerTexts();
    expect(items.length === 2 && items.every((text) => /ready to review/i.test(text)), 'both pages are read and ready to review');
    expect(items.some((text) => /MS-02/.test(text) && /2 to check/.test(text)) && items.some((text) => /MS-04/.test(text)), 'the list shows the sheet and how many cells to check');

    // ---- review: phase-group sheet first (merged cells span the group)
    await page.locator('.scan-item', { hasText: 'MS-04' }).locator('[data-action="scan-open"]').click();
    await page.waitForSelector('[data-scan-review]');
    expect(await page.locator('[data-scan-table] tbody tr').count() === 12, 'the loop sheet shows its 12 printed rows');
    const merged = page.locator('[data-cell-box="r0.location"]');
    expect(await merged.getAttribute('rowspan') === '3' && (await merged.locator('input').inputValue()) === 'MCC-1', 'the merged cells span the three phase rows and hold what was read');
    expect(await page.locator('[data-scan-table] [data-cell-box="r1.location"]').count() === 0, 'and appear once per group');
    expect(await page.locator('[data-scan-image]').evaluate((img) => img.complete && img.naturalWidth > 500), 'the scan is shown beside the cells');
    await page.click('[data-action="scan-accept"]');
    await page.waitForFunction(() => state.view === 'workspace');
    expect((await page.innerText('.toast')).includes('Accepted') && (await page.innerText('.toast')).includes('new daily report'), 'Accept starts the day and opens the project workspace');

    // ---- review: flat sheet with flagged cells
    await page.click('[data-action="open-scans"]');
    await page.locator('.scan-item', { hasText: 'MS-02' }).locator('[data-action="scan-open"]').click();
    await page.waitForSelector('[data-scan-review]');
    expect(await page.locator('[data-scan-table] tbody tr').count() === 10, 'the electrode sheet shows its 10 printed rows');
    const flagged = await page.locator('[data-scan-table] .scan-flag').count();
    expect(flagged === 2, `the two doubtful cells are highlighted (${flagged})`);
    expect((await page.locator('[data-cell-box="r0.resistanceWithGrid"]').getAttribute('title')).includes('disagree'), 'hovering a flagged cell gives the reason');
    expect(await page.isChecked('[data-scan-tick="h.instrument"][data-value="mi3290"]') && await page.isChecked('[data-scan-tick="r0.conditionCodes"][data-value="C1"]') && await page.isChecked('[data-scan-tick="r0.conditionCodes"][data-value="C6"]') && !(await page.isChecked('[data-scan-tick="r0.conditionCodes"][data-value="C2"]')), 'ticks are shown as ticks, several condition codes included');
    await page.click('[data-action="scan-zoom"][data-step="0.25"]');
    expect((await page.locator('[data-scan-image]').getAttribute('style')).includes('125%'), 'zoom enlarges the scan');
    await page.fill('[data-scan-cell="r0.resistanceWithGrid"]', '0,53');
    await page.press('[data-scan-cell="r0.resistanceWithGrid"]', 'Tab');
    await page.waitForSelector('[data-cell-box="r0.resistanceWithGrid"].scan-edited');
    expect((await page.locator('[data-scan-cell="r0.resistanceWithGrid"]').inputValue()) === '0.53' && !(await page.locator('[data-cell-box="r0.resistanceWithGrid"]').getAttribute('class')).includes('scan-flag'), 'an edited cell is marked as edited and is no longer flagged');
    await page.fill('[data-scan-cell="r1.length"]', '0.6');
    await page.press('[data-scan-cell="r1.length"]', 'Tab');
    await page.waitForFunction(() => !document.querySelector('[data-scan-table] .scan-flag'));
    await page.uncheck('[data-scan-tick="r0.conditionCodes"][data-value="C6"]');
    await page.waitForFunction(() => !document.querySelector('[data-scan-tick="r0.conditionCodes"][data-value="C6"]').checked);
    await page.click('[data-action="scan-accept"]');
    await page.waitForFunction(() => state.view === 'workspace');
    const dayId = (await api('GET', `/api/projects/${workspace}`)).days[0].reportId;
    const day = await api('GET', `/api/reports/${dayId}`);
    expect(day.electrodeResistance.map((row) => `${row.tag}:${row.resistanceWithGrid}:${row.length}`).join() === 'P-1:0.53:3,P-2:1.20:0.6' && JSON.stringify(day.electrodeResistance[0].conditionCodes) === '["C1"]', 'the accepted day holds the corrected values, not the original reading');
    expect(day.loopImpedanceTest.length === 3, 'and the loop page joined the same day');

    // ---- overwrite confirmation
    setTruth({ [electrode.issued.sheetId]: { primary: { ...electrode.truth, 'r0.resistanceWithGrid': { value: '0.47' } } } });
    await page.click('[data-action="go-dashboard"]');
    await page.click('[data-action="open-scans"]');
    await page.setInputFiles('[data-scan-files]', { name: 'again.pdf', mimeType: 'application/pdf', buffer: electrode.pdf });
    await page.click('[data-action="scan-upload"]');
    await page.waitForSelector('.scan-item[data-scan-status="read"] >> nth=0', { timeout: 20000 });
    await page.locator('.scan-item[data-scan-status="read"]').first().locator('[data-action="scan-open"]').click();
    await page.waitForSelector('[data-scan-review]');
    await page.click('[data-action="scan-accept"]');
    await page.waitForSelector('[data-scan-confirm]');
    expect((await page.innerText('[data-scan-confirm]')).includes('0.53 → 0.47') && (await page.innerText('[data-scan-confirm]')).includes('2 existing values') && (await page.innerText('[data-scan-confirm]')).includes('C1 → C1, C6'), 'the overwrite is listed with before and after (the reading, and the condition code the engineer had unticked), nothing written yet');
    await page.click('[data-action="scan-accept-confirm"]');
    await page.waitForFunction(() => state.view === 'workspace');
    expect((await api('GET', `/api/reports/${dayId}`)).electrodeResistance[0].resistanceWithGrid === '0.47', 'confirming overwrites');

    // ---- a page whose QR will not scan: the engineer chooses the sheet
    const [rendered] = await splitUpload({ buffer: electrode.pdf, mimetype: 'application/pdf' });
    setTruth({ electrodeResistance: { primary: { ...electrode.truth, 'r1.tag': { value: 'P-9' } } } });
    await page.click('[data-action="go-dashboard"]');
    await page.click('[data-action="open-scans"]');
    await page.setInputFiles('[data-scan-files]', { name: 'photo.png', mimeType: 'image/png', buffer: withoutQr(rendered) });
    await page.click('[data-action="scan-upload"]');
    await page.waitForSelector('.scan-item[data-scan-status="needs_sheet"]', { timeout: 20000 });
    expect((await page.locator('.scan-item[data-scan-status="needs_sheet"]').innerText()).includes('QR code could not be read'), 'a page without a readable code says so');
    await page.click('[data-action="scan-assign-open"]');
    await page.waitForSelector('[data-scan-assign]');
    expect(await page.locator('[data-action="scan-assign"]').isDisabled(), 'the choice needs project, date and sheet');
    await page.selectOption('[data-scan-assign-field="project"]', workspace);
    await page.fill('[data-scan-assign-field="workDate"]', '2026-10-13');
    await page.selectOption('[data-scan-assign-field="sheetKey"]', 'electrodeResistance');
    await page.fill('[data-scan-assign-field="start"]', '11');
    await page.click('[data-action="scan-assign"]');
    await page.waitForSelector('.scan-item[data-scan-status="read"]', { timeout: 20000 });
    await page.locator('.scan-item[data-scan-status="read"]').first().locator('[data-action="scan-open"]').click();
    await page.waitForSelector('[data-scan-review]');
    expect((await page.innerText('[data-scan-review]')).includes('found by chosen by the engineer') && (await page.innerText('[data-scan-review]')).includes('rows from 11'), 'the review shows it was chosen by hand and where its rows start');
    await page.click('[data-action="scan-discard-current"]');
    await page.waitForSelector('[data-scans]');

    // ---- the suggest-wording helper on a flagged row
    const html = await page.evaluate(() => [suggestButton('electrodeResistance', { rowId: 'r1' }, { status: { tone: 'critical' }, priority: 'P1' }), suggestButton('electrodeResistance', { rowId: 'r1' }, { status: { tone: 'healthy' }, priority: 'P3' }), suggestButton('electrodeResistance', {}, { status: { tone: 'critical' } })]);
    expect(html[0].includes('data-action="suggest-wording"') && html[1] === '' && html[2] === '', 'the helper is offered on flagged rows only');
    const inserted = await page.evaluate(async () => {
      state.draft = createDraft();
      const row = state.draft.electrodeResistance[0];
      row.conditionCodes = ['C1', 'C6'];
      handleScanAction('suggest-wording', { dataset: { section: 'electrodeResistance', rowId: row.rowId } });
      for (let i = 0; i < 50 && !(suggestUi[row.rowId] && suggestUi[row.rowId].text); i += 1) await new Promise((resolve) => setTimeout(resolve, 100));
      const suggested = suggestUi[row.rowId] && suggestUi[row.rowId].text;
      const before = row.observation;
      handleScanAction('suggest-insert', { dataset: { rowId: row.rowId } });
      return { suggested, before, after: row.observation, cleared: !suggestUi[row.rowId] };
    });
    expect(/corrosion is visible and a connection is loose/.test(inserted.suggested) && !inserted.before && inserted.after === inserted.suggested && inserted.cleared, 'Suggest wording drafts text from the condition codes and Insert puts it in the remark field only on the click');

    // ---- admin: purge and usage
    const admin = await browser.newContext({ viewport: { width: 1400, height: 1000 } });
    await browserLogin(admin, base, 'admin@elegrow.com');
    const adminPage = await admin.newPage();
    await adminPage.goto(`${base}/`);
    await adminPage.waitForSelector('.user-bar');
    await adminPage.click('[data-action="open-admin"]');
    await adminPage.click('[data-tab="scans"]');
    await adminPage.waitForSelector('[data-ai-usage]');
    expect((await adminPage.innerText('[data-ai-usage]')).includes('gemini'), 'the admin sees the reader calls');
    expect(await adminPage.locator('[data-action="scan-purge"]').isDisabled(), 'purge needs a date range');
    await adminPage.fill('[data-scan-purge-field="from"]', '2001-01-01');
    await adminPage.fill('[data-scan-purge-field="to"]', '2001-12-31');
    await adminPage.click('[data-action="scan-purge"]');
    await adminPage.waitForSelector('[data-scan-purge-message]');
    expect((await adminPage.innerText('[data-scan-purge-message]')).includes('0 scans deleted'), 'a purge says how many it deleted');
    expect(errors.length === 0, `no page errors (${errors.join('; ')})`);
  } finally {
    await browser.close();
    server.kill();
  }
  console.log(`\nscan-ui: ${checks} checks`);
  if (failures.length) {
    console.log(`${failures.length} FAILED`);
    process.exit(1);
  }
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
