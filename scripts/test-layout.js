// Layout and data-entry test (Sessions 6B.1c and 6B.1d): opens every measurement sheet screen
// MS-01..MS-09 in a real browser at 1920, 1440, 1288, 1280, 1024, 768 and 390 px wide, filled with
// worst-case content (long text, every condition code ticked, an assessment on every row).
//
// Page:    the page never scrolls sideways (document.scrollWidth <= window.innerWidth).
// Tables:  MS-01, 03, 06, 07 fit the page at >= 1024 px (no horizontal scroll). The dense sheets
//          MS-02, 04, 05, 08, 09 may scroll inside their frame at any width; every column is at
//          least its minimum width (tick options on one line, numbers fully visible, Remark >= 160 px,
//          Photo ref >= 80 px) and Sr/TAG stay left, assessment/actions stay right while it scrolls.
// Entry:   no pre-filled values, empty dates read dd/mm/yyyy, Enter/Tab flow, a new row when the
//          last one is used, copy previous row, fill down, paste from Excel, step dots, set-up
//          diagrams, explanation popover, WCAG AA contrast on chips and card labels.
// Also saves review screenshots to reference/screenshots/ (git-ignored: they may show client data).
//
// Needs Playwright's Chromium (npx playwright install chromium). Starts the real server on a free
// port; nothing is saved to the database.
const fs = require('fs');
const net = require('net');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');

const ROOT = path.join(__dirname, '..');
const SHOTS = path.join(ROOT, 'reference', 'screenshots');
const WIDTHS = [1920, 1440, 1288, 1280, 1024, 768, 390];
const HEIGHT = 900;
const SHEET_IDS = ['soilResistivity', 'electrodeResistance', 'continuityTest', 'loopImpedanceTest', 'prospectiveFaultCurrent', 'riserIntegrityTest', 'groundResistance', 'stepTouchVoltage', 'towerFootingResistance'];
const DENSE = new Set(['electrodeResistance', 'loopImpedanceTest', 'prospectiveFaultCurrent', 'stepTouchVoltage', 'towerFootingResistance']);
const SHOT_SHEETS = { electrodeResistance: 'MS-02', loopImpedanceTest: 'MS-04', stepTouchVoltage: 'MS-08', towerFootingResistance: 'MS-09' };
const SHOT_WIDTHS = [1288, 1440, 1920, 390];

const failures = [];
let checks = 0;
function expect(condition, message) {
  checks += 1;
  if (!condition) {
    failures.push(message);
    console.log(`  FAIL - ${message}`);
  }
}

function freePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address();
      server.close(() => resolve(port));
    });
    server.on('error', reject);
  });
}

async function startServer() {
  const port = await freePort();
  const child = spawn(process.execPath, ['src/server.js'], {
    cwd: ROOT,
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_HOST: '127.0.0.1', ZOHO_USE_MOCK: 'true' },
    stdio: ['ignore', 'pipe', 'pipe']
  });
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('server did not start')), 15000);
    child.stdout.on('data', (chunk) => {
      if (String(chunk).includes('running on')) {
        clearTimeout(timer);
        resolve();
      }
    });
    child.on('exit', (code) => reject(new Error(`server exited with ${code}`)));
  });
  return { child, url: `http://127.0.0.1:${port}` };
}

// Runs inside the page: select every sheet and fill it with worst-case content.
const FILL_DRAFT = () => {
  const longText = 'Reactor Bay R-201 north wall';
  const fillValue = (field) => {
    if (field.kind === 'number') return '123.45';
    if (field.kind === 'ticks') {
      const options = optionList(field).map((option) => option.value);
      return field.multi ? options : options[options.length - 1];
    }
    return longText;
  };
  const fillTarget = (target, field) => {
    if (field.kind === 'serial' || field.kind === 'fixed' || field.legacy || field.extra) return;
    target[field.key] = fillValue(field);
  };
  state.draft.project.projectNo = 'PRJ-2026-001';
  state.draft.project.clientName = 'Sample Client';
  allSheets().forEach((sheet) => {
    state.draft.tests[sheet.id] = !sheet.retired;
    if (sheet.retired || sheet.id === 'soilResistivity') return;
    if (sheet.layout === 'towerGroups') {
      state.draft[sheet.id] = [newTowerGroup('1'), newTowerGroup('2')];
      state.draft[sheet.id].forEach((group) => {
        sheet.rowFields.forEach((field) => (field.merge ? fillTarget(group, field) : (group.readings || []).forEach((reading) => fillTarget(reading, field))));
      });
      return;
    }
    if (sheet.layout === 'phaseGroups') {
      state.draft[sheet.id] = [...newPhaseGroup(sheet.id, 1), ...newPhaseGroup(sheet.id, 2)];
    } else {
      state.draft[sheet.id] = [newSheetRow(sheet.id, { srNo: '1' }), newSheetRow(sheet.id, { srNo: '2' }), newSheetRow(sheet.id, { srNo: '3' })];
    }
    state.draft[sheet.id].forEach((row) => sheet.rowFields.forEach((field) => fillTarget(row, field)));
  });
  const preset = soilPresetFor(state.draft.project.soilSpacingPreset);
  state.draft.soilResistivity.locations[0].inputUnit = 'ohm-m';
  preset.spacings.forEach((spacing) => {
    ['direction1', 'direction2', 'direction3', 'direction4'].forEach((direction) => {
      applyFieldInput({ dataset: { soilCell: 'true', locationIndex: '0', direction, spacing, field: 'resistivity' }, type: 'text', value: '123.45' });
    });
  });
};

// Runs inside the page: how far does anything stick out?
const MEASURE = () => {
  const doc = document.documentElement;
  const tables = [...document.querySelectorAll('.builder-main table.data-table')];
  return {
    pageScrollWidth: Math.max(doc.scrollWidth, document.body.scrollWidth),
    innerWidth: window.innerWidth,
    tables: tables.map((table) => {
      const wrap = table.closest('.table-wrap');
      const wrapRect = wrap.getBoundingClientRect();
      const tableRect = table.getBoundingClientRect();
      let widestCell = 0;
      table.querySelectorAll('td, th').forEach((cell) => {
        const rect = cell.getBoundingClientRect();
        widestCell = Math.max(widestCell, rect.right);
        cell.querySelectorAll('input, button, .tick, .pill, .table-input-static').forEach((child) => {
          const childRect = child.getBoundingClientRect();
          if (childRect.width > 0 && childRect.right > rect.right + 1 && getComputedStyle(child).display !== 'none') {
            widestCell = Math.max(widestCell, childRect.right);
          }
        });
      });
      return {
        wrapScrolls: wrap.scrollWidth > wrap.clientWidth + 1,
        tableRight: tableRect.right,
        wrapRight: wrapRect.right,
        tableOverflow: table.scrollWidth - wrap.clientWidth,
        contentRight: widestCell
      };
    })
  };
};

// Runs inside the page: every column of the first dense table against its minimum.
const MEASURE_COLUMNS = () => {
  const table = document.querySelector('.builder-main table.sheet-dense');
  const problems = [];
  const cols = [...table.querySelectorAll('col')];
  [...table.querySelectorAll('thead th')].forEach((head, index) => {
    const label = (head.querySelector('.th-label') || head).textContent.trim();
    const width = head.getBoundingClientRect().width;
    if (/^Remark$/.test(label) && width < 159.5) problems.push(`Remark column ${Math.round(width)} px (min 160)`);
    if (/^Photo ref\.?$/.test(label) && width < 79.5) problems.push(`Photo ref column ${Math.round(width)} px (min 80)`);
    if (cols[index] && cols[index].classList.contains('col-number') && width < 87.5) problems.push(`number column "${label}" ${Math.round(width)} px (min 88)`);
  });
  // Tick options on one line; numbers whole.
  table.querySelectorAll('.tick-group').forEach((group) => {
    const tops = [...group.querySelectorAll('.tick')].map((tick) => Math.round(tick.getBoundingClientRect().top));
    if (new Set(tops).size > 1) problems.push(`tick options wrap in "${group.getAttribute('aria-label')}"`);
  });
  table.querySelectorAll('td.sheet-cell-number input').forEach((input) => {
    if (input.scrollWidth > input.clientWidth + 1) problems.push(`number "${input.getAttribute('aria-label')}" is cut off`);
  });
  return { problems };
};

async function openBuilder(page, base) {
  await page.goto(base, { waitUntil: 'networkidle' });
  await page.click('[data-action="new-report"]');
  await page.evaluate(FILL_DRAFT);
  await page.evaluate(() => render());
  await page.waitForLoadState('networkidle');
}

async function showStep(page, sheetId) {
  await page.evaluate((id) => {
    state.stepIndex = getSteps().findIndex((step) => step.id === id);
    render();
  }, sheetId);
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(250);
}

async function checkSheet(page, width, sheetId, label) {
  await showStep(page, sheetId);
  const measure = await page.evaluate(MEASURE);
  expect(measure.pageScrollWidth <= measure.innerWidth, `${label} @${width}: page scrolls sideways (${measure.pageScrollWidth} > ${measure.innerWidth})`);
  expect(measure.tables.length > 0, `${label} @${width}: no table found`);
  const dense = DENSE.has(sheetId);
  measure.tables.forEach((table, index) => {
    if (width >= 1024 && !dense) {
      expect(!table.wrapScrolls && table.tableOverflow <= 1, `${label} @${width}: table ${index + 1} overflows its container by ${Math.round(table.tableOverflow)} px`);
      expect(table.tableRight <= table.wrapRight + 1, `${label} @${width}: table ${index + 1} sticks out ${Math.round(table.tableRight - table.wrapRight)} px`);
      expect(table.contentRight <= table.wrapRight + 1, `${label} @${width}: cell content sticks out ${Math.round(table.contentRight - table.wrapRight)} px (table ${index + 1})`);
    }
  });
  if (dense && width >= 721) {
    const columns = await page.evaluate(MEASURE_COLUMNS);
    columns.problems.forEach((problem) => expect(false, `${label} @${width}: ${problem}`));
    expect(columns.problems.length === 0, `${label} @${width}: columns below their minimum width`);
  }
  return measure;
}

// Non-dense sheets: the header row stays on screen while the page scrolls.
async function checkHeaderSticksOnPage(page, width, label) {
  const result = await page.evaluate(() => {
    const table = document.querySelector('.builder-main table.sheet-table');
    const head = table.querySelector('thead th');
    window.scrollTo({ top: window.scrollY + table.getBoundingClientRect().top + 140, behavior: 'instant' });
    return new Promise((resolve) => requestAnimationFrame(() => resolve({ headTop: head.getBoundingClientRect().top, tableTop: table.getBoundingClientRect().top })));
  });
  expect(result.tableTop < -20 && result.headTop >= -1 && result.headTop <= 40, `${label} @${width}: header row does not stay on screen (table top ${Math.round(result.tableTop)}, header top ${Math.round(result.headTop)})`);
  await page.evaluate(() => window.scrollTo({ top: 0, behavior: 'instant' }));
}

// Dense sheets scroll inside their frame: header, Sr/TAG and assessment/actions stay put.
async function checkDenseSticky(page, width, label) {
  await page.evaluate(() => {
    state.draft.electrodeResistance = Array.from({ length: 16 }, (_, index) => newSheetRow('electrodeResistance', { srNo: String(index + 1), tag: `EP-${index + 1}`, location: 'Reactor Bay' }));
    render();
  });
  await page.waitForTimeout(250);
  const result = await page.evaluate(() => {
    const wrap = document.querySelector('.builder-main .sheet-dense-wrap');
    const table = wrap.querySelector('table');
    const sample = (selector) => table.querySelector(selector);
    const first = sample('tbody tr:nth-child(8) .sticky-left-1');
    const second = sample('tbody tr:nth-child(8) .sticky-left-2');
    const status = sample('tbody tr:nth-child(8) .sticky-right-2');
    const action = sample('tbody tr:nth-child(8) .sticky-right-1');
    const head = sample('thead th');
    const rect = (node) => node.getBoundingClientRect();
    const before = [first, second, status, action].map((node) => rect(node).left - rect(wrap).left);
    wrap.scrollLeft = 400;
    wrap.scrollTop = 120;
    const after = [first, second, status, action].map((node) => rect(node).left - rect(wrap).left);
    return {
      scrollsX: wrap.scrollWidth > wrap.clientWidth + 1,
      scrollsY: wrap.scrollHeight > wrap.clientHeight + 1,
      movedX: wrap.scrollLeft,
      leftDrift: Math.max(Math.abs(before[0] - after[0]), Math.abs(before[1] - after[1])),
      rightDrift: Math.max(Math.abs(before[2] - after[2]), Math.abs(before[3] - after[3])),
      headTop: rect(head).top - rect(wrap).top,
      positions: [first, second, status, action].map((node) => getComputedStyle(node).position)
    };
  });
  expect(result.scrollsX, `${label} @${width}: a dense sheet should scroll sideways inside its frame`);
  expect(result.positions.every((position) => position === 'sticky'), `${label} @${width}: Sr/TAG/assessment/action cells are not sticky (${result.positions})`);
  expect(result.movedX > 0 && result.leftDrift <= 1, `${label} @${width}: Sr/TAG columns drifted ${result.leftDrift.toFixed(1)} px while scrolling`);
  expect(result.rightDrift <= 1, `${label} @${width}: assessment/action columns drifted ${result.rightDrift.toFixed(1)} px while scrolling`);
  expect(!result.scrollsY || Math.abs(result.headTop) <= 1, `${label} @${width}: header row left the frame (${Math.round(result.headTop)} px)`);
  const pageOverflow = await page.evaluate(() => Math.max(document.documentElement.scrollWidth, document.body.scrollWidth) - window.innerWidth);
  expect(pageOverflow <= 0, `${label} @${width}: the page scrolls sideways while the table scrolls`);
}

async function checkPicker(page, width) {
  await showStep(page, 'groundResistance');
  const state = await page.evaluate(() => {
    const cell = document.querySelector('.builder-main [data-tick-cell*="-method-"]');
    return { summaryShown: getComputedStyle(cell.querySelector('.tick-summary')).display !== 'none', groupShown: getComputedStyle(cell.querySelector('.tick-group')).display !== 'none' };
  });
  expect(state.summaryShown && !state.groupShown, `MS-07 @${width}: tick groups are not collapsed to a summary (summary ${state.summaryShown}, group ${state.groupShown})`);
  await page.click('.builder-main [data-tick-cell*="-method-"] .tick-summary');
  const open = await page.evaluate(() => {
    const group = document.querySelector('.builder-main [data-tick-cell*="-method-"] .tick-group');
    const rect = group.getBoundingClientRect();
    return { open: group.matches(':popover-open'), inView: rect.left >= 0 && rect.right <= window.innerWidth && rect.top >= 0 && rect.bottom <= window.innerHeight, options: group.querySelectorAll('input').length };
  });
  expect(open.open && open.inView && open.options === 3, `MS-07 @${width}: method picker did not open on tap (open ${open.open}, in view ${open.inView}, options ${open.options})`);
  await page.click('.builder-main [data-tick-cell*="-method-"] .tick-group .tick:has(input[value="Slope"])');
  await page.waitForTimeout(150);
  const after = await page.evaluate(() => {
    const cell = document.querySelector('.builder-main [data-tick-cell*="-method-"]');
    return { text: cell.querySelector('.tick-summary').textContent.trim(), open: cell.querySelector('.tick-group').matches(':popover-open'), value: state.draft.groundResistance[0].method };
  });
  expect(after.value === 'Slope' && after.text.startsWith('Slope') && !after.open, `MS-07 @${width}: picking a method did not update and close (value ${after.value}, text "${after.text}", open ${after.open})`);
  await page.click('.builder-main [data-tick-cell*="conditionCodes"] .tick-summary');
  await page.click('.builder-main [data-tick-cell*="conditionCodes"] .tick-group .tick:has(input[data-code="C6"])');
  await page.waitForTimeout(150);
  const codes = await page.evaluate(() => {
    const cell = document.querySelector('.builder-main [data-tick-cell*="conditionCodes"]');
    return { open: cell.querySelector('.tick-group').matches(':popover-open'), codes: state.draft.groundResistance[0].conditionCodes };
  });
  expect(codes.open && !codes.codes.includes('C6'), `MS-07 @${width}: condition-code picker should stay open and toggle C6 (open ${codes.open}, codes ${codes.codes})`);
  await page.keyboard.press('Escape');
  const closed = await page.evaluate(() => !document.querySelector('.builder-main [data-tick-cell*="conditionCodes"] .tick-group').matches(':popover-open'));
  expect(closed, `MS-07 @${width}: Escape did not close the picker`);
}

// A: nothing is pre-filled; empty dates read dd/mm/yyyy; a reading without its selection is flagged.
async function checkNoDefaults(page) {
  const defaults = await page.evaluate(() => {
    const draft = createDraft();
    const row = (id) => newSheetRow(id, { srNo: '1' });
    return {
      electrode: [row('electrodeResistance').electrodeType, row('electrodeResistance').materialType],
      loop: [row('loopImpedanceTest').deviceType, row('loopImpedanceTest').voltage],
      pfc: [row('prospectiveFaultCurrent').deviceType, row('prospectiveFaultCurrent').voltage],
      foot: newTowerGroup('1').readings.map((reading) => reading.footToEarthingConnectionStatus),
      unit: draft.soilResistivity.locations[0].inputUnit
    };
  });
  expect(defaults.electrode.every((value) => value === ''), `electrode type/material are pre-filled: ${defaults.electrode}`);
  expect(defaults.loop.every((value) => value === '') && defaults.pfc.every((value) => value === ''), `device type / voltage are pre-filled: ${defaults.loop} ${defaults.pfc}`);
  expect(defaults.foot.every((value) => value === ''), `tower foot connection is pre-filled: ${defaults.foot}`);
  expect(defaults.unit === '', `soil reading unit is pre-filled: ${defaults.unit}`);
  await page.evaluate(() => {
    state.draft.tests.electrodeResistance = true;
    state.draft.electrodeResistance = [newSheetRow('electrodeResistance', { srNo: '1', resistanceWithoutGrid: '2.5' })];
    sheetHeaderFor('electrodeResistance').calibrationDue = '';
    state.stepIndex = getSteps().findIndex((step) => step.id === 'electrodeResistance');
    render();
  });
  await page.waitForTimeout(900);
  await showStep(page, 'electrodeResistance');
  const placeholder = await page.evaluate(() => {
    const field = [...document.querySelectorAll('.builder-main .date-field')][0];
    const text = field.querySelector('.date-placeholder');
    return { empty: field.classList.contains('is-empty'), shown: getComputedStyle(text).display !== 'none', text: text.textContent };
  });
  expect(placeholder.empty && placeholder.shown && placeholder.text === 'dd/mm/yyyy', `empty date does not show dd/mm/yyyy (${JSON.stringify(placeholder)})`);
  const warned = await page.evaluate(() => (document.querySelector('.builder-main .sheet-problems') || { textContent: '' }).textContent);
  expect(/Electrode type and Material are not ticked/.test(warned), `a reading without electrode type/material gets no warning (problems: "${warned.slice(0, 160)}")`);
}

// D: Enter / Tab flow, a new row when the last is used, copy previous row, fill down, paste.
async function checkFastEntry(page) {
  await page.evaluate(() => {
    state.draft.electrodeResistance = [newSheetRow('electrodeResistance', { srNo: '1' }), newSheetRow('electrodeResistance', { srNo: '2' })];
    render();
  });
  await showStep(page, 'electrodeResistance');
  const field = (index, key) => `.builder-main [data-section="electrodeResistance"][data-index="${index}"][data-field="${key}"]`;
  await page.focus(field(0, 'length'));
  await page.keyboard.type('3');
  await page.keyboard.press('Enter');
  const focused = await page.evaluate(() => ({ index: document.activeElement.dataset.index, field: document.activeElement.dataset.field }));
  expect(focused.index === '1' && focused.field === 'length', `Enter did not move down the same column (${JSON.stringify(focused)})`);
  await page.keyboard.type('4');
  await page.waitForTimeout(150);
  const rowsAfterTyping = await page.evaluate(() => state.draft.electrodeResistance.length);
  expect(rowsAfterTyping === 3, `typing in the last row did not add a row (rows ${rowsAfterTyping})`);
  await page.focus(field(1, 'tag'));
  await page.keyboard.press('Tab');
  const tabTarget = await page.evaluate(() => ({ index: document.activeElement.dataset.index, field: document.activeElement.dataset.field }));
  expect(tabTarget.index === '1' && tabTarget.field === 'location', `Tab did not move across the row (${JSON.stringify(tabTarget)})`);

  await page.evaluate(() => {
    Object.assign(state.draft.electrodeResistance[0], { electrodeType: 'Pipe', materialType: 'GI', length: '3', diameter: '25' });
    render();
  });
  await page.click('.builder-main [data-action="copy-prev-row"][data-index="1"]');
  const copied = await page.evaluate(() => state.draft.electrodeResistance[1]);
  expect(copied.electrodeType === 'Pipe' && copied.materialType === 'GI' && copied.length === '3' && copied.diameter === '25', `copy previous row did not copy type/material/dimensions (${JSON.stringify([copied.electrodeType, copied.materialType, copied.length, copied.diameter])})`);

  await page.evaluate(() => {
    state.draft.electrodeResistance.forEach((row) => (row.tag = ''));
    state.draft.electrodeResistance[0].tag = 'EP-1';
    render();
  });
  await page.click('.builder-main .fill-down[data-field="tag"]');
  const filled = await page.evaluate(() => state.draft.electrodeResistance.map((row) => row.tag));
  expect(filled.every((tag) => tag === 'EP-1'), `fill down did not fill the empty cells (${filled})`);

  await page.evaluate(() => {
    state.draft.electrodeResistance = [newSheetRow('electrodeResistance', { srNo: '1' })];
    render();
  });
  await page.evaluate(() => {
    const target = document.querySelector('[data-section="electrodeResistance"][data-index="0"][data-field="tag"]');
    const data = new DataTransfer();
    data.setData('text/plain', 'T-1\tLoc A\tRod\tCu\nT-2\tLoc B\tPipe\tGI\nT-3\tLoc C\tStrip\tSS\n');
    target.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await page.waitForTimeout(150);
  const pasted = await page.evaluate(() => state.draft.electrodeResistance.map((row) => [row.tag, row.location, row.electrodeType, row.materialType].join('|')));
  expect(pasted.length >= 3 && pasted[0] === 'T-1|Loc A|Rod|Cu' && pasted[1] === 'T-2|Loc B|Pipe|GI' && pasted[2] === 'T-3|Loc C|Strip|SS', `pasting a tab-separated block did not fill the rows (${pasted})`);
}

// C: assessment explanation, tower totals, header grid, autofill bar, set-up diagrams, step dots.
async function checkSheetFeatures(page) {
  await page.evaluate(FILL_DRAFT);
  await showStep(page, 'loopImpedanceTest');
  const explain = await page.evaluate(() => ({ buttons: document.querySelectorAll('.builder-main .explain-btn').length, groups: state.draft.loopImpedanceTest.length / 3, inline: [...document.querySelectorAll('.builder-main .status-explain-inline')].some((node) => getComputedStyle(node).display !== 'none') }));
  expect(explain.buttons === explain.groups, `MS-04: expected one explanation button per feeder (${explain.groups}), found ${explain.buttons}`);
  expect(!explain.inline, 'MS-04: the full explanation is still printed inside the assessment cell');
  await page.hover('.builder-main .explain-btn');
  await page.waitForTimeout(150);
  const popover = await page.evaluate(() => {
    const pop = document.getElementById('explainPop');
    return { open: Boolean(pop) && pop.matches(':popover-open'), text: pop ? pop.textContent.length : 0 };
  });
  expect(popover.open && popover.text > 10, `hovering the "i" button did not open the explanation (${JSON.stringify(popover)})`);

  await showStep(page, 'towerFootingResistance');
  const tower = await page.evaluate(() => {
    const totals = document.querySelector('.builder-main .tower-totals');
    const remove = document.querySelector('.builder-main [data-action="remove-tower-group"]');
    return { totals: totals ? totals.textContent.trim() : '', removeIsIcon: Boolean(remove) && remove.classList.contains('icon-button') && /Remove tower/.test(remove.getAttribute('aria-label')) && remove.textContent.trim() === '×' };
  });
  expect(/^Zt .+ Ω · Zsat .+ Ω · .+/.test(tower.totals), `MS-09: totals line is not "Zt … Ω · Zsat … Ω · status" (${tower.totals})`);
  expect(tower.removeIsIcon, 'MS-09: Remove tower is not an icon button');

  await showStep(page, 'stepTouchVoltage');
  const header = await page.evaluate(() => [...document.querySelectorAll('.builder-main .header-grid > .header-cell:not(.header-cell-wide)')].map((cell) => Math.round(cell.getBoundingClientRect().width)));
  expect(header.length > 0 && header.every((value) => value >= 179 && value <= 281), `MS-08: header fields are not 180–280 px wide (${header})`);

  await showStep(page, 'electrodeResistance');
  const bar = await page.evaluate(() => Math.round(document.querySelector('.builder-main .soil-ocr-assist').getBoundingClientRect().height));
  expect(bar <= 70, `MS-02: the autofill block is ${bar} px tall, expected a slim bar (<= 70)`);

  for (const sheetId of SHEET_IDS) {
    await showStep(page, sheetId);
    await page.click('.builder-main details.sheet-notes summary');
    const diagram = await page.evaluate(() => {
      const svg = document.querySelector('.builder-main details.sheet-notes[open] svg.setup-diagram');
      return svg ? { shapes: svg.querySelectorAll('line, polyline, rect, circle').length, texts: svg.querySelectorAll('text').length, items: document.querySelectorAll('.builder-main .howto-steps li').length } : null;
    });
    expect(diagram && diagram.shapes > 5 && diagram.texts > 2 && diagram.items > 0, `${sheetId}: "How to record this sheet" has no SET-UP diagram and instructions (${JSON.stringify(diagram)})`);
  }

  const dots = await page.evaluate(() => [...document.querySelectorAll('.stepper .step-pill')].map((pill) => ({ label: pill.querySelector('strong').textContent, dot: (pill.querySelector('.step-dot') || { className: '' }).className })));
  expect(dots.length > 3 && dots.every((entry) => /step-dot-(not-started|in-progress|issues|complete)/.test(entry.dot)), `a step tab has no status dot (${JSON.stringify(dots.filter((entry) => !entry.dot))})`);
  const known = await page.evaluate(() => {
    const saved = state.draft.electrodeResistance;
    state.draft.electrodeResistance = [newSheetRow('electrodeResistance', { srNo: '1' })];
    const empty = stepStatus('electrodeResistance');
    state.draft.electrodeResistance = saved;
    return { selection: stepStatus('selection'), empty };
  });
  expect(known.selection === 'complete' && known.empty === 'not-started', `step status dots are wrong (${JSON.stringify(known)})`);
}

function contrastRatio(a, b) {
  const lum = ([r, g, bl]) => {
    const channel = (v) => {
      const s = v / 255;
      return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
    };
    return 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(bl);
  };
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x);
  return (hi + 0.05) / (lo + 0.05);
}

// B: dashboard, Project and Selection screens, observation drawer.
async function checkFormScreens(browser, base, width) {
  const context = await browser.newContext({ viewport: { width, height: HEIGHT } });
  const page = await context.newPage();
  await page.goto(base, { waitUntil: 'networkidle' });
  let measure = await page.evaluate(MEASURE);
  expect(measure.pageScrollWidth <= measure.innerWidth, `dashboard @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  await page.click('[data-action="new-report"]');
  await page.evaluate(() => {
    state.draft.project.projectNo = 'PRJ-2026-001';
    state.draft.project.clientName = 'Sample Client';
    state.draft.project.zohoProjectId = 'z-1';
    state.draft.project.zohoProjectName = 'Sample Zoho project';
    state.stepIndex = 0;
    render();
  });
  measure = await page.evaluate(MEASURE);
  expect(measure.pageScrollWidth <= measure.innerWidth, `project step @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  const project = await page.evaluate(() => {
    const labels = [...document.querySelectorAll('main.app-shell span, main.app-shell label, main.app-shell strong')].filter((node) => node.children.length === 0 && /^project number$/i.test(node.textContent.trim()));
    const rail = document.querySelector('.builder-layout-form .builder-aside');
    const layout = document.querySelector('.builder-layout-form');
    return { count: labels.length, railWidth: Math.round(rail.getBoundingClientRect().width), railSticky: getComputedStyle(rail).position, layoutWidth: layout.getBoundingClientRect().width, formRight: document.querySelector('.builder-main').getBoundingClientRect().right - layout.getBoundingClientRect().left };
  });
  expect(project.count === 1, `project step @${width}: "Project number" appears ${project.count} times`);
  if (width >= 1024) {
    expect(project.railWidth >= 339 && project.railWidth <= 381, `project step @${width}: right rail is ${project.railWidth} px (340–380)`);
    expect(project.railSticky === 'sticky', `project step @${width}: right rail is not sticky`);
    expect(project.formRight > project.layoutWidth * 0.55, `project step @${width}: the form does not use the width`);
  }
  await page.evaluate(() => {
    state.stepIndex = 1;
    render();
  });
  measure = await page.evaluate(MEASURE);
  expect(measure.pageScrollWidth <= measure.innerWidth, `selection step @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  if (width === 1440) {
    const contrast = await page.evaluate(() => {
      const parse = (value) => (value.match(/[\d.]+/g) || []).slice(0, 3).map(Number);
      const rows = [];
      const measureNode = (selector, what) => {
        const node = document.querySelector(selector);
        if (!node) return;
        const style = getComputedStyle(node);
        let background = style.backgroundColor;
        let walker = node;
        while ((/, 0\)$/.test(background) || background === 'transparent') && walker.parentElement) {
          walker = walker.parentElement;
          background = getComputedStyle(walker).backgroundColor;
        }
        // Gradient backgrounds have no solid colour: the card sits on white / pale blue (#dbeafe at worst).
        const solid = !/, 0\)$/.test(background) && background !== 'transparent';
        rows.push({ what, color: parse(style.color), background: solid ? parse(background) : [219, 234, 254] });
      };
      measureNode('.selection-tile .soft-chip', 'MS-0x card label');
      measureNode('.selection-tile p', 'card description');
      measureNode('.selection-tile .toggle-mark', 'card toggle');
      measureNode('.selection-tile-active .toggle-mark', 'selected card toggle');
      measureNode('.builder-followup .soft-chip', 'scope chip');
      return rows;
    });
    contrast.forEach((row) => expect(contrastRatio(row.color, row.background) >= 4.5, `selection step: ${row.what} contrast ${contrastRatio(row.color, row.background).toFixed(2)} < 4.5`));
    expect(contrast.length === 5, `selection step: contrast sample missing (${contrast.length} of 5)`);
  }
  await page.evaluate(FILL_DRAFT);
  await showStep(page, 'electrodeResistance');
  await page.evaluate(() => openObservationEditor('electrodeResistance', 0));
  await page.waitForTimeout(200);
  measure = await page.evaluate(MEASURE);
  expect(measure.pageScrollWidth <= measure.innerWidth, `observation drawer @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  await context.close();
}

async function shootForms(browser, base, width) {
  const context = await browser.newContext({ viewport: { width, height: HEIGHT } });
  const page = await context.newPage();
  await page.goto(base, { waitUntil: 'networkidle' });
  await page.click('[data-action="new-report"]');
  for (const [index, name] of [[0, 'Project'], [1, 'Selection']]) {
    await page.evaluate((i) => {
      state.draft.project.projectNo = 'PRJ-2026-001';
      state.draft.project.clientName = 'Sample Client';
      state.stepIndex = i;
      render();
    }, index);
    await page.evaluate(() => window.scrollTo({ top: 0, behavior: 'instant' }));
    await page.screenshot({ path: path.join(SHOTS, `${name}_${width}.png`), fullPage: true });
  }
  await context.close();
}

async function shoot(page, width, sheetId, code) {
  await showStep(page, sheetId);
  await page.evaluate(() => window.scrollTo({ top: 0, behavior: 'instant' }));
  await page.screenshot({ path: path.join(SHOTS, `${code}_${width}.png`), fullPage: true });
}

(async () => {
  fs.mkdirSync(SHOTS, { recursive: true });
  const { child, url } = await startServer();
  const browser = await chromium.launch();
  const pageErrors = [];
  try {
    for (const width of WIDTHS) {
      console.log(`width ${width}px`);
      const context = await browser.newContext({ viewport: { width, height: HEIGHT } });
      const page = await context.newPage();
      page.on('pageerror', (error) => pageErrors.push(`@${width}: ${error.message}`));
      await openBuilder(page, url);

      for (const sheetId of SHEET_IDS) {
        const code = (await page.evaluate((id) => sheetById(id).code, sheetId)) || sheetId;
        await checkSheet(page, width, sheetId, code);
        if (width >= 1280 && sheetId === 'groundResistance') {
          await checkHeaderSticksOnPage(page, width, code);
        }
      }
      if (width >= 1024 && width <= 1279) {
        await checkPicker(page, width);
      } else if (width >= 1280) {
        await showStep(page, 'groundResistance');
        const visible = await page.evaluate(() => {
          const cell = document.querySelector('.builder-main [data-tick-cell*="-method-"]');
          return getComputedStyle(cell.querySelector('.tick-group')).display !== 'none' && getComputedStyle(cell.querySelector('.tick-summary')).display === 'none';
        });
        expect(visible, `MS-07 @${width}: ticks should be shown, not collapsed`);
      }
      if (SHOT_WIDTHS.includes(width)) {
        await page.evaluate(FILL_DRAFT);
        for (const sheetId of Object.keys(SHOT_SHEETS)) {
          await shoot(page, width, sheetId, SHOT_SHEETS[sheetId]);
        }
        await shootForms(browser, url, width);
      }
      if (width >= 721) {
        await showStep(page, 'electrodeResistance');
        await checkDenseSticky(page, width, 'MS-02');
        await page.evaluate(FILL_DRAFT);
      }
      if (width === 1440) {
        await checkSheetFeatures(page);
        await checkNoDefaults(page);
        await checkFastEntry(page);
      }
      await checkFormScreens(browser, url, width);
      await context.close();
    }
  } finally {
    await browser.close();
    child.kill();
  }
  pageErrors.forEach((message) => expect(false, `script error ${message}`));

  if (failures.length) {
    console.log(`\n${failures.length} of ${checks} layout checks failed`);
    process.exit(1);
  }
  console.log(`\nlayout: ${checks} checks passed across ${WIDTHS.length} widths x ${SHEET_IDS.length} sheets`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
