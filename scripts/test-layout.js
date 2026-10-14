// Layout test (Session 6B.1c): opens every measurement sheet screen MS-01..MS-09 in a real
// browser at 1920, 1440, 1280, 1024, 768 and 390 px wide, filled with worst-case content
// (long text, every condition code ticked, an assessment on every row), and fails if
//   - the page scrolls sideways (document.scrollWidth > window.innerWidth) at any width,
//   - at >= 1024 px the table overflows its container or any cell content leaves the table,
//   - at >= 1024 px the header row does not stay on screen while the page scrolls,
//   - at 721-1023 px the Sr / TAG columns do not stay put while the table scrolls sideways,
//   - at 1024-1279 px the tick groups are not collapsed to a picker that works.
// It also checks the dashboard, the project / selection steps and the row observation drawer
// for sideways page scroll, and saves review screenshots to reference/screenshots/.
//
// Needs Playwright's Chromium (npx playwright install chromium). Starts the real server on
// a free port; nothing is saved to the database.
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const path = require('path');
const { spawn } = require('child_process');
const { chromium } = require('@playwright/test');

const ROOT = path.join(__dirname, '..');
const SHOTS = path.join(ROOT, 'reference', 'screenshots');
const WIDTHS = [1920, 1440, 1280, 1024, 768, 390];
const HEIGHT = 900;
const SHEET_IDS = [
  'soilResistivity',
  'electrodeResistance',
  'continuityTest',
  'loopImpedanceTest',
  'prospectiveFaultCurrent',
  'riserIntegrityTest',
  'groundResistance',
  'stepTouchVoltage',
  'towerFootingResistance'
];
const SHOT_SHEETS = { electrodeResistance: 'MS-02', loopImpedanceTest: 'MS-04', stepTouchVoltage: 'MS-08' };
const SHOT_WIDTHS = [1440, 1024, 390];

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
  const location = state.draft.soilResistivity.locations[0];
  const preset = soilPresetFor(state.draft.project.soilSpacingPreset);
  preset.spacings.forEach((spacing) => {
    ['direction1', 'direction2', 'direction3', 'direction4'].forEach((direction) => {
      applyFieldInput({ dataset: { soilCell: 'true', locationIndex: '0', direction, spacing, field: 'resistivity' }, type: 'text', value: '123.45' });
    });
  });
  return location.name || '';
};

// Runs inside the page: how far does anything stick out?
const MEASURE = () => {
  const doc = document.documentElement;
  const tables = [...document.querySelectorAll('.builder-main table.data-table')];
  const report = {
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
        wrapLeft: wrapRect.left,
        tableOverflow: table.scrollWidth - wrap.clientWidth,
        contentRight: widestCell,
        className: table.className
      };
    })
  };
  return report;
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
  measure.tables.forEach((table, index) => {
    if (width >= 1024) {
      expect(!table.wrapScrolls && table.tableOverflow <= 1, `${label} @${width}: table ${index + 1} overflows its container by ${Math.round(table.tableOverflow)} px`);
      expect(table.tableRight <= table.wrapRight + 1, `${label} @${width}: table ${index + 1} sticks out ${Math.round(table.tableRight - table.wrapRight)} px`);
      expect(table.contentRight <= table.wrapRight + 1, `${label} @${width}: cell content sticks out ${Math.round(table.contentRight - table.wrapRight)} px (table ${index + 1})`);
    }
  });
  return measure;
}

async function checkHeaderSticks(page, width, label) {
  const result = await page.evaluate(() => {
    const table = document.querySelector('.builder-main table.sheet-table');
    const head = table.querySelector('thead th');
    const rect = table.getBoundingClientRect();
    window.scrollTo({ top: window.scrollY + rect.top + 140, behavior: 'instant' });
    return new Promise((resolve) =>
      requestAnimationFrame(() => {
        const top = head.getBoundingClientRect().top;
        const tableTop = table.getBoundingClientRect().top;
        resolve({ headTop: top, tableTop, scrolled: window.scrollY });
      })
    );
  });
  expect(result.tableTop < -20 && result.headTop >= -1 && result.headTop <= 40, `${label} @${width}: header row does not stay on screen (table top ${Math.round(result.tableTop)}, header top ${Math.round(result.headTop)})`);
  await page.evaluate(() => window.scrollTo({ top: 0, behavior: 'instant' }));
}

async function checkStickyColumns(page, width, label) {
  const result = await page.evaluate(() => {
    const wrap = document.querySelector('.builder-main .sheet-table-wrap');
    const table = wrap.querySelector('table');
    const firstCell = table.querySelector('tbody tr > :nth-child(1)');
    const secondCell = table.querySelector('tbody tr > :nth-child(2)');
    const scrollable = wrap.scrollWidth > wrap.clientWidth + 1;
    const wrapLeft = wrap.getBoundingClientRect().left;
    const before = [firstCell, secondCell].map((cell) => cell.getBoundingClientRect().left - wrapLeft);
    wrap.scrollLeft = 300;
    const after = [firstCell, secondCell].map((cell) => cell.getBoundingClientRect().left - wrapLeft);
    return {
      scrollable,
      moved: wrap.scrollLeft,
      drift: Math.max(Math.abs(before[0] - after[0]), Math.abs(before[1] - after[1])),
      firstSticky: getComputedStyle(firstCell).position,
      secondSticky: getComputedStyle(secondCell).position
    };
  });
  expect(result.scrollable, `${label} @${width}: expected sideways scroll inside the table at this width`);
  expect(result.firstSticky === 'sticky' && result.secondSticky === 'sticky', `${label} @${width}: Sr and TAG columns are not sticky`);
  expect(result.moved > 0 && result.drift <= 1, `${label} @${width}: Sr/TAG columns moved ${result.drift.toFixed(1)} px while the table scrolled ${result.moved} px`);
}

async function checkPicker(page, width) {
  await showStep(page, 'electrodeResistance');
  const state = await page.evaluate(() => {
    const cell = document.querySelector('.builder-main [data-tick-cell*="materialType"]');
    const summary = cell.querySelector('.tick-summary');
    const group = cell.querySelector('.tick-group');
    return { summaryShown: getComputedStyle(summary).display !== 'none', groupShown: getComputedStyle(group).display !== 'none', text: summary.textContent.trim() };
  });
  expect(state.summaryShown && !state.groupShown, `MS-02 @${width}: tick groups are not collapsed to a summary (summary ${state.summaryShown}, group ${state.groupShown})`);
  await page.click('.builder-main [data-tick-cell*="materialType"] .tick-summary');
  const open = await page.evaluate(() => {
    const group = document.querySelector('.builder-main [data-tick-cell*="materialType"] .tick-group');
    const rect = group.getBoundingClientRect();
    return { open: group.matches(':popover-open'), inView: rect.left >= 0 && rect.right <= window.innerWidth && rect.top >= 0 && rect.bottom <= window.innerHeight, options: group.querySelectorAll('input').length };
  });
  expect(open.open && open.inView && open.options === 6, `MS-02 @${width}: material picker did not open on tap (open ${open.open}, in view ${open.inView}, options ${open.options})`);
  await page.click('.builder-main [data-tick-cell*="materialType"] .tick-group .tick:has(input[value="Cu"])');
  await page.waitForTimeout(150);
  const after = await page.evaluate(() => {
    const cell = document.querySelector('.builder-main [data-tick-cell*="materialType"]');
    return { text: cell.querySelector('.tick-summary').textContent.trim(), open: cell.querySelector('.tick-group').matches(':popover-open'), value: state.draft.electrodeResistance[0].materialType };
  });
  expect(after.value === 'Cu' && after.text.startsWith('Cu') && !after.open, `MS-02 @${width}: picking a material did not update and close (value ${after.value}, text "${after.text}", open ${after.open})`);

  // Condition codes stay open so several can be ticked.
  await page.click('.builder-main [data-tick-cell*="conditionCodes"] .tick-summary');
  await page.click('.builder-main [data-tick-cell*="conditionCodes"] .tick-group .tick:has(input[data-code="C4"])');
  await page.waitForTimeout(150);
  const codes = await page.evaluate(() => {
    const cell = document.querySelector('.builder-main [data-tick-cell*="conditionCodes"]');
    return { open: cell.querySelector('.tick-group').matches(':popover-open'), codes: state.draft.electrodeResistance[0].conditionCodes };
  });
  expect(codes.open && !codes.codes.includes('C4'), `MS-02 @${width}: condition-code picker should stay open and toggle C4 (open ${codes.open}, codes ${codes.codes})`);
  await page.keyboard.press('Escape');
  const closed = await page.evaluate(() => !document.querySelector('.builder-main [data-tick-cell*="conditionCodes"] .tick-group').matches(':popover-open'));
  expect(closed, `MS-02 @${width}: Escape did not close the picker`);
}

async function checkOtherScreens(browser, base, width) {
  // A fresh browser context: the builder autosaves a draft and would reopen on reload.
  const context = await browser.newContext({ viewport: { width, height: HEIGHT } });
  const page = await context.newPage();
  await page.goto(base, { waitUntil: 'networkidle' });
  let measure = await page.evaluate(MEASURE);
  expect(measure.pageScrollWidth <= measure.innerWidth, `dashboard @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  await page.click('[data-action="new-report"]');
  for (const [index, name] of [[0, 'project step'], [1, 'selection step']]) {
    await page.evaluate((i) => {
      state.stepIndex = i;
      render();
    }, index);
    measure = await page.evaluate(MEASURE);
    expect(measure.pageScrollWidth <= measure.innerWidth, `${name} @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  }
  await page.evaluate(FILL_DRAFT);
  await showStep(page, 'electrodeResistance');
  await page.evaluate(() => openObservationEditor('electrodeResistance', 0));
  await page.waitForTimeout(200);
  measure = await page.evaluate(MEASURE);
  const drawer = await page.evaluate(() => {
    const node = document.querySelector('.observation-drawer, .drawer, [class*="observation"][class*="panel"], .modal');
    if (!node) return null;
    const rect = node.getBoundingClientRect();
    return { left: rect.left, right: rect.right, width: rect.width };
  });
  expect(measure.pageScrollWidth <= measure.innerWidth, `observation drawer @${width}: page scrolls sideways (${measure.pageScrollWidth})`);
  if (drawer) {
    expect(drawer.left >= -1 && drawer.right <= width + 1, `observation drawer @${width}: sticks out of the window (${Math.round(drawer.left)}..${Math.round(drawer.right)})`);
  }
  await page.evaluate(() => closeObservationEditor());
  await context.close();
}

async function shoot(page, width, sheetId, code) {
  await showStep(page, sheetId);
  await page.evaluate(() => window.scrollTo({ top: 0, behavior: 'instant' }));
  const target = path.join(SHOTS, `${code}_${width}.png`);
  await page.screenshot({ path: target, fullPage: true });
  return target;
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
        if (width >= 1280 && sheetId === 'electrodeResistance') {
          await checkHeaderSticks(page, width, code);
        }
        if (width >= 721 && width <= 1023 && sheetId === 'electrodeResistance') {
          await checkStickyColumns(page, width, code);
        }
      }
      if (width === 1024) {
        await checkHeaderSticks(page, width, 'MS-02');
      }
      if (width >= 1024 && width <= 1279) {
        await checkPicker(page, width);
      } else if (width >= 1280) {
        // Wide screens show the ticks themselves.
        await showStep(page, 'electrodeResistance');
        const visible = await page.evaluate(() => {
          const cell = document.querySelector('.builder-main [data-tick-cell*="materialType"]');
          return getComputedStyle(cell.querySelector('.tick-group')).display !== 'none' && getComputedStyle(cell.querySelector('.tick-summary')).display === 'none';
        });
        expect(visible, `MS-02 @${width}: ticks should be shown, not collapsed`);
      }

      for (const sheetId of Object.keys(SHOT_SHEETS)) {
        if (SHOT_WIDTHS.includes(width)) {
          await shoot(page, width, sheetId, SHOT_SHEETS[sheetId]);
        }
      }
      await checkOtherScreens(browser, url, width);
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
  assert.ok(checks > 0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
