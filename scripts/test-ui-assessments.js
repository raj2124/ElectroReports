// Browser smoke test: runs public/sheets.js + public/app.js in a stubbed DOM against the
// real server modules (definitions, assessments) and checks that
//   - the entry screens are generated from the sheet definitions, in printed order,
//   - every status, limit, priority, flag and problem shown comes from the server,
//   - the old client-side rules are gone,
//   - inputs (text, tick boxes, condition codes, group cells, soil matrix) update the draft.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { assessReport } = require('../src/assessmentService');
const { normalizeReportInput } = require('../src/reportModel');
const { publicDefinitions } = require('../src/sheetDefinitions');
const { listProfiles } = require('../src/acceptance');

const read = (file) => fs.readFileSync(path.join(__dirname, '..', 'public', file), 'utf8');
const appSource = read('app.js');
const sheetsSource = read('sheets.js');
const printSource = read('print-sheets.js');
const reviewSource = read('review.js');
const workspaceSource = read('workspace.js');
let passed = 0;
function check(name, fn) {
  return Promise.resolve(fn()).then(() => {
    passed += 1;
    console.log(`  ok - ${name}`);
  });
}

// A DOM stub that tolerates anything the scripts touch.
function tolerant(target = {}) {
  return new Proxy(target, {
    get(obj, key) {
      if (key in obj) return obj[key];
      if (key === 'classList') return { add() {}, remove() {}, toggle() {}, contains: () => false };
      if (key === 'style' || key === 'dataset') {
        obj[key] = tolerant();
        return obj[key];
      }
      if (key === 'querySelector') return () => null;
      if (key === 'querySelectorAll') return () => [];
      if (key === 'getBoundingClientRect') return () => ({ left: 0, top: 0, width: 0, height: 0 });
      return typeof key === 'string' ? () => undefined : undefined;
    },
    set(obj, key, value) {
      obj[key] = value;
      return true;
    }
  });
}

async function createBrowser() {
  const calls = [];
  const appEl = tolerant({ innerHTML: '' });
  const storage = new Map();
  const timers = [];
  const window = tolerant({
    setTimeout: (fn) => {
      timers.push(fn);
      return timers.length;
    },
    clearTimeout() {},
    addEventListener() {},
    confirm: () => true,
    matchMedia: () => ({ matches: false, addEventListener() {} }),
    innerWidth: 390,
    innerHeight: 800,
    localStorage: { getItem: (key) => (storage.has(key) ? storage.get(key) : null), setItem: (key, value) => storage.set(key, String(value)), removeItem: (key) => storage.delete(key) }
  });
  const document = tolerant({ getElementById: (id) => (id === 'app' ? appEl : null), createElement: () => tolerant({ click() {}, remove() {} }), addEventListener() {}, activeElement: null, body: tolerant(), documentElement: tolerant(), hidden: false });
  const fetchStub = async (url, options = {}) => {
    calls.push({ url, method: options.method || 'GET', body: options.body });
    const reply = (data, status = 200) => ({ ok: status < 400, status, json: async () => data });
    if (url === '/api/sheets/print') return { ok: true, status: 200, headers: { get: (name) => (name === 'X-Sheet-Ids' ? 'ER-AAAA-AAAAA,ER-BBBB-BBBBB' : '') }, blob: async () => ({}) };
    if (url === '/api/sheets') return reply(JSON.parse(JSON.stringify(publicDefinitions())));
    if (url === '/api/acceptance/profiles') return reply(listProfiles());
    if (url === '/api/assessments') {
      try {
        return reply(JSON.parse(JSON.stringify(assessReport(JSON.parse(options.body)))));
      } catch (error) {
        return reply({ message: error.message }, 400);
      }
    }
    if (url === '/api/catalog') return reply({ tests: [] });
    if (url.startsWith('/api/reports') || url.startsWith('/api/zoho')) return reply([]);
    return reply({});
  };
  const context = vm.createContext({
    window, document, localStorage: window.localStorage, fetch: fetchStub, console, JSON, Math, Date, Intl, Promise, WeakMap, Map, Set, Array, Object, String, Number, RegExp, Error, Boolean,
    setTimeout: window.setTimeout, clearTimeout: () => {}, requestAnimationFrame: () => 0, FormData: class {}, FileReader: class {},
    Image: class {}, URL: { createObjectURL: () => '', revokeObjectURL: () => {} }, structuredClone: (value) => JSON.parse(JSON.stringify(value)), navigator: {}
  });
  vm.runInContext(sheetsSource, context, { filename: 'sheets.js' });
  vm.runInContext(printSource, context, { filename: 'print-sheets.js' });
  vm.runInContext(reviewSource, context, { filename: 'review.js' });
  vm.runInContext(workspaceSource, context, { filename: 'workspace.js' });
  vm.runInContext(appSource, context, { filename: 'app.js' });
  const run = (code) => vm.runInContext(code, context);
  const flush = async () => {
    for (let round = 0; round < 6; round += 1) {
      while (timers.length) {
        const fn = timers.shift();
        await fn();
      }
      await new Promise((resolve) => setImmediate(resolve));
    }
  };
  await flush(); // init(): loads the sheet definitions and builds the draft
  assert.ok(run('SHEET_DEFS !== null'), 'sheet definitions loaded');
  return { run, flush, calls, appEl, html: () => appEl.innerHTML };
}

// A fake input element for applyFieldInput.
const input = (dataset, extra = {}) => ({ dataset, type: 'text', value: '', checked: false, ...extra });

const project = { projectNo: 'P-1', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-13', engineerName: 'E', assumedSoilResistivity: '30', assumedSoilResistivitySource: 'GT-12' };
const saved = () => {
  const report = {
    ...normalizeReportInput({
      project,
      tests: { electrodeResistance: true, loopImpedanceTest: true, groundResistance: true, stepTouchVoltage: true, towerFootingResistance: true },
      sheetHeaders: { electrodeResistance: { instrument: 'mi3290', serialNo: '18480898', calibrationDue: '2026-09-01', currentElectrodeDistanceM: '30' }, stepTouchVoltage: { faultCurrentKA: '10', faultClearingTimeS: '0.5', bodyWeightBasis: 'both' } },
      electrodeResistance: [
        { tag: 'EP-1', electrodeType: 'Rod', materialType: 'Cu', length: '3', diameter: '40', resistanceWithoutGrid: '30', checkR50: '30', checkR70: '36', conditionCodes: ['C8'] },
        { tag: 'EP-2', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '3' }
      ],
      loopImpedanceTest: [{ location: 'P', loopImpedance: '0.5' }, { location: 'P', loopImpedance: '0.5' }, { location: 'P', loopImpedance: '2.5' }],
      groundResistance: [{ tag: 'EP-2', r618: '3' }],
      stepTouchVoltage: [{ location: 'Gate', measurementType: 'Touch-gate', testCurrentA: '5', v1: '0.05' }],
      towerFootingResistance: [{ mainLocationTower: 'T1', readings: [1, 2, 3, 4].map((n) => ({ footToEarthingConnectionStatus: n === 3 ? 'Not given' : 'Given', measuredCurrentMa: '10', measuredImpedance: '2' })) }]
    }),
    id: 'r1',
    createdAt: '2026-10-13'
  };
  report.assessment = JSON.parse(JSON.stringify(assessReport(report)));
  return report;
};

async function main() {
  await check('the scripts hold no copy of the old rules or their wording', () => {
    ['getSoilCategory', 'getElectrodeStatus', 'getContinuityStatus', 'getLoopStatus', 'getRiserStatus', 'STANDARD_GUIDANCE', 'deriveElectrodeAssessment', 'buildModuleInsights', 'countActionRows', 'calculateSoilSummary', 'summarizeTowerGroups', 'TEST_STANDARD_REFERENCES'].forEach((name) => {
      assert.ok(!appSource.includes(name) && !sheetsSource.includes(name), `${name} is still in the browser code`);
    });
    ['4.60', '4.6 ', 'Exceeds Permissible', 'Needs Attention', "label: 'Low'", "label: 'Medium'", "label: 'High'", '<= 0.5', '<= 1.5', 'Fixed Zsat 10'].forEach((text) => {
      assert.ok(!appSource.includes(text) && !sheetsSource.includes(text), `"${text}" is still in the browser code`);
    });
    assert.ok(!/Math\.(sqrt|log|PI)|\* 0\.9|\* 1\.2/.test(sheetsSource + appSource), 'no formulas or factors in the browser code');
  });

  await check('the electrode screen is generated from the definitions: printed labels, printed order, tick groups', async () => {
    const browser = await createBrowser();
    browser.run(`state.view = 'builder'; state.draft.tests.soilResistivity = false; state.stepIndex = 2; render();`);
    const html = browser.html();
    const printed = ['Pit TAG No.', 'Location / name of pit', 'Electrode type', 'Material', 'Length l (m)', 'Dia d / strip width (mm)', 'R without grid at 61.8% (Ω)', 'Check R at 50% (Ω)', 'Check R at 70% (Ω)', 'R with grid (Ω)', 'Method', 'Condition codes', 'Photo ref.', 'Remark'];
    let position = html.indexOf('<thead>');
    printed.forEach((label) => {
      // Header cells carry a fill-down button and sticky classes, so match the label inside <th ...>.
      const escaped = label.replace(/&/g, '&amp;').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const pattern = new RegExp(`<th[^>]*>(<span class="th-label">)?${escaped}(</span>|</th>)`, 'g');
      pattern.lastIndex = position;
      const match = pattern.exec(html);
      const next = match ? match.index : -1;
      assert.ok(next > position, `column "${label}" is missing or out of order`);
      position = next;
    });
    ['Plate', 'Pipe', 'Rod', 'Strip', 'Cu-bonded', 'GI', 'MS', 'CI', 'SS', '61.8%', 'Clamp', 'Slope'].forEach((option) => assert.ok(html.includes(`<span class="tick-label">${option}</span>`), `tick option ${option}`));
    ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8'].forEach((code) => assert.ok(html.includes(`data-code="${code}"`), `condition code ${code}`));
    assert.ok(/data-kind="codes"/.test(html) && /type="checkbox"/.test(html) && /type="radio"/.test(html));
    assert.ok(html.includes('Instrument used') && html.includes('Instrument Sr. No.') && html.includes('Calibration due'), 'printed instrument row');
    ['Current electrode distance C (m)', 'Auxiliary rods', '90°/180° check done'].forEach((label) => assert.ok(html.includes(label), `header cell ${label}`));
    assert.ok(html.includes('inputmode="decimal"'), 'numeric inputs use the decimal keypad');
    assert.ok(!html.includes('Plate length (m)'), 'plate size columns appear only for plates');
    assert.ok(html.includes('MS-02 Earth Electrode (Pit) Resistance'));
  });

  await check('inputs update the draft: text, radio, condition-code ticks, plate columns, header cells', async () => {
    const browser = await createBrowser();
    browser.run(`state.view = 'builder'; state.draft.tests.soilResistivity = false; state.stepIndex = 2;`);
    const apply = (dataset, extra) => browser.run(`applyFieldInput(${JSON.stringify({ dataset, ...extra })})`);
    const dataset = (field, more = {}) => ({ section: 'electrodeResistance', index: '0', field, ...more });
    assert.ok(apply(dataset('tag'), { type: 'text', value: 'EP-9' }));
    assert.ok(apply(dataset('electrodeType'), { type: 'radio', value: 'Plate', checked: true }));
    assert.ok(apply(dataset('conditionCodes', { kind: 'codes', code: 'C8' }), { type: 'checkbox', checked: true }));
    assert.ok(apply(dataset('conditionCodes', { kind: 'codes', code: 'C1' }), { type: 'checkbox', checked: true }));
    assert.ok(apply(dataset('conditionCodes', { kind: 'codes', code: 'C8' }), { type: 'checkbox', checked: false }));
    assert.ok(apply({ headerSheet: 'electrodeResistance', field: 'instrument' }, { type: 'radio', value: 'mi3290', checked: true }));
    assert.ok(apply({ headerSheet: 'electrodeResistance', field: 'auxRodsOk' }, { type: 'checkbox', checked: true }));
    const row = browser.run('state.draft.electrodeResistance[0]');
    assert.strictEqual(row.tag, 'EP-9');
    assert.strictEqual(row.electrodeType, 'Plate');
    assert.deepStrictEqual(JSON.parse(JSON.stringify(row.conditionCodes)), ['C1'], 'codes toggle on and off');
    const header = browser.run('state.draft.sheetHeaders.electrodeResistance');
    assert.strictEqual(header.instrument, 'mi3290');
    assert.strictEqual(header.auxRodsOk, true);
    browser.run('render();');
    assert.ok(browser.html().includes('Plate length (m)') && browser.html().includes('Plate width (m)'), 'plate columns appear once a plate is selected');
    assert.strictEqual(browser.run("applyFieldInput({ dataset: { section: 'nope', index: '0', field: 'x' }, type: 'text', value: '' })"), false);
  });

  await check('loop / PFC groups share merged cells across the three phases; tower feet take codes and connection', async () => {
    const browser = await createBrowser();
    browser.run(`state.view = 'builder'; state.draft.tests.loopImpedanceTest = true; state.draft.tests.towerFootingResistance = true;`);
    const apply = (dataset, extra) => browser.run(`applyFieldInput(${JSON.stringify({ dataset, ...extra })})`);
    apply({ section: 'loopImpedanceTest', index: '1', field: 'feederTag', groupSync: 'true' }, { type: 'text', value: 'O/G to AC' });
    apply({ section: 'loopImpedanceTest', index: '0', field: 'tripCurve', groupSync: 'true' }, { type: 'radio', value: 'C', checked: true });
    apply({ section: 'loopImpedanceTest', index: '2', field: 'loopImpedance' }, { type: 'text', value: '0.9' });
    assert.deepStrictEqual(browser.run('state.draft.loopImpedanceTest.map((row) => row.feederTag)').slice(), ['O/G to AC', 'O/G to AC', 'O/G to AC']);
    assert.deepStrictEqual(browser.run('state.draft.loopImpedanceTest.map((row) => row.tripCurve)').slice(), ['C', 'C', 'C']);
    assert.deepStrictEqual(browser.run('state.draft.loopImpedanceTest.map((row) => row.loopImpedance)').slice(), ['', '', '0.9'], 'readings are per phase');
    browser.run("handleSheetAction('add-row', { dataset: { section: 'loopImpedanceTest' } })");
    assert.strictEqual(browser.run('state.draft.loopImpedanceTest.length'), 6);
    assert.deepStrictEqual(browser.run('state.draft.loopImpedanceTest.slice(3).map((row) => row.measuredPoints)').slice(), ['R-E', 'Y-E', 'B-E']);
    browser.run("handleSheetAction('remove-row', { dataset: { section: 'loopImpedanceTest', index: '4' } })");
    assert.strictEqual(browser.run('state.draft.loopImpedanceTest.length'), 3, 'removing a row removes its whole 3-phase group');
    apply({ section: 'towerFootingResistance', index: '0', readingIndex: '1', field: 'footToEarthingConnectionStatus' }, { type: 'radio', value: 'Not given', checked: true });
    apply({ section: 'towerFootingResistance', index: '0', readingIndex: '1', field: 'conditionCodes', kind: 'codes', code: 'C5' }, { type: 'checkbox', checked: true });
    apply({ section: 'towerFootingResistance', index: '0', field: 'mainLocationTower' }, { type: 'text', value: 'T-12' });
    const tower = browser.run('state.draft.towerFootingResistance[0]');
    assert.strictEqual(tower.mainLocationTower, 'T-12');
    assert.strictEqual(tower.readings[1].footToEarthingConnectionStatus, 'Not given');
    assert.deepStrictEqual(JSON.parse(JSON.stringify(tower.readings[1].conditionCodes)), ['C5']);
    assert.strictEqual(tower.readings.length, 4);
  });

  await check('the soil matrix creates readings per spacing and direction; the substation preset gives 8 directions', async () => {
    const browser = await createBrowser();
    browser.run("state.view = 'builder'; state.stepIndex = 2;");
    browser.run('render();');
    let html = browser.html();
    assert.ok(html.includes('MS-01 Soil Resistivity') && html.includes('Probe spacing a (m)') && html.includes('Direction 4 reading'));
    assert.ok(!html.includes('Direction 5 reading'), 'ETPL default shows 4 directions');
    assert.ok(html.includes('Reading unit') && html.includes('Probe depth (cm)') && html.includes('Direction 1 / 2 bearing') && html.includes('Weather'));
    const cell = (direction, spacing, field, value) => browser.run(`applyFieldInput(${JSON.stringify({ dataset: { soilCell: 'true', locationIndex: '0', direction, spacing, field }, type: 'text', value })})`);
    cell('direction1', '1.0', 'resistivity', '4.6');
    cell('direction2', '1', 'resistivity', '4.9');
    cell('direction1', '1.00', 'remark', 'dry');
    assert.strictEqual(browser.run("state.draft.soilResistivity.locations[0].direction1.find((row) => row.spacing === '1.0').resistivity"), '4.6');
    assert.strictEqual(browser.run("state.draft.soilResistivity.locations[0].direction1.find((row) => row.spacing === '1.0').remark"), 'dry', 'spacings match numerically');
    assert.ok(browser.run("applyFieldInput({ dataset: { soilLocation: '0', field: 'inputUnit' }, type: 'radio', value: 'ohm', checked: true })"));
    assert.ok(browser.run("applyFieldInput({ dataset: { soilLocation: '0', bearing: 'direction3' }, type: 'text', value: '45' })"));
    assert.strictEqual(browser.run('state.draft.soilResistivity.locations[0].inputUnit'), 'ohm');
    assert.strictEqual(browser.run('state.draft.soilResistivity.locations[0].bearings.direction3'), '45');
    // substation preset: 2, 5, 10, 15, 25, 50 m and 8 directions (two pages of 4)
    browser.run("state.draft.soilResistivity.locations = [newSoilLocation('Switchyard', 'is3043-substation')]; state.draft.project.soilSpacingPreset = 'is3043-substation'; render();");
    html = browser.html();
    assert.ok(html.includes('Directions 1–4') && html.includes('Directions 5–8'), 'two pages of directions');
    ['2', '5', '10', '15', '25', '50'].forEach((spacing) => assert.ok(html.includes(`<span class="table-input-static">${spacing}</span>`), `spacing ${spacing}`));
    assert.strictEqual(browser.run("SOIL_DIRECTION_KEYS.filter((key) => state.draft.soilResistivity.locations[0][key].length).length"), 8);
    // custom spacing
    browser.run("handleSheetAction('add-soil-spacing', { dataset: { locationIndex: '0' } })");
    assert.strictEqual(browser.run("state.draft.soilResistivity.locations[0].direction1.length"), 6, 'no input value: nothing added');
  });

  await check('a saved report shows every sheet from the definitions with server statuses, flags, limits and the reviewer checklist', async () => {
    const browser = await createBrowser();
    const report = saved();
    browser.run(`state.activeReport = ${JSON.stringify(report)}; state.view = 'detail'; render();`);
    const html = browser.html();
    ['MS-02 Earth Electrode (Pit) Resistance', 'MS-04 Earth Fault Loop Impedance', 'MS-07 Ground (Grid) Resistance', 'MS-08 Step &amp; Touch Voltage', 'MS-09 Tower Footing Impedance'].forEach((title) => assert.ok(html.includes(title), `${title} section`));
    assert.ok(html.includes('&gt;Permissible Limit'), 'over-limit status from the server');
    assert.ok(html.includes('Re-test: 61.8% unconfirmed'), '50/70% flag chip');
    assert.ok(html.includes('Foot without earth connection'), 'tower defect chip');
    const limit = report.assessment.tests.electrodeResistance.rows[0].limit.toFixed(2);
    assert.ok(html.includes(`Limit ${limit}`), `limit ${limit} shown`);
    assert.ok(html.includes('Reviewer checklist') && html.includes('calibration expired on 2026-09-01'), 'calibration item on the reviewer checklist');
    assert.ok(html.includes('Blocks final approval until acknowledged'), 'blocking item marked');
    assert.ok(html.includes('The PDF does not include MS-07 and MS-08 yet'), 'visible note next to the export button');
    assert.ok(html.includes('18480898') && html.includes('Metrel MI 3290 GX'), 'instrument and serial from the sheet header');
    assert.ok(html.includes('Work date') && html.includes('Acceptance profile'));
    assert.ok(!html.includes('4.60'));
  });

  await check('a report without MS-07 / MS-08 shows no export note', async () => {
    const browser = await createBrowser();
    const report = saved();
    report.tests.groundResistance = false;
    report.tests.stepTouchVoltage = false;
    browser.run(`state.activeReport = ${JSON.stringify(report)}; state.view = 'detail'; render();`);
    assert.ok(!browser.html().includes('The PDF does not include'));
  });

  await check('the builder asks the server for the draft assessment and shows its problems and statuses', async () => {
    const browser = await createBrowser();
    browser.run(`
      state.view = 'builder';
      state.draft.project.assumedSoilResistivity = '30';
      state.draft.project.assumedSoilResistivitySource = 'GT-12';
      state.draft.tests.soilResistivity = false;
      state.draft.electrodeResistance = [Object.assign(newSheetRow('electrodeResistance'), { tag: 'EP-1', length: '3', diameter: 'forty', resistanceWithoutGrid: '3' })];
      state.stepIndex = 2;
      render();
    `);
    assert.ok(browser.html().includes('Pending'), 'Pending until the server answers');
    await browser.flush();
    assert.ok(browser.calls.some((call) => call.url === '/api/assessments' && call.method === 'POST'));
    let html = browser.html();
    assert.ok(html.includes('assumed: GT-12'), 'banner states the assumed rho and its source');
    assert.ok(html.includes('Fix before saving') && html.includes('Dia d / strip width (mm) must be a number'), 'validation error from the server shown on the sheet');
    browser.run("state.draft.electrodeResistance[0].diameter = '40'; render();");
    await browser.flush();
    html = browser.html();
    assert.ok(!html.includes('Fix before saving'), 'error gone once fixed');
    assert.ok(html.includes('Healthy') && html.includes('>P3<'), 'status and priority from the server');
    const before = browser.calls.length;
    browser.run('render();');
    await browser.flush();
    assert.strictEqual(browser.calls.length, before, 'an unchanged draft is not re-requested');
  });

  await check('the project step carries the job-level fields: profile from the register, spacing set, ground condition, assumed rho', async () => {
    const browser = await createBrowser();
    browser.run("state.view = 'builder'; state.stepIndex = 0; state.draft.tests.soilResistivity = false; render();");
    await browser.flush();
    const html = browser.html();
    ['Work date', 'Client representative', 'Acceptance profile', 'Soil probe-spacing set', 'Ground condition', 'Assumed soil resistivity (Ω·m)', 'Source of the assumption'].forEach((label) => assert.ok(html.includes(label), label));
    assert.ok(html.includes('ETPL-DEFAULT') && !html.includes('CLIENT-TEMPLATE'), 'profiles come from the register; the template is not offered');
    ['Dry', 'Damp', 'Wet'].forEach((condition) => assert.ok(html.includes(`<span class="tick-label">${condition}</span>`), condition));
    assert.ok(html.includes('ETPL default') && html.includes('IS 3043 substation'));
  });

  await check('the selection step shows sheet codes, offers the new tests and hides the retired one', async () => {
    const browser = await createBrowser();
    browser.run(`state.catalog = ${JSON.stringify(require('../src/reportModel').TEST_LIBRARY)}; state.view = 'builder'; state.stepIndex = 1; render();`);
    const html = browser.html();
    ['MS-07', 'MS-08', 'MS-09', 'Ground (Grid) Resistance Test', 'Step &amp; Touch Voltage Test'].forEach((text) => assert.ok(html.includes(text), text));
    assert.ok(!html.includes('data-test="earthContinuityTest"'), 'retired test is not offered for new reports');
    assert.ok(html.includes('data-test="groundResistance"') && html.includes('data-test="stepTouchVoltage"'));
    assert.ok(/aria-pressed="(true|false)"/.test(html));
  });

  await check('every sheet renders in the builder and on a saved report, showing all its printed columns', async () => {
    const { SHEETS } = require('../src/sheetDefinitions');
    const browser = await createBrowser();
    const labelOf = (label) => label.replace(/&/g, '&amp;');
    for (const sheet of SHEETS.filter((entry) => !entry.retired)) {
      browser.run(`state.view = 'builder'; Object.keys(state.draft.tests).forEach((id) => { state.draft.tests[id] = id === ${JSON.stringify(sheet.id)}; }); state.stepIndex = 2; render();`);
      await browser.flush();
      const html = browser.html();
      assert.ok(html.includes(`${sheet.code} ${labelOf(sheet.title)}`), `${sheet.code}: title`);
      const printed = sheet.rowFields.filter((field) => !field.extra && !field.legacy && field.kind !== 'serial').map((field) => field.label);
      printed.forEach((label) => assert.ok(html.includes(`>${labelOf(label)}<`) || html.includes(`${labelOf(label)}`), `${sheet.code}: column "${label}"`));
      sheet.headerFields.filter((field) => !field.extra).forEach((field) => assert.ok(html.includes(labelOf(field.label)), `${sheet.code}: header cell "${field.label}"`));
      assert.ok(html.includes('How to record this sheet') || sheet.id === 'soilResistivity' || !sheet.notes.length, `${sheet.code}: notes`);
    }
    const report = saved();
    report.tests = Object.fromEntries(SHEETS.map((sheet) => [sheet.id, false]));
    for (const sheet of SHEETS.filter((entry) => !entry.retired && entry.id !== 'soilResistivity')) {
      browser.run(`state.activeReport = ${JSON.stringify({ ...report, tests: { [sheet.id]: true } })}; state.view = 'detail'; render();`);
      const html = browser.html();
      assert.ok(html.includes(`${sheet.code} ${labelOf(sheet.title)}`), `${sheet.code}: detail title`);
    }
  });

  await check('tick boxes are 44 px tap targets and tables turn into cards on a phone', () => {
    const css = read('styles.css');
    const tick = /\.tick \{[^}]*\}/.exec(css)[0];
    assert.ok(/min-height:\s*44px/.test(tick) && /min-width:\s*44px/.test(tick), 'ticks are at least 44 x 44 px');
    assert.ok(/\.tick input \{[^}]*inset:\s*0/.test(css), 'the whole tick label is the hit area');
    assert.ok(/@media \(max-width: 720px\)[\s\S]*attr\(data-label\)/.test(css), 'mobile card layout labels cells by column');
    assert.ok(/inputmode/.test(sheetsSource) && /autocomplete="off"/.test(sheetsSource));
    assert.ok(read('index.html').includes('sheets.js') && read('index.html').indexOf('sheets.js') < read('index.html').indexOf('app.js'), 'sheets.js loads before app.js');
  });

  await check('a draft saved before these fields existed restores with the new fields filled in', async () => {
    const browser = await createBrowser();
    const old = {
      project: { projectNo: 'OLD', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-01-01', engineerName: 'E' },
      tests: { soilResistivity: true, electrodeResistance: true, earthContinuityTest: false },
      soilResistivity: { locations: [{ name: 'Yard', direction1: [{ spacing: '0.5', resistivity: '9.8' }], direction2: [{ spacing: '0.5', resistivity: '10.4' }] }] },
      electrodeResistance: [{ tag: 'E1', electrodeType: 'Rod', materialType: 'Copper Bonded', length: '3', diameter: '40', resistanceWithGrid: '2' }]
    };
    const restored = browser.run(`restoreDraftShape(${JSON.stringify(old)})`);
    const location = restored.soilResistivity.locations[0];
    assert.deepStrictEqual(Object.keys(location.bearings), ['direction1', 'direction2', 'direction3', 'direction4', 'direction5', 'direction6', 'direction7', 'direction8']);
    assert.strictEqual(location.direction3.length, 0);
    assert.strictEqual(location.direction1[0].resistivity, '9.8');
    assert.strictEqual(restored.electrodeResistance[0].materialType, 'Cu-bonded', 'legacy material maps to the printed option');
    assert.deepStrictEqual(JSON.parse(JSON.stringify(restored.electrodeResistance[0].conditionCodes)), []);
    assert.strictEqual(restored.project.soilSpacingPreset, 'etpl-default');
    assert.ok(Array.isArray(restored.groundResistance) && Array.isArray(restored.stepTouchVoltage));
  });

  await check('"Print sheets for site" is on the project step and every sheet screen, and sends the project, date and tests', async () => {
    const html = read('index.html');
    assert.ok(html.indexOf('print-sheets.js') > html.indexOf('sheets.js') && html.indexOf('print-sheets.js') < html.indexOf('app.js'), 'print-sheets.js loads between sheets.js and app.js');
    const browser = await createBrowser();
    browser.run(`state.draft.project = { ...state.draft.project, projectNo: 'P-9', clientName: 'Client', siteLocation: 'Site', workOrder: 'WO', reportDate: '2026-10-14', engineerName: 'Eng', zohoProjectId: '' }; state.draft.tests.electrodeResistance = true; state.draft.tests.loopImpedanceTest = true; state.view = 'builder'; state.stepIndex = 0; render();`);
    assert.ok(browser.html().includes('data-action="open-print-sheets"'), 'project step');
    const sheetIds = browser.run('allSheets().filter((sheet) => !sheet.retired).map((sheet) => sheet.id)');
    assert.strictEqual(sheetIds.length, 9);
    sheetIds.forEach((id) => assert.ok(browser.run(`renderSheetIntro(sheetById('${id}'))`).includes('Print sheets for site'), `${id} sheet screen`));
    assert.ok(!browser.run("renderSheetIntro(sheetById('earthContinuityTest'))").includes('Print sheets for site'), 'the retired sheet has no print button');
    browser.run("openPrintSheets('stepTouchVoltage')");
    await browser.flush();
    assert.deepStrictEqual(JSON.parse(browser.run('JSON.stringify(Object.keys(printState.selected).filter((id) => printState.selected[id]).sort())')), ['electrodeResistance', 'loopImpedanceTest', 'soilResistivity', 'stepTouchVoltage'], 'the report\'s selected tests (soil is on by default) plus the sheet it was opened from');
    assert.ok(browser.html().includes('print-dialog') && browser.html().includes('Download PDF'));
    assert.strictEqual(browser.run('printState.workDate'), '2026-10-14');
    assert.strictEqual(browser.run('printState.engineerName'), 'Eng');
    browser.run("handlePrintInput({ dataset: { printSheet: 'loopImpedanceTest' }, checked: false }); handlePrintInput({ dataset: { printField: 'engineerName' }, value: 'Typed' }); handlePrintInput({ dataset: { printField: 'workDate' }, value: '2026-10-20' });");
    await browser.run('downloadPrintSheets()');
    await browser.flush();
    const call = browser.calls.find((item) => item.url === '/api/sheets/print');
    assert.ok(call && call.method === 'POST', 'POSTs to /api/sheets/print');
    const body = JSON.parse(call.body);
    assert.deepStrictEqual(body.sheets.sort(), ['electrodeResistance', 'soilResistivity', 'stepTouchVoltage']);
    assert.strictEqual(body.workDate, '2026-10-20');
    assert.strictEqual(body.engineerName, 'Typed');
    assert.strictEqual(body.project.projectNo, 'P-9');
    assert.strictEqual(browser.run('printState.open'), false, 'the dialog closes after the download');
    // Nothing ticked: refused in the dialog, no request.
    browser.run("openPrintSheets(''); Object.keys(printState.selected).forEach((id) => { printState.selected[id] = false; });");
    const before = browser.calls.length;
    await browser.run('downloadPrintSheets()');
    assert.strictEqual(browser.calls.length, before);
    assert.ok(/at least one test/i.test(browser.run('printState.error')));
  });

  console.log(`\n${passed} UI checks passed.`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
