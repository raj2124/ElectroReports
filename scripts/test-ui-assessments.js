// Browser smoke test: runs public/app.js in a stubbed DOM and checks that what the UI
// renders comes from the server's assessments (src/assessmentService.js), that the old
// client-side rules are gone, and that an in-progress draft is assessed by the server.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { assessReport } = require('../src/assessmentService');
const { normalizeReportInput } = require('../src/reportModel');

const source = fs.readFileSync(path.join(__dirname, '..', 'public', 'app.js'), 'utf8');
let passed = 0;
function check(name, fn) {
  return Promise.resolve(fn()).then(() => {
    passed += 1;
    console.log(`  ok - ${name}`);
  });
}

// A DOM stub that tolerates anything app.js touches.
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

function createBrowser() {
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
    matchMedia: () => ({ matches: false, addEventListener() {} }),
    innerWidth: 1280,
    innerHeight: 800,
    localStorage: { getItem: (key) => (storage.has(key) ? storage.get(key) : null), setItem: (key, value) => storage.set(key, String(value)), removeItem: (key) => storage.delete(key) }
  });
  const document = tolerant({ getElementById: () => appEl, addEventListener() {}, activeElement: null, body: tolerant(), documentElement: tolerant(), hidden: false });
  const fetchStub = async (url, options = {}) => {
    calls.push({ url, method: options.method || 'GET' });
    const reply = (data, status = 200) => ({ ok: status < 400, status, json: async () => data });
    if (url === '/api/assessments') {
      try {
        return reply(assessReport(JSON.parse(options.body)));
      } catch (error) {
        return reply({ message: error.message }, 400);
      }
    }
    if (url === '/api/catalog') return reply([]);
    return reply({});
  };
  const context = vm.createContext({
    window, document, localStorage: window.localStorage, fetch: fetchStub, console, JSON, Math, Date, Intl, Promise, WeakMap, Map, Set,
    setTimeout: window.setTimeout, clearTimeout: () => {}, requestAnimationFrame: () => 0, FormData: class {}, FileReader: class {},
    Image: class {}, URL: { createObjectURL: () => '' }, structuredClone: (value) => JSON.parse(JSON.stringify(value)), navigator: {}
  });
  vm.runInContext(source, context, { filename: 'app.js' });
  const run = (code) => vm.runInContext(code, context);
  const flush = async () => {
    while (timers.length) {
      const fn = timers.shift();
      await fn();
    }
    await new Promise((resolve) => setImmediate(resolve));
  };
  return { run, flush, calls, appEl };
}

const project = { projectNo: 'P-1', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-13', engineerName: 'E' };
const savedPayload = {
  project,
  tests: { soilResistivity: true, electrodeResistance: true, loopImpedanceTest: true },
  soilResistivity: { locations: [{ name: 'L1', direction1: [{ spacing: '1', resistivity: '40' }], direction2: [{ spacing: '1', resistivity: '40' }] }] },
  electrodeResistance: [
    { tag: 'EP-1', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '30', conditionCodes: ['C8'] },
    { tag: 'EP-2', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '3' }
  ],
  loopImpedanceTest: [
    { location: 'P', loopImpedance: '0.5' },
    { location: 'P', loopImpedance: '0.5' },
    { location: 'P', loopImpedance: '2.5' }
  ]
};

async function main() {
  await check('app.js holds no copy of the old rules or their wording', () => {
    ['getSoilCategory', 'getElectrodeStatus', 'getContinuityStatus', 'getLoopStatus', 'getRiserStatus', 'STANDARD_GUIDANCE', 'deriveElectrodeAssessment', 'buildModuleInsights', 'countActionRows', 'calculateSoilSummary', 'summarizeTowerGroups'].forEach((name) => {
      assert.ok(!source.includes(name), `${name} is still in app.js`);
    });
    ['4.60', '4.6 ', 'Exceeds Permissible', 'Needs Attention', "label: 'Low'", "label: 'Medium'", "label: 'High'", '<= 0.5', '<= 1.5', 'Fixed Zsat 10'].forEach((text) => {
      assert.ok(!source.includes(text), `"${text}" is still in app.js`);
    });
  });

  await check('a saved report renders the server assessment: status, priority, limit, soil class, no stale text', async () => {
    const report = { ...normalizeReportInput(savedPayload), id: 'r1', createdAt: '2026-10-13' };
    report.assessment = assessReport(report);
    const browser = createBrowser();
    browser.run(`state.activeReport = ${JSON.stringify(report)}; state.view = 'detail'; render();`);
    const html = browser.appEl.innerHTML;
    assert.ok(html.includes('Moderately corrosive'), 'soil class from server (40 ohm-m is 25-50)');
    assert.ok(html.includes('&gt;Permissible Limit'), 'over-limit status label from server');
    assert.ok(html.includes('>P1<'), 'priority from server (over limit => P1)');
    const limit = report.assessment.tests.electrodeResistance.rows[0].limit.toFixed(2);
    assert.ok(html.includes(`${limit} ohm`), `calculated limit ${limit} ohm shown`);
    assert.ok(!html.includes('4.60'), 'no hardcoded 4.60');
    assert.ok(!/>Low</.test(html), 'no Low/Medium/High soil class');
  });

  await check('the builder asks the server for the draft assessment and re-renders from the answer', async () => {
    const browser = createBrowser();
    browser.run(`
      state.view = 'builder';
      state.draft.project.assumedSoilResistivity = '30';
      state.draft.project.assumedSoilResistivitySource = 'GT-12';
      state.draft.tests.soilResistivity = false;
      state.draft.tests.electrodeResistance = true;
      state.draft.electrodeResistance = [Object.assign(createDefaultElectrodeRowForTest(), { tag: 'EP-1', length: '3', diameter: '40', resistanceWithoutGrid: '3' })];
      state.stepIndex = 2;
    `.replace('createDefaultElectrodeRowForTest()', '{ rowId: "e1", electrodeType: "Rod", materialType: "Copper", resistanceWithGrid: "", rowObservation: "", rowPhotos: [] }'));
    browser.run('render();');
    assert.ok(browser.appEl.innerHTML.includes('Pending'), 'Pending until the server answers');
    await browser.flush();
    assert.ok(browser.calls.some((call) => call.url === '/api/assessments' && call.method === 'POST'), 'POST /api/assessments was made');
    const html = browser.appEl.innerHTML;
    assert.ok(html.includes('assumed: GT-12'), 'banner from server states the assumed rho and its source');
    assert.ok(html.includes('Healthy'), 'status from server');
    assert.ok(html.includes('P3'), 'priority from server');
    const before = browser.calls.length;
    browser.run('render();');
    await browser.flush();
    assert.strictEqual(browser.calls.length, before, 'an unchanged draft is not re-requested');
  });

  await check('legacy generated remark text is replaced by the live remark; an engineer\'s text is kept', async () => {
    const legacy = 'Measured Zs is in the healthy range, but final compliance still depends on the circuit/device-specific maximum Zs under IEC 60364 / device-specific max Zs.';
    const report = { ...normalizeReportInput(savedPayload), id: 'r2', createdAt: '2026-10-13' };
    report.loopImpedanceTest[0].remarks = 'Feeder labelled differently on site.';
    report.loopImpedanceTest[2].remarks = legacy; // saved by an older version, contradicts the 2.5 ohm status
    const { scrubGeneratedRemarks } = require('../src/remarks');
    const scrubbed = scrubGeneratedRemarks(report);
    scrubbed.assessment = assessReport(scrubbed);
    const rows = scrubbed.assessment.tests.loopImpedanceTest.rows;
    assert.strictEqual(rows[0].remark, 'Feeder labelled differently on site.');
    assert.ok(!rows[2].remark.includes('healthy range') && rows[2].remark.includes('exceeds'), rows[2].remark);
  });

  console.log(`\n${passed} UI assessment checks passed.`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
