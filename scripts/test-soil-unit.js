// A soil location whose reading unit (R or rho) is not selected cannot be read: it is Pending
// ("reading unit not selected"), gives no rho and no corrosivity class, and stays out of the site
// average and out of the electrode limit calculation. The SELECTION_MISSING warning stays.
const assert = require('assert');
const { normalizeReportInput, buildAssessmentContext, deriveElectrodeAssessment, calculateSoilSummary } = require('../src/reportModel');
const { assessReport } = require('../src/assessmentService');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const project = { projectNo: 'P-1', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-14', engineerName: 'E', acceptanceProfile: 'ETPL-DEFAULT' };
const readings = (value) => [{ spacing: '1', resistivity: value }, { spacing: '2', resistivity: value }];
const location = (name, inputUnit, value) => ({ name, inputUnit, direction1: readings(value), direction2: readings(value) });
const pit = { tag: 'EP-1', electrodeType: 'Rod', materialType: 'Cu', length: '3', diameter: '40', resistanceWithoutGrid: '2' };
const report = (locations, extra = {}) => normalizeReportInput({ project, tests: { soilResistivity: true, electrodeResistance: true }, soilResistivity: { locations }, electrodeResistance: [pit], ...extra });

check('a blank unit is kept blank, not turned into rho', () => {
  const saved = report([location('A', '', '40')]);
  assert.strictEqual(saved.soilResistivity.locations[0].inputUnit, '');
});

check('a location with no unit is Pending: no rho, no averages, no corrosivity class', () => {
  const summary = calculateSoilSummary(report([location('A', '', '40')]));
  const a = summary.locations[0];
  assert.strictEqual(a.pending, true);
  assert.strictEqual(a.pendingReason, 'reading unit not selected');
  assert.strictEqual(a.overallAverage, null);
  assert.strictEqual(a.direction1Average, null);
  assert.deepStrictEqual(Object.values(a.directionAverages).filter((value) => value !== null), []);
  assert.strictEqual(a.category.label, 'Insufficient Data');
  assert.strictEqual(summary.overallAverage, null);
  assert.deepStrictEqual(summary.pendingLocations, ['A']);
});

check('it is left out of the site average; the other locations still count', () => {
  const summary = calculateSoilSummary(report([location('A', 'ohm-m', '40'), location('B', '', '5000')]));
  assert.strictEqual(summary.overallAverage, 40, 'site average is the unit-selected location only');
  assert.strictEqual(summary.locations[0].overallAverage, 40);
  assert.strictEqual(summary.locations[1].overallAverage, null);
  assert.deepStrictEqual(summary.pendingLocations, ['B']);
});

check('it is left out of the electrode limit: no unit means no rho and no calculated limit', () => {
  const pending = report([location('A', '', '40')]);
  const context = buildAssessmentContext(pending);
  assert.strictEqual(context.rho, null);
  const row = deriveElectrodeAssessment(pending.electrodeResistance[0], context);
  assert.strictEqual(row.limit, null, 'no limit from an unreadable soil location');
  const withBoth = buildAssessmentContext(report([location('A', 'ohm-m', '40'), location('B', '', '5000')]));
  assert.strictEqual(withBoth.rho, 40, 'the limit uses the unit-selected location only');
  const alone = buildAssessmentContext(report([location('A', 'ohm-m', '40')]));
  assert.strictEqual(withBoth.rho, alone.rho);
  assert.strictEqual(deriveElectrodeAssessment(pit, withBoth).limit, deriveElectrodeAssessment(pit, alone).limit);
});

check('the soil module says why it is pending, and the warning is kept', () => {
  const result = assessReport(report([location('A', '', '40')]));
  const soil = result.modules.find((module) => module.id === 'soilResistivity');
  assert.match(soil.detail, /reading unit not selected/);
  assert.strictEqual(soil.complete, false);
  assert.ok(result.validation.warnings.some((warning) => warning.code === 'SELECTION_MISSING' && warning.sheet === 'soilResistivity'), 'SELECTION_MISSING warning stays');
  const mixed = assessReport(report([location('A', 'ohm-m', '40'), location('B', '', '5000')]));
  assert.match(mixed.modules.find((module) => module.id === 'soilResistivity').detail, /B excluded \(reading unit not selected\)/);
});

check('the two real units still calculate as before', () => {
  const rho = calculateSoilSummary(report([location('A', 'ohm-m', '40')]));
  assert.strictEqual(rho.overallAverage, 40);
  const resistance = calculateSoilSummary(report([location('A', 'ohm', '10')]));
  assert.ok(Math.abs(resistance.locations[0].direction1Average - ((2 * Math.PI * 1 * 10 + 2 * Math.PI * 2 * 10) / 2)) < 0.01, 'rho = 2 pi a R for R readings');
  assert.strictEqual(resistance.locations[0].pending, false);
});

check('a report in the old single-location shape (no unit field) is still read as rho', () => {
  const legacy = calculateSoilSummary({ soilResistivity: { direction1: readings('90'), direction2: readings('90') }, project });
  assert.strictEqual(legacy.overallAverage, 90);
});

console.log(`\n${passed} soil-unit checks passed.`);
