// Golden test: the sample data and expected results from
// reference/generators/make_report.py (ETPL ESHA standard v0.4, ET-D03-REG-002).
// If the app's calculation rules drift from the approved reference, this fails.
const assert = require('assert');
const {
  normalizeReportInput,
  calculateSoilSummary,
  buildAssessmentContext,
  deriveElectrodeAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  summarizePanels,
  buildTestSummaries
} = require('../src/reportModel');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const { buildPayload } = require('./fixtures/golden-esha');

const report = normalizeReportInput(buildPayload());
const context = buildAssessmentContext(report);
const pits = report.electrodeResistance.map((row) => ({ tag: row.tag, ...deriveElectrodeAssessment(row, report) }));
const byTag = (tag) => pits.find((pit) => pit.tag === tag);
const pct = (summary) => Math.round(summary.passRate * 1000) / 10;
const summaries = buildTestSummaries(report);

check('site rho 25.68 ohm-m, Table 4 corrosivity, measured not assumed', () => {
  assert.strictEqual(calculateSoilSummary(report).overallAverage, 25.68);
  assert.strictEqual(context.rho, 25.68);
  assert.strictEqual(context.rhoAssumed, false);
  assert.strictEqual(calculateSoilSummary(report).category.label, 'Moderately corrosive');
  assert.strictEqual(context.profile.profileId, 'ETPL-DEFAULT');
});

check('rod limit 6.83 ohm (Cl. 14.2.2) and strip limit 4.21 ohm (Cl. 14.2.3)', () => {
  assert.strictEqual(byTag('EP-01').limit, 6.83);
  assert.strictEqual(byTag('EP-01').limitCriterionId, 'PIT-ROD');
  assert.strictEqual(byTag('EP-10').limit, 4.21);
  assert.strictEqual(byTag('EP-10').limitCriterionId, 'PIT-STRIP');
});

check('EP-03 / EP-07 / EP-08 are P1; EP-05 / EP-09 are P2; the rest P3', () => {
  const priorities = Object.fromEntries(pits.map((pit) => [pit.tag, pit.priority]));
  assert.deepStrictEqual(priorities, {
    'EP-01': 'P3', 'EP-02': 'P3', 'EP-03': 'P1', 'EP-04': 'P3', 'EP-05': 'P2',
    'EP-06': 'P3', 'EP-07': 'P1', 'EP-08': 'P1', 'EP-09': 'P2', 'EP-10': 'P3'
  });
});

check('electrode statuses: EP-03 healthy but P1 by code, EP-07 over, EP-08 not measurable, EP-09 warning', () => {
  assert.strictEqual(byTag('EP-03').status.label, 'Healthy');
  assert.strictEqual(byTag('EP-07').status.label, '>Permissible Limit');
  assert.strictEqual(byTag('EP-08').status.label, 'Not Measurable');
  assert.strictEqual(byTag('EP-09').status.label, 'Warning');
});

check('electrode pass 88.9% => >Attention Required; warning counts as within limit; not measurable excluded', () => {
  const electrode = summaries.electrodeResistance;
  assert.strictEqual(pct(electrode), 88.9);
  assert.strictEqual(electrode.label, '>Attention Required');
  assert.strictEqual(electrode.total, 9);
  assert.strictEqual(electrode.healthy, 8);
  assert.strictEqual(electrode.unhealthy, 1);
  assert.strictEqual(electrode.notAssessed, 1);
});

check('loop 75%: panel healthy only if all 3 phases <= 1.8 ohm', () => {
  assert.strictEqual(pct(summaries.loopImpedanceTest), 75);
  const panels = summarizePanels(report.loopImpedanceTest, deriveLoopImpedanceAssessment, context);
  assert.deepStrictEqual(panels.map((panel) => panel.status.label), ['Healthy', '>Permissible Limit', 'Healthy', 'Healthy']);
});

check('riser 83.3%: R-202 (R1) towards equipment 1.42 ohm is over 1 ohm', () => {
  assert.strictEqual(pct(summaries.riserIntegrityTest), 83.3);
  assert.strictEqual(deriveRiserAssessment(report.riserIntegrityTest[2], context).status.label, '>Permissible Limit');
});

check('PFC: APDB-2 Not Assessable and excluded from pass %; the others healthy', () => {
  const panels = summarizePanels(report.prospectiveFaultCurrent, deriveFaultCurrentAssessment, context);
  assert.deepStrictEqual(panels.map((panel) => panel.status.label), ['Healthy', 'Healthy', 'Healthy', 'Not Assessable']);
  assert.strictEqual(summaries.prospectiveFaultCurrent.total, 3);
  assert.strictEqual(summaries.prospectiveFaultCurrent.notAssessed, 1);
  assert.strictEqual(pct(summaries.prospectiveFaultCurrent), 100);
});

console.log(`\n${passed} golden checks passed.`);
