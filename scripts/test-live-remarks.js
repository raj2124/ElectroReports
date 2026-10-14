// Live remarks: generated text is never saved, and a saved sentence from an older
// version is recognised and replaced, so a remark can never contradict its status.
const assert = require('assert');
const { isGeneratedRemark, engineerRemark, scrubGeneratedRemarks } = require('../src/remarks');
const { resolveProfile } = require('../src/acceptance');
const {
  normalizeReportInput,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveEarthContinuityAssessment,
  deriveTowerFootingAssessment
} = require('../src/reportModel');
const { assessReport } = require('../src/assessmentService');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

// Every sentence the pre-6A code (commit 8f332ad) could save into a remark field,
// generated from that code across all of its status branches.
const LEGACY = require('./fixtures/legacy-remarks.json');
const LEGACY_ENTER_ZS = LEGACY[7];
const LEGACY_ZS_HEALTHY = LEGACY[8];
const LEGACY_RISER_WITHIN = LEGACY[16];
const LEGACY_EARTH_ENTER = LEGACY[19];

const project = { projectNo: 'P', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-13', engineerName: 'E' };

check('every sentence saved by the pre-6A code is recognised as generated', () => {
  assert.strictEqual(LEGACY.length, 23);
  LEGACY.forEach((text) => assert.ok(isGeneratedRemark(text), `not recognised: ${text}`));
});

check('every sentence the current rules generate is recognised as generated', () => {
  const context = { profile: resolveProfile(), rho: 25, rhoAssumed: false, rhoSource: '' };
  const texts = [];
  ['', '0.3', '1.4', '2.5'].forEach((value) => {
    const lead = value ? '0.1' : '';
    texts.push(deriveContinuityAssessment({ resistance: value, impedance: value ? '0.2' : '', testLeadResistance: lead }, context).comment);
    texts.push(deriveEarthContinuityAssessment({ measuredValue: value, testLeadResistance: lead }, context).comment);
    texts.push(deriveLoopImpedanceAssessment({ loopImpedance: value }, context).comment);
  });
  [['', ''], ['5', '10'], ['12', '10'], ['5', '']].forEach(([pfc, kA]) => {
    texts.push(deriveFaultCurrentAssessment({ prospectiveFaultCurrent: pfc, breakingCapacity: kA }, context).comment);
    texts.push(deriveFaultCurrentAssessment({ prospectiveFaultCurrent: pfc, breakingCapacity: kA, shortCircuitCurrentLN: '3' }, context).comment);
  });
  [['', ''], ['0.3', '0.4'], ['1.4', '0.4'], ['0.3', '']].forEach(([equipment, grid]) => {
    texts.push(deriveRiserAssessment({ resistanceTowardsEquipment: equipment, resistanceTowardsGrid: grid, testLeadResistance: '0.1' }, context).comment);
  });
  [null, 5, 11, 13].forEach((zt) => texts.push(deriveTowerFootingAssessment({}, { totalImpedanceZt: zt }, context).comment));
  assert.strictEqual(texts.length, 28);
  texts.forEach((text) => assert.ok(isGeneratedRemark(text), `not recognised: ${text}`));
});

check("an engineer's own text is kept, including text that mentions the same subjects", () => {
  [
    'Feeder labelled differently on site.',
    'Continuity checked again after tightening the lug.',
    'Zs high because the cable run is long; client informed.',
    'Healthy plant, but tag missing',
    'PFC measured at the bus, not the feeder.',
    'Riser R2 bolt replaced.'
  ].forEach((text) => {
    assert.strictEqual(engineerRemark(text), text);
    assert.ok(!isGeneratedRemark(text));
  });
  assert.strictEqual(engineerRemark('  Typed note  '), 'Typed note');
  assert.strictEqual(engineerRemark(''), '');
  assert.strictEqual(engineerRemark(undefined), '');
});

check('normalizing no longer saves generated text; typed text survives', () => {
  const report = normalizeReportInput({
    project,
    tests: { continuityTest: true, loopImpedanceTest: true, prospectiveFaultCurrent: true, riserIntegrityTest: true, earthContinuityTest: true, towerFootingResistance: true },
    continuityTest: [{ resistance: '0.3' }, { resistance: '0.4', comment: 'Re-tested after cleaning.' }],
    loopImpedanceTest: [{ loopImpedance: '0.5' }, { loopImpedance: '0.5', remarks: LEGACY_ENTER_ZS }, { loopImpedance: '0.5', remarks: 'Long cable run.' }],
    prospectiveFaultCurrent: [{ prospectiveFaultCurrent: '5', breakingCapacity: '10' }],
    riserIntegrityTest: [{ resistanceTowardsEquipment: '0.3', resistanceTowardsGrid: '0.3' }],
    earthContinuityTest: [{ measuredValue: '0.3', remark: LEGACY_EARTH_ENTER }],
    towerFootingResistance: [{ mainLocationTower: 'T1', readings: [1, 2, 3, 4].map(() => ({ measuredCurrentMa: '10', measuredImpedance: '1' })), remarks: 'Marginal' }]
  }, { allowRetired: true });
  assert.strictEqual(report.continuityTest[0].comment, '');
  assert.strictEqual(report.continuityTest[1].comment, 'Re-tested after cleaning.');
  assert.deepStrictEqual(report.loopImpedanceTest.map((row) => row.remarks), ['', '', 'Long cable run.']);
  assert.strictEqual(report.prospectiveFaultCurrent[0].comment, '');
  assert.strictEqual(report.riserIntegrityTest[0].comment, '');
  assert.strictEqual(report.earthContinuityTest[0].remark, '');
  assert.strictEqual(report.towerFootingResistance[0].remarks, '');
});

check('an old report with a contradicting saved remark shows the live remark instead', () => {
  const stored = {
    project: { ...project, reportDate: '2026-01-01' },
    tests: { loopImpedanceTest: true, riserIntegrityTest: true },
    // saved long ago: the "healthy range" sentence on a row that is over the Zs limit today
    loopImpedanceTest: [{ loopImpedance: '0.5', remarks: LEGACY_ZS_HEALTHY }, { loopImpedance: '0.6', remarks: LEGACY_ZS_HEALTHY }, { loopImpedance: '2.4', remarks: LEGACY_ZS_HEALTHY }],
    riserIntegrityTest: [{ resistanceTowardsEquipment: '1.4', resistanceTowardsGrid: '0.2', comment: LEGACY_RISER_WITHIN }]
  };
  const scrubbed = scrubGeneratedRemarks(stored);
  assert.strictEqual(scrubbed.loopImpedanceTest[2].remarks, '');
  assert.strictEqual(stored.loopImpedanceTest[2].remarks, LEGACY_ZS_HEALTHY, 'the stored report itself is not mutated');
  const assessment = assessReport(scrubbed);
  const loop = assessment.tests.loopImpedanceTest.rows[2];
  assert.strictEqual(loop.status.label, '>Permissible Limit');
  assert.ok(loop.remark.includes('exceeds') && !loop.remark.includes('healthy range'), loop.remark);
  const riser = assessment.tests.riserIntegrityTest.rows[0];
  assert.strictEqual(riser.status.label, '>Permissible Limit');
  assert.ok(riser.remark.includes('exceed'), riser.remark);
});

check('a typed remark wins over the generated one; PSC-not-recorded shows in the live remark', () => {
  const report = {
    project: {},
    tests: { prospectiveFaultCurrent: true },
    prospectiveFaultCurrent: [
      { prospectiveFaultCurrent: '5', breakingCapacity: '10', comment: 'Nameplate photographed.' },
      { prospectiveFaultCurrent: '5', breakingCapacity: '10' }
    ]
  };
  const rows = assessReport(report).tests.prospectiveFaultCurrent.rows;
  assert.strictEqual(rows[0].remark, 'Nameplate photographed.');
  assert.ok(rows[1].remark.includes('PSC not recorded'), rows[1].remark);
});

console.log(`\n${passed} live-remark checks passed.`);
