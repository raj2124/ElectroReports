// Rule-level tests for the ETPL ESHA standard v0.4 beyond the golden sample:
// acceptance profiles, electrode types, assumptions, test-lead subtraction, PFC/PSC,
// step & touch, and condition-code priority.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { resolveProfile, listProfiles, describeCriterion, loadRegister } = require('../src/acceptance');
const {
  normalizeReportInput,
  buildAssessmentContext,
  calculateElectrodeLimit,
  getSoilCategory,
  deriveElectrodeAssessment,
  deriveElectrodePriority,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveTowerFootingAssessment,
  calculateTolerableStepTouch,
  assessStepTouchReading,
  summarizeOverallStatus,
  getCorrosionAllowancePct
} = require('../src/reportModel');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}
const near = (actual, expected, tolerance = 0.005) => assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} is not within ${tolerance} of ${expected}`);
const def = resolveProfile();
const ctx = { profile: def, rho: 25.68, rhoAssumed: false, rhoSource: '' };

check('profile resolves ETPL-DEFAULT with value, basis, source and clause per criterion', () => {
  assert.strictEqual(def.profileId, 'ETPL-DEFAULT');
  const loop = def.get('LOOP-LIMIT');
  assert.deepStrictEqual([loop.value, loop.basis, loop.clause], [1.8, 'etpl_acceptance', 'IS 3043:2018 Cl. 42 (method)']);
  assert.ok(loop.source.includes('IS/IEC 60364-4-41'));
  assert.strictEqual(describeCriterion(def.get('TOWER-ZSAT')).basisLabel, 'ETPL acceptance criterion');
  assert.strictEqual(resolveProfile('').profileId, 'ETPL-DEFAULT');
});

check('CLIENT-TEMPLATE and unknown profiles are not selectable', () => {
  assert.throws(() => resolveProfile('CLIENT-TEMPLATE'), /template/);
  assert.throws(() => resolveProfile('NOPE'), /Unknown acceptance profile/);
  assert.deepStrictEqual(listProfiles().map((profile) => profile.id), ['ETPL-DEFAULT']);
});

check('a client profile overrides only the criteria it lists', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'esha-reg-')), 'register.json');
  const register = JSON.parse(JSON.stringify(loadRegister()));
  register.profiles['CLIENT-X'] = {
    description: 'Test client',
    client_spec_document: 'X-SPEC-1 R2',
    criteria: [
      { id: 'LOOP-LIMIT', value: 0.8, unit: 'ohm', basis: 'client_spec', source: 'X-SPEC-1', clause: '5.2' },
      { id: 'ST-LIMIT', stepLimitV: 300, touchLimitV: 200, basis: 'client_spec', source: 'X-SPEC-1', clause: '7.1' }
    ]
  };
  fs.writeFileSync(file, JSON.stringify(register));
  const client = resolveProfile('CLIENT-X', { registerPath: file });
  assert.strictEqual(client.value('LOOP-LIMIT'), 0.8);
  assert.strictEqual(client.get('LOOP-LIMIT').basis, 'client_spec');
  assert.strictEqual(client.get('LOOP-LIMIT').overridden, true);
  assert.strictEqual(client.value('RISER-LIMIT'), 1.0); // falls back to ETPL-DEFAULT
  assert.strictEqual(client.get('RISER-LIMIT').overridden, false);
  const row = { loopImpedance: '1.0' };
  assert.strictEqual(deriveLoopImpedanceAssessment(row, { profile: client }).status.label, '>Permissible Limit');
  assert.strictEqual(deriveLoopImpedanceAssessment(row, { profile: def }).status.label, 'Healthy');
  const step = assessStepTouchReading({ kind: 'step', measuredVolts: 10, testCurrent: 1, faultCurrent: 40, tolerable: { step: 355 } }, { profile: client });
  assert.deepStrictEqual([step.limit, step.status.label, step.limitSource], [300, '>Permissible Limit', 'X-SPEC-1, Cl. 7.1']);
});

check('IS 3043 Table 4 corrosivity classes; no pass/fail tone', () => {
  assert.strictEqual(getSoilCategory(24.9).label, 'Severely corrosive');
  assert.strictEqual(getSoilCategory(25).label, 'Moderately corrosive');
  assert.strictEqual(getSoilCategory(50).label, 'Mildly corrosive');
  assert.strictEqual(getSoilCategory(100).label, 'Very mildly corrosive');
  assert.strictEqual(getSoilCategory(10).tone, 'neutral');
});

check('plate limit (rho/4)*sqrt(pi/A); strip uses 4l/d, rod 2l/d; unusable inputs give no limit', () => {
  near(calculateElectrodeLimit({ type: 'Plate', areaM2: 2, rho: 25.68 }).limit, (25.68 / 4) * Math.sqrt(Math.PI / 2));
  const rod = calculateElectrodeLimit({ type: 'Rod', lengthM: 3, diameterMm: 40, rho: 25.68 }).limit;
  const strip = calculateElectrodeLimit({ type: 'Strip', lengthM: 3, diameterMm: 40, rho: 25.68 }).limit;
  near(strip - rod, ((100 * 25.68) / (2 * Math.PI * 300)) * Math.log(2), 1e-9);
  assert.strictEqual(calculateElectrodeLimit({ type: 'Pipe', lengthM: 3, diameterMm: 40, rho: 25.68 }).criterion.id, 'PIT-ROD');
  assert.strictEqual(calculateElectrodeLimit({ type: 'Rod', lengthM: 3, diameterMm: 40, rho: null }), null);
  assert.strictEqual(calculateElectrodeLimit({ type: 'Rod', lengthM: 0.01, diameterMm: 40, rho: 25 }), null); // ln(2l/d) <= 0
  assert.strictEqual(calculateElectrodeLimit({ type: 'Grid', lengthM: 3, diameterMm: 40, rho: 25 }), null);
});

check('electrode bands: <= 0.9 x limit Healthy, <= limit Warning, > limit over; limit itself is within', () => {
  const rod = (r) => deriveElectrodeAssessment({ electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: r }, ctx);
  const limit = calculateElectrodeLimit({ type: 'Rod', lengthM: 3, diameterMm: 40, rho: 25.68 }).limit;
  assert.strictEqual(rod(String(0.9 * limit - 0.001)).status.label, 'Healthy');
  assert.strictEqual(rod(String(0.9 * limit + 0.001)).status.label, 'Warning');
  assert.strictEqual(rod(String(limit)).status.label, 'Warning');
  assert.strictEqual(rod(String(limit + 0.001)).status.label, '>Permissible Limit');
  assert.strictEqual(rod(String(limit)).status.withinLimit, true);
});

check('without-grid reading is assessed; with-grid is only the fallback', () => {
  const both = deriveElectrodeAssessment({ electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '7', resistanceWithGrid: '1' }, ctx);
  assert.strictEqual(both.measured, 7);
  assert.strictEqual(both.measuredBasis, 'without_grid');
  const fallback = deriveElectrodeAssessment({ electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithGrid: '1' }, ctx);
  assert.strictEqual(fallback.measuredBasis, 'with_grid');
});

check('priority: healthy + no code is P3; codes lift to P1/P2; pending has none', () => {
  const healthy = { label: 'Healthy', tone: 'healthy' };
  assert.strictEqual(deriveElectrodePriority(healthy, []), 'P3');
  ['C1', 'C2', 'C4', 'C5'].forEach((code) => assert.strictEqual(deriveElectrodePriority(healthy, [code]), 'P1', code));
  ['C3', 'C6', 'C7', 'C8'].forEach((code) => assert.strictEqual(deriveElectrodePriority(healthy, [code]), 'P2', code));
  assert.strictEqual(deriveElectrodePriority(healthy, ['C6', 'C1']), 'P1');
  assert.strictEqual(deriveElectrodePriority(healthy, ['OTH']), 'P3');
  assert.strictEqual(deriveElectrodePriority({ label: 'Pending', tone: 'neutral' }, []), null);
  assert.strictEqual(deriveElectrodePriority({ label: 'Pending', tone: 'neutral' }, ['C1']), 'P1');
});

check('no soil data: assumed rho is used, flagged and sourced; without it the electrode is Pending', () => {
  const base = {
    project: { projectNo: 'P', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-01', engineerName: 'E' },
    tests: { electrodeResistance: true },
    electrodeResistance: [{ tag: 'E1', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '2' }]
  };
  const none = normalizeReportInput(base);
  assert.strictEqual(deriveElectrodeAssessment(none.electrodeResistance[0], none).status.label, 'Pending');
  const assumed = normalizeReportInput({ ...base, project: { ...base.project, assumedSoilResistivity: '30', assumedSoilResistivitySource: 'Client geotech report GT-12' } });
  const result = deriveElectrodeAssessment(assumed.electrodeResistance[0], assumed);
  assert.strictEqual(result.rhoAssumed, true);
  assert.strictEqual(result.rhoSource, 'Client geotech report GT-12');
  assert.ok(result.comment.includes('assumed soil resistivity of 30.00 ohm-m (Client geotech report GT-12)'));
  assert.throws(() => normalizeReportInput({ ...base, project: { ...base.project, assumedSoilResistivity: '30' } }), /must state its source/);
  assert.throws(() => normalizeReportInput({ ...base, project: { ...base.project, acceptanceProfile: 'NOPE' } }), /Unknown acceptance profile/);
  assert.strictEqual(buildAssessmentContext(assumed).profile.profileId, 'ETPL-DEFAULT');
});

check('soil entered as R gives rho = 2*pi*a*R; entered as rho is used as read', () => {
  const soil = (inputUnit, value) => ({ name: 'L', inputUnit, direction1: [{ spacing: '2', resistivity: value }], direction2: [{ spacing: '2', resistivity: value }] });
  const make = (location) => normalizeReportInput({
    project: { projectNo: 'P', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-01', engineerName: 'E' },
    tests: { soilResistivity: true },
    soilResistivity: { locations: [location] }
  });
  assert.strictEqual(buildAssessmentContext(make(soil('ohm', '2'))).rho, 25.13); // 2*pi*2*2
  assert.strictEqual(buildAssessmentContext(make(soil('ohm-m', '40'))).rho, 40);
});

check('continuity < 1 ohm is strict; test-lead resistance is subtracted before comparing', () => {
  const cont = (r, lead) => deriveContinuityAssessment({ resistance: r, testLeadResistance: lead }, ctx);
  assert.strictEqual(cont('0.99', '').status.label, 'Healthy');
  assert.strictEqual(cont('1.0', '').status.label, '>Permissible Limit');
  assert.strictEqual(cont('1.2', '0.3').status.label, 'Healthy'); // net 0.9
  assert.strictEqual(cont('1.2', '0.3').measured, 0.9);
  assert.strictEqual(cont('0.2', '0.5').measured, 0); // net never negative
});

check('riser: both readings net of lead must be <= 1 ohm; one reading alone cannot pass', () => {
  const riser = (eq, grid, lead) => deriveRiserAssessment({ resistanceTowardsEquipment: eq, resistanceTowardsGrid: grid, testLeadResistance: lead }, ctx).status.label;
  assert.strictEqual(riser('1.0', '0.5', ''), 'Healthy');
  assert.strictEqual(riser('1.2', '0.5', ''), '>Permissible Limit');
  assert.strictEqual(riser('1.2', '0.5', '0.3'), 'Healthy');
  assert.strictEqual(riser('0.5', '', ''), 'Pending');
  assert.strictEqual(riser('1.5', '', ''), '>Permissible Limit');
});

check('PFC: device kA must exceed max(PFC L-PE, PSC L-N); missing kA is Not Assessable', () => {
  const pfc = (row) => deriveFaultCurrentAssessment(row, ctx).status.label;
  assert.strictEqual(pfc({ breakingCapacity: '25', prospectiveFaultCurrent: '20' }), 'Healthy');
  assert.strictEqual(pfc({ breakingCapacity: '25', prospectiveFaultCurrent: '20', shortCircuitCurrentLN: '26' }), '>Permissible Limit');
  assert.strictEqual(pfc({ breakingCapacity: '25', prospectiveFaultCurrent: '25' }), '>Permissible Limit'); // must exceed
  assert.strictEqual(pfc({ breakingCapacity: '', prospectiveFaultCurrent: '5' }), 'Not Assessable');
  assert.strictEqual(pfc({ breakingCapacity: '25' }), 'Pending');
  // PFC assessed alone must say so wherever it is reported
  const alone = deriveFaultCurrentAssessment({ breakingCapacity: '25', prospectiveFaultCurrent: '20' }, ctx);
  assert.strictEqual(alone.pscNotRecorded, true);
  assert.ok(alone.comment.includes('PSC not recorded'));
  const both = deriveFaultCurrentAssessment({ breakingCapacity: '25', prospectiveFaultCurrent: '20', shortCircuitCurrentLN: '10' }, ctx);
  assert.strictEqual(both.pscNotRecorded, false);
  assert.ok(!both.comment.includes('PSC not recorded'));
  const normalized = normalizeReportInput({
    project: { projectNo: 'P', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-01', engineerName: 'E' },
    tests: { prospectiveFaultCurrent: true },
    prospectiveFaultCurrent: [{ breakingCapacity: '25', prospectiveFaultCurrent: '20' }]
  });
  // generated text is no longer saved on the row; the live remark carries the PSC note
  assert.strictEqual(normalized.prospectiveFaultCurrent[0].comment, '');
  assert.ok(deriveFaultCurrentAssessment(normalized.prospectiveFaultCurrent[0], normalized).comment.includes('PSC not recorded'));
});

check('tower: Zt vs Zsat from the profile, Marginal <= 1.2 x Zsat, Marginal is not within limit', () => {
  const tower = (zt) => deriveTowerFootingAssessment({}, { totalImpedanceZt: zt }, ctx);
  assert.strictEqual(tower(10).status.label, 'Healthy');
  assert.strictEqual(tower(12).status.label, 'Marginal');
  assert.strictEqual(tower(12.01).status.label, 'Not Acceptable');
  assert.strictEqual(tower(10).zsat, 10);
  assert.strictEqual(summarizeOverallStatus([tower(10).status, tower(12).status]).unhealthy, 1);
});

check('step & touch: IS 3043 Cl. 10.3 tolerable limits and If/It scaling', () => {
  const tolerable = calculateTolerableStepTouch({ surfaceResistivity: 100, surfaceFactor: 1, faultClearingTime: 0.5, bodyWeightKg: 70 });
  near(tolerable.step, ((1000 + 6 * 100) * 0.157) / Math.sqrt(0.5), 1e-9);
  near(tolerable.touch, ((1000 + 1.5 * 100) * 0.157) / Math.sqrt(0.5), 1e-9);
  assert.strictEqual(calculateTolerableStepTouch({ surfaceResistivity: 100, faultClearingTime: 0.5, bodyWeightKg: 50 }).k, 0.116);
  assert.strictEqual(calculateTolerableStepTouch({ surfaceResistivity: 100, faultClearingTime: 0 }), null);
  // 2 V measured at 1 A test current scales to 100 V at a 50 A fault; 100 V < ~255 V touch limit
  const touch = assessStepTouchReading({ kind: 'touch', measuredVolts: 2, testCurrent: 1, faultCurrent: 50, tolerable }, ctx);
  assert.strictEqual(touch.scaledVolts, 100);
  assert.strictEqual(touch.status.label, 'Healthy');
  // the same reading unscaled would pass; scaled at 150 A it must fail (300 V > 255 V)
  assert.strictEqual(assessStepTouchReading({ kind: 'touch', measuredVolts: 2, testCurrent: 1, faultCurrent: 150, tolerable }, ctx).status.label, '>Permissible Limit');
  assert.strictEqual(assessStepTouchReading({ kind: 'touch', measuredVolts: 2, testCurrent: 0, faultCurrent: 50, tolerable }, ctx).status.label, 'Pending');
});

check('overall: 0 unhealthy Healthy; >= 90% Minor Actions; else >Attention Required; warning within limit', () => {
  const unit = (tone, withinLimit) => ({ tone, ...(withinLimit ? { withinLimit } : {}) });
  assert.strictEqual(summarizeOverallStatus([unit('healthy'), unit('warning', true)]).label, 'Healthy');
  assert.strictEqual(summarizeOverallStatus([unit('warning')]).label, '>Attention Required'); // warning NOT within limit counts against
  assert.strictEqual(summarizeOverallStatus([...Array(9).fill(unit('healthy')), unit('critical')]).label, 'Healthy - Minor Actions');
});

check('register params drive the rules: factors, bands, labels, k-factors and override field names', () => {
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'esha-params-')), 'register.json');
  const register = JSON.parse(JSON.stringify(loadRegister()));
  register.profiles['CLIENT-P'] = {
    description: 'Params test client',
    client_spec_document: 'P-SPEC-1',
    criteria: [
      { id: 'PIT-BAND', params: { warning_factor: 0.5 }, basis: 'client_spec', source: 'P-SPEC-1', clause: '1' },
      { id: 'TOWER-ZSAT', params: { marginal_factor: 1.5 }, value: 10, unit: 'ohm', basis: 'client_spec', source: 'P-SPEC-1', clause: '2' },
      { id: 'OVERALL', params: { minor_actions_min_pass: 0.5, labels: { healthy: 'OK', minor: 'OK-ish', attention: 'BAD' } }, basis: 'client_spec', source: 'P-SPEC-1', clause: '3' },
      {
        id: 'SOIL-CLASS',
        params: { bands: [{ max: 10, class: 'Low band' }, { max: null, class: 'High band' }], corrosion_allowance: [{ max: 10, steel_pct: 50 }, { max: null, steel_pct: 0 }] },
        basis: 'client_spec', source: 'P-SPEC-1', clause: '4'
      },
      { id: 'ST-LIMIT', params: { k_50kg: 0.2, k_70kg: 0.3, client_override_fields: ['stepV', 'touchV'] }, stepV: 111, touchV: 222, basis: 'client_spec', source: 'P-SPEC-1', clause: '5' }
    ]
  };
  fs.writeFileSync(file, JSON.stringify(register));
  const client = resolveProfile('CLIENT-P', { registerPath: file });
  const clientCtx = { profile: client, rho: 25.68, rhoAssumed: false, rhoSource: '' };
  const limit = calculateElectrodeLimit({ type: 'Rod', lengthM: 3, diameterMm: 40, rho: 25.68 }).limit;
  const rod = (r, c) => deriveElectrodeAssessment({ electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: String(r) }, c).status.label;
  // 0.6 x limit: Healthy by default (<= 0.9), Warning when the band is 0.5
  assert.strictEqual(rod(0.6 * limit, ctx), 'Healthy');
  assert.strictEqual(rod(0.6 * limit, clientCtx), 'Warning');
  // tower Zt 14: Marginal at 1.2 x Zsat would be <= 12 only, so default is Not Acceptable; 1.5 makes it Marginal
  assert.strictEqual(deriveTowerFootingAssessment({}, { totalImpedanceZt: 14 }, ctx).status.label, 'Not Acceptable');
  assert.strictEqual(deriveTowerFootingAssessment({}, { totalImpedanceZt: 14 }, clientCtx).status.label, 'Marginal');
  // overall: 1 of 2 passing is >Attention Required by default, minor actions at a 0.5 threshold; labels come from params
  const units = [{ tone: 'healthy' }, { tone: 'critical' }];
  assert.strictEqual(summarizeOverallStatus(units, def).label, '>Attention Required');
  assert.strictEqual(summarizeOverallStatus(units, client).label, 'OK-ish');
  assert.strictEqual(summarizeOverallStatus([{ tone: 'healthy' }], client).label, 'OK');
  assert.strictEqual(summarizeOverallStatus([{ tone: 'critical' }], client).label, 'BAD');
  // soil bands and corrosion allowance
  assert.strictEqual(getSoilCategory(30, def).label, 'Moderately corrosive');
  assert.strictEqual(getSoilCategory(30, client).label, 'High band');
  assert.strictEqual(getCorrosionAllowancePct(20, def), 30);
  assert.strictEqual(getCorrosionAllowancePct(5, client), 50);
  assert.strictEqual(getCorrosionAllowancePct(30, def), 15);
  assert.strictEqual(getCorrosionAllowancePct(200, def), 0);
  // step & touch: k-factors and the override field names both come from ST-LIMIT params
  near(calculateTolerableStepTouch({ surfaceResistivity: 100, faultClearingTime: 1, bodyWeightKg: 70 }, client).step, (1000 + 600) * 0.3, 1e-9);
  near(calculateTolerableStepTouch({ surfaceResistivity: 100, faultClearingTime: 1, bodyWeightKg: 50 }, client).touch, (1000 + 150) * 0.2, 1e-9);
  assert.strictEqual(calculateTolerableStepTouch({ surfaceResistivity: 100, faultClearingTime: 1, bodyWeightKg: 60 }, def), null);
  assert.strictEqual(assessStepTouchReading({ kind: 'step', measuredVolts: 1, testCurrent: 1, faultCurrent: 100, tolerable: { step: 999 } }, { profile: client }).limit, 111);
  assert.strictEqual(assessStepTouchReading({ kind: 'touch', measuredVolts: 1, testCurrent: 1, faultCurrent: 100, tolerable: { touch: 999 } }, { profile: client }).limit, 222);
  // a profile without params fails loudly instead of defaulting silently
  const broken = JSON.parse(JSON.stringify(register));
  delete broken.profiles['ETPL-DEFAULT'].criteria.find((criterion) => criterion.id === 'PIT-BAND').params;
  const brokenFile = path.join(path.dirname(file), 'broken.json');
  fs.writeFileSync(brokenFile, JSON.stringify(broken));
  assert.throws(() => resolveProfile('', { registerPath: brokenFile }).params('PIT-BAND'), /has no params/);
});

console.log(`\n${passed} ESHA rule checks passed.`);
