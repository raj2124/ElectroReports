// Rules and validation behind the approved sheets MS-01..MS-09 (Session 6B.1a):
// field validation, header-level lead resistance, plate size, the 50/70% agreement
// flag, ground resistance, step & touch (Cs), tower defects, retired tests,
// calibration warnings and the FINAL approval gate.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createReportStore } = require('../src/reportStore');
const { assessReport } = require('../src/assessmentService');
const { buildReviewState } = require('../src/review');
const { validateReport } = require('../src/sheetValidation');
const {
  normalizeReportInput,
  ReportValidationError,
  buildAssessmentContext,
  deriveElectrodeAssessment,
  deriveContinuityAssessment,
  deriveRiserAssessment,
  deriveGroundResistanceAssessment,
  deriveStepTouchAssessment,
  deriveTowerFootingAssessment,
  summarizeTowerGroups,
  buildTowerGroupKey,
  checkFopAgreement,
  plateAreaOf,
  resolveSurfaceFactor
} = require('../src/reportModel');
const { resolveProfile } = require('../src/acceptance');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}
const near = (actual, expected, tolerance = 1e-6) => assert.ok(Math.abs(actual - expected) <= tolerance, `${actual} is not within ${tolerance} of ${expected}`);

const project = { projectNo: 'P-1', clientName: 'C', siteLocation: 'S', workOrder: 'W', reportDate: '2026-10-13', engineerName: 'E', assumedSoilResistivity: '30', assumedSoilResistivitySource: 'GT-12' };
const base = (extra = {}) => ({ project, tests: { electrodeResistance: true }, electrodeResistance: [{ tag: 'EP-1', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '2' }], ...extra });
const profile = resolveProfile();
const ctxFor = (report) => buildAssessmentContext(report);
const errorCodes = (payload) => {
  try {
    normalizeReportInput(payload);
    return [];
  } catch (error) {
    assert.ok(error instanceof ReportValidationError, `expected a ReportValidationError, got ${error.message}`);
    return error.errors.map((item) => item.code);
  }
};

check('new fields normalize: job header, soil directions/bearings, plate size, codes, tower reading fields', () => {
  const report = normalizeReportInput({
    ...base(),
    project: { ...project, assumedSoilResistivity: '', assumedSoilResistivitySource: '', groundCondition: 'Damp', clientRepresentative: 'B. Thakkar', soilSpacingPreset: 'is3043-substation' },
    tests: { soilResistivity: true, electrodeResistance: true, towerFootingResistance: true },
    soilResistivity: {
      locations: [{ name: 'Yard', inputUnit: 'ohm', probeDepthCm: '12', weather: 'Sunny', bearings: { direction1: '0', direction2: '90', direction3: '45' },
        direction1: [{ spacing: '2', resistivity: '9.1' }, { spacing: '5', resistivity: '' }], direction2: [{ spacing: '2', resistivity: '9.4', remark: 'dry top soil' }],
        direction3: [{ spacing: '2', resistivity: '9.0' }] }]
    },
    electrodeResistance: [{ tag: 'EP-2', electrodeType: 'Plate', materialType: 'Copper Bonded', plateLengthM: '1', plateWidthM: '2', resistanceWithoutGrid: '2', checkR50: '2.02', method: '61.8%', conditionCodes: ['c8', 'C3', 'X9'], photoRef: 'IMG_0412' }],
    sheetHeaders: { electrodeResistance: { instrument: 'mi3290', serialNo: '18480898', calibrationDue: '2027-03-01', currentElectrodeDistanceM: '30', auxRodsOk: 'true', check90180: 'Yes' } },
    towerFootingResistance: [{ mainLocationTower: 'T1', readings: [1, 2, 3, 4].map((n) => ({ footToEarthingConnectionStatus: n === 2 ? 'Not Given' : 'Given', measuredCurrentMa: '10', measuredImpedance: '1', conditionCodes: n === 3 ? ['C5'] : [], photoRef: n === 3 ? 'IMG_9' : '' })) }]
  });
  assert.strictEqual(report.project.groundCondition, 'Damp');
  assert.strictEqual(report.project.soilSpacingPreset, 'is3043-substation');
  const location = report.soilResistivity.locations[0];
  assert.strictEqual(location.probeDepthCm, '12');
  assert.deepStrictEqual([location.bearings.direction1, location.bearings.direction2, location.bearings.direction3, location.bearings.direction4], ['0', '90', '45', '']);
  assert.strictEqual(location.direction1.length, 1, 'a spacing with no reading is not a measurement row');
  assert.strictEqual(location.direction2[0].remark, 'dry top soil');
  assert.strictEqual(location.direction3.length, 1);
  assert.deepStrictEqual(location.direction8, []);
  const pit = report.electrodeResistance[0];
  assert.strictEqual(pit.materialType, 'Cu-bonded', 'legacy "Copper Bonded" maps to the printed option');
  assert.deepStrictEqual(pit.conditionCodes, ['C3', 'C8'], 'unknown codes are dropped, order is canonical');
  assert.strictEqual(pit.photoRef, 'IMG_0412');
  assert.strictEqual(report.sheetHeaders.electrodeResistance.auxRodsOk, true);
  assert.strictEqual(report.sheetHeaders.electrodeResistance.serialNo, '18480898');
  const reading = report.towerFootingResistance[0].readings;
  assert.deepStrictEqual(reading.map((r) => r.footToEarthingConnectionStatus), ['Given', 'Not given', 'Given', 'Given']);
  assert.deepStrictEqual(reading[2].conditionCodes, ['C5']);
});

check('invalid data is an error that blocks the save', () => {
  assert.deepStrictEqual(errorCodes(base({ electrodeResistance: [{ tag: 'A', length: '3', diameter: '40', resistanceWithoutGrid: 'abc' }] })), ['NOT_NUMBER']);
  assert.deepStrictEqual(errorCodes(base({ electrodeResistance: [{ tag: 'A', length: '-3', diameter: '40' }] })), ['OUT_OF_RANGE']);
  assert.deepStrictEqual(errorCodes(base({ electrodeResistance: [{ tag: 'A', electrodeType: 'Grid', length: '3', diameter: '40' }] })), ['BAD_OPTION']);
  assert.deepStrictEqual(errorCodes(base({ electrodeResistance: [{ tag: 'A', method: 'Guess', length: '3', diameter: '40' }] })), ['BAD_OPTION']);
  assert.deepStrictEqual(errorCodes({ ...base(), tests: { continuityTest: true }, continuityTest: [{ tag: 'B', resistance: '0.3', conditionCodes: ['C3'] }] }), ['CODE_NOT_ALLOWED']);
  assert.deepStrictEqual(errorCodes({ ...base(), project: { ...project, groundCondition: 'Soggy' } }), ['BAD_OPTION']);
  assert.deepStrictEqual(errorCodes({ ...base(), project: { ...project, soilSpacingPreset: 'x' } }), ['BAD_OPTION']);
  assert.deepStrictEqual(errorCodes({ ...base(), sheetHeaders: { electrodeResistance: { instrument: 'kyoritsu4118a', serialNo: '1' } } }), ['INSTRUMENT_NOT_ALLOWED']);
  assert.deepStrictEqual(errorCodes({ ...base(), sheetHeaders: { electrodeResistance: { instrument: 'mi3290', serialNo: '1', calibrationDue: '01/03/2027' } } }), ['BAD_DATE']);
  assert.deepStrictEqual(errorCodes({ ...base(), tests: { loopImpedanceTest: true }, loopImpedanceTest: [{ location: 'P', deviceType: 'Breaker', loopImpedance: '0.5' }] }), ['BAD_OPTION']);
  assert.deepStrictEqual(errorCodes({ ...base(), sheetHeaders: { electrodeResistance: { currentElectrodeDistanceM: 'far' } } }), ['NOT_NUMBER']);
  assert.deepStrictEqual(errorCodes({ ...base(), project: { ...project, assumedSoilResistivity: '', assumedSoilResistivitySource: '' }, tests: { soilResistivity: true }, soilResistivity: { locations: [{ name: 'L', bearings: { direction1: '400' }, direction1: [{ spacing: '1', resistivity: '5' }], direction2: [{ spacing: '1', resistivity: '5' }] }] } }), ['OUT_OF_RANGE']);
  // a dash means "not taken" (clamp method), not an error
  assert.deepStrictEqual(errorCodes({ ...base(), tests: { groundResistance: true }, groundResistance: [{ tag: 'EP-1', method: 'Clamp', currentProbeDistanceM: '—', r50: '—', r618: '3.2', r70: '—' }] }), []);
});

check('an assumed rho is rejected when soil was measured, and needs a number', () => {
  const soil = { locations: [{ name: 'L', inputUnit: 'ohm-m', direction1: [{ spacing: '1', resistivity: '40' }], direction2: [{ spacing: '1', resistivity: '40' }] }] };
  assert.deepStrictEqual(errorCodes({ ...base(), tests: { soilResistivity: true, electrodeResistance: true }, soilResistivity: soil }), ['ASSUMED_RHO_WITH_MEASURED']);
  const measured = { ...base(), project: { ...project, assumedSoilResistivity: '', assumedSoilResistivitySource: '' }, tests: { soilResistivity: true, electrodeResistance: true }, soilResistivity: soil };
  assert.doesNotThrow(() => normalizeReportInput(measured));
});

check('warnings are saved and shown, never blocking: missing dimensions, probe depth, serial, calibration', () => {
  const draft = {
    ...base(),
    project: { ...project, reportDate: '2026-10-13' },
    tests: { soilResistivity: true, electrodeResistance: true },
    soilResistivity: { locations: [{ name: 'L', inputUnit: 'ohm-m', probeDepthCm: '25', direction1: [{ spacing: '1', resistivity: '40' }], direction2: [{ spacing: '1', resistivity: '40' }] }] },
    electrodeResistance: [{ tag: 'EP-9', electrodeType: 'Rod', materialType: 'Cu', resistanceWithoutGrid: '2' }, { tag: 'EP-8', electrodeType: 'Plate', materialType: 'Cu', resistanceWithGrid: '2' }],
    sheetHeaders: { electrodeResistance: { instrument: 'mi3290', calibrationDue: '2026-09-01' } }
  };
  draft.project.assumedSoilResistivity = '';
  draft.project.assumedSoilResistivitySource = '';
  assert.doesNotThrow(() => normalizeReportInput(draft), 'warnings never block the save');
  const normalized = normalizeReportInput(draft);
  const codes = assessReport(normalized).validation.warnings.map((warning) => warning.code).sort();
  assert.deepStrictEqual(codes, ['CAL_EXPIRED', 'MISSING_DIMS', 'MISSING_DIMS', 'PROBE_DEPTH', 'SERIAL_MISSING']);
  const expired = assessReport(normalized).reviewItems.find((item) => item.code === 'CAL_EXPIRED');
  assert.strictEqual(expired.blocksApproval, true);
  assert.ok(expired.text.includes('2026-09-01') && expired.text.includes('2026-10-13'));
  // calibration due on the work date is still valid
  const ok = { ...draft, sheetHeaders: { electrodeResistance: { instrument: 'mi3290', serialNo: '1', calibrationDue: '2026-10-13' } } };
  assert.ok(!assessReport(normalizeReportInput(ok)).validation.warnings.some((warning) => warning.code === 'CAL_EXPIRED'));
});

check('plate area is 2 x length x width (both faces); older plateArea still works', () => {
  assert.strictEqual(plateAreaOf({ plateLengthM: '1', plateWidthM: '2' }), 4);
  assert.strictEqual(plateAreaOf({ plateArea: '5' }), 5);
  const ctx = { profile, rho: 30, rhoAssumed: false, rhoSource: '' };
  const plate = deriveElectrodeAssessment({ electrodeType: 'Plate', plateLengthM: '1', plateWidthM: '2', resistanceWithoutGrid: '3' }, ctx);
  near(plate.limit, Math.round((30 / 4) * Math.sqrt(Math.PI / 4) * 100) / 100, 0.005);
  assert.strictEqual(plate.limitCriterionId, 'PIT-PLATE');
});

check('50% / 70% check readings: flagged beyond the 5% register tolerance, status and priority unchanged', () => {
  const rod = (extra) => deriveElectrodeAssessment({ electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '1', ...extra }, { profile, rho: 30, rhoAssumed: false, rhoSource: '' });
  const agree = rod({ checkR50: '1.04', checkR70: '1.05' });
  assert.deepStrictEqual(agree.flags, []);
  assert.strictEqual(agree.fop.agrees, true);
  const disagree = rod({ checkR50: '1.04', checkR70: '1.06' });
  assert.strictEqual(disagree.fop.agrees, false);
  assert.strictEqual(disagree.fop.tolerancePct, 5);
  assert.strictEqual(disagree.fop.r70DevPct, 6);
  assert.strictEqual(disagree.flags[0].code, 'FOP_DISAGREE');
  assert.ok(disagree.flags[0].review);
  assert.ok(disagree.comment.includes('re-test – 61.8% value unconfirmed'), disagree.comment);
  assert.deepStrictEqual([disagree.status.label, disagree.priority], [agree.status.label, agree.priority], 'flag only');
  assert.strictEqual(rod({ checkR50: '1.05' }).fop.agrees, true, 'exactly 5% agrees');
  assert.strictEqual(rod({ checkR50: '1.0501' }).fop.agrees, false);
  assert.strictEqual(rod({ checkR50: '2', method: 'Clamp' }).fop, null, 'clamp takes no 50/70% readings');
  assert.strictEqual(rod({ checkR50: '2', method: 'Slope' }).fop, null);
  assert.strictEqual(rod({}).fop, null, 'no check readings, nothing to compare');
  assert.strictEqual(checkFopAgreement({ r618: 0, r50: 0.1, r70: 0 }, profile).agrees, false);
  assert.strictEqual(checkFopAgreement({ r618: 0, r50: 0, r70: 0 }, profile).agrees, true);
  assert.ok(profile.params('PIT-METHOD').agreement_tolerance_source.includes('gives no number'));
  assert.strictEqual(profile.params('PIT-METHOD').agreement_tolerance_basis, 'etpl_acceptance');
  // it reaches the reviewer's checklist
  const report = normalizeReportInput(base({ electrodeResistance: [{ tag: 'EP-1', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '1', checkR50: '1.2', checkR70: '1.0' }] }));
  const items = assessReport(report).reviewItems;
  assert.ok(items.some((item) => item.code === 'FOP_DISAGREE' && !item.blocksApproval));
});

check('test-lead resistance comes from the sheet header; a value on the row (older reports, stitched finals) wins', () => {
  const report = { project, sheetHeaders: { continuityTest: { testLeadResistance: '0.3' }, riserIntegrityTest: { testLeadResistance: '0.2' } } };
  const ctx = ctxFor(report);
  assert.strictEqual(deriveContinuityAssessment({ resistance: '1.2' }, ctx).measured, 0.9);
  assert.strictEqual(deriveContinuityAssessment({ resistance: '1.2' }, ctx).status.label, 'Healthy');
  assert.strictEqual(deriveContinuityAssessment({ resistance: '1.2', testLeadResistance: '0.1' }, ctx).measured, 1.1);
  assert.strictEqual(deriveRiserAssessment({ resistanceTowardsEquipment: '1.15', resistanceTowardsGrid: '0.5' }, ctx).status.label, 'Healthy');
  assert.strictEqual(deriveRiserAssessment({ resistanceTowardsEquipment: '1.25', resistanceTowardsGrid: '0.5' }, ctx).status.label, '>Permissible Limit', '1.25 - 0.2 = 1.05 ohm is over the 1 ohm limit');
});

check('ground resistance (MS-07) is compared with the limit of the pit sharing its tag', () => {
  const report = normalizeReportInput({
    ...base(),
    tests: { electrodeResistance: true, groundResistance: true },
    groundResistance: [
      { tag: 'EP-1', method: '61.8%', currentProbeDistanceM: '60', r50: '7.0', r618: '7.0', r70: '7.1' },
      { tag: 'ep-1', r618: '9.5' },
      { tag: 'EP-77', r618: '1.0' },
      { tag: '', r618: '1.0' }
    ]
  });
  const ctx = ctxFor(report);
  const [within, over, unknown, untagged] = report.groundResistance.map((row) => deriveGroundResistanceAssessment(row, ctx));
  assert.strictEqual(within.status.label, 'Healthy');
  assert.strictEqual(within.limitSource, 'limit of pit EP-1');
  near(within.limit, 7.97, 0.005);
  assert.strictEqual(over.status.label, '>Permissible Limit', 'tag match ignores case');
  assert.strictEqual(unknown.status.label, 'Pending');
  assert.ok(unknown.comment.includes('EP-77'));
  assert.strictEqual(untagged.status.label, 'Pending');
  assert.strictEqual(summaryFor(report, 'groundResistance').healthy, 1);
});
function summaryFor(report, testId) {
  return assessReport(report).summaries[testId];
}

check('step & touch (MS-08): scaled by If/It, Cs by the IEEE 80 formula, client Cs wins, stricter body weight governs', () => {
  const header = { faultCurrentKA: '10', faultClearingTimeS: '0.5', bodyWeightBasis: 'both' };
  const row = (extra) => ({ measurementType: 'Touch-mesh', testCurrentA: '5', v1: '0.04', v2: '0.05', v3: '0.03', ...extra });
  const assess = (rowExtra, headerExtra = {}, rho = 30) => deriveStepTouchAssessment(row(rowExtra), { profile, rho, rhoAssumed: false, rhoSource: '', headers: { stepTouchVoltage: { ...header, ...headerExtra } } });

  // no surface layer: Cs = 1 and rho_s = site rho; 50 kg is stricter than 70 kg
  const plain = assess({});
  assert.strictEqual(plain.cs, 1);
  assert.strictEqual(plain.csSource, 'no surface layer (Cs = 1)');
  assert.strictEqual(plain.bodyWeightKg, 50);
  near(plain.limit, ((1000 + 1.5 * 30) * 0.116) / Math.sqrt(0.5), 0.005);
  assert.strictEqual(plain.vmax, 0.05, 'Vmax defaults to the largest of V1-V3');
  near(plain.scaledVolts, 0.05 * (10000 / 5), 1e-9); // 100 V
  assert.strictEqual(plain.status.label, 'Healthy');
  assert.strictEqual(assess({ vmax: '0.1' }).status.label, '>Permissible Limit', '200 V is above the ~171 V touch limit');
  assert.strictEqual(assess({ measurementType: 'Step' }).limit > plain.limit, true, 'step limit is higher than touch');
  assert.strictEqual(assess({}, { bodyWeightBasis: '70' }).bodyWeightKg, 70);
  assert.ok(assess({}, { bodyWeightBasis: '70' }).limit > plain.limit);

  // surface layer: Cs = 1 - 0.09 (1 - rho/rho_s) / (2 hs + 0.09)
  const layered = assess({}, { surfaceResistivity: '2500', surfaceThicknessM: '0.1' });
  near(layered.cs, 1 - (0.09 * (1 - 30 / 2500)) / (2 * 0.1 + 0.09));
  assert.strictEqual(layered.csSource, 'IEEE 80-2013, surface layer derating factor Cs');
  near(layered.limit, ((1000 + 1.5 * layered.cs * 2500) * 0.116) / Math.sqrt(0.5), 0.005);

  // client design Cs wins and is printed as client-specified
  const client = assess({}, { surfaceResistivity: '2500', surfaceThicknessM: '0.1', csOverride: '0.8', csSource: 'GFL design calc 12' });
  assert.strictEqual(client.cs, 0.8);
  assert.strictEqual(client.csSource, 'client-specified (GFL design calc 12)');
  assert.ok(client.comment.includes('client-specified (GFL design calc 12)'));

  // missing inputs: Pending, with the list of what is missing
  const missing = assess({}, { faultCurrentKA: '', faultClearingTimeS: '', bodyWeightBasis: '' });
  assert.strictEqual(missing.status.label, 'Pending');
  assert.deepStrictEqual(missing.missing, ['earth fault current If (header)', 'fault clearing time ts (header)', 'body weight basis (header)']);
  assert.strictEqual(assess({}, { surfaceResistivity: '2500', surfaceThicknessM: '0.1' }, null).status.label, 'Pending', 'a surface layer needs the site rho for Cs');
  assert.strictEqual(deriveStepTouchAssessment({ measurementType: '', v1: '1' }, { profile, rho: 30, headers: {} }).status.label, 'Pending');
  assert.strictEqual(resolveSurfaceFactor({}, 30, profile).cs, 1);
  // a final report's rows carry their own day's header
  const stitched = deriveStepTouchAssessment(row({ dayHeader: { ...header, faultCurrentKA: '20' } }), { profile, rho: 30, headers: { stepTouchVoltage: header } });
  near(stitched.scaledVolts, 0.05 * (20000 / 5), 1e-9);
});

check("tower: a 'Not given' foot is a defect flag with a remark; the status stays on Zt", () => {
  const groups = (status) => [{ groupId: 'g1', mainLocationTower: 'T1', readings: [1, 2, 3, 4].map((n) => ({ footToEarthingConnectionStatus: n === 2 ? status : 'Given', measuredCurrentMa: '10', measuredImpedance: '2' })) }];
  const run = (status) => {
    const list = groups(status);
    return deriveTowerFootingAssessment(list[0], summarizeTowerGroups(list).get(buildTowerGroupKey(list[0])), { profile });
  };
  const given = run('Given');
  const notGiven = run('Not given');
  assert.strictEqual(given.flags.length, 0);
  assert.deepStrictEqual(notGiven.flags.map((flag) => [flag.code, flag.feet]), [['FOOT_NOT_GIVEN', ['F-2']]]);
  assert.strictEqual(notGiven.status.label, given.status.label, 'status unchanged');
  assert.strictEqual(notGiven.totalImpedanceZt, given.totalImpedanceZt);
  assert.strictEqual(notGiven.comment, given.comment, 'the status word is unchanged');
  const assessed = assessReport({ project, tests: { towerFootingResistance: true }, towerFootingResistance: groups('Not given') });
  assert.ok(assessed.tests.towerFootingResistance.rows[0].remark.includes('Foot F-2 has no earth connection (defect)'));
  assert.strictEqual(assessed.tests.towerFootingResistance.rows[0].typedRemark, '');
});

check('earthContinuityTest is retired for new reports and still reads for old ones', () => {
  const payload = { ...base(), tests: { earthContinuityTest: true }, earthContinuityTest: [{ tag: 'E', measuredValue: '0.3' }] };
  assert.throws(() => normalizeReportInput(payload), /retired/);
  const old = normalizeReportInput(payload, { allowRetired: true });
  assert.strictEqual(assessReport(old).tests.earthContinuityTest.rows[0].status.label, 'Healthy');
  assert.strictEqual(assessReport(old).summaries.earthContinuityTest.healthy, 1);
});

check('validation runs on drafts too, live, with field paths', () => {
  const draft = { ...base(), electrodeResistance: [{ tag: 'A', length: 'x', diameter: '40' }] };
  const { errors } = assessReport(draft).validation;
  assert.deepStrictEqual(errors.map((item) => [item.code, item.path]), [['NOT_NUMBER', 'electrodeResistance.0.length']]);
  assert.ok(errors[0].message.includes('MS-02 row 1'));
  assert.deepStrictEqual(validateReport({ tests: {}, project: {} }), { errors: [], warnings: [] });
});

check('stitching snapshots each day\'s header onto its rows (lead resistance, step & touch header)', () => {
  const { assembleFinalReport } = require('../src/reportCollectionModel');
  const day = (id, date, lead, ifKA) => ({
    id, createdAt: date, collectionDate: date, project: { ...project, reportDate: date }, tests: { continuityTest: true, stepTouchVoltage: true },
    sheetHeaders: { continuityTest: { testLeadResistance: lead }, stepTouchVoltage: { faultCurrentKA: ifKA, faultClearingTimeS: '0.5', bodyWeightBasis: '70' } },
    continuityTest: [{ tag: `T-${id}`, mainLocation: 'A', measurementPoint: id, resistance: '1.2' }, { tag: 'own', mainLocation: 'A', measurementPoint: 'own-lead', resistance: '1.2', testLeadResistance: '0.05' }],
    stepTouchVoltage: [{ location: id, measurementType: 'Step', surface: 'Soil', distanceM: '1', testCurrentA: '5', v1: '0.05' }]
  });
  const final = assembleFinalReport([day('d1', '2026-10-10', '0.3', '10'), day('d2', '2026-10-11', '0.1', '20')]);
  const byPoint = Object.fromEntries(final.continuityTest.map((row) => [row.measurementPoint, row.testLeadResistance]));
  assert.strictEqual(byPoint.d1, '0.3');
  assert.strictEqual(byPoint.d2, '0.1');
  assert.strictEqual(byPoint['own-lead'], '0.05', 'a lead value already on the row is kept');
  assert.deepStrictEqual(final.stepTouchVoltage.map((row) => row.dayHeader.faultCurrentKA), ['10', '20']);
  const ctx = ctxFor({ project, sheetHeaders: {} });
  assert.strictEqual(deriveContinuityAssessment(final.continuityTest[0], ctx).measured, 0.9);
  assert.strictEqual(deriveContinuityAssessment(final.continuityTest.find((row) => row.measurementPoint === 'd2'), ctx).measured, 1.1);
});

// ---- FINAL approval gate -------------------------------------------------------------
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-rules-'));
const store = createReportStore({ filePath: path.join(dir, 'gate.db') });
const prj = store.addProject({ projectNo: 'G-1', clientName: 'C', siteLocation: 'S', workOrder: 'W', engineerName: 'E' });
const daily = (date, calibrationDue) => store.addReport({
  kind: 'daily', projectId: prj.id, collectionDate: date,
  project: { ...project, reportDate: date }, tests: { electrodeResistance: true },
  sheetHeaders: { electrodeResistance: { instrument: 'mi3290', serialNo: '18480898', calibrationDue } },
  electrodeResistance: [{ tag: `EP-${date}`, electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '2' }]
});

// Review round: submit, tick the checklist, mark as checked. Approval is then a separate call.
function submitAndCheck(id, actor = 'reviewer') {
  store.transitionReport(id, 'in_review', { actor: 'preparer' });
  const ticks = Object.fromEntries(buildReviewState(store, store.getReport(id)).checklist.filter((item) => item.manual).map((item) => [item.id, true]));
  store.setReviewTicks(id, ticks);
  store.markChecked(id, { actor });
}

check('final reports carry per-day calibration records and each row keeps its own day header', () => {
  const day1 = daily('2026-10-10', '2026-09-01'); // expired on the day
  const day2 = daily('2026-10-11', '2027-03-01');
  [day1, day2].forEach((report) => {
    submitAndCheck(report.id);
    store.transitionReport(report.id, 'approved', { actor: 'reviewer' }); // dailies are not gated on calibration
  });
  const { report: final } = store.assembleFinal(prj.id);
  assert.strictEqual(final.calibrationRecords.length, 2);
  assert.deepStrictEqual(final.calibrationRecords.map((record) => [record.workDate, record.calibrationDue]), [['2026-10-10', '2026-09-01'], ['2026-10-11', '2027-03-01']]);
  const items = assessReport(final).reviewItems.filter((item) => item.code === 'CAL_EXPIRED');
  assert.strictEqual(items.length, 1);
  assert.ok(items[0].text.includes('2026-10-10'));
});

check('a FINAL report with expired calibration cannot be approved until acknowledged; the note is recorded', () => {
  const final = store.listProjectReports(prj.id, { kind: 'final' })[0];
  submitAndCheck(final.id);
  const sanity = 'Sanity check unavailable (provider billing); narrative read against the tables by the reviewer.';
  assert.throws(() => store.transitionReport(final.id, 'approved', { actor: 'approver', acknowledgements: { sanityCheck: sanity } }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED' && /2026-09-01/.test(error.message) && error.items.length === 1);
  assert.throws(() => store.transitionReport(final.id, 'approved', { actor: 'approver', acknowledgement: 'Certificate on file.' }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED' && /Sanity check not run/.test(error.message));
  assert.throws(() => store.transitionReport(final.id, 'approved', { actor: 'approver', acknowledgements: { calibration: '   ', sanityCheck: sanity } }), /Acknowledge/);
  assert.strictEqual(store.getReport(final.id).reviewStatus, 'in_review', 'still in review after the refused approvals');
  const result = store.transition(final.id, 'approved', { actor: 'approver', acknowledgements: { calibration: 'Certificate re-issued 2026-09-20, held on file (CAL-4471).', sanityCheck: sanity }, note: 'Final check done' });
  assert.strictEqual(result.report.reviewStatus, 'approved');
  assert.deepStrictEqual(result.warnings, []);
  const last = store.listTransitions(final.id).pop();
  assert.deepStrictEqual([last.action, last.fromStatus, last.toStatus, last.actor], ['approve', 'in_review', 'approved', 'approver']);
  assert.ok(last.note.startsWith('Calibration acknowledged: Certificate re-issued 2026-09-20, held on file (CAL-4471).'));
  assert.ok(last.note.includes(`Sanity check not run - acknowledged: ${sanity}`));
  assert.ok(last.note.endsWith('| Final check done'));
});

check('a FINAL report with valid calibration still needs the sanity-check acknowledgement', () => {
  const clean = store.addProject({ projectNo: 'G-2', clientName: 'C', siteLocation: 'S', workOrder: 'W', engineerName: 'E' });
  const day = store.addReport({
    kind: 'daily', projectId: clean.id, collectionDate: '2026-10-12', project: { ...project, reportDate: '2026-10-12' }, tests: { electrodeResistance: true },
    sheetHeaders: { electrodeResistance: { instrument: 'mi3290', serialNo: '1', calibrationDue: '2027-01-01' } }, electrodeResistance: [{ tag: 'E', electrodeType: 'Rod', length: '3', diameter: '40', resistanceWithoutGrid: '2' }]
  });
  submitAndCheck(day.id);
  store.transitionReport(day.id, 'approved', { actor: 'reviewer' });
  const { report: final } = store.assembleFinal(clean.id);
  submitAndCheck(final.id);
  assert.throws(() => store.transitionReport(final.id, 'approved', { actor: 'approver' }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED' && error.items.length === 1 && error.items[0].key === 'sanityCheck');
  assert.strictEqual(store.transitionReport(final.id, 'approved', { actor: 'approver', acknowledgements: { sanityCheck: 'Reviewed by hand.' } }).reviewStatus, 'approved');
  assert.strictEqual(store.listTransitions(final.id).pop().note, 'Sanity check not run - acknowledged: Reviewed by hand.');
});

console.log(`\n${passed} sheet-rule checks passed.`);
