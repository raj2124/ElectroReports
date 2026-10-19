// Final-report computed results and the consistency check (Session 6C).
// Uses the make_report.py sample data, so the counts match test:esha.
const assert = require('assert');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { buildAllTestsPayload } = require('./fixtures/golden-all-tests');
const { deterministicNarrative } = require('../src/finalReport/narrative');
const { buildFinalReportData, checkConsistency, assertConsistent, ConsistencyError } = require('../src/finalReport/data');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const fresh = () => buildFinalReportData(normalizeReportInput(buildPayload()));
const data = fresh();
const summary = (id) => data.summaryTable.find((row) => row.testId === id);
const cap = (testId) => data.cap.filter((line) => line.testId === testId);

check('summary table carries the golden counts and statuses', () => {
  const electrode = summary('electrodeResistance');
  assert.deepStrictEqual([electrode.total, electrode.healthy, electrode.unhealthy, electrode.notAssessed, electrode.passPct, electrode.status], [9, 8, 1, 1, 88.9, '>Attention Required']);
  assert.deepStrictEqual([summary('loopImpedanceTest').passPct, summary('loopImpedanceTest').status], [75, '>Attention Required']);
  assert.deepStrictEqual([summary('prospectiveFaultCurrent').total, summary('prospectiveFaultCurrent').notAssessed, summary('prospectiveFaultCurrent').passPct], [3, 1, 100]);
  assert.strictEqual(summary('riserIntegrityTest').passPct, 83.3);
});

check('electrode priority block: P1 EP-03/07/08, P2 EP-05/09, P3 the other five', () => {
  assert.deepStrictEqual(data.electrodePriority.P1.tags, ['EP-03', 'EP-07', 'EP-08']);
  assert.deepStrictEqual(data.electrodePriority.P2.tags, ['EP-05', 'EP-09']);
  assert.strictEqual(data.electrodePriority.P3.count, 5);
});

check('CAP: one line per real defect point, in test order', () => {
  const defects = (id) => cap(id).filter((line) => line.kind === 'point');
  assert.deepStrictEqual(defects('electrodeResistance').map((line) => [line.point, line.priorityOrStatus]), [['EP-03', 'P1'], ['EP-05', 'P2'], ['EP-07', 'P1'], ['EP-08', 'P1'], ['EP-09', 'P2']]);
  assert.ok(defects('electrodeResistance')[0].observation.includes('C1: Corroded'));
  assert.ok(defects('electrodeResistance')[1].observation.includes('C8: No tag / label'));
  assert.deepStrictEqual(defects('loopImpedanceTest').map((line) => line.point), ['MCC-2 7F2 Feeder']);
  assert.ok(defects('loopImpedanceTest')[0].observation.includes('Zs 1.92 ohm on B-E exceeds 1.8 ohm (ETPL criterion)'));
  assert.strictEqual(defects('prospectiveFaultCurrent').length, 0);
  assert.ok(defects('riserIntegrityTest').some((line) => line.observation.includes('C6: Loose connection')));
});

check('CAP: data gaps are one line per test listing the affected points', () => {
  const gaps = cap('prospectiveFaultCurrent').filter((line) => line.kind === 'gap');
  assert.deepStrictEqual(gaps.map((line) => line.gap).sort(), ['NOT_ASSESSABLE', 'PSC_NOT_RECORDED']);
  const psc = gaps.find((line) => line.gap === 'PSC_NOT_RECORDED');
  assert.strictEqual(psc.affectedPoints.length, 4);
  assert.strictEqual(psc.priorityOrStatus, 'Data gap');
  assert.deepStrictEqual(gaps.find((line) => line.gap === 'NOT_ASSESSABLE').affectedPoints, ['APDB-2 Incomer']);
  assert.strictEqual(data.cap.filter((line) => line.kind === 'gap' && line.testId !== 'prospectiveFaultCurrent').length, 0);
});

check('CAP: unconfirmed 50/70% readings group into one FOP line', () => {
  const payload = buildPayload();
  [0, 1].forEach((i) => { Object.assign(payload.electrodeResistance[i], { checkR50: '1.50', checkR70: '1.60' }); });
  const result = buildFinalReportData(normalizeReportInput(payload));
  const lines = result.cap.filter((line) => line.gap === 'FOP_UNCONFIRMED');
  assert.strictEqual(lines.length, 1);
  assert.deepStrictEqual(lines[0].affectedPoints, ['EP-01', 'EP-02']);
  assertConsistent(result);
});

check('document number is <SEQ>_<unit>_ESHA', () => {
  const payload = buildPayload();
  const report = normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant' } });
  assert.strictEqual(buildFinalReportData(report).project.documentNo, '024_FIPLANT_ESHA');
});

check('the golden report is consistent; unset planned quantities warn but do not block', () => {
  const result = checkConsistency(data);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.warnings.length, 5);
  assert.ok(result.warnings.some((warning) => warning.startsWith('Unit code not set')));
  assertConsistent(data);
});

check('a mismatched summary count blocks release', () => {
  const bad = fresh();
  bad.summaryTable.find((row) => row.testId === 'electrodeResistance').healthy += 1;
  assert.throws(() => assertConsistent(bad), (error) => error instanceof ConsistencyError && error.code === 'REPORT_INCONSISTENT' && /Electrode Resistance Test within limit/.test(error.message));
});

check('a mismatched priority block, scope count and CAP all block release', () => {
  const priority = fresh();
  priority.electrodePriority.P1.count = 4;
  assert.ok(checkConsistency(priority).problems.some((problem) => problem.startsWith('Electrode P1')));
  const scope = fresh();
  scope.scope[0].completed = 9;
  assert.ok(checkConsistency(scope).problems.some((problem) => problem.includes('completed')));
  const missingCap = fresh();
  missingCap.cap = missingCap.cap.filter((line) => line.point !== 'EP-07');
  assert.ok(checkConsistency(missingCap).problems.some((problem) => problem.includes('CAP defect lines')));
});

check('expired calibration becomes a CAP note', () => {
  const report = normalizeReportInput(buildPayload());
  report.calibrationRecords = [{ instrument: 'Metrel MI 3290 GX', serialNo: '18480898', calibrationDue: '2026-09-01', workDate: '2026-09-14' }];
  const result = buildFinalReportData(report);
  assert.ok(result.cap.some((line) => line.kind === 'calibration' && line.point.includes('2026-09-01') && line.affectedPoints.length === 1));
  assertConsistent(result);
});

check('project step fields: unit code is cleaned, planned quantities are kept and drive the scope table', () => {
  const payload = buildPayload();
  payload.project = { ...payload.project, unit: 'FI Plant', unitCode: ' f-i!x ', plannedQuantities: { electrodeResistance: '10', riserIntegrityTest: 6, bogus: 3, loopImpedanceTest: 'x' } };
  const report = normalizeReportInput(payload);
  assert.strictEqual(report.project.unitCode, 'FIX');
  assert.deepStrictEqual(report.project.plannedQuantities, { electrodeResistance: 10, riserIntegrityTest: 6 });
  const result = buildFinalReportData(report);
  assert.strictEqual(result.project.documentNo, '024_FIX_ESHA');
  assert.strictEqual(result.scope.find((row) => row.testId === 'electrodeResistance').planned, 10);
  assert.strictEqual(result.scope.find((row) => row.testId === 'loopImpedanceTest').planned, null);
  assert.ok(!result.warnings.some((warning) => warning.startsWith('Unit code')));
});

// ---- Session 6C.2b: report fixes, on the all-tests golden fixture --------------------------------
const all = buildFinalReportData(normalizeReportInput(buildAllTestsPayload()));
const allTest = (id) => all.tests.find((test) => test.id === id);
const allCap = (testId) => all.cap.filter((line) => line.testId === testId && line.kind === 'point');

check('electrode Status is the resistance result only; Priority carries the verdict', () => {
  const electrode = allTest('electrodeResistance');
  const header = electrode.columns.map((column) => column[0]);
  const status = header.indexOf('Status (R vs limit)');
  const priority = header.indexOf('Priority');
  const rowOf = (tag) => electrode.points.find((point) => point.label === tag).cells;
  assert.deepStrictEqual([rowOf('EP-01')[status], rowOf('EP-01')[priority]], ['Within limit', 'P3']);
  assert.deepStrictEqual([rowOf('EP-03')[status], rowOf('EP-03')[priority]], ['Within limit', 'P1']); // P1 by condition code C1 only
  assert.deepStrictEqual([rowOf('EP-07')[status], rowOf('EP-07')[priority]], ['Above limit', 'P1']);
  assert.deepStrictEqual([rowOf('EP-08')[status], rowOf('EP-08')[priority]], ['Not measurable', 'P1']);
  assert.deepStrictEqual([rowOf('EP-09')[status], rowOf('EP-09')[priority]], ['Warning', 'P2']);
  assert.ok(electrode.points.every((point) => !/Healthy|Permissible/.test(point.cells[status])));
});

check('electrode table shows the test condition and the 50/70% check', () => {
  const electrode = allTest('electrodeResistance');
  const header = electrode.columns.map((column) => column[0]);
  const condition = header.indexOf('Test condition');
  const fop = header.indexOf('50/70% check');
  const rowOf = (tag) => electrode.points.find((point) => point.label === tag).cells;
  assert.strictEqual(rowOf('EP-01')[condition], 'Isolated');
  assert.strictEqual(rowOf('EP-01')[fop], 'Agree');
  assert.strictEqual(rowOf('EP-06')[fop].text, 'Disagree');
  assert.deepStrictEqual([rowOf('EP-08')[condition], rowOf('EP-08')[fop]], ['-', '-']);
});

check('CAP clause follows the condition code; Cl. 14.2 only on resistance lines', () => {
  const clause = (tag) => allCap('electrodeResistance').find((line) => line.point === tag).clause;
  assert.strictEqual(clause('EP-03'), 'IS 3043:2018 Cl. 38.4.3'); // C1
  assert.strictEqual(clause('EP-08'), 'IS 3043:2018 Cl. 38.4.3'); // C4, not measurable
  assert.strictEqual(clause('EP-05'), 'ETPL criterion'); // C8
  assert.strictEqual(clause('EP-07'), 'IS 3043:2018 Cl. 14.2, 25.6.3.4'); // above limit + C2
  assert.strictEqual(clause('EP-09'), 'IS 3043:2018 Cl. 14.2'); // warning band, resistance only
  assert.ok(!allCap('continuityTest')[0].clause.includes('14.2'));
  assert.ok(!allCap('electrodeResistance').find((line) => line.point === 'EP-03').action.includes('14.2'));
});

check('CAP action text is not duplicated (EP-07)', () => {
  const action = allCap('electrodeResistance').find((line) => line.point === 'EP-07').action;
  assert.strictEqual((action.match(/test link and bonding/g) || []).length, 1);
  assert.strictEqual((action.match(/30 days/g) || []).length, 1);
});

check('loop and riser results have a Limit column and name the limit in the observation', () => {
  ['loopImpedanceTest', 'riserIntegrityTest'].forEach((id) => assert.ok(allTest(id).columns.some((column) => column[0] === 'Limit (ohm)')));
  assert.strictEqual(allTest('loopImpedanceTest').points[0].cells[allTest('loopImpedanceTest').columns.findIndex((column) => column[0] === 'Limit (ohm)')], '1.80');
  assert.ok(allCap('loopImpedanceTest')[0].observation.startsWith('Zs 1.92 ohm on B-E exceeds 1.8 ohm (ETPL criterion)'));
  assert.ok(allCap('riserIntegrityTest')[0].observation.startsWith('1.42 ohm towards equipment exceeds 1 ohm (ETPL criterion)'));
});

check('pit remarks say "x ohm = n% of the y ohm limit", never "within 90%"', () => {
  const pit = allTest('electrodeResistance').points.find((point) => point.label === 'EP-01');
  assert.ok(/Resistance 0\.82 ohm = 12% of the 6\.83 ohm limit/.test(pit.comment), pit.comment);
  allTest('electrodeResistance').points.forEach((point) => assert.ok(!/within 90%/.test(point.comment)));
});

check('instruments carry range and accuracy from the config catalogue', () => {
  assert.ok(all.instruments.length >= 4);
  assert.ok(all.instruments.every((row) => row.catalogued && row.range && row.accuracy));
  const names = all.instruments.map((row) => row.instrument).join(' | ');
  ['MI 3290', 'MI 3152', 'MI 3295', 'Kyoritsu 4118A'].forEach((model) => assert.ok(names.includes(model), `${model} in ${names}`));
  const catalogue = require('../src/config').instrumentCatalogue;
  assert.deepStrictEqual(catalogue.map((entry) => entry.match).sort(), ['kyoritsu4105a', 'kyoritsu4118a', 'mi3152', 'mi3290', 'mi3295']);
});

check('soil locations carry measured R beside rho, test conditions and the two chart series', () => {
  const location = all.soil.locations[0];
  assert.strictEqual(location.weather, 'Sunny');
  assert.deepStrictEqual(location.rows[0].rs, [9.8, 10.4]);
  assert.strictEqual(location.series.length, 2);
  assert.strictEqual(location.series[0].points.length, 6);
  assert.strictEqual(all.project.groundCondition, 'Dry');
});

check('executive findings: one bullet per test, no point repeated across bullets', () => {
  const findings = deterministicNarrative(all).keyFindings;
  const pfc = findings.filter((line) => line.startsWith('Prospective Fault Current'));
  assert.strictEqual(pfc.length, 1);
  assert.ok(pfc[0].includes('APDB-2 Incomer could not be assessed'));
  assert.strictEqual(findings.filter((line) => /could not be assessed/.test(line)).length, 1);
  data.summaryTable.forEach((row) => assert.ok(findings.filter((line) => line.startsWith(row.test)).length <= 1));
});

check('the all-tests fixture is consistent and has all nine tests', () => {
  assert.strictEqual(all.tests.length, 8); // eight pass/fail sheets + the soil section = MS-01 to MS-09
  assert.deepStrictEqual(all.tests.map((test) => test.sheetCode), ['MS-02', 'MS-03', 'MS-04', 'MS-05', 'MS-06', 'MS-07', 'MS-08', 'MS-09']);
  assert.ok(all.soil.locations.length === 2); // MS-01 is the soil section
  assertConsistent(all);
});

check('step & touch: the limit of each point uses its own surface; a note lists If, If/It, ts, rho_s per surface, Cs and body weight', () => {
  const step = allTest('stepTouchVoltage');
  const limitColumn = step.columns.findIndex((column) => column[0] === 'Limit (V)');
  const limits = Object.fromEntries(step.points.map((point) => [point.surface, point.cells[limitColumn]]));
  assert.strictEqual(new Set(Object.values(limits)).size, 3, JSON.stringify(limits));
  assert.notStrictEqual(limits.Gravel, limits.Concrete);
  assert.ok(step.note.startsWith('Basis of the limits:'));
  ['If = 20 kA (Client protection study)', 'If/It = 4000.00 (It = 5 A)', 'ts = 0.5 s (Client protection study)', 'Gravel: rho_s = 3000.00 ohm-m', 'Concrete: rho_s = 5000.00 ohm-m', 'Soil: rho_s = 25.68 ohm-m (site soil resistivity, no surface layer), Cs = 1.00', 'body weight 70 kg (k = 0.157)'].forEach((text) => assert.ok(step.note.includes(text), `note lacks "${text}": ${step.note}`));
});

check('tower footing: four feet under each tower (connection, I, Z), then the Zt / Zsat / status row', () => {
  const tower = allTest('towerFootingResistance');
  assert.deepStrictEqual(tower.columns.map((column) => column[0]), ['Sr', 'Tower / foot', 'Foot connection', 'I (mA)', 'Z (ohm)', 'Zt (ohm)', 'Zsat (ohm)', 'Status']);
  const point = tower.points[0];
  assert.strictEqual(point.subRows.length, 4);
  assert.deepStrictEqual(point.subRows.map((row) => [row[1].trim(), row[2], row[3], row[4]]), [['Foot-1', 'Given', '10.00', '2.20'], ['Foot-2', 'Given', '10.00', '2.40'], ['Foot-3', 'Given', '10.00', '2.60'], ['Foot-4', 'Given', '10.00', '2.80']]);
  assert.strictEqual(point.cells[1].text, 'Tower T1');
  assert.strictEqual(point.cells[5].text, '10.00');
  const notGiven = normalizeReportInput({ ...buildAllTestsPayload(), towerFootingResistance: [{ mainLocationTower: 'Tower T1', readings: [1, 2, 3, 4].map((n) => ({ footToEarthingConnectionStatus: n === 2 ? 'Not given' : 'Given', measuredCurrentMa: '10', measuredImpedance: '2', conditionCodes: [] })) }] });
  assert.strictEqual(buildFinalReportData(notGiven).tests.find((test) => test.id === 'towerFootingResistance').points[0].subRows[1][2].text, 'Not given');
});

check('continuity table has a Limit column', () => {
  const continuity = allTest('continuityTest');
  const column = continuity.columns.findIndex((entry) => entry[0] === 'Limit (ohm)');
  assert.ok(column > 0);
  assert.ok(continuity.points.every((point) => point.cells[column] === '1.00'));
});

check('key findings: singular and plural read correctly', () => {
  const findings = deterministicNarrative(all).keyFindings;
  assert.ok(findings.some((line) => line.startsWith('Tower Footing Resistance Measurement & Analysis: the 1 tower assessed is within limit')), findings.join(' | '));
  assert.ok(findings.some((line) => /: all 3 points assessed are within limit/.test(line)) || findings.some((line) => /all \d+ \w+ assessed are within limit/.test(line)));
  assert.ok(!findings.some((line) => /\b1 (towers|panels|points|pits)\b/.test(line)));
  const one = JSON.parse(JSON.stringify(all));
  one.summaryTable = one.summaryTable.filter((row) => row.testId === 'towerFootingResistance');
  one.tests = one.tests.filter((test) => test.id === 'towerFootingResistance');
  one.cap = [];
  one.electrodePriority = { P1: { count: 0, tags: [] }, P2: { count: 0, tags: [] }, P3: { count: 0, tags: [] } };
  one.tests[0].points[0].counted = true;
  one.tests[0].points[0].withinLimit = false;
  assert.ok(deterministicNarrative(one).keyFindings.some((line) => line.includes('the 1 tower assessed is above limit')));
});

console.log(`\n${passed} report-data checks passed.`);
