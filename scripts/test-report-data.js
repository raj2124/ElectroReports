// Final-report computed results and the consistency check (Session 6C).
// Uses the make_report.py sample data, so the counts match test:esha.
const assert = require('assert');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
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
  assert.ok(defects('loopImpedanceTest')[0].observation.includes('B-E (1.92)'));
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

console.log(`\n${passed} report-data checks passed.`);
