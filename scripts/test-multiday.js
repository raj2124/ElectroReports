// Verification for the multi-day data model (store + model). Run: node scripts/test-multiday.js
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const Database = require('better-sqlite3');
const { createReportStore } = require('../src/reportStore');
const {
  getElectrodeStatus, deriveElectrodeAssessment, summarizeOverallStatus, summarizePanels,
  deriveLoopImpedanceAssessment, buildOverallStatus, normalizeReportInput
} = require('../src/reportModel');

let passed = 0;
function check(name, fn) { fn(); passed += 1; console.log(`  ok - ${name}`); }

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'multiday-'));
const project = { projectNo: 'P1', clientName: 'Acme', siteLocation: 'Yard', workOrder: 'W1', engineerName: 'E' };

// --- Migration: v1 DB with a standalone row survives untouched -------------
const v1Path = path.join(dir, 'v1.db');
{
  const db = new Database(v1Path);
  db.exec(`CREATE TABLE reports (id TEXT PRIMARY KEY, created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
    project_no TEXT, client_name TEXT, site_location TEXT, work_order TEXT, engineer_name TEXT, data TEXT NOT NULL);`);
  db.prepare('INSERT INTO reports VALUES (?,?,?,?,?,?,?,?,?)')
    .run('ER-OLD', 't', 't', 'P0', 'Old', 'S', 'W', 'E', JSON.stringify({ id: 'ER-OLD', project: { projectNo: 'P0' } }));
  db.close();
}
check('v1 database migrates in place; legacy row reads as standalone', () => {
  const store = createReportStore({ filePath: v1Path });
  const old = store.getReport('ER-OLD');
  assert.deepStrictEqual(old, { id: 'ER-OLD', project: { projectNo: 'P0' } });
  const raw = new Database(v1Path).prepare("SELECT kind, project_id, review_status FROM reports WHERE id='ER-OLD'").get();
  assert.deepStrictEqual({ ...raw }, { kind: 'standalone', project_id: null, review_status: null });
  createReportStore({ filePath: v1Path }); // second open: idempotent
});

const store = createReportStore({ filePath: path.join(dir, 'new.db') });

check('standalone add/update/list unchanged (no kind, no review fields)', () => {
  const r = store.addReport({ project });
  assert.strictEqual(r.kind, undefined);
  assert.strictEqual(r.reviewStatus, undefined);
  const u = store.updateReport(r.id, { project: { ...project, clientName: 'Beta' } });
  assert.strictEqual(u.project.clientName, 'Beta');
  assert.strictEqual(store.listReports('beta').length, 1);
  assert.throws(() => store.transitionReport(r.id, 'in_review'), /Standalone/);
});

const prj = store.addProject(project);
const mkDaily = (date, rows, extra = {}) => store.addReport({
  kind: 'daily', projectId: prj.id, collectionDate: date, project, tests: { electrodeResistance: true },
  electrodeResistance: rows, ...extra
});

check('daily/final require an existing project', () => {
  assert.throws(() => store.addReport({ kind: 'daily', project }), /projectId/);
  assert.throws(() => store.addReport({ kind: 'daily', projectId: 'nope', project }), /projectId/);
});

const d1 = mkDaily('2026-07-10', [
  { rowId: 'a1', tag: 'E1', resistanceWithGrid: '9.0' },
  { rowId: 'a2', tag: 'E2', resistanceWithGrid: '2.0' }
]);
const d2 = mkDaily('2026-07-11', [{ rowId: 'b1', tag: 'e1', resistanceWithGrid: '3.0' }]);
const d3 = mkDaily('2026-07-12', [{ rowId: 'c1', tag: 'E3', resistanceWithGrid: '1.0' }]);

check('workflow: draft -> in_review -> checked -> approved, logged; illegal jumps rejected', () => {
  assert.strictEqual(d1.reviewStatus, 'draft');
  assert.throws(() => store.transitionReport(d1.id, 'approved'), /Invalid transition/);
  store.transitionReport(d1.id, 'in_review', { actor: 'eng', note: 'ready' });
  store.transitionReport(d1.id, 'draft', { actor: 'rev', note: 'redo' });
  store.transitionReport(d1.id, 'in_review', { actor: 'eng' });
  assert.throws(() => store.transitionReport(d1.id, 'approved', { actor: 'rev' }), (error) => error.code === 'NOT_CHECKED');
  assert.throws(() => store.markChecked(d1.id, { actor: 'chk' }), (error) => error.code === 'CHECKLIST_INCOMPLETE');
  store.setReviewTicks(d1.id, { counts: true });
  store.markChecked(d1.id, { actor: 'chk', note: 'checked' });
  assert.throws(() => store.markChecked(d1.id, { actor: 'chk' }), (error) => error.code === 'ALREADY_CHECKED');
  const a = store.transitionReport(d1.id, 'approved', { actor: 'rev' });
  assert.strictEqual(a.reviewStatus, 'approved');
  assert.deepStrictEqual(store.listTransitions(d1.id).map((t) => `${t.action}:${t.fromStatus}>${t.toStatus}`),
    ['submit:draft>in_review', 'return:in_review>draft', 'submit:draft>in_review', 'check:in_review>in_review', 'approve:in_review>approved']);
  assert.throws(() => store.transitionReport(d1.id, 'draft'), /Invalid transition/);
});

check('generic updateReport cannot change lifecycle fields', () => {
  const u = store.updateReport(d2.id, { reviewStatus: 'approved', kind: 'final', projectId: 'x' });
  assert.strictEqual(u.reviewStatus, 'draft');
  assert.strictEqual(u.kind, 'daily');
  assert.strictEqual(u.projectId, prj.id);
});

check('final assembles ALL dailies, lists dates, flags unapproved', () => {
  const { report } = store.assembleFinal(prj.id);
  assert.strictEqual(report.kind, 'final');
  assert.strictEqual(report.reviewStatus, 'draft');
  assert.deepStrictEqual(report.collectionManifest.includedDates, ['2026-07-10', '2026-07-11', '2026-07-12']);
  assert.deepStrictEqual(report.collectionManifest.flaggedDates, ['2026-07-11', '2026-07-12']);
  assert.strictEqual(report.collectionManifest.hasUnapproved, true);
  assert.deepStrictEqual(report.electrodeResistance.map((r) => r.rowId), ['b1', 'a2', 'c1']);
});

check('re-test: latest day wins; superseded reading kept in provenance', () => {
  const finals = store.listProjectReports(prj.id, { kind: 'final' });
  const prov = store.listAssemblyProvenance(finals[0].id);
  assert.strictEqual(prov.length, 1);
  assert.strictEqual(prov[0].resolution, 'latest_day');
  assert.strictEqual(prov[0].supersededRow.resistanceWithGrid, '9.0');
  assert.strictEqual(prov[0].supersededReportId, d1.id);
  assert.strictEqual(prov[0].keptReportId, d2.id);
});

check('engineer override keeps the chosen day; the other becomes provenance', () => {
  const { report, provenance } = store.assembleFinal(prj.id, {
    overrides: { 'electrodeResistance::e1': d1.id }, actor: 'lead'
  });
  assert.strictEqual(report.electrodeResistance.find((r) => r.tag.toLowerCase() === 'e1').rowId, 'a1');
  assert.strictEqual(provenance[0].resolution, 'engineer_override');
  assert.strictEqual(provenance[0].overriddenBy, 'lead');
  assert.strictEqual(provenance[0].supersededRow.rowId, 'b1');
});

check('final follows the same review workflow; delete cleans side tables', () => {
  const finals = store.listProjectReports(prj.id, { kind: 'final' });
  store.transitionReport(finals[0].id, 'in_review');
  assert.strictEqual(store.listTransitions(finals[0].id).length, 1);
  store.deleteReport(finals[0].id);
  assert.strictEqual(store.listTransitions(finals[0].id).length, 0);
  assert.strictEqual(store.listAssemblyProvenance(finals[0].id).length, 0);
});

// --- reportModel additions --------------------------------------------------
check('Not measurable electrode status is neutral and survives normalization', () => {
  assert.deepStrictEqual(getElectrodeStatus(9, { notMeasurable: true }), { label: 'Not Measurable', tone: 'neutral' });
  assert.strictEqual(getElectrodeStatus(9, { limit: 6 }).tone, 'critical'); // over the calculated limit
  assert.strictEqual(getElectrodeStatus(9).label, 'Pending'); // no calculated limit, so no verdict
  assert.strictEqual(deriveElectrodeAssessment({ notMeasurable: true }).status.label, 'Not Measurable');
  const n = normalizeReportInput({
    project: { ...project, reportDate: '2026-07-10' }, tests: { electrodeResistance: true },
    electrodeResistance: [{ tag: 'E1', notMeasurable: true }, { tag: 'E2', resistanceWithGrid: '2' }]
  });
  assert.strictEqual(n.electrodeResistance[0].notMeasurable, true);
  assert.strictEqual('notMeasurable' in n.electrodeResistance[1], false);
});

check('overall status thresholds', () => {
  const u = (healthy, bad) => [...Array(healthy).fill({ tone: 'healthy' }), ...Array(bad).fill({ tone: 'critical' })];
  assert.strictEqual(summarizeOverallStatus(u(10, 0)).label, 'Healthy');
  assert.strictEqual(summarizeOverallStatus(u(9, 1)).label, 'Healthy - Minor Actions'); // exactly 90%
  assert.strictEqual(summarizeOverallStatus(u(8, 2)).label, '>Attention Required');
  assert.strictEqual(summarizeOverallStatus([{ tone: 'neutral' }]).label, 'Pending');
  assert.strictEqual(summarizeOverallStatus([...u(10, 0), { tone: 'neutral' }]).label, 'Healthy'); // not-measurable excluded
});

check('panel healthy only if all 3 phases healthy', () => {
  const row = (z) => ({ loopImpedance: z });
  const ok = summarizePanels([row('0.5'), row('0.5'), row('0.5')], deriveLoopImpedanceAssessment);
  assert.strictEqual(ok[0].status.tone, 'healthy');
  const one = summarizePanels([row('0.5'), row('0.5'), row('2.0')], deriveLoopImpedanceAssessment);
  assert.strictEqual(one[0].status.tone, 'critical'); // one phase over the 1.8 ohm Zs limit fails the panel
  const bad = summarizePanels([row('0.5'), row('0.5'), row('3')], deriveLoopImpedanceAssessment);
  assert.strictEqual(bad[0].status.tone, 'critical');
  const short = summarizePanels([row('0.5'), row('0.5')], deriveLoopImpedanceAssessment);
  assert.strictEqual(short[0].status.tone, 'warning'); // incomplete panel is not healthy
  assert.strictEqual(buildOverallStatus({ tests: { loopImpedanceTest: true }, loopImpedanceTest: [row('0.5'), row('0.5'), row('2.0')] }).label, '>Attention Required');
});

console.log(`\n${passed} checks passed.`);
