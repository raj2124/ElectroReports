// Standalone verification for the multi-day collection model (step 1).
// Run: node scripts/test-report-collection.js
const assert = require('assert');
const {
  REPORT_KINDS,
  DAILY_STATUSES,
  getReportKind,
  isStandaloneReport,
  getDailyStatus,
  getCollectionDate,
  summarizeDailyCollection,
  assembleFinalReport
} = require('../src/reportCollectionModel');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

// --- Standalone-preservation defaults --------------------------------------
check('report with no kind reads as standalone', () => {
  assert.strictEqual(getReportKind({ project: {} }), REPORT_KINDS.STANDALONE);
  assert.strictEqual(isStandaloneReport({}), true);
});

check('unknown kind falls back to standalone', () => {
  assert.strictEqual(getReportKind({ kind: 'weekly' }), REPORT_KINDS.STANDALONE);
});

check('missing/unknown status reads as draft', () => {
  assert.strictEqual(getDailyStatus({}), DAILY_STATUSES.DRAFT);
  assert.strictEqual(getDailyStatus({ dailyStatus: 'nonsense' }), DAILY_STATUSES.DRAFT);
});

check('collectionDate falls back to project.reportDate', () => {
  assert.strictEqual(getCollectionDate({ project: { reportDate: '2026-07-10' } }), '2026-07-10');
  assert.strictEqual(getCollectionDate({ collectionDate: '2026-07-12', project: { reportDate: '2026-07-10' } }), '2026-07-12');
});

// --- Locked decision: ALL days included, unapproved flagged ----------------
const daily = [
  { id: 'D2', kind: 'daily', dailyStatus: 'approved', collectionDate: '2026-07-11' },
  { id: 'D1', kind: 'daily', dailyStatus: 'draft', collectionDate: '2026-07-10' },
  { id: 'D3', kind: 'daily', dailyStatus: 'in_review', collectionDate: '2026-07-12', project: { clientName: 'Acme' } }
];

check('manifest includes every daily report, sorted by date', () => {
  const m = summarizeDailyCollection(daily);
  assert.strictEqual(m.totalDailyReports, 3);
  assert.deepStrictEqual(m.entries.map((e) => e.reportId), ['D1', 'D2', 'D3']);
  assert.deepStrictEqual(m.includedDates, ['2026-07-10', '2026-07-11', '2026-07-12']);
});

check('unapproved days are flagged; approved is not', () => {
  const m = summarizeDailyCollection(daily);
  assert.strictEqual(m.hasUnapproved, true);
  assert.deepStrictEqual(m.flaggedDates, ['2026-07-10', '2026-07-12']);
  assert.deepStrictEqual(m.statusCounts, { draft: 1, in_review: 1, approved: 1 });
});

check('all-approved collection is not flagged', () => {
  const m = summarizeDailyCollection([
    { id: 'A', kind: 'daily', dailyStatus: 'approved', collectionDate: '2026-07-10' }
  ]);
  assert.strictEqual(m.hasUnapproved, false);
  assert.deepStrictEqual(m.flaggedDates, []);
});

// --- Assembly skeleton ------------------------------------------------------
check('assembleFinalReport unions tests and carries latest project', () => {
  const withTests = [
    { id: 'D1', kind: 'daily', collectionDate: '2026-07-10', tests: { soilResistivity: true, continuityTest: false }, project: { clientName: 'Old' } },
    { id: 'D2', kind: 'daily', collectionDate: '2026-07-11', tests: { continuityTest: true }, project: { clientName: 'New' } }
  ];
  const final = assembleFinalReport(withTests);
  assert.strictEqual(final.kind, REPORT_KINDS.FINAL);
  assert.strictEqual(final.tests.soilResistivity, true);
  assert.strictEqual(final.tests.continuityTest, true);
  assert.strictEqual(final.project.clientName, 'New'); // latest wins
  assert.deepStrictEqual(final.constituentReportIds, ['D1', 'D2']);
  assert.strictEqual(final.collectionManifest.totalDailyReports, 2);
});

check('empty collection assembles safely', () => {
  const final = assembleFinalReport([]);
  assert.strictEqual(final.collectionManifest.totalDailyReports, 0);
  assert.deepStrictEqual(final.constituentReportIds, []);
});

console.log(`\n${passed} checks passed.`);
