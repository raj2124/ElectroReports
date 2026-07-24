#!/usr/bin/env node
'use strict';

// Proves the SQLite store:
//   (A) supports the full create -> get -> list -> delete lifecycle behind the
//       same interface as the flat-file store, and
//   (B) survives two rapid CONCURRENT writers from separate processes — the exact
//       case the flat-file store loses. For contrast, we first reproduce the
//       flat-file lost-update, then show SQLite keeps both writes.
//
// Single file, self-forking: run normally to drive the suite; it re-invokes itself
// with `--worker <mode> <path>` for the concurrent-writer child processes.

const fs = require('fs');
const os = require('os');
const path = require('path');
const { fork } = require('child_process');

const { createReportStore } = require('../src/reportStore');

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(`ASSERTION FAILED: ${message}`);
  }
  console.log(`  ok - ${message}`);
}

function sampleReport(label) {
  return {
    project: {
      projectNo: `PRJ-${label}`,
      clientName: `Client ${label}`,
      siteLocation: `Site ${label}`,
      workOrder: `WO-${label}`,
      engineerName: `Engineer ${label}`
    },
    measurements: { earthResistance: 1.23, note: `payload-${label}` }
  };
}

// ---------------------------------------------------------------------------
// Worker branch: performs a single write, then exits. Two of these run at once.
// ---------------------------------------------------------------------------
async function runWorker(mode, targetPath, label) {
  if (mode === 'flatfile') {
    // Faithful reproduction of the flat-file store's non-atomic read-modify-write.
    // The sleep only widens the naturally-existing race window so the lost-update
    // is deterministic in a test rather than timing-dependent.
    const raw = fs.existsSync(targetPath) ? fs.readFileSync(targetPath, 'utf8') : '[]';
    const records = raw.trim() ? JSON.parse(raw) : [];
    await sleep(250);
    records.unshift({ id: `flat-${label}`, label });
    fs.writeFileSync(targetPath, JSON.stringify(records, null, 2), 'utf8');
    return;
  }
  if (mode === 'sqlite') {
    const store = createReportStore({ filePath: targetPath });
    await sleep(250);
    store.addReport(sampleReport(label));
    return;
  }
  throw new Error(`Unknown worker mode: ${mode}`);
}

function forkWorker(mode, targetPath, label) {
  return new Promise((resolve, reject) => {
    const child = fork(__filename, ['--worker', mode, targetPath, label], {
      stdio: 'ignore'
    });
    child.on('exit', (code) =>
      code === 0 ? resolve() : reject(new Error(`worker ${label} exited with code ${code}`))
    );
    child.on('error', reject);
  });
}

// ---------------------------------------------------------------------------
// Main suite.
// ---------------------------------------------------------------------------
async function main() {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'electroreports-sqlite-'));
  const storeFile = path.join(tempDir, 'reports.json'); // passed as a .json path on purpose

  console.log('\n[1] Lifecycle: create -> get -> list -> delete (same interface)');
  const store = createReportStore({ filePath: storeFile });

  assert(store.listReports().length === 0, 'store starts empty');

  const created = store.addReport(sampleReport('A'));
  assert(/^ER-\d{14}-\d{4}$/.test(created.id), `id matches ER-format (${created.id})`);
  assert(!!created.createdAt && !!created.updatedAt, 'createdAt/updatedAt set on create');

  const fetched = store.getReport(created.id);
  assert(fetched && fetched.id === created.id, 'getReport returns the created report');
  assert(
    fetched.measurements.note === 'payload-A',
    'nested payload round-trips unchanged through JSON column'
  );
  assert(store.getReport('does-not-exist') === null, 'getReport returns null when absent');

  store.addReport(sampleReport('B'));
  assert(store.listReports().length === 2, 'listReports returns both reports');
  assert(store.listReports('client a').length === 1, 'listReports text search filters (case-insensitive)');
  assert(store.listReports('Site B')[0].project.clientName === 'Client B', 'search matches site field');
  assert(store.listReports('nomatch').length === 0, 'search with no match returns empty');

  // updateReport: function form (as server.js uses) and preservation rules.
  await sleep(5);
  const updated = store.updateReport(created.id, (current) => ({
    ...current,
    aiNarrative: 'generated text'
  }));
  assert(updated.aiNarrative === 'generated text', 'updateReport (function form) applies changes');
  assert(updated.id === created.id, 'updateReport preserves id');
  assert(updated.createdAt === created.createdAt, 'updateReport preserves createdAt');
  assert(updated.updatedAt !== created.updatedAt, 'updateReport bumps updatedAt');
  assert(store.updateReport('missing', {}) === null, 'updateReport returns null when absent');

  const del = store.deleteReport(created.id);
  assert(del.deleted === true && del.report.id === created.id, 'deleteReport returns removed report');
  assert(store.getReport(created.id) === null, 'deleted report is gone');
  assert(store.deleteReport('missing').deleted === false, 'deleteReport reports not-found');
  assert(store.listReports().length === 1, 'list reflects deletion');

  console.log('\n[2] Concurrency — flat file (reproduce the failure)');
  const flatPath = path.join(tempDir, 'flat.json');
  fs.writeFileSync(flatPath, '[]', 'utf8');
  await Promise.all([
    forkWorker('flatfile', flatPath, 'X'),
    forkWorker('flatfile', flatPath, 'Y')
  ]);
  const flatCount = JSON.parse(fs.readFileSync(flatPath, 'utf8')).length;
  console.log(`  flat file rows after 2 concurrent writes: ${flatCount} (expected 1 — LOST UPDATE)`);
  assert(flatCount === 1, 'flat file LOSES one write (this is the bug we are fixing)');

  console.log('\n[3] Concurrency — SQLite (both writes persist)');
  const sqlitePath = path.join(tempDir, 'concurrent.json'); // -> concurrent.db
  createReportStore({ filePath: sqlitePath }); // create schema up front
  await Promise.all([
    forkWorker('sqlite', sqlitePath, 'X'),
    forkWorker('sqlite', sqlitePath, 'Y')
  ]);
  const verifyStore = createReportStore({ filePath: sqlitePath });
  const sqliteCount = verifyStore.listReports().length;
  console.log(`  sqlite rows after 2 concurrent writes: ${sqliteCount} (expected 2)`);
  assert(sqliteCount === 2, 'SQLite persists BOTH concurrent writes');
  const labels = verifyStore.listReports().map((r) => r.project.projectNo).sort();
  assert(labels.join(',') === 'PRJ-X,PRJ-Y', `both distinct records present (${labels.join(', ')})`);

  fs.rmSync(tempDir, { recursive: true, force: true });
  console.log('\n[sqlite-store] ALL TESTS PASS\n');
}

// Entry point: worker vs. driver.
const workerIdx = process.argv.indexOf('--worker');
if (workerIdx !== -1) {
  const [mode, targetPath, label] = process.argv.slice(workerIdx + 1);
  runWorker(mode, targetPath, label).then(
    () => process.exit(0),
    (err) => {
      console.error(err);
      process.exit(1);
    }
  );
} else {
  main().catch((err) => {
    console.error('\n[sqlite-store] TEST FAILURE');
    console.error(err);
    process.exit(1);
  });
}
