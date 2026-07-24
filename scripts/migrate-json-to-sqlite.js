#!/usr/bin/env node
'use strict';

// One-time migration: import data/reports.json into the SQLite store.
//
// Safe to run against an empty ([]) or missing JSON file, and safe to run more
// than once: existing rows are preserved via INSERT OR IGNORE keyed on the report
// id. Original id / createdAt / updatedAt are preserved exactly (NOT regenerated),
// so migrated records keep their identity and history.
//
// This script only WRITES to the SQLite DB. It never modifies or deletes
// data/reports.json — that stays as the fallback until you confirm the switchover.
//
// Usage:
//   node scripts/migrate-json-to-sqlite.js               # uses config paths
//   node scripts/migrate-json-to-sqlite.js <json> <db>   # explicit override

const fs = require('fs');
const path = require('path');
const {
  openDatabase,
  resolveDbPath,
  extractSearchColumns
} = require('../src/reportStore');

function loadConfigPaths() {
  try {
    const config = require('../src/config');
    return {
      jsonPath: config.app.dataFilePath,
      dbPath: resolveDbPath(config.app.dataFilePath)
    };
  } catch (error) {
    const fallbackJson = path.resolve(__dirname, '..', 'data', 'reports.json');
    return { jsonPath: fallbackJson, dbPath: resolveDbPath(fallbackJson) };
  }
}

function readJsonReports(jsonPath) {
  if (!fs.existsSync(jsonPath)) {
    console.log(`[migrate] No JSON file at ${jsonPath} — nothing to import.`);
    return [];
  }
  const raw = fs.readFileSync(jsonPath, 'utf8');
  if (!raw.trim()) {
    console.log('[migrate] JSON file is empty — nothing to import.');
    return [];
  }
  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error(`[migrate] Could not parse ${jsonPath}: ${error.message}`);
  }
  if (!Array.isArray(parsed)) {
    throw new Error(`[migrate] Expected a JSON array in ${jsonPath}.`);
  }
  return parsed;
}

function main() {
  const [, , jsonArg, dbArg] = process.argv;
  const configPaths = loadConfigPaths();
  const jsonPath = jsonArg ? path.resolve(jsonArg) : configPaths.jsonPath;
  const dbPath = dbArg ? resolveDbPath(path.resolve(dbArg)) : configPaths.dbPath;

  console.log(`[migrate] Source JSON: ${jsonPath}`);
  console.log(`[migrate] Target DB:   ${dbPath}`);

  const reports = readJsonReports(jsonPath);
  const db = openDatabase(dbPath);

  const before = db.prepare('SELECT COUNT(*) AS n FROM reports').get().n;

  const insert = db.prepare(`
    INSERT OR IGNORE INTO reports
      (id, created_at, updated_at, project_no, client_name, site_location, work_order, engineer_name, data)
    VALUES
      (@id, @created_at, @updated_at, @project_no, @client_name, @site_location, @work_order, @engineer_name, @data)
  `);

  const nowIso = new Date().toISOString();
  let imported = 0;
  let skippedNoId = 0;

  const runAll = db.transaction((records) => {
    for (const report of records) {
      if (!report || report.id == null) {
        skippedNoId += 1;
        continue;
      }
      const createdAt = report.createdAt || nowIso;
      const updatedAt = report.updatedAt || createdAt;
      // Persist the record with its original timestamps embedded in the payload too,
      // so the stored object shape matches what the store returns.
      const normalized = { ...report, createdAt, updatedAt };
      const result = insert.run({
        id: String(report.id),
        created_at: createdAt,
        updated_at: updatedAt,
        ...extractSearchColumns(normalized),
        data: JSON.stringify(normalized)
      });
      imported += result.changes;
    }
  });

  runAll(reports);

  const after = db.prepare('SELECT COUNT(*) AS n FROM reports').get().n;

  console.log(`[migrate] Records in JSON:        ${reports.length}`);
  console.log(`[migrate] Rows before:            ${before}`);
  console.log(`[migrate] Newly imported:         ${imported}`);
  console.log(`[migrate] Skipped (no id):        ${skippedNoId}`);
  console.log(`[migrate] Skipped (already there): ${reports.length - imported - skippedNoId}`);
  console.log(`[migrate] Rows after:             ${after}`);
  console.log('[migrate] Done. data/reports.json left untouched as fallback.');
}

main();
