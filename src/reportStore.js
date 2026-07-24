const fs = require('fs');
const path = require('path');
const Database = require('better-sqlite3');

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

// Keep the constructor signature identical to the previous flat-file store.
// Callers pass `{ filePath: '.../data/reports.json' }`; we transparently map that
// to a sibling `.../data/reports.db`. An explicit `.db`/`.sqlite` path is used
// as-is. This lets server.js and the verify script keep their call unchanged.
function resolveDbPath(filePath) {
  if (!filePath) {
    return path.join(process.cwd(), 'data', 'reports.db');
  }
  if (/\.(db|sqlite|sqlite3)$/i.test(filePath)) {
    return filePath;
  }
  const dir = path.dirname(filePath);
  const base = path.basename(filePath).replace(/\.[^.]+$/, '');
  return path.join(dir, `${base}.db`);
}

// Open (or create) the database and apply the schema + concurrency pragmas.
// WAL lets multiple processes read/write safely; busy_timeout makes a writer wait
// briefly for a lock instead of throwing SQLITE_BUSY. Exported so the migration
// script and tests share exactly one schema definition.
function openDatabase(dbPath) {
  ensureDir(path.dirname(dbPath));
  const db = new Database(dbPath);
  db.pragma('journal_mode = WAL');
  db.pragma('synchronous = NORMAL');
  db.pragma('busy_timeout = 5000');
  db.exec(`
    CREATE TABLE IF NOT EXISTS reports (
      id            TEXT PRIMARY KEY,
      created_at    TEXT NOT NULL,
      updated_at    TEXT NOT NULL,
      project_no    TEXT,
      client_name   TEXT,
      site_location TEXT,
      work_order    TEXT,
      engineer_name TEXT,
      data          TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_reports_updated_at ON reports(updated_at DESC);
  `);
  return db;
}

// The mirror columns used for list sorting/searching. Everything else lives in `data`.
function extractSearchColumns(report) {
  const project = (report && report.project) || {};
  return {
    project_no: project.projectNo != null ? String(project.projectNo) : null,
    client_name: project.clientName != null ? String(project.clientName) : null,
    site_location: project.siteLocation != null ? String(project.siteLocation) : null,
    work_order: project.workOrder != null ? String(project.workOrder) : null,
    engineer_name: project.engineerName != null ? String(project.engineerName) : null
  };
}

function buildId() {
  const timestamp = new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
  const suffix = Math.floor(1000 + Math.random() * 9000);
  return `ER-${timestamp}-${suffix}`;
}

function createReportStore({ filePath }) {
  const dbPath = resolveDbPath(filePath);
  const db = openDatabase(dbPath);

  const insertStmt = db.prepare(`
    INSERT INTO reports
      (id, created_at, updated_at, project_no, client_name, site_location, work_order, engineer_name, data)
    VALUES
      (@id, @created_at, @updated_at, @project_no, @client_name, @site_location, @work_order, @engineer_name, @data)
  `);
  const updateStmt = db.prepare(`
    UPDATE reports SET
      updated_at = @updated_at,
      project_no = @project_no,
      client_name = @client_name,
      site_location = @site_location,
      work_order = @work_order,
      engineer_name = @engineer_name,
      data = @data
    WHERE id = @id
  `);
  const selectByIdStmt = db.prepare('SELECT data FROM reports WHERE id = ?');
  const deleteStmt = db.prepare('DELETE FROM reports WHERE id = ?');
  const selectAllStmt = db.prepare(
    'SELECT data FROM reports ORDER BY COALESCE(updated_at, created_at) DESC'
  );
  // Mirror the flat-file filter: join the 5 project fields with a space, lowercase,
  // and substring-match. COALESCE keeps NULLs from breaking the concatenation.
  const searchStmt = db.prepare(`
    SELECT data FROM reports
    WHERE lower(
      COALESCE(project_no, '') || ' ' ||
      COALESCE(client_name, '') || ' ' ||
      COALESCE(site_location, '') || ' ' ||
      COALESCE(work_order, '') || ' ' ||
      COALESCE(engineer_name, '')
    ) LIKE @needle ESCAPE '\\'
    ORDER BY COALESCE(updated_at, created_at) DESC
  `);

  function rowToReport(row) {
    return row ? JSON.parse(row.data) : null;
  }

  function rowsToReports(rows) {
    return rows.map((row) => JSON.parse(row.data));
  }

  function escapeLike(text) {
    return text.replace(/[\\%_]/g, (ch) => `\\${ch}`);
  }

  function listReports(query = '') {
    const text = String(query || '').trim().toLowerCase();
    if (!text) {
      return rowsToReports(selectAllStmt.all());
    }
    return rowsToReports(searchStmt.all({ needle: `%${escapeLike(text)}%` }));
  }

  function getReport(id) {
    return rowToReport(selectByIdStmt.get(String(id)));
  }

  function addReport(reportInput) {
    const now = new Date().toISOString();
    const report = {
      ...reportInput,
      id: buildId(),
      createdAt: now,
      updatedAt: now
    };
    insertStmt.run({
      id: report.id,
      created_at: report.createdAt,
      updated_at: report.updatedAt,
      ...extractSearchColumns(report),
      data: JSON.stringify(report)
    });
    return report;
  }

  function updateReport(id, updater) {
    const current = getReport(id);
    if (!current) {
      return null;
    }
    const nextValue =
      typeof updater === 'function'
        ? updater(current)
        : {
            ...current,
            ...(updater || {})
          };
    const updated = {
      ...current,
      ...(nextValue || {}),
      id: current.id,
      createdAt: current.createdAt,
      updatedAt: new Date().toISOString()
    };
    updateStmt.run({
      id: updated.id,
      updated_at: updated.updatedAt,
      ...extractSearchColumns(updated),
      data: JSON.stringify(updated)
    });
    return updated;
  }

  function deleteReport(id) {
    const target = getReport(id);
    if (!target) {
      return { deleted: false, report: null };
    }
    deleteStmt.run(String(id));
    return { deleted: true, report: target };
  }

  return {
    listReports,
    getReport,
    addReport,
    updateReport,
    deleteReport
  };
}

module.exports = {
  createReportStore,
  // Exported for the migration script and tests (shared schema, no duplication).
  openDatabase,
  resolveDbPath,
  extractSearchColumns
};
