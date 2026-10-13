const fs = require('fs');
const path = require('path');
const Database = require('better-sqlite3');
const {
  REPORT_KINDS,
  getReportKind,
  getReviewStatus,
  getCollectionDate,
  canTransition,
  assembleFinalReport
} = require('./reportCollectionModel');

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
  migrateMultiDaySchema(db);
  return db;
}

function hasColumn(db, table, column) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((col) => col.name === column);
}

// Multi-day data model (ADDITIVE). Idempotent: safe on fresh, v1 and already-migrated
// databases. Existing `reports` rows are untouched; they read as kind='standalone'
// through the column default, so the standalone path behaves exactly as before.
function migrateMultiDaySchema(db) {
  const apply = db.transaction(() => {
    db.exec(`
      CREATE TABLE IF NOT EXISTS projects (
        id            TEXT PRIMARY KEY,
        project_no    TEXT,
        client_name   TEXT,
        site_location TEXT,
        work_order    TEXT,
        engineer_name TEXT,
        created_at    TEXT NOT NULL,
        updated_at    TEXT NOT NULL,
        data          TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS report_transitions (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        report_id   TEXT NOT NULL,
        from_status TEXT,
        to_status   TEXT NOT NULL,
        actor       TEXT,
        note        TEXT,
        at          TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_transitions_report ON report_transitions(report_id, id);
      CREATE TABLE IF NOT EXISTS assembly_provenance (
        id                         INTEGER PRIMARY KEY AUTOINCREMENT,
        final_report_id            TEXT NOT NULL,
        sheet                      TEXT NOT NULL,
        point_key                  TEXT NOT NULL,
        kept_report_id             TEXT,
        kept_collection_date       TEXT,
        superseded_report_id       TEXT NOT NULL,
        superseded_collection_date TEXT,
        superseded_row             TEXT NOT NULL,
        resolution                 TEXT NOT NULL,
        overridden_by              TEXT,
        created_at                 TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_provenance_final ON assembly_provenance(final_report_id);
    `);
    [
      ['project_id', 'TEXT'],
      ['kind', "TEXT NOT NULL DEFAULT 'standalone'"],
      ['collection_date', 'TEXT'],
      ['review_status', 'TEXT']
    ].forEach(([column, type]) => {
      if (!hasColumn(db, 'reports', column)) {
        db.exec(`ALTER TABLE reports ADD COLUMN ${column} ${type}`);
      }
    });
    db.exec('CREATE INDEX IF NOT EXISTS idx_reports_project ON reports(project_id, kind, collection_date)');
    db.pragma('user_version = 2');
  });
  // IMMEDIATE takes the write lock before the column checks, so two processes opening
  // a fresh database serialise (busy_timeout applies) instead of both ALTERing.
  apply.immediate();
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

// Multi-day mirror columns. A report with no `kind` is standalone and carries no
// project link or review status, which keeps legacy rows byte-for-byte equivalent.
function extractCollectionColumns(report) {
  const kind = getReportKind(report);
  const standalone = kind === REPORT_KINDS.STANDALONE;
  return {
    project_id: standalone || !report.projectId ? null : String(report.projectId),
    kind,
    collection_date: kind === REPORT_KINDS.DAILY ? getCollectionDate(report) || null : null,
    review_status: standalone ? null : getReviewStatus(report)
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
      (id, created_at, updated_at, project_no, client_name, site_location, work_order, engineer_name,
       project_id, kind, collection_date, review_status, data)
    VALUES
      (@id, @created_at, @updated_at, @project_no, @client_name, @site_location, @work_order, @engineer_name,
       @project_id, @kind, @collection_date, @review_status, @data)
  `);
  const updateStmt = db.prepare(`
    UPDATE reports SET
      updated_at = @updated_at,
      project_no = @project_no,
      client_name = @client_name,
      site_location = @site_location,
      work_order = @work_order,
      engineer_name = @engineer_name,
      project_id = @project_id,
      kind = @kind,
      collection_date = @collection_date,
      review_status = @review_status,
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
    const kind = getReportKind(reportInput);
    if (kind !== REPORT_KINDS.STANDALONE) {
      if (!reportInput.projectId || !getProject(reportInput.projectId)) {
        throw new Error(`A ${kind} report requires an existing projectId.`);
      }
    }
    const report = {
      ...reportInput,
      ...(kind === REPORT_KINDS.STANDALONE ? {} : { kind, reviewStatus: getReviewStatus(reportInput) }),
      id: buildId(),
      createdAt: now,
      updatedAt: now
    };
    insertStmt.run({
      id: report.id,
      created_at: report.createdAt,
      updated_at: report.updatedAt,
      ...extractSearchColumns(report),
      ...extractCollectionColumns(report),
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
    // kind / project link / review status are lifecycle fields: only addReport and
    // transitionReport may set them, never a generic update.
    ['kind', 'projectId', 'reviewStatus', 'dailyStatus'].forEach((field) => {
      if (current[field] === undefined) {
        delete updated[field];
      } else {
        updated[field] = current[field];
      }
    });
    updateStmt.run({
      id: updated.id,
      updated_at: updated.updatedAt,
      ...extractSearchColumns(updated),
      ...extractCollectionColumns(updated),
      data: JSON.stringify(updated)
    });
    return updated;
  }

  function deleteReport(id) {
    const target = getReport(id);
    if (!target) {
      return { deleted: false, report: null };
    }
    db.transaction(() => {
      deleteStmt.run(String(id));
      db.prepare('DELETE FROM report_transitions WHERE report_id = ?').run(String(id));
      db.prepare('DELETE FROM assembly_provenance WHERE final_report_id = ?').run(String(id));
    })();
    return { deleted: true, report: target };
  }


  // ---- Projects (multi-day) -------------------------------------------------
  function buildProjectId() {
    return buildId().replace(/^ER-/, 'PRJ-');
  }

  function getProject(id) {
    const row = db.prepare('SELECT data FROM projects WHERE id = ?').get(String(id));
    return row ? JSON.parse(row.data) : null;
  }

  function listProjects() {
    return db
      .prepare('SELECT data FROM projects ORDER BY updated_at DESC')
      .all()
      .map((row) => JSON.parse(row.data));
  }

  function writeProject(project, isNew) {
    const columns = extractSearchColumns({ project });
    const params = {
      id: project.id,
      created_at: project.createdAt,
      updated_at: project.updatedAt,
      ...columns,
      data: JSON.stringify(project)
    };
    if (isNew) {
      db.prepare(`INSERT INTO projects (id, project_no, client_name, site_location, work_order, engineer_name, created_at, updated_at, data)
        VALUES (@id, @project_no, @client_name, @site_location, @work_order, @engineer_name, @created_at, @updated_at, @data)`).run(params);
    } else {
      db.prepare(`UPDATE projects SET project_no=@project_no, client_name=@client_name, site_location=@site_location,
        work_order=@work_order, engineer_name=@engineer_name, updated_at=@updated_at, data=@data WHERE id=@id`).run(params);
    }
    return project;
  }

  // `projectInput` is a plain project-details object (same fields as report.project).
  function addProject(projectInput) {
    const now = new Date().toISOString();
    return writeProject({ ...projectInput, id: buildProjectId(), createdAt: now, updatedAt: now }, true);
  }

  function updateProject(id, patch) {
    const current = getProject(id);
    if (!current) {
      return null;
    }
    return writeProject(
      { ...current, ...(patch || {}), id: current.id, createdAt: current.createdAt, updatedAt: new Date().toISOString() },
      false
    );
  }

  function listProjectReports(projectId, { kind } = {}) {
    const rows = kind
      ? db
          .prepare('SELECT data FROM reports WHERE project_id = ? AND kind = ? ORDER BY COALESCE(collection_date, created_at), created_at')
          .all(String(projectId), kind)
      : db
          .prepare('SELECT data FROM reports WHERE project_id = ? ORDER BY kind, COALESCE(collection_date, created_at), created_at')
          .all(String(projectId));
    return rowsToReports(rows);
  }

  // ---- Review workflow: draft -> in_review -> approved (daily AND final) ------
  function transitionReport(id, toStatus, { actor = '', note = '' } = {}) {
    const current = getReport(id);
    if (!current) {
      return null;
    }
    if (getReportKind(current) === REPORT_KINDS.STANDALONE) {
      throw new Error('Standalone reports have no review workflow.');
    }
    const fromStatus = getReviewStatus(current);
    if (!canTransition(fromStatus, toStatus)) {
      throw new Error(`Invalid transition ${fromStatus} -> ${toStatus}.`);
    }
    const at = new Date().toISOString();
    const updated = { ...current, reviewStatus: toStatus, updatedAt: at };
    db.transaction(() => {
      updateStmt.run({
        id: updated.id,
        updated_at: at,
        ...extractSearchColumns(updated),
        ...extractCollectionColumns(updated),
        data: JSON.stringify(updated)
      });
      db.prepare(
        'INSERT INTO report_transitions (report_id, from_status, to_status, actor, note, at) VALUES (?, ?, ?, ?, ?, ?)'
      ).run(updated.id, fromStatus, toStatus, String(actor || ''), String(note || ''), at);
    })();
    return updated;
  }

  function listTransitions(reportId) {
    return db
      .prepare('SELECT id, report_id AS reportId, from_status AS fromStatus, to_status AS toStatus, actor, note, at FROM report_transitions WHERE report_id = ? ORDER BY id')
      .all(String(reportId));
  }

  // ---- Final assembly -------------------------------------------------------
  function listAssemblyProvenance(finalReportId) {
    return db
      .prepare('SELECT * FROM assembly_provenance WHERE final_report_id = ? ORDER BY id')
      .all(String(finalReportId))
      .map((row) => ({
        id: row.id,
        finalReportId: row.final_report_id,
        sheet: row.sheet,
        pointKey: row.point_key,
        keptReportId: row.kept_report_id,
        keptCollectionDate: row.kept_collection_date,
        supersededReportId: row.superseded_report_id,
        supersededCollectionDate: row.superseded_collection_date,
        supersededRow: JSON.parse(row.superseded_row),
        resolution: row.resolution,
        overriddenBy: row.overridden_by,
        createdAt: row.created_at
      }));
  }

  // Stitch ALL daily reports of a project into one final report (draft). Superseded
  // re-test readings are persisted to assembly_provenance, never discarded.
  // options.overrides: { 'sheet::pointKey': dailyReportId } (engineer override).
  function assembleFinal(projectId, { overrides = {}, actor = '' } = {}) {
    const project = getProject(projectId);
    if (!project) {
      throw new Error('Project not found.');
    }
    const dailies = listProjectReports(projectId, { kind: REPORT_KINDS.DAILY });
    if (!dailies.length) {
      throw new Error('Project has no daily reports to assemble.');
    }
    const assembled = assembleFinalReport(dailies, { overrides, overriddenBy: actor });
    const { assemblyProvenance, ...reportBody } = assembled;
    const now = new Date().toISOString();
    return db.transaction(() => {
      const final = addReport({ ...reportBody, projectId: project.id });
      const insert = db.prepare(`INSERT INTO assembly_provenance
        (final_report_id, sheet, point_key, kept_report_id, kept_collection_date, superseded_report_id,
         superseded_collection_date, superseded_row, resolution, overridden_by, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);
      assemblyProvenance.forEach((entry) => {
        insert.run(
          final.id, entry.sheet, entry.pointKey, entry.keptReportId, entry.keptCollectionDate,
          entry.supersededReportId, entry.supersededCollectionDate, JSON.stringify(entry.supersededRow),
          entry.resolution, entry.overriddenBy || null, now
        );
      });
      return { report: final, provenance: listAssemblyProvenance(final.id) };
    })();
  }

  return {
    listReports,
    getReport,
    addReport,
    updateReport,
    deleteReport,
    // multi-day (additive)
    addProject,
    getProject,
    listProjects,
    updateProject,
    listProjectReports,
    transitionReport,
    listTransitions,
    assembleFinal,
    listAssemblyProvenance
  };
}

module.exports = {
  createReportStore,
  // Exported for the migration script and tests (shared schema, no duplication).
  openDatabase,
  resolveDbPath,
  extractSearchColumns
};
