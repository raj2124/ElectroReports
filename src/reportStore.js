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
const { requiredAcknowledgements, approvalBlockers, ACK_LABEL, currentRound, lastOf, buildReviewState } = require('./review');
const { buildFinalReportData, assertConsistent } = require('./finalReport/data');

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
  migrateSheetSchema(db);
  migrateReviewSchema(db);
  migrateNoticeSchema(db);
  migrateAuthSchema(db);
  migrateScheduleSchema(db);
  migratePeopleSyncSchema(db);
  migrateCalendarSyncSchema(db);
  migrateScanSchema(db);
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

// Review (ADDITIVE): every transition records which step it was (submit / check / approve /
// return), and per-client logos are kept for the report covers. Logos live only in the database,
// which is not committed to the repository.
function migrateReviewSchema(db) {
  db.transaction(() => {
    if (!hasColumn(db, 'report_transitions', 'action')) {
      db.exec('ALTER TABLE report_transitions ADD COLUMN action TEXT');
    }
    if (!hasColumn(db, 'report_transitions', 'designation')) {
      db.exec('ALTER TABLE report_transitions ADD COLUMN designation TEXT');
    }
    db.exec(`
      CREATE TABLE IF NOT EXISTS client_logos (
        key         TEXT PRIMARY KEY,
        client_name TEXT,
        mime        TEXT NOT NULL,
        data        BLOB NOT NULL,
        updated_at  TEXT NOT NULL
      );
    `);
  }).immediate();
}

// In-app notices (ADDITIVE): "a new D03 project needs two answers". One row per project and
// recipient, never deleted, so an answered or dismissed notice cannot come back.
function migrateNoticeSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS notices (
      id              TEXT PRIMARY KEY,
      kind            TEXT NOT NULL,
      zoho_project_id TEXT NOT NULL,
      recipient_key   TEXT NOT NULL,
      recipient_name  TEXT,
      status          TEXT NOT NULL,
      payload         TEXT NOT NULL,
      answers         TEXT,
      created_at      TEXT NOT NULL,
      resolved_at     TEXT,
      resolved_by     TEXT,
      UNIQUE (kind, zoho_project_id, recipient_key)
    );
    CREATE INDEX IF NOT EXISTS idx_notices_recipient ON notices(recipient_key, status);
  `);
}

// Scheduling (ADDITIVE, Session 6F): what was last read from Zoho (with when and the last error) and the
// settings an admin edits. Nothing here is a second copy of a date: a refresh replaces it.
function migrateScheduleSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS schedule_cache (
      key        TEXT PRIMARY KEY,
      fetched_at TEXT,
      data       TEXT,
      error      TEXT,
      error_at   TEXT
    );
    CREATE TABLE IF NOT EXISTS schedule_settings (
      key        TEXT PRIMARY KEY,
      value      TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      updated_by TEXT
    );
  `);
}

// Site Work Log in Zoho People (ADDITIVE, Session 6G): one row per log entry, keyed by the Record Key
// that is also written to People. It remembers the People record id (so a later revision updates the same
// record), the neutral entry that was or is to be written (a retry re-resolves the engineer), and how the
// last attempt went.
function migratePeopleSyncSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS people_sync (
      record_key       TEXT PRIMARY KEY,
      root_id          TEXT NOT NULL,
      report_id        TEXT,
      test_id          TEXT,
      work_date        TEXT,
      status           TEXT NOT NULL,
      sync_state       TEXT NOT NULL,
      people_record_id TEXT,
      hash             TEXT,
      entry            TEXT,
      error            TEXT,
      attempts         INTEGER NOT NULL DEFAULT 0,
      synced_at        TEXT,
      updated_at       TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_people_sync_root ON people_sync(root_id);
    CREATE INDEX IF NOT EXISTS idx_people_sync_state ON people_sync(sync_state);
  `);
}

// Scanned sheets and AI calls (ADDITIVE, Session 6I). A scan is one page of an upload: its image is a file in
// data/scans/, everything else (the issued sheet it matched, what was read, the engineer's edits, what an accept
// overwrote) is JSON in `data`. ai_usage counts provider calls (no content), for budgeting.
function migrateScanSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS scans (
      id          TEXT PRIMARY KEY,
      batch_id    TEXT NOT NULL,
      created_by  TEXT NOT NULL,
      created_at  TEXT NOT NULL,
      updated_at  TEXT NOT NULL,
      status      TEXT NOT NULL,
      file_name   TEXT,
      page_index  INTEGER NOT NULL DEFAULT 0,
      image_path  TEXT,
      sheet_id    TEXT,
      sheet_key   TEXT,
      project_no  TEXT,
      work_date   TEXT,
      accepted_at TEXT,
      data        TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_scans_status ON scans(status);
    CREATE INDEX IF NOT EXISTS idx_scans_created ON scans(created_at);
    CREATE TABLE IF NOT EXISTS ai_usage (
      id       INTEGER PRIMARY KEY AUTOINCREMENT,
      at       TEXT NOT NULL,
      provider TEXT NOT NULL,
      model    TEXT,
      kind     TEXT NOT NULL,
      ok       INTEGER NOT NULL,
      code     TEXT,
      ms       INTEGER,
      units    INTEGER
    );
    CREATE INDEX IF NOT EXISTS idx_ai_usage_at ON ai_usage(at);
  `);
}

// Outlook calendar sync (ADDITIVE, Session 6H): the events the app created, one per Zoho task and owner. The
// event id and a hash of what was sent let the next run update only what changed; Outlook itself (the events
// tagged with the app's extended property) is what a run reconciles against.
function migrateCalendarSyncSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS calendar_sync (
      event_key   TEXT PRIMARY KEY,
      task_id     TEXT NOT NULL,
      owner_email TEXT NOT NULL,
      event_id    TEXT,
      hash        TEXT,
      sync_state  TEXT NOT NULL,
      error       TEXT,
      synced_at   TEXT,
      updated_at  TEXT NOT NULL
    );
  `);
}

// Sign-in (ADDITIVE): the people who sign in, their sessions, and the audit log. Every transition
// also records the user id; the name and designation stay as a snapshot of that moment.
function migrateAuthSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id            TEXT PRIMARY KEY,
      email         TEXT NOT NULL UNIQUE,
      name          TEXT NOT NULL,
      designation   TEXT NOT NULL DEFAULT '',
      roles         TEXT NOT NULL,
      zuid          TEXT,
      active        INTEGER NOT NULL DEFAULT 1,
      created_at    TEXT NOT NULL,
      last_login_at TEXT
    );
    CREATE TABLE IF NOT EXISTS sessions (
      id_hash      TEXT PRIMARY KEY,
      user_id      TEXT NOT NULL,
      created_at   TEXT NOT NULL,
      last_seen_at TEXT NOT NULL,
      checked_at   TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_sessions_user ON sessions(user_id);
    CREATE TABLE IF NOT EXISTS audit_log (
      id          INTEGER PRIMARY KEY AUTOINCREMENT,
      at          TEXT NOT NULL,
      user_id     TEXT,
      user_email  TEXT,
      user_name   TEXT,
      action      TEXT NOT NULL,
      entity_type TEXT,
      entity_id   TEXT,
      detail      TEXT
    );
    CREATE INDEX IF NOT EXISTS idx_audit_at ON audit_log(at);
  `);
  if (!hasColumn(db, 'report_transitions', 'user_id')) {
    db.exec('ALTER TABLE report_transitions ADD COLUMN user_id TEXT');
  }
}

// Printed measurement sheets (ADDITIVE). Every printed page is issued with a unique sheet ID that
// is also its QR code; this table is what lets a scan be matched back to exactly what was
// printed: project, work date, test, page, and the pre-printed rows.
function migrateSheetSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS issued_sheets (
      sheet_id        TEXT PRIMARY KEY,
      batch_id        TEXT NOT NULL,
      project_no      TEXT,
      zoho_project_id TEXT,
      report_id       TEXT,
      work_date       TEXT NOT NULL,
      sheet_key       TEXT NOT NULL,
      sheet_code      TEXT NOT NULL,
      doc_no          TEXT NOT NULL,
      page_no         INTEGER NOT NULL,
      page_count      INTEGER NOT NULL,
      row_offset      INTEGER NOT NULL,
      row_count       INTEGER NOT NULL,
      engineer        TEXT,
      issued_at       TEXT NOT NULL,
      data            TEXT NOT NULL
    );
    CREATE INDEX IF NOT EXISTS idx_issued_sheets_project ON issued_sheets(project_no, work_date);
    CREATE INDEX IF NOT EXISTS idx_issued_sheets_batch ON issued_sheets(batch_id);
  `);
  db.pragma('user_version = 3');
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
    ['kind', 'projectId', 'reviewStatus', 'dailyStatus', 'createdBy'].forEach((field) => {
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

  // ---- Review workflow: draft -> in_review -> (checked) -> approved (daily AND final) ------
  // A report is submitted (Prepared), marked as checked by the checker (Checked), then approved
  // (Approved); each step is its own row in report_transitions. Approval is blocked until the
  // report is checked, its counts are consistent and, on a FINAL report, the expired-calibration
  // and sanity-check-not-run notes are acknowledged. The same person checking and approving is
  // allowed but returned as a warning.
  function requireWorkflow(id) {
    const current = getReport(id);
    if (!current) {
      return null;
    }
    if (getReportKind(current) === REPORT_KINDS.STANDALONE) {
      throw new Error('Standalone reports have no review workflow.');
    }
    return current;
  }

  function failure(message, code, extra = {}) {
    const error = new Error(message);
    error.code = code;
    return Object.assign(error, extra);
  }

  function writeTransition(updated, fromStatus, toStatus, action, actor, note, at, designation = '', userId = '') {
    db.transaction(() => {
      updateStmt.run({
        id: updated.id,
        updated_at: at,
        ...extractSearchColumns(updated),
        ...extractCollectionColumns(updated),
        data: JSON.stringify(updated)
      });
      db.prepare(
        'INSERT INTO report_transitions (report_id, from_status, to_status, actor, note, at, action, designation, user_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
      ).run(updated.id, fromStatus, toStatus, String(actor || ''), String(note || ''), at, action, String(designation || ''), String(userId || ''));
    })();
  }

  // acknowledgements: { calibration, sanityCheck } notes. `acknowledgement` (a single string)
  // is the older form and covers the calibration note only.
  function transition(id, toStatus, { actor = '', designation = '', userId = '', note = '', acknowledgement = '', acknowledgements = {} } = {}) {
    const current = requireWorkflow(id);
    if (!current) {
      return null;
    }
    const fromStatus = getReviewStatus(current);
    if (!canTransition(fromStatus, toStatus)) {
      throw failure(`Invalid transition ${fromStatus} -> ${toStatus}.`, 'INVALID_TRANSITION');
    }
    const warnings = [];
    const recorded = [];
    let action = toStatus === 'in_review' ? 'submit' : toStatus === 'approved' ? 'approve' : 'return';
    if (toStatus === 'approved') {
      const round = currentRound(listTransitions(id));
      const checked = lastOf(round, 'check');
      if (!checked) {
        throw failure('Mark the report as checked before approving it.', 'NOT_CHECKED');
      }
      assertConsistent(buildFinalReportData(current));
      const blockers = approvalBlockers(current);
      if (blockers.length) {
        throw failure(blockers.map((item) => item.text).join(' '), blockers[0].code, { items: blockers });
      }
      const given = Object.fromEntries(Object.entries(acknowledgements || {}).map(([key, value]) => [key, String(value || '').trim()]));
      if (!given.calibration && String(acknowledgement || '').trim()) {
        given.calibration = String(acknowledgement).trim();
      }
      const missing = requiredAcknowledgements(current).filter((item) => !given[item.key]);
      if (missing.length) {
        throw failure(`Acknowledge before approving: ${missing.map((item) => item.text).join(' ')}`, 'ACKNOWLEDGEMENT_REQUIRED', { items: missing });
      }
      requiredAcknowledgements(current).forEach((item) => {
        recorded.push(`${ACK_LABEL[item.key] || item.key}: ${given[item.key]}`);
      });
      if (String(actor || '').trim() && String(actor).trim().toLowerCase() === String(checked.actor || '').trim().toLowerCase()) {
        warnings.push(`${actor} both checked and approved this report.`);
      }
    }
    const at = new Date().toISOString();
    const updated = { ...current, reviewStatus: toStatus, updatedAt: at };
    if (toStatus === 'approved') {
      updated.approvedAt = at;
    }
    if (toStatus === 'draft') {
      updated.review = { ...(current.review || {}), ticks: {} };
    }
    const fullNote = [...recorded, String(note || '')].filter(Boolean).join(' | ');
    writeTransition(updated, fromStatus, toStatus, action, actor, fullNote, at, designation, userId);
    return { report: updated, warnings };
  }

  function transitionReport(id, toStatus, options) {
    const result = transition(id, toStatus, options);
    return result ? result.report : null;
  }

  // The checker's step: stays in_review, recorded as its own transition.
  function markChecked(id, { actor = '', designation = '', userId = '', note = '' } = {}) {
    const current = requireWorkflow(id);
    if (!current) {
      return null;
    }
    if (getReviewStatus(current) !== 'in_review') {
      throw failure('Only a report in review can be marked as checked.', 'INVALID_STATE');
    }
    if (lastOf(currentRound(listTransitions(id)), 'check')) {
      throw failure('This report is already marked as checked; return it to draft to change it.', 'ALREADY_CHECKED');
    }
    assertConsistent(buildFinalReportData(current));
    const unticked = buildReviewState({ getReport, listTransitions }, current).checklist.filter((item) => item.manual && !item.ticked);
    if (unticked.length) {
      throw failure(`Tick every checklist item first: ${unticked.map((item) => item.text).join('; ')}.`, 'CHECKLIST_INCOMPLETE', { items: unticked });
    }
    const at = new Date().toISOString();
    const updated = { ...current, updatedAt: at };
    writeTransition(updated, 'in_review', 'in_review', 'check', actor, note, at, designation, userId);
    return updated;
  }

  // Checklist ticks are reviewer state on the report, not a transition.
  function setReviewTicks(id, ticks) {
    const current = requireWorkflow(id);
    if (!current) {
      return null;
    }
    if (getReviewStatus(current) === 'approved') {
      throw failure('An approved report cannot be changed; re-issue it as the next revision.', 'INVALID_STATE');
    }
    return updateReport(id, (report) => ({ ...report, review: { ...(report.review || {}), ticks: { ...Object.fromEntries(Object.entries(ticks || {}).map(([key, value]) => [key, Boolean(value)])) } } }));
  }

  // Re-issue after approval: the approved report stays as issued; a copy becomes the next
  // revision in draft, with its own review round.
  function reissueReport(id, { actor = '', userId = '', reason = '' } = {}) {
    const current = requireWorkflow(id);
    if (!current) {
      return null;
    }
    if (getReviewStatus(current) !== 'approved') {
      throw failure('Only an approved report can be re-issued.', 'INVALID_STATE');
    }
    const existing = listProjectReports(current.projectId, { kind: getReportKind(current) }).filter((entry) => entry.revisionOf === current.id);
    if (existing.length) {
      throw failure('This revision has already been re-issued.', 'ALREADY_REISSUED', { reportId: existing[0].id });
    }
    const { id: _id, createdAt: _created, updatedAt: _updated, approvedAt: _approved, review: _review, dailyStatus: _daily, ...body } = current;
    const copy = addReport({
      ...body,
      reviewStatus: 'draft',
      revision: (Number.isInteger(current.revision) ? current.revision : 0) + 1,
      revisionOf: current.id,
      revisionReason: String(reason || '').trim() || 'Re-issue',
      revisedBy: String(actor || ''),
      createdBy: userId ? String(userId) : current.createdBy
    });
    return copy;
  }

  function listTransitions(reportId) {
    return db
      .prepare('SELECT id, report_id AS reportId, from_status AS fromStatus, to_status AS toStatus, actor, designation, user_id AS userId, note, at, action FROM report_transitions WHERE report_id = ? ORDER BY id')
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
  function assembleFinal(projectId, { overrides = {}, actor = '', userId = '' } = {}) {
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
      const final = addReport({ ...reportBody, projectId: project.id, ...(userId ? { createdBy: String(userId) } : {}) });
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

  // ---- Issued (printed) sheets ----------------------------------------------------------------
  const issuedColumns = (sheet) => ({
    sheet_id: sheet.sheetId,
    batch_id: sheet.batchId,
    project_no: sheet.project.projectNo || null,
    zoho_project_id: sheet.project.zohoProjectId || null,
    report_id: sheet.reportId || null,
    work_date: sheet.workDate,
    sheet_key: sheet.sheetKey,
    sheet_code: sheet.sheetCode,
    doc_no: sheet.docNo,
    page_no: sheet.page,
    page_count: sheet.pages,
    row_offset: sheet.rowOffset,
    row_count: sheet.rowCount,
    engineer: sheet.engineer || null,
    issued_at: sheet.issuedAt,
    data: JSON.stringify({ project: sheet.project, preRows: sheet.preRows, itemOffset: sheet.itemOffset, group: sheet.group })
  });

  function rowToIssuedSheet(row) {
    if (!row) {
      return null;
    }
    const data = JSON.parse(row.data);
    return {
      sheetId: row.sheet_id,
      batchId: row.batch_id,
      project: data.project,
      reportId: row.report_id,
      workDate: row.work_date,
      sheetKey: row.sheet_key,
      sheetCode: row.sheet_code,
      docNo: row.doc_no,
      page: row.page_no,
      pages: row.page_count,
      rowOffset: row.row_offset,
      rowCount: row.row_count,
      itemOffset: data.itemOffset,
      group: data.group,
      engineer: row.engineer,
      issuedAt: row.issued_at,
      preRows: data.preRows
    };
  }

  function addIssuedSheets(sheets) {
    const insert = db.prepare(`INSERT INTO issued_sheets
      (sheet_id, batch_id, project_no, zoho_project_id, report_id, work_date, sheet_key, sheet_code, doc_no,
       page_no, page_count, row_offset, row_count, engineer, issued_at, data)
      VALUES (@sheet_id, @batch_id, @project_no, @zoho_project_id, @report_id, @work_date, @sheet_key, @sheet_code, @doc_no,
       @page_no, @page_count, @row_offset, @row_count, @engineer, @issued_at, @data)`);
    db.transaction(() => sheets.forEach((sheet) => insert.run(issuedColumns(sheet))))();
    return sheets;
  }

  function getIssuedSheet(sheetId) {
    return rowToIssuedSheet(db.prepare('SELECT * FROM issued_sheets WHERE sheet_id = ?').get(String(sheetId)));
  }

  function listIssuedSheets({ batchId, projectNo, workDate } = {}) {
    const where = [];
    const params = {};
    if (batchId) { where.push('batch_id = @batchId'); params.batchId = batchId; }
    if (projectNo) { where.push('project_no = @projectNo'); params.projectNo = projectNo; }
    if (workDate) { where.push('work_date = @workDate'); params.workDate = workDate; }
    const sql = `SELECT * FROM issued_sheets ${where.length ? `WHERE ${where.join(' AND ')}` : ''} ORDER BY issued_at, sheet_code, page_no`;
    return db.prepare(sql).all(params).map(rowToIssuedSheet);
  }

  // ---- Client logos (cover pages) ------------------------------------------------------------
  function getClientLogo(key) {
    const row = db.prepare('SELECT key, client_name AS clientName, mime, data, updated_at AS updatedAt FROM client_logos WHERE key = ?').get(String(key));
    return row || null;
  }

  function setClientLogo({ key, clientName = '', mime, data }) {
    db.prepare(`INSERT INTO client_logos (key, client_name, mime, data, updated_at) VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(key) DO UPDATE SET client_name = excluded.client_name, mime = excluded.mime, data = excluded.data, updated_at = excluded.updated_at`)
      .run(String(key), String(clientName), mime, data, new Date().toISOString());
    return getClientLogo(key);
  }

  function deleteClientLogo(key) {
    return db.prepare('DELETE FROM client_logos WHERE key = ?').run(String(key)).changes > 0;
  }

  // ---- Users, sessions, audit ---------------------------------------------------------------
  function rowToUser(row) {
    return row
      ? { id: row.id, email: row.email, name: row.name, designation: row.designation || '', roles: JSON.parse(row.roles), zuid: row.zuid || '', active: Boolean(row.active), createdAt: row.created_at, lastLoginAt: row.last_login_at || '' }
      : null;
  }

  function getUser(id) {
    return rowToUser(db.prepare('SELECT * FROM users WHERE id = ?').get(String(id)));
  }

  function getUserByEmail(email) {
    return rowToUser(db.prepare('SELECT * FROM users WHERE email = ?').get(String(email).trim().toLowerCase()));
  }

  function listUsers() {
    return db.prepare('SELECT * FROM users ORDER BY name COLLATE NOCASE').all().map(rowToUser);
  }

  function addUser({ email, name, designation = '', roles = ['engineer'], zuid = '' }) {
    const id = `U-${buildId().replace(/^ER-/, '')}`;
    db.prepare('INSERT INTO users (id, email, name, designation, roles, zuid, active, created_at) VALUES (?, ?, ?, ?, ?, ?, 1, ?)').run(id, String(email).trim().toLowerCase(), name, designation, JSON.stringify(roles), zuid || null, new Date().toISOString());
    return getUser(id);
  }

  function updateUser(id, patch) {
    const current = getUser(id);
    if (!current) return null;
    const next = { ...current, ...patch };
    db.prepare('UPDATE users SET name = ?, designation = ?, roles = ?, zuid = ?, active = ?, last_login_at = ? WHERE id = ?')
      .run(next.name, next.designation || '', JSON.stringify(next.roles), next.zuid || null, next.active ? 1 : 0, next.lastLoginAt || null, String(id));
    return getUser(id);
  }

  function createSession(idHash, userId) {
    const now = new Date().toISOString();
    db.prepare('INSERT INTO sessions (id_hash, user_id, created_at, last_seen_at, checked_at) VALUES (?, ?, ?, ?, ?)').run(idHash, String(userId), now, now, now);
  }

  function getSession(idHash) {
    const row = db.prepare('SELECT * FROM sessions WHERE id_hash = ?').get(String(idHash));
    return row ? { idHash: row.id_hash, userId: row.user_id, createdAt: row.created_at, lastSeenAt: row.last_seen_at, checkedAt: row.checked_at } : null;
  }

  function touchSession(idHash, { checked = false } = {}) {
    const now = new Date().toISOString();
    db.prepare(`UPDATE sessions SET last_seen_at = ?${checked ? ', checked_at = ?' : ''} WHERE id_hash = ?`).run(...(checked ? [now, now, String(idHash)] : [now, String(idHash)]));
  }

  function deleteSession(idHash) {
    db.prepare('DELETE FROM sessions WHERE id_hash = ?').run(String(idHash));
  }

  function deleteUserSessions(userId) {
    db.prepare('DELETE FROM sessions WHERE user_id = ?').run(String(userId));
  }

  function addAudit({ user = null, action, entityType = '', entityId = '', detail = null }) {
    db.prepare('INSERT INTO audit_log (at, user_id, user_email, user_name, action, entity_type, entity_id, detail) VALUES (?, ?, ?, ?, ?, ?, ?, ?)')
      .run(new Date().toISOString(), user ? user.id : null, user ? user.email : null, user ? user.name : null, action, entityType || null, entityId ? String(entityId) : null, detail ? JSON.stringify(detail) : null);
  }

  function listAudit({ userId = '', action = '', limit = 100, offset = 0 } = {}) {
    const clauses = [];
    const params = [];
    if (userId) { clauses.push('user_id = ?'); params.push(String(userId)); }
    if (action) { clauses.push('action LIKE ?'); params.push(`${String(action).replace(/[%_]/g, '')}%`); }
    const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
    const total = db.prepare(`SELECT COUNT(*) AS n FROM audit_log ${where}`).get(...params).n;
    const rows = db.prepare(`SELECT * FROM audit_log ${where} ORDER BY id DESC LIMIT ? OFFSET ?`).all(...params, Math.min(500, Math.max(1, limit)), Math.max(0, offset));
    return { total, entries: rows.map((row) => ({ id: row.id, at: row.at, userId: row.user_id || '', userEmail: row.user_email || '', userName: row.user_name || '', action: row.action, entityType: row.entity_type || '', entityId: row.entity_id || '', detail: row.detail ? JSON.parse(row.detail) : null })) };
  }

  // ---- Schedule cache and settings --------------------------------------------------------------
  function getScheduleCache(key) {
    const row = db.prepare('SELECT * FROM schedule_cache WHERE key = ?').get(String(key));
    return row ? { key: row.key, fetchedAt: row.fetched_at || '', data: row.data ? JSON.parse(row.data) : null, error: row.error || '', errorAt: row.error_at || '' } : null;
  }

  // A good read replaces the data and clears the error. A failed read (data undefined) keeps the last
  // good data and records the error, so the screen can show old data with its age.
  function setScheduleCache(key, data, { error = '', at = '' } = {}) {
    const now = at || new Date().toISOString();
    if (data === undefined) {
      db.prepare('INSERT INTO schedule_cache (key, error, error_at) VALUES (?, ?, ?) ON CONFLICT(key) DO UPDATE SET error = excluded.error, error_at = excluded.error_at').run(String(key), String(error), now);
    } else {
      db.prepare("INSERT INTO schedule_cache (key, fetched_at, data, error, error_at) VALUES (?, ?, ?, '', NULL) ON CONFLICT(key) DO UPDATE SET fetched_at = excluded.fetched_at, data = excluded.data, error = '', error_at = NULL").run(String(key), now, JSON.stringify(data));
    }
    return getScheduleCache(key);
  }

  function getScheduleSetting(key) {
    const row = db.prepare('SELECT value FROM schedule_settings WHERE key = ?').get(String(key));
    return row ? JSON.parse(row.value) : null;
  }

  function setScheduleSetting(key, value, by = '') {
    db.prepare('INSERT INTO schedule_settings (key, value, updated_at, updated_by) VALUES (?, ?, ?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value, updated_at = excluded.updated_at, updated_by = excluded.updated_by')
      .run(String(key), JSON.stringify(value), new Date().toISOString(), String(by));
    return value;
  }

  // ---- Site Work Log sync (6G) --------------------------------------------------------------
  function rowToPeopleSync(row) {
    return row ? {
      key: row.record_key, rootId: row.root_id, reportId: row.report_id || '', testId: row.test_id || '', workDate: row.work_date || '',
      status: row.status, state: row.sync_state, peopleRecordId: row.people_record_id || '', hash: row.hash || '',
      entry: row.entry ? JSON.parse(row.entry) : null, error: row.error || '', attempts: row.attempts, syncedAt: row.synced_at || '', updatedAt: row.updated_at
    } : null;
  }

  function getPeopleSync(key) {
    return rowToPeopleSync(db.prepare('SELECT * FROM people_sync WHERE record_key = ?').get(String(key)));
  }

  function listPeopleSync({ rootId = '', state = '', limit = 500 } = {}) {
    const where = [];
    const args = [];
    if (rootId) { where.push('root_id = ?'); args.push(String(rootId)); }
    if (state) { where.push('sync_state = ?'); args.push(String(state)); }
    const rows = db.prepare(`SELECT * FROM people_sync ${where.length ? `WHERE ${where.join(' AND ')}` : ''} ORDER BY updated_at DESC LIMIT ?`).all(...args, limit);
    return rows.map(rowToPeopleSync);
  }

  function peopleSyncCounts() {
    return Object.fromEntries(db.prepare('SELECT sync_state AS state, COUNT(*) AS n FROM people_sync GROUP BY sync_state').all().map((row) => [row.state, row.n]));
  }

  // Insert or replace the row of `key`; fields not given keep their value.
  function savePeopleSync(key, fields) {
    const current = getPeopleSync(key) || { attempts: 0 };
    const next = { ...current, ...fields, updatedAt: new Date().toISOString() };
    db.prepare(`INSERT INTO people_sync (record_key, root_id, report_id, test_id, work_date, status, sync_state, people_record_id, hash, entry, error, attempts, synced_at, updated_at)
      VALUES (@key, @rootId, @reportId, @testId, @workDate, @status, @state, @peopleRecordId, @hash, @entry, @error, @attempts, @syncedAt, @updatedAt)
      ON CONFLICT(record_key) DO UPDATE SET root_id = @rootId, report_id = @reportId, test_id = @testId, work_date = @workDate, status = @status, sync_state = @state,
        people_record_id = @peopleRecordId, hash = @hash, entry = @entry, error = @error, attempts = @attempts, synced_at = @syncedAt, updated_at = @updatedAt`).run({
      key: String(key), rootId: next.rootId || '', reportId: next.reportId || '', testId: next.testId || '', workDate: next.workDate || '', status: next.status || '',
      state: next.state || 'pending', peopleRecordId: next.peopleRecordId || '', hash: next.hash || '', entry: next.entry ? JSON.stringify(next.entry) : null,
      error: next.error || '', attempts: next.attempts || 0, syncedAt: next.syncedAt || '', updatedAt: next.updatedAt
    });
    return getPeopleSync(key);
  }

  // The workspaces a person worked on: they created a daily report there or are credited on a row (6G stamps).
  function projectIdsWorkedBy(userId) {
    const id = String(userId || '');
    if (!id) return [];
    return db.prepare("SELECT DISTINCT project_id FROM reports WHERE project_id IS NOT NULL AND (data LIKE ? OR data LIKE ?)").all(`%"createdBy":"${id.replace(/[%_"\\]/g, '')}"%`, `%"userId":"${id.replace(/[%_"\\]/g, '')}"%`).map((row) => row.project_id);
  }

  // ---- Scans (6I) -------------------------------------------------------------------------------
  const rowToScan = (row) => (row ? { id: row.id, batchId: row.batch_id, createdBy: row.created_by, createdAt: row.created_at, updatedAt: row.updated_at, status: row.status, fileName: row.file_name || '', pageIndex: row.page_index, imagePath: row.image_path || '', sheetId: row.sheet_id || '', sheetKey: row.sheet_key || '', projectNo: row.project_no || '', workDate: row.work_date || '', acceptedAt: row.accepted_at || '', ...JSON.parse(row.data) } : null);
  const getScan = (id) => rowToScan(db.prepare('SELECT * FROM scans WHERE id = ?').get(String(id)));
  function listScans({ createdBy = '', status = '', limit = 200 } = {}) {
    const where = [];
    const args = [];
    if (createdBy) { where.push('created_by = ?'); args.push(String(createdBy)); }
    if (status) { where.push('status = ?'); args.push(String(status)); }
    return db.prepare(`SELECT * FROM scans ${where.length ? `WHERE ${where.join(' AND ')}` : ''} ORDER BY created_at DESC, page_index LIMIT ?`).all(...args, limit).map(rowToScan);
  }
  const SCAN_COLUMNS = { status: 'status', fileName: 'file_name', imagePath: 'image_path', sheetId: 'sheet_id', sheetKey: 'sheet_key', projectNo: 'project_no', workDate: 'work_date', acceptedAt: 'accepted_at' };
  // Columns and the JSON `data` are both merged: pass only what changes.
  function saveScan(id, fields) {
    const current = getScan(id);
    const now = new Date().toISOString();
    if (!current) {
      const scan = { id: String(id), batchId: fields.batchId || '', createdBy: fields.createdBy || '', createdAt: now, pageIndex: fields.pageIndex || 0, ...fields };
      const { batchId, createdBy, createdAt, pageIndex, ...rest } = scan;
      const columns = {};
      Object.entries(SCAN_COLUMNS).forEach(([key, column]) => { columns[column] = rest[key] === undefined ? null : rest[key]; delete rest[key]; });
      delete rest.id;
      db.prepare('INSERT INTO scans (id, batch_id, created_by, created_at, updated_at, status, file_name, page_index, image_path, sheet_id, sheet_key, project_no, work_date, accepted_at, data) VALUES (@id, @batch_id, @created_by, @created_at, @updated_at, @status, @file_name, @page_index, @image_path, @sheet_id, @sheet_key, @project_no, @work_date, @accepted_at, @data)')
        .run({ id: scan.id, batch_id: batchId, created_by: createdBy, created_at: createdAt, updated_at: now, page_index: pageIndex, ...columns, status: columns.status || 'queued', data: JSON.stringify(rest) });
      return getScan(id);
    }
    const columns = {};
    const rest = {};
    Object.entries(fields).forEach(([key, value]) => { if (SCAN_COLUMNS[key]) columns[SCAN_COLUMNS[key]] = value; else rest[key] = value; });
    const { id: _id, batchId, createdBy, createdAt, updatedAt, pageIndex, imagePath: _p, status: _s, fileName: _f, sheetId: _si, sheetKey: _sk, projectNo: _pn, workDate: _wd, acceptedAt: _aa, ...data } = current;
    const sets = Object.keys(columns).map((column) => `${column} = @${column}`);
    db.prepare(`UPDATE scans SET ${[...sets, 'updated_at = @updated_at', 'data = @data'].join(', ')} WHERE id = @id`).run({ id: String(id), updated_at: now, data: JSON.stringify({ ...data, ...rest }), ...columns });
    return getScan(id);
  }
  // Deletes the scans created on [from, to] (ISO dates); returns what was deleted, so the caller can remove the images.
  function deleteScans({ from, to }) {
    const rows = db.prepare('SELECT id, image_path FROM scans WHERE substr(created_at, 1, 10) >= ? AND substr(created_at, 1, 10) <= ?').all(String(from), String(to));
    db.transaction(() => rows.forEach((row) => db.prepare('DELETE FROM scans WHERE id = ?').run(row.id)))();
    return rows.map((row) => ({ id: row.id, imagePath: row.image_path || '' }));
  }
  function addAiUsage({ provider, model = '', kind, ok, code = '', ms = 0, units = 1 }) {
    db.prepare('INSERT INTO ai_usage (at, provider, model, kind, ok, code, ms, units) VALUES (?, ?, ?, ?, ?, ?, ?, ?)').run(new Date().toISOString(), provider, model, kind, ok ? 1 : 0, code, ms, units);
  }
  const aiUsageSummary = ({ from = '0000-01-01', to = '9999-12-31' } = {}) => db.prepare("SELECT provider, kind, COUNT(*) AS calls, SUM(ok) AS ok, SUM(units) AS units FROM ai_usage WHERE substr(at, 1, 10) >= ? AND substr(at, 1, 10) <= ? GROUP BY provider, kind").all(String(from), String(to));

  // ---- Calendar sync (6H) -------------------------------------------------------------------
  const rowToCalendar = (row) => (row ? { key: row.event_key, taskId: row.task_id, ownerEmail: row.owner_email, eventId: row.event_id || '', hash: row.hash || '', state: row.sync_state, error: row.error || '', syncedAt: row.synced_at || '' } : null);
  const getCalendarSync = (key) => rowToCalendar(db.prepare('SELECT * FROM calendar_sync WHERE event_key = ?').get(String(key)));
  const listCalendarSync = () => db.prepare('SELECT * FROM calendar_sync ORDER BY event_key').all().map(rowToCalendar);
  function saveCalendarSync(key, fields) {
    const next = { ...(getCalendarSync(key) || {}), ...fields };
    db.prepare(`INSERT INTO calendar_sync (event_key, task_id, owner_email, event_id, hash, sync_state, error, synced_at, updated_at) VALUES (@key, @taskId, @ownerEmail, @eventId, @hash, @state, @error, @syncedAt, @updatedAt)
      ON CONFLICT(event_key) DO UPDATE SET task_id = @taskId, owner_email = @ownerEmail, event_id = @eventId, hash = @hash, sync_state = @state, error = @error, synced_at = @syncedAt, updated_at = @updatedAt`)
      .run({ key: String(key), taskId: next.taskId || '', ownerEmail: next.ownerEmail || '', eventId: next.eventId || '', hash: next.hash || '', state: next.state || 'done', error: next.error || '', syncedAt: next.syncedAt || '', updatedAt: new Date().toISOString() });
    return getCalendarSync(key);
  }
  const deleteCalendarSync = (key) => db.prepare('DELETE FROM calendar_sync WHERE event_key = ?').run(String(key));

  // ---- Notices ------------------------------------------------------------------------------
  function rowToNotice(row) {
    return {
      id: row.id,
      kind: row.kind,
      zohoProjectId: row.zoho_project_id,
      recipientKey: row.recipient_key,
      recipientName: row.recipient_name || '',
      status: row.status,
      payload: JSON.parse(row.payload),
      answers: row.answers ? JSON.parse(row.answers) : null,
      createdAt: row.created_at,
      resolvedAt: row.resolved_at || '',
      resolvedBy: row.resolved_by || ''
    };
  }

  // Returns the notice, or null when this recipient already has one for the project (any status).
  function addNotice({ kind = 'new_project', zohoProjectId, recipientKey, recipientName = '', payload = {} }) {
    const id = `NT-${buildId().replace(/^ER-/, '')}`;
    const result = db
      .prepare('INSERT OR IGNORE INTO notices (id, kind, zoho_project_id, recipient_key, recipient_name, status, payload, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)')
      .run(id, kind, String(zohoProjectId), String(recipientKey).toLowerCase(), recipientName, 'open', JSON.stringify(payload), new Date().toISOString());
    return result.changes ? getNotice(id) : null;
  }

  function getNotice(id) {
    const row = db.prepare('SELECT * FROM notices WHERE id = ?').get(String(id));
    return row ? rowToNotice(row) : null;
  }

  function listNotices({ recipientKey, status, zohoProjectId } = {}) {
    const clauses = [];
    const params = [];
    if (recipientKey !== undefined) { clauses.push('recipient_key = ?'); params.push(String(recipientKey).toLowerCase()); }
    if (status) { clauses.push('status = ?'); params.push(status); }
    if (zohoProjectId) { clauses.push('zoho_project_id = ?'); params.push(String(zohoProjectId)); }
    const where = clauses.length ? `WHERE ${clauses.join(' AND ')}` : '';
    return db.prepare(`SELECT * FROM notices ${where} ORDER BY created_at, id`).all(...params).map(rowToNotice);
  }

  function resolveNotice(id, { status, answers = null, by = '' }) {
    db.prepare("UPDATE notices SET status = ?, answers = ?, resolved_at = ?, resolved_by = ? WHERE id = ? AND status = 'open'")
      .run(status, answers ? JSON.stringify(answers) : null, new Date().toISOString(), String(by), String(id));
    return getNotice(id);
  }

  // When one owner answers, the other owners' open notices for the project are closed.
  function closeOpenNotices(zohoProjectId, { exceptId = '', by = '' } = {}) {
    return db
      .prepare("UPDATE notices SET status = 'answered', resolved_at = ?, resolved_by = ? WHERE zoho_project_id = ? AND status = 'open' AND id != ?")
      .run(new Date().toISOString(), String(by), String(zohoProjectId), String(exceptId)).changes;
  }

  return {
    getUser,
    getUserByEmail,
    listUsers,
    addUser,
    updateUser,
    createSession,
    getSession,
    touchSession,
    deleteSession,
    deleteUserSessions,
    addAudit,
    listAudit,
    getScheduleCache,
    setScheduleCache,
    getScheduleSetting,
    setScheduleSetting,
    projectIdsWorkedBy,
    getScan,
    listScans,
    saveScan,
    deleteScans,
    addAiUsage,
    aiUsageSummary,
    getCalendarSync,
    listCalendarSync,
    saveCalendarSync,
    deleteCalendarSync,
    getPeopleSync,
    listPeopleSync,
    peopleSyncCounts,
    savePeopleSync,
    addNotice,
    getNotice,
    listNotices,
    resolveNotice,
    closeOpenNotices,
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
    transition,
    markChecked,
    setReviewTicks,
    reissueReport,
    listTransitions,
    getClientLogo,
    setClientLogo,
    deleteClientLogo,
    assembleFinal,
    listAssemblyProvenance,
    // printed sheets (additive)
    addIssuedSheets,
    getIssuedSheet,
    listIssuedSheets
  };
}

module.exports = {
  createReportStore,
  // Exported for the migration script and tests (shared schema, no duplication).
  openDatabase,
  resolveDbPath,
  extractSearchColumns
};
