// SharePoint knowledge base (Session 6H.3 / 6H.4): a local full-text index (SQLite FTS5, its own file) of the
// PDF, DOCX and XLSX files of the configured sites, kept current with delta queries.
//
// HARD RULE: strictly READ-ONLY toward SharePoint. This module only calls graph.get / graph.getAll (the client
// refuses any other method on /sites and /drives), never renames, moves, uploads or deletes anything, and does
// not depend on the folder structure being changed. Files are streamed into memory, their text extracted, and
// the bytes dropped; only title, path, dates, link and text chunks are stored.
//
// What is indexed:
//   client sites   a project folder at the root of Shared Documents, named NNN-D03_TS-YY-CLIENT or
//                  NNN_MM_YY_TS_CLIENT_ESHA; only its Submitted, Reference and Input folders (Working = drafts is
//                  skipped, wherever it is). A file in a folder that matches neither name is indexed with no
//                  project, and only reviewer and above can find it.
//   general sites  only the configured folder(s); visible to every signed-in user.
//
// Who finds what (enforced in the SQL, not in the browser): admin / approver / reviewer search everything; an
// engineer finds the general documents and the client documents of projects they worked on.

const fs = require('fs');
const path = require('path');
const Database = require('better-sqlite3');
const { GraphError } = require('./graphClient');
const { extractText, chunkText } = require('./kbExtract');
const { hasRole } = require('./permissions');

const EXTENSIONS = new Set(['pdf', 'docx', 'xlsx']);
const INDEXED_SECTIONS = new Set(['submitted', 'reference', 'input']);
const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const lower = (value) => text(value).toLowerCase();
const safeDecode = (value) => { try { return decodeURIComponent(value); } catch (_error) { return value; } };

// The project a folder name stands for, in a form that compares equal across both naming styles, or ''.
function projectKeyOf(name) {
  const value = text(name);
  const fresh = /^(\d{3})-D03_TS-(\d{2})-([A-Za-z0-9]+)/i.exec(value);
  if (fresh) return `${fresh[1]}-D03_TS-${fresh[2]}-${fresh[3]}`.toUpperCase();
  const old = /^(\d{3})_(\d{2})_(\d{2})_TS_([A-Za-z0-9]+)_ESHA/i.exec(value);
  if (old) return `${old[1]}_${old[2]}_${old[3]}_TS_${old[4]}_ESHA`.toUpperCase();
  return '';
}

// Decides whether a file is indexed. `segments` = [library, folder..., file name].
function classify(site, segments) {
  const ext = lower(segments[segments.length - 1].split('.').pop());
  if (segments.length < 2 || !EXTENSIONS.has(ext)) return { include: false };
  const folders = segments.slice(1, -1);
  if (site.scope === 'general') {
    const inside = !site.folders.length || site.folders.some((entry) => {
      const wanted = entry.split('/').map(lower);
      return wanted.every((part, index) => lower(segments[index]) === part);
    });
    return inside ? { include: true, ext, projectKey: '', section: '' } : { include: false };
  }
  if (folders.some((folder) => lower(folder) === 'working')) return { include: false };
  const projectKey = folders.length ? projectKeyOf(folders[0]) : '';
  if (!projectKey) return { include: true, ext, projectKey: '', section: '' };
  const section = folders[1] ? lower(folders[1]) : '';
  return INDEXED_SECTIONS.has(section) ? { include: true, ext, projectKey, section } : { include: false };
}

function createKnowledgeBase({ graph, config, store, log = console, now = () => new Date() }) {
  const dbPath = config.kb.dbPath;
  if (dbPath !== ':memory:') fs.mkdirSync(path.dirname(dbPath), { recursive: true });
  const db = new Database(dbPath);
  db.pragma('journal_mode = WAL');
  db.exec(`
    CREATE TABLE IF NOT EXISTS kb_files (
      id INTEGER PRIMARY KEY, drive_key TEXT NOT NULL, item_id TEXT NOT NULL, site_name TEXT NOT NULL, scope TEXT NOT NULL,
      path TEXT NOT NULL, title TEXT NOT NULL, ext TEXT NOT NULL, modified TEXT, size INTEGER, web_url TEXT,
      project_key TEXT NOT NULL DEFAULT '', section TEXT NOT NULL DEFAULT '', etag TEXT, state TEXT NOT NULL, note TEXT, indexed_at TEXT,
      UNIQUE (drive_key, item_id)
    );
    CREATE INDEX IF NOT EXISTS idx_kb_files_project ON kb_files(project_key);
    CREATE TABLE IF NOT EXISTS kb_state (drive_key TEXT PRIMARY KEY, site_name TEXT, delta_link TEXT, refreshed_at TEXT, error TEXT);
    CREATE VIRTUAL TABLE IF NOT EXISTS kb_chunks USING fts5(title, text, file_id UNINDEXED, chunk_no UNINDEXED, tokenize = 'porter unicode61');
  `);

  const q = {
    file: db.prepare('SELECT * FROM kb_files WHERE drive_key = ? AND item_id = ?'),
    fileById: db.prepare('SELECT * FROM kb_files WHERE id = ?'),
    insertFile: db.prepare(`INSERT INTO kb_files (drive_key, item_id, site_name, scope, path, title, ext, modified, size, web_url, project_key, section, etag, state, note, indexed_at)
      VALUES (@drive_key, @item_id, @site_name, @scope, @path, @title, @ext, @modified, @size, @web_url, @project_key, @section, @etag, @state, @note, @indexed_at)`),
    updateFile: db.prepare(`UPDATE kb_files SET site_name=@site_name, scope=@scope, path=@path, title=@title, ext=@ext, modified=@modified, size=@size, web_url=@web_url, project_key=@project_key,
      section=@section, etag=@etag, state=@state, note=@note, indexed_at=@indexed_at WHERE id=@id`),
    deleteChunks: db.prepare('DELETE FROM kb_chunks WHERE file_id = ?'),
    deleteFile: db.prepare('DELETE FROM kb_files WHERE id = ?'),
    addChunk: db.prepare('INSERT INTO kb_chunks (title, text, file_id, chunk_no) VALUES (?, ?, ?, ?)'),
    state: db.prepare('SELECT * FROM kb_state WHERE drive_key = ?'),
    saveState: db.prepare('INSERT INTO kb_state (drive_key, site_name, delta_link, refreshed_at, error) VALUES (@drive_key, @site_name, @delta_link, @refreshed_at, @error) ON CONFLICT(drive_key) DO UPDATE SET site_name=@site_name, delta_link=@delta_link, refreshed_at=@refreshed_at, error=@error'),
    driveFiles: db.prepare('SELECT id, item_id FROM kb_files WHERE drive_key = ?')
  };

  function removeFile(row) {
    q.deleteChunks.run(row.id);
    q.deleteFile.run(row.id);
  }

  function saveFile(existing, meta, chunks) {
    const run = db.transaction(() => {
      let id = existing && existing.id;
      if (id) q.updateFile.run({ ...meta, id });
      else id = q.insertFile.run(meta).lastInsertRowid;
      q.deleteChunks.run(id);
      chunks.forEach((chunk, index) => q.addChunk.run(meta.title, chunk, id, index));
    });
    run();
  }

  // ---- reading a site (graph.get only) ----------------------------------------------------------
  async function resolveDrives(site) {
    const url = new URL(site.url);
    const { data: info } = await graph.get(`/sites/${url.host}:${url.pathname}`);
    const { items: drives } = await graph.getAll(`/sites/${encodeURIComponent(info.id)}/drives`);
    const libraries = [...new Set(site.folders.length ? site.folders.map((folder) => folder.split('/')[0]) : ['Shared Documents'])];
    return libraries.map((library) => {
      const drive = drives.find((entry) => safeDecode(entry.webUrl || '').toLowerCase().endsWith(`/${library.toLowerCase()}`) || lower(entry.name) === library.toLowerCase());
      return drive ? { library, id: drive.id } : null;
    }).filter(Boolean);
  }

  function segmentsOf(item, library) {
    const parent = safeDecode(((item.parentReference || {}).path || '').replace(/^.*?root:/, ''));
    return [library, ...parent.split('/').filter(Boolean), item.name];
  }

  async function refreshDrive(site, drive, tally) {
    const driveKey = `${site.name}|${drive.id}`;
    const saved = q.state.get(driveKey);
    let started = saved && saved.delta_link ? saved.delta_link : `/drives/${drive.id}/root/delta`;
    let full = !(saved && saved.delta_link);
    let result;
    try {
      result = await graph.getAll(started);
    } catch (error) {
      if (!(error instanceof GraphError && error.status === 410 && !full)) throw error;
      full = true;
      started = `/drives/${drive.id}/root/delta`;
      result = await graph.getAll(started);
    }
    const seen = new Set();
    for (const item of result.items) {
      if (!item.id) continue;
      const existing = q.file.get(driveKey, item.id);
      if (item.deleted) {
        if (existing) { removeFile(existing); tally.removed += 1; }
        continue;
      }
      if (item.folder || !item.file) continue;
      seen.add(item.id);
      const segments = segmentsOf(item, drive.library);
      const verdict = classify(site, segments);
      if (!verdict.include) {
        if (existing) { removeFile(existing); tally.removed += 1; }
        continue;
      }
      const meta = {
        drive_key: driveKey, item_id: item.id, site_name: site.name, scope: site.scope, path: segments.join('/'), title: segments[segments.length - 1].replace(/\.[^.]+$/, ''),
        ext: verdict.ext, modified: text(item.lastModifiedDateTime), size: Number(item.size) || 0, web_url: text(item.webUrl), project_key: verdict.projectKey, section: verdict.section,
        etag: text(item.eTag), state: 'indexed', note: '', indexed_at: now().toISOString()
      };
      if (existing && existing.etag === meta.etag && existing.state !== 'error') {
        if (existing.path !== meta.path || existing.web_url !== meta.web_url) { q.updateFile.run({ ...meta, state: existing.state, note: existing.note, id: existing.id }); tally.updated += 1; }
        continue;
      }
      if (meta.size > config.kb.maxMb * 1024 * 1024) {
        saveFile(existing, { ...meta, state: 'skipped', note: `larger than ${config.kb.maxMb} MB` }, []);
        tally.skipped += 1;
        continue;
      }
      try {
        const { data } = await graph.get(`/drives/${drive.id}/items/${encodeURIComponent(item.id)}/content`, { binary: true });
        const body = await extractText(verdict.ext, Buffer.from(data));
        if (!body) {
          saveFile(existing, { ...meta, state: 'no_text', note: 'no text layer (a scan?)' }, [meta.title]);
          tally.noText += 1;
        } else {
          saveFile(existing, meta, chunkText(body));
          tally[existing ? 'updated' : 'added'] += 1;
        }
      } catch (error) {
        saveFile(existing, { ...meta, state: 'error', note: error.status ? `Graph ${error.status}` : 'could not be read' }, []);
        tally.errors += 1;
      }
    }
    if (full) {
      q.driveFiles.all(driveKey).filter((row) => !seen.has(row.item_id)).forEach((row) => { removeFile(row); tally.removed += 1; });
    }
    q.saveState.run({ drive_key: driveKey, site_name: site.name, delta_link: result.deltaLink || '', refreshed_at: now().toISOString(), error: '' });
  }

  let running = null;
  // Reads every configured site. A site that cannot be read keeps its index and records the reason.
  function refresh() {
    if (running) return running;
    running = (async () => {
      const summary = { at: now().toISOString(), sites: [] };
      for (const site of config.kb.sites) {
        const tally = { name: site.name, scope: site.scope, added: 0, updated: 0, removed: 0, skipped: 0, noText: 0, errors: 0, error: '' };
        try {
          const drives = await resolveDrives(site);
          if (!drives.length) throw Object.assign(new Error('library not found'), { code: 'library_not_found' });
          for (const drive of drives) await refreshDrive(site, drive, tally);
        } catch (error) {
          tally.error = `${error.status || ''} ${error.code || error.message}`.trim();
          q.saveState.run({ drive_key: `${site.name}|`, site_name: site.name, delta_link: '', refreshed_at: now().toISOString(), error: tally.error });
          log.error(`Knowledge base: ${site.name} could not be read (${tally.error})`);
        }
        summary.sites.push(tally);
      }
      db.prepare("INSERT INTO kb_state (drive_key, site_name, delta_link, refreshed_at, error) VALUES ('__last__', '', ?, ?, '') ON CONFLICT(drive_key) DO UPDATE SET delta_link = excluded.delta_link, refreshed_at = excluded.refreshed_at").run(JSON.stringify(summary), summary.at);
      return summary;
    })().finally(() => { running = null; });
    return running;
  }

  function start() {
    const minutes = config.kb.refreshMinutes;
    if (!(minutes > 0)) return { stop() {} };
    const run = () => refresh().catch((error) => log.error(`Knowledge base refresh failed: ${error.status || ''} ${error.code || error.message}`.trim()));
    const first = setTimeout(run, 30000);
    const timer = setInterval(run, minutes * 60 * 1000);
    first.unref && first.unref();
    timer.unref && timer.unref();
    return { stop() { clearTimeout(first); clearInterval(timer); } };
  }

  // ---- who finds what ------------------------------------------------------------------------------
  // { all: true } or { projectKeys: [...] } for a signed-in user.
  function accessFor(user) {
    if (hasRole(user, 'admin', 'approver', 'reviewer')) return { all: true };
    const keys = new Set();
    store.projectIdsWorkedBy(user.id).forEach((id) => {
      const project = store.getProject(id);
      const key = project && projectKeyOf(project.projectNo);
      if (key) keys.add(key);
    });
    const tasks = store.getScheduleCache('tasks');
    ((tasks && tasks.data && tasks.data.projects) || []).forEach((project) => {
      if ((project.tasks || []).some((task) => (task.owners || []).some((owner) => lower(owner.email) === lower(user.email)))) {
        const key = projectKeyOf(project.name) || projectKeyOf(project.projectNumber);
        if (key) keys.add(key);
      }
    });
    return { projectKeys: [...keys] };
  }

  const publicFile = (row) => ({ id: row.id, title: row.title, path: row.path, source: row.site_name, scope: row.scope, modified: row.modified, webUrl: row.web_url, projectKey: row.project_key });

  function accessClause(access) {
    if (access.all) return { sql: '1 = 1', args: [] };
    const keys = (access.projectKeys || []).filter(Boolean);
    return keys.length ? { sql: `(f.scope = 'general' OR f.project_key IN (${keys.map(() => '?').join(',')}))`, args: keys } : { sql: "f.scope = 'general'", args: [] };
  }

  // access: from accessFor; restrictTo (optional): extra filter for a report's references.
  function search({ query, access, limit = 20, projectKey = null }) {
    const tokens = (String(query || '').match(/[\p{L}\p{N}]+/gu) || []).slice(0, 8);
    if (!tokens.length) return [];
    const clause = accessClause(access);
    const extra = projectKey === null ? { sql: '1 = 1', args: [] } : { sql: "(f.scope = 'general' OR (f.project_key != '' AND f.project_key = ?))", args: [projectKey] };
    const rows = db.prepare(`SELECT f.*, snippet(kb_chunks, 1, '', '', ' ... ', 24) AS snip, bm25(kb_chunks, 5.0, 1.0) AS score
      FROM kb_chunks JOIN kb_files f ON f.id = kb_chunks.file_id
      WHERE kb_chunks MATCH ? AND ${clause.sql} AND ${extra.sql} AND f.state = 'indexed' ORDER BY score LIMIT ?`).all(tokens.map((token) => `"${token}"*`).join(' '), ...clause.args, ...extra.args, limit * 6);
    const seen = new Set();
    const hits = [];
    rows.forEach((row) => {
      if (seen.has(row.id) || hits.length >= limit) return;
      seen.add(row.id);
      hits.push({ ...publicFile(row), snippet: row.snip });
    });
    return hits;
  }

  // One file by id if `access` allows it (else null): used when a reviewer picks a reference.
  function getFile(id, access) {
    const row = q.fileById.get(Number(id));
    if (!row || row.state !== 'indexed') return null;
    const clause = accessClause(access);
    const allowed = db.prepare(`SELECT 1 FROM kb_files f WHERE f.id = ? AND ${clause.sql}`).get(row.id, ...clause.args);
    return allowed ? publicFile(row) : null;
  }

  function status() {
    const bySite = db.prepare("SELECT site_name AS site, scope, state, COUNT(*) AS n FROM kb_files GROUP BY site_name, scope, state").all();
    const states = db.prepare("SELECT drive_key AS key, site_name AS site, refreshed_at AS at, error FROM kb_state WHERE drive_key != '__last__'").all();
    const last = db.prepare("SELECT delta_link AS summary FROM kb_state WHERE drive_key = '__last__'").get();
    const sites = config.kb.sites.map((site) => ({
      name: site.name, scope: site.scope, folders: site.folders,
      files: bySite.filter((row) => row.site === site.name && row.state === 'indexed').reduce((sum, row) => sum + row.n, 0),
      noText: bySite.filter((row) => row.site === site.name && row.state === 'no_text').reduce((sum, row) => sum + row.n, 0),
      problems: bySite.filter((row) => row.site === site.name && ['error', 'skipped'].includes(row.state)).reduce((sum, row) => sum + row.n, 0),
      refreshedAt: (states.filter((row) => row.site === site.name && row.at).map((row) => row.at).sort().pop()) || '',
      error: (states.find((row) => row.site === site.name && row.error) || {}).error || ''
    }));
    return { sites, last: last && last.summary ? JSON.parse(last.summary) : null, running: Boolean(running) };
  }

  return { refresh, start, search, getFile, accessFor, status, close: () => db.close(), db };
}

module.exports = { createKnowledgeBase, projectKeyOf, classify };
