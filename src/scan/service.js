// Scanned sheets, end to end (Session 6I): upload -> pages -> which sheet it is (QR, else the printed id, else the
// engineer says) -> the provider transcribes -> the engineer reviews -> Accept writes into the project's daily report.
//
// Nothing reaches a report before Accept, and Accept goes through the same path as a hand-typed save: the report is
// built, normalizeReportInput and the field validation run, the rows are credited to the person who accepts (6G).
// A provider failure leaves the scan 'failed' with the reason; it can be retried and never blocks anything else.

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const { getReviewStatus, getCollectionDate } = require('../reportCollectionModel');
const { normalizeReportInput, ReportValidationError } = require('../reportModel');
const { stampReport } = require('../entryCredit');
const { buildWorkspace, findDay } = require('../workspace');
const { mayEditReport, hasRole } = require('../permissions');
const { parseSheetId } = require('../sheetPrint/ids');
const { SPECS } = require('../sheetPrint/spec');
const { pageCapacity } = require('../sheetPrint/render');
const { ProviderError } = require('../aiGuard');
const { splitUpload } = require('./split');
const { findSheetId } = require('./qr');
const { buildForm } = require('./form');
const { readPage, applyEdits } = require('./reader');
const { applyScan } = require('./apply');

class ScanError extends Error {
  constructor(status, code, message, extra = {}) {
    super(message);
    this.name = 'ScanError';
    this.status = status;
    this.code = code;
    this.extra = extra;
  }
}

const ISO_DAY = /^\d{4}-\d{2}-\d{2}$/;
const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const lower = (value) => text(value).toLowerCase();

function createScanService({ store, config, provider, guard, models, log = console, now = () => new Date() }) {
  const dir = config.dir;
  const imagePath = (scan) => path.join(dir, scan.imagePath);
  let chain = Promise.resolve();
  const enqueue = (job) => {
    const run = chain.then(job, job);
    chain = run.catch(() => {});
    return run;
  };
  const guarded = {
    read: (args) => guard.run('gemini', 'ocr', args.model, () => provider.read(args)),
    readSheetId: (args) => (provider.readSheetId ? guard.run('gemini', 'ocr-id', args.model, () => provider.readSheetId(args)) : Promise.resolve(''))
  };

  const formOf = (scan) => buildForm({ ...scan.issued, sheetId: scan.issued.sheetId || '' });
  const flaggedCount = (scan) => Object.values(scan.cells || {}).filter((cell) => cell.flags && cell.flags.length).length;

  const summary = (scan) => ({
    id: scan.id,
    batchId: scan.batchId,
    status: scan.status,
    error: scan.error || null,
    fileName: scan.fileName,
    pageIndex: scan.pageIndex,
    sheetId: scan.sheetId,
    sheetKey: scan.sheetKey,
    sheetCode: scan.issued && SPECS[scan.issued.sheetKey] ? SPECS[scan.issued.sheetKey].code : '',
    projectNo: scan.projectNo,
    clientName: (scan.issued && scan.issued.project && scan.issued.project.clientName) || '',
    workDate: scan.workDate,
    createdAt: scan.createdAt,
    acceptedAt: scan.acceptedAt,
    flagged: scan.status === 'read' ? flaggedCount(scan) : 0,
    note: scan.note || ''
  });

  const canView = (user, scan) => scan.createdBy === user.id || hasRole(user, 'reviewer', 'approver', 'admin');
  const mayChange = (user, scan) => scan.createdBy === user.id || hasRole(user, 'admin');

  function scanFor(id, user, { change = false } = {}) {
    const scan = store.getScan(id);
    if (!scan || !canView(user, scan)) throw new ScanError(404, 'NOT_FOUND', 'Scan not found.');
    if (change && !mayChange(user, scan)) throw new ScanError(403, 'NOT_OWNER', 'Only the person who uploaded this scan (or an admin) can change it.');
    return scan;
  }

  // ---- the background job -------------------------------------------------------------------------
  async function identify(scan) {
    const png = fs.readFileSync(imagePath(scan));
    const [page] = await splitUpload({ buffer: png, mimetype: 'image/png' });
    const found = findSheetId(page.rgba);
    let sheetId = found ? found.sheetId : '';
    let how = found ? `QR (${found.attempt})` : '';
    if (!sheetId) {
      try {
        sheetId = parseSheetId(await guarded.readSheetId({ png, model: models.primary })) || '';
        how = sheetId ? 'printed sheet id' : '';
      } catch (_error) {
        sheetId = ''; // the provider is not needed to ask the engineer
      }
    }
    if (!sheetId) return { reason: 'The QR code could not be read.' };
    const issued = store.getIssuedSheet(sheetId);
    return issued ? { issued, how } : { reason: `Sheet ${sheetId} was not issued by this app.` };
  }

  async function process(id) {
    const scan = store.getScan(id);
    if (!scan || scan.status !== 'queued') return;
    store.saveScan(id, { status: 'reading', error: null });
    try {
      let issued = scan.issued;
      let how = scan.identified || '';
      if (!issued) {
        const found = await identify(scan);
        if (!found.issued) {
          store.saveScan(id, { status: 'needs_sheet', note: found.reason });
          return;
        }
        issued = found.issued;
        how = found.how;
      }
      const form = buildForm({ ...issued, sheetId: issued.sheetId || '' });
      const result = await readPage({ form, png: fs.readFileSync(imagePath(scan)), provider: guarded, models, threshold: config.minConfidence });
      store.saveScan(id, { status: 'read', issued, identified: how, sheetId: issued.sheetId || '', sheetKey: issued.sheetKey, projectNo: (issued.project && issued.project.projectNo) || '', workDate: issued.workDate, cells: result.cells, models: result.models, escalated: result.escalated, note: '', error: null });
    } catch (error) {
      const failure = error instanceof ProviderError ? { code: error.code, message: error.message, retryable: error.retryable } : { code: 'INTERNAL', message: 'The sheet could not be read.', retryable: true };
      if (!(error instanceof ProviderError)) log.error(`Scan ${id} failed: ${error && error.name}`);
      store.saveScan(id, { status: 'failed', error: failure });
    }
  }

  // ---- uploads ----------------------------------------------------------------------------------
  async function ingest({ buffer, mimetype, fileName, user }) {
    const pages = await splitUpload({ buffer, mimetype, maxPages: config.maxPages });
    fs.mkdirSync(dir, { recursive: true });
    const batchId = `SB-${crypto.randomBytes(5).toString('hex').toUpperCase()}`;
    const scans = pages.map((page) => {
      const id = `SC-${now().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14)}-${crypto.randomBytes(3).toString('hex').toUpperCase()}`;
      fs.writeFileSync(path.join(dir, `${id}.png`), page.png);
      return store.saveScan(id, { batchId, createdBy: user.id, pageIndex: page.index, fileName: text(fileName).slice(0, 120), imagePath: `${id}.png`, status: 'queued', width: page.width, height: page.height });
    });
    scans.forEach((scan) => enqueue(() => process(scan.id)));
    return scans.map(summary);
  }

  // A page with no readable QR: the engineer says which sheet it is.
  function assign(id, user, { project, workDate, sheetKey, rowOffset = 0 }) {
    const scan = scanFor(id, user, { change: true });
    if (!['needs_sheet', 'failed', 'queued'].includes(scan.status) || (scan.status === 'failed' && scan.issued)) throw new ScanError(409, 'NOT_ASSIGNABLE', 'This scan already knows which sheet it is.');
    if (!ISO_DAY.test(text(workDate))) throw new ScanError(400, 'BAD_DATE', 'Choose the work date.');
    const spec = SPECS[sheetKey];
    if (!spec) throw new ScanError(400, 'BAD_SHEET', 'Choose which sheet this is.');
    const workspace = project && store.getProject(String(project));
    if (!workspace) throw new ScanError(400, 'BAD_PROJECT', 'Choose the project.');
    const offset = Math.max(0, Number.parseInt(rowOffset, 10) || 0);
    const issued = {
      sheetId: '', sheetKey, sheetCode: spec.code, workDate: text(workDate), page: sheetKey === 'soilResistivity' ? offset + 1 : 1, pages: 1,
      project: { projectNo: workspace.projectNo, clientName: workspace.clientName, zohoProjectId: workspace.zohoProjectId || '', siteLocation: workspace.siteLocation || '' },
      rowOffset: offset, rowCount: pageCapacity(spec).rows, preRows: [], manual: true
    };
    store.saveScan(id, { status: 'queued', issued, identified: 'chosen by the engineer', note: '' });
    enqueue(() => process(id));
    return summary(store.getScan(id));
  }

  function retry(id, user) {
    const scan = scanFor(id, user, { change: true });
    if (scan.status !== 'failed') throw new ScanError(409, 'NOT_FAILED', 'Only a failed scan can be read again.');
    store.saveScan(id, { status: 'queued', error: null });
    enqueue(() => process(id));
    return summary(store.getScan(id));
  }

  // ---- reading and editing ----------------------------------------------------------------------
  const slimCell = (def) => ({ path: def.path, section: def.section, row: def.row, item: def.item, key: def.key, label: def.label, kind: def.kind, options: def.options || null, multi: Boolean(def.multi), preprinted: Boolean(def.preprinted), value: def.value || '', unit: def.unit || '', date: Boolean(def.date) });

  function detail(id, user) {
    const scan = scanFor(id, user);
    const base = { ...summary(scan), identified: scan.identified || '', models: scan.models || null, escalated: Boolean(scan.escalated), acceptedReportId: scan.acceptedReportId || '', overwritten: (scan.overwrites || []).length, canChange: mayChange(user, scan) };
    if (!scan.issued) return base;
    const form = formOf(scan);
    return { ...base, form: { sheetKey: form.sheetKey, code: form.code, title: form.title, layout: form.layout, group: form.group, rowCount: form.rowCount, rowOffset: form.rowOffset, landscape: form.landscape, spacings: form.spacings, columns: form.columns, cells: form.cells.map(slimCell) }, cells: scan.cells || null, threshold: config.minConfidence };
  }

  function edit(id, user, edits) {
    const scan = scanFor(id, user, { change: true });
    if (scan.status !== 'read') throw new ScanError(409, 'NOT_READ', 'Only a scan that has been read can be edited.');
    const form = formOf(scan);
    const cells = applyEdits(form, scan.cells, edits, config.minConfidence);
    store.saveScan(id, { cells });
    return detail(id, user);
  }

  function image(id, user) {
    const scan = scanFor(id, user);
    return fs.readFileSync(imagePath(scan));
  }

  function discard(id, user) {
    const scan = scanFor(id, user, { change: true });
    if (scan.status === 'accepted') throw new ScanError(409, 'ACCEPTED', 'An accepted scan is kept as the record of what was read.');
    store.saveScan(id, { status: 'discarded' });
    return summary(store.getScan(id));
  }

  function list(user, { all = false } = {}) {
    const scans = store.listScans({ createdBy: all && hasRole(user, 'reviewer', 'approver', 'admin') ? '' : user.id, limit: 200 });
    return scans.map(summary);
  }

  // ---- Accept -----------------------------------------------------------------------------------
  const findWorkspace = (info = {}) => {
    const all = store.listProjects();
    const byZoho = text(info.zohoProjectId) && all.find((project) => text(project.zohoProjectId) === text(info.zohoProjectId));
    return byZoho || all.find((project) => text(info.projectNo) && lower(project.projectNo) === lower(info.projectNo)) || null;
  };

  // Without `confirm`, a scan that would overwrite existing values is not written: it answers NEEDS_CONFIRM with the list.
  function accept(id, user, { confirm = false } = {}) {
    const scan = scanFor(id, user, { change: true });
    if (scan.status !== 'read') throw new ScanError(409, 'NOT_READ', scan.status === 'accepted' ? 'This scan was already accepted.' : 'Only a scan that has been read can be accepted.');
    const form = formOf(scan);
    const issued = scan.issued;
    const project = findWorkspace(issued.project);
    if (!project) throw new ScanError(409, 'NO_WORKSPACE', 'This project has no workspace yet: open it from Home first, then accept the scan.');
    const day = findDay(store, project.id, issued.workDate);
    let base;
    if (day) {
      if (!mayEditReport(user, day)) throw new ScanError(403, 'NOT_OWNER', day.createdBy ? 'Only the engineer who created this day (or an admin) can add to it.' : 'This day has no creator: only an admin can add to it.');
      if (getReviewStatus(day) !== 'draft') throw new ScanError(409, 'NOT_DRAFT', 'This day is no longer a draft. Return it to draft from the review screen first.');
      base = day;
    } else {
      if (!hasRole(user, 'engineer')) throw new ScanError(403, 'NOT_ENGINEER', 'Only an engineer can start a daily report.');
      const template = buildWorkspace(store, project.id).dayTemplate;
      base = { project: { ...template.project, reportDate: issued.workDate }, tests: { ...template.tests }, sheetHeaders: {}, soilResistivity: { locations: [] } };
    }
    const { candidate, overwrites, previous, written } = applyScan({ report: base, form, cells: scan.cells });
    if (!written) throw new ScanError(422, 'NOTHING_TO_ACCEPT', 'Nothing was read from this sheet: every cell is blank.');
    if (overwrites.length && !confirm) throw new ScanError(409, 'NEEDS_CONFIRM', `${overwrites.length} existing value${overwrites.length === 1 ? '' : 's'} would be overwritten.`, { overwrites });
    let normalized;
    try {
      normalized = normalizeReportInput({ ...candidate, project: { ...candidate.project, reportDate: issued.workDate } });
    } catch (error) {
      throw new ScanError(422, 'INVALID_REPORT', error instanceof Error ? error.message : 'The report would not be valid.', { errors: error instanceof ReportValidationError ? error.errors : undefined });
    }
    const actor = { id: user.id, name: user.name, email: user.email };
    let saved;
    if (day) {
      saved = store.updateReport(day.id, (current) => stampReport({ ...normalized, collectionDate: getCollectionDate(day) }, current, actor));
    } else {
      saved = store.addReport({ ...stampReport(normalized, null, actor), kind: 'daily', projectId: project.id, collectionDate: issued.workDate, createdBy: user.id });
    }
    store.saveScan(id, { status: 'accepted', acceptedAt: now().toISOString(), acceptedBy: user.id, acceptedReportId: saved.id, createdDay: !day, previous, overwrites });
    return { reportId: saved.id, projectId: project.id, workDate: issued.workDate, createdDay: !day, written, overwrites: overwrites.length, flaggedRemaining: flaggedCount(scan) };
  }

  // Admin: deletes the scans created in a date range, images included. Returns the number removed.
  function purge({ from, to }) {
    if (!ISO_DAY.test(text(from)) || !ISO_DAY.test(text(to)) || from > to) throw new ScanError(400, 'BAD_RANGE', 'Give a from and a to date (YYYY-MM-DD), from not after to.');
    const removed = store.deleteScans({ from, to });
    removed.forEach((entry) => {
      if (entry.imagePath) fs.rmSync(path.join(dir, entry.imagePath), { force: true });
    });
    return { removed: removed.length, from, to };
  }

  return { ingest, assign, retry, detail, edit, image, discard, list, accept, purge, idle: () => chain, summary };
}

module.exports = { createScanService, ScanError };
