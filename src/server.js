const fs = require('fs');
const path = require('path');
const express = require('express');
const multer = require('multer');

const config = require('./config');
const { createReportStore } = require('./reportStore');
const { normalizeReportInput, TEST_LIBRARY, ReportValidationError } = require('./reportModel');
const { publicDefinitions } = require('./sheetDefinitions');
const { listProfiles } = require('./acceptance');
const { printSheets, SheetPrintError } = require('./sheetPrint');
const { generateElectroReportPdf } = require('./pdfService');
const { assessReport } = require('./assessmentService');
const { scrubGeneratedRemarks } = require('./remarks');
const {
  getAiStatus,
  generateReportNarrative,
  generateMeasurementObservations,
  sanityCheckNarrative
} = require('./aiService');
const { getReportKind, getReviewStatus } = require('./reportCollectionModel');
const { ensureProject, buildWorkspace, listProjectSummaries, findDay, assembleFor, pickProjectFields } = require('./workspace');
const { buildReviewState } = require('./review');
const { exportReportPdf } = require('./finalReport/service');
const { ConsistencyError } = require('./finalReport/data');
const { LogoError, parseLogoDataUrl, clientKeys } = require('./finalReport/logos');
const { getProjects, getProject, getProjectUsers, getPortalUsers, filterProjectsForPicker, getProjectClientUsers, getProjectTasks } = require('./zohoClient');
const { isD03Name, buildTaskPlan, nameOwners } = require('./zohoTasks');
const { pollNewProjects, answerNotice, dismissNotice, startProjectWatch } = require('./projectWatch');
const { createScheduleService } = require('./scheduleService');
const { createPeopleClient } = require('./zohoPeople');
const { createPeopleSync } = require('./peopleSync');
const { stampReport } = require('./entryCredit');
const { createAuth } = require('./auth');
const { reviewGate, reviewAccess, mayEditReport, hasRole } = require('./permissions');
const {
  previewSoilSheetWithGemini,
  previewElectrodeSheetWithGemini,
  previewContinuitySheetWithGemini,
  previewLoopImpedanceSheetWithGemini,
  previewProspectiveFaultSheetWithGemini,
  previewRiserIntegritySheetWithGemini,
  previewEarthContinuitySheetWithGemini,
  previewTowerFootingSheetWithGemini
} = require('./geminiOcrService');
const {
  mapSoilOcrPreviewToDraft,
  mapElectrodeOcrPreviewToDraft,
  mapContinuityOcrPreviewToDraft,
  mapLoopImpedanceOcrPreviewToDraft,
  mapProspectiveFaultOcrPreviewToDraft,
  mapRiserIntegrityOcrPreviewToDraft,
  mapEarthContinuityOcrPreviewToDraft,
  mapTowerFootingOcrPreviewToDraft
} = require('./ocrMappers');

const app = express();
const rootDir = config.app.rootDir;
const publicDir = config.app.publicDir;
const generatedDir = config.app.generatedDir;
const dataFilePath = config.app.dataFilePath;
const store = createReportStore({ filePath: dataFilePath });
const port = config.app.port;
const host = config.app.host;
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: config.gemini.maxUploadBytes
  }
});

if (!fs.existsSync(generatedDir)) {
  fs.mkdirSync(generatedDir, { recursive: true });
}

app.use(express.json({ limit: '25mb' }));

// Scheduling (6F): Zoho Projects tasks and Zoho People leave, read only and cached.
const peopleClient = createPeopleClient();
const schedule = createScheduleService({ store, zoho: require('./zohoClient'), people: peopleClient });
// Site Work Log (6G): the one thing written to Zoho People, on approval only. Never blocks a review step.
const peopleSync = createPeopleSync({ store, people: peopleClient, config, getWorkingDays: () => schedule.getSettings().workingDays });

// Sign-in (Session 6E): the session user on every API / auth / PDF request, then the route rules.
const auth = createAuth({ store, config, zoho: require('./zohoClient') });
app.use(['/api', '/auth', '/generated-pdfs'], auth.attachUser, auth.sameOrigin);
app.use('/generated-pdfs', auth.requireSignIn, express.static(generatedDir));
app.use(
  express.static(publicDir, {
    setHeaders(res, filePath) {
      if (filePath.endsWith('index.html') || filePath.endsWith('app.js') || filePath.endsWith('styles.css')) {
        res.setHeader('Cache-Control', 'no-store');
      }
    }
  })
);
auth.register(app);
app.use('/api', auth.requireSignIn, auth.authorize);
auth.registerProtected(app); // profile and admin routes: behind the guard

app.get('/api/health', (_req, res) => {
  res.json({
    ok: true,
    service: 'electroreports',
    date: new Date().toISOString()
  });
});

// The approved measurement sheets (MS-01..MS-09) as data: header fields and row columns
// in printed order. The browser renders every entry screen from these.
app.get('/api/sheets', (_req, res) => {
  res.json(publicDefinitions());
});

// Printable measurement sheets for a site visit: project + work date + selected tests in, one
// PDF out. Every printed page is recorded with its sheet ID (the QR code) in `issued_sheets`.
app.post('/api/sheets/print', async (req, res) => {
  try {
    const result = await printSheets(req.body, { store, getTasks: (zohoProjectId) => getProjectTasks(zohoProjectId) });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `inline; filename="ESHA_Sheets_${String(req.body?.workDate || '').replace(/[^0-9-]/g, '')}.pdf"`);
    res.setHeader('X-Sheet-Batch', result.batchId);
    res.setHeader('X-Sheet-Ids', result.pages.map((page) => page.sheetId).join(','));
    res.setHeader('X-Print-Engineer', encodeURIComponent(result.engineer));
    if (result.warnings.length) {
      res.setHeader('X-Print-Warnings', encodeURIComponent(result.warnings.join(' ')));
    }
    res.setHeader('Access-Control-Expose-Headers', 'X-Sheet-Batch, X-Sheet-Ids, X-Print-Engineer, X-Print-Warnings');
    return res.send(result.buffer);
  } catch (error) {
    if (error instanceof SheetPrintError) {
      return res.status(error.status).json({ message: error.message });
    }
    console.error('Failed to print sheets', error);
    return res.status(500).json({ message: 'Failed to generate the sheets.' });
  }
});

// What was issued: by sheet ID (what a scan matches against), or by batch / project / date.
app.get('/api/sheets/issued/:id', (req, res) => {
  const sheet = store.getIssuedSheet(req.params.id);
  return sheet ? res.json(sheet) : res.status(404).json({ message: 'Sheet ID not found.' });
});

app.get('/api/sheets/issued', (req, res) => {
  res.json(store.listIssuedSheets({
    batchId: String(req.query.batch || '').trim(),
    projectNo: String(req.query.projectNo || '').trim(),
    workDate: String(req.query.workDate || '').trim()
  }));
});

// Acceptance profiles a project can pick (from the register; templates excluded).
app.get('/api/acceptance/profiles', (_req, res) => {
  res.json(listProfiles());
});

app.get('/api/catalog', (_req, res) => {
  res.json({ tests: TEST_LIBRARY });
});

app.get('/api/ai/status', (_req, res) => {
  res.json(getAiStatus());
});

app.get('/api/ocr/status', (_req, res) => {
  res.json({
    configured: Boolean(config.gemini.enabled && config.gemini.apiKey),
    enabled: Boolean(config.gemini.enabled),
    model: String(config.gemini.model || '').trim(),
    escalationModel: String(config.gemini.escalationModel || '').trim(),
    supportedSheets: [
      'soilResistivity',
      'electrodeResistance',
      'continuityTest',
      'loopImpedanceTest',
      'prospectiveFaultCurrent',
      'riserIntegrityTest',
      'earthContinuityTest',
      'towerFootingResistance'
    ]
  });
});

function sendOcrError(res, error, fallbackMessage) {
  const status = Number.isInteger(error?.status) ? error.status : 500;
  return res.status(status).json({
    message: error instanceof Error ? error.message : fallbackMessage,
    code: error?.code || undefined,
    details: error?.details || undefined
  });
}

function createOcrPreviewHandler(previewFn, mapFn, fallbackMessage) {
  return async (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({
          message: 'Please choose a sheet image or PDF to scan.',
          code: 'OCR_FILE_MISSING'
        });
      }

      const preview = await previewFn(config, req.file);
      const draftPatch = mapFn(preview);

      return res.json({
        success: true,
        preview,
        draftPatch,
        fileName: req.file.originalname || 'uploaded-sheet',
        mimeType: req.file.mimetype || '',
        fileSize: req.file.size || 0,
        scannedAt: new Date().toISOString()
      });
    } catch (error) {
      console.error(fallbackMessage, error);
      return sendOcrError(res, error, fallbackMessage);
    }
  };
}

app.get('/api/zoho/projects', async (req, res) => {
  try {
    const query = String(req.query.q || '').trim();
    const projects = await getProjects(query);
    return res.json(filterProjectsForPicker(projects));
  } catch (error) {
    console.error('Failed to load Zoho projects', error);
    return res.status(500).json({ message: 'Failed to load Zoho projects.' });
  }
});

app.get('/api/zoho/projects/:id/users', async (req, res) => {
  try {
    const users = await getProjectUsers(req.params.id);
    return res.json(users);
  } catch (error) {
    console.error('Failed to load Zoho project users', error);
    return res.status(500).json({ message: 'Failed to load Zoho project users.' });
  }
});

app.get('/api/zoho/users', async (_req, res) => {
  try {
    const users = await getPortalUsers();
    return res.json(users);
  } catch (error) {
    console.error('Failed to load Zoho users', error);
    return res.status(500).json({ message: 'Failed to load Zoho users.' });
  }
});

// The step 1 task plan of a project: tasklists and tasks, the measurement tasks (pre-selected), the
// planned quantity per test with its hint / warnings, and the engineers (task owners).
async function loadTaskPlan(projectRef) {
  const project = await getProject(projectRef);
  const d03 = Boolean(project) && isD03Name(project.name);
  const tasks = await getProjectTasks(projectRef);
  let users = [];
  try {
    users = await getProjectUsers(projectRef);
  } catch (_error) {
    users = [];
  }
  return { project: project ? { id: project.id, name: project.name, stage: project.stage } : { id: projectRef, name: '' }, plan: nameOwners(buildTaskPlan(tasks, { d03 }), users) };
}

app.get('/api/zoho/projects/:id/plan', async (req, res) => {
  try {
    return res.json(await loadTaskPlan(req.params.id));
  } catch (error) {
    console.error('Failed to load the Zoho task plan', error);
    return res.status(500).json({ message: 'Failed to load the Zoho project tasks.' });
  }
});

app.get('/api/zoho/projects/:id/client-users', async (req, res) => {
  try {
    const users = await getProjectClientUsers(req.params.id);
    return res.json(users);
  } catch (error) {
    console.error('Failed to load Zoho client users', error);
    return res.status(500).json({ message: 'Failed to load Zoho client users.' });
  }
});

app.get('/api/zoho/projects/:id/tasks', async (req, res) => {
  try {
    const tasks = await getProjectTasks(req.params.id, {
      query: String(req.query.q || '').trim(),
      status: String(req.query.status || '').trim(),
      index: req.query.index,
      range: req.query.range
    });
    return res.json(tasks);
  } catch (error) {
    console.error('Failed to load Zoho project tasks', error);
    return res.status(500).json({ message: 'Failed to load Zoho project tasks.' });
  }
});

// Saved reports may carry remark text generated by older versions; clear it on load so
// a stored sentence can never contradict the status computed from the same row.
function loadReport(id) {
  const report = store.getReport(id);
  return report ? scrubGeneratedRemarks(report) : report;
}

// What the browser shows: the report plus its computed assessment. The assessment is
// attached on the way out and is never stored.
function presentReport(report) {
  const scrubbed = scrubGeneratedRemarks(report);
  try {
    return { ...scrubbed, assessment: assessReport(scrubbed) };
  } catch (error) {
    return { ...scrubbed, assessment: { error: error instanceof Error ? error.message : 'Assessment failed.' } };
  }
}

app.get('/api/reports', (req, res) => {
  const query = String(req.query.q || '').trim();
  res.json(store.listReports(query).map(presentReport));
});

app.get('/api/reports/:id', (req, res) => {
  const report = store.getReport(req.params.id);
  if (!report) {
    return res.status(404).json({ message: 'Report not found.' });
  }
  return res.json(presentReport(report));
});

// Computed assessments for a report or an unsaved draft (statuses, priorities, limits,
// basis, remarks, soil class, summaries). The UI renders these and holds no rules.
app.post('/api/assessments', (req, res) => {
  try {
    return res.json(assessReport(req.body));
  } catch (error) {
    return res.status(400).json({
      message: error instanceof Error ? error.message : 'Invalid report payload.'
    });
  }
});

app.post('/api/reports', (req, res) => {
  try {
    const normalized = normalizeReportInput(req.body);
    const created = store.addReport({ ...normalized, createdBy: req.user.id });
    return res.status(201).json(presentReport(created));
  } catch (error) {
    return res.status(400).json({
      message: error instanceof Error ? error.message : 'Invalid report payload.',
      errors: error instanceof ReportValidationError ? error.errors : undefined
    });
  }
});

app.post('/api/reports/:id/ai/generate', async (req, res) => {
  const report = loadReport(req.params.id);
  if (!report) {
    return res.status(404).json({ message: 'Report not found.' });
  }

  try {
    // PRIMARY: Claude writes the narrative (falls back to OpenAI inside aiService).
    const narrative = await generateReportNarrative(report);

    // SECONDARY (OpenAI): observations/recommendations for measurement sheets and
    // a sanity-check of Claude's narrative. Both are best-effort — a failure here
    // must not block the report from receiving its narrative.
    let observations = null;
    let sanityCheck = null;
    try {
      observations = await generateMeasurementObservations(report);
    } catch (obsError) {
      console.error('OpenAI observation generation failed (non-blocking).', obsError);
    }
    try {
      sanityCheck = await sanityCheckNarrative(report, narrative);
    } catch (checkError) {
      console.error('OpenAI narrative sanity-check failed (non-blocking).', checkError);
    }

    const updated = store.updateReport(req.params.id, (current) => ({
      ...current,
      aiNarrative: narrative,
      aiObservations: observations,
      aiSanityCheck: sanityCheck
    }));
    return res.json(presentReport(updated));
  } catch (error) {
    console.error('Failed to generate ElectroReports AI narrative', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to generate AI narrative.'
    });
  }
});

app.delete('/api/reports/:id', (req, res) => {
  const existing = store.getReport(req.params.id);
  if (existing && !mayEditReport(req.user, existing)) {
    return res.status(403).json({ message: existing.createdBy ? 'Only the engineer who created this report (or an admin) can delete it.' : 'This report has no creator: only an admin can delete it.', code: 'NOT_OWNER' });
  }
  const result = store.deleteReport(req.params.id);
  if (!result.deleted) {
    return res.status(404).json({ message: 'Report not found.' });
  }
  return res.status(204).send();
});

app.post('/api/reports/:id/pdf', async (req, res) => {
  const report = loadReport(req.params.id);
  if (!report) {
    return res.status(404).json({ message: 'Report not found.' });
  }

  try {
    // Daily and final reports use the new renderers; standalone reports keep the legacy layout.
    const pdf =
      getReportKind(report) === 'standalone'
        ? await generateElectroReportPdf(report, { generatedDir, publicPathPrefix: '/generated-pdfs' })
        : await exportReportPdf(store, report, { generatedDir, publicPathPrefix: '/generated-pdfs' });
    return res.json(pdf);
  } catch (error) {
    if (error instanceof ConsistencyError) {
      return res.status(409).json({ message: error.message, code: error.code, problems: error.problems });
    }
    console.error('Failed to generate ElectroReports PDF', error);
    return res.status(500).json({ message: 'Failed to generate PDF.' });
  }
});

// ---- Review and approval (daily and final reports) -------------------------------------------
function reviewError(res, error) {
  const status = error instanceof ConsistencyError ? 409 : error.code ? 409 : 400;
  return res.status(status).json({
    message: error instanceof Error ? error.message : 'Review step failed.',
    code: error.code,
    items: error.items,
    problems: error.problems
  });
}

function requireWorkflowReport(req, res) {
  const report = loadReport(req.params.id);
  if (!report) {
    res.status(404).json({ message: 'Report not found.' });
    return null;
  }
  if (getReportKind(report) === 'standalone') {
    res.status(400).json({ message: 'Standalone reports have no review workflow.' });
    return null;
  }
  return report;
}

// The review state plus what this user may do with it (the screen shows only those buttons).
function reviewPayload(req, report, extra = {}) {
  return { ...buildReviewState(store, report), access: reviewAccess(store, req.user, report), ...extra };
}

// A step the gate refuses: 403 with the reason, and whether an admin could override it.
function refuse(res, gate) {
  return res.status(403).json({ message: gate.message, code: gate.code, canOverride: gate.canOverride });
}

app.get('/api/reports/:id/review', (req, res) => {
  const report = requireWorkflowReport(req, res);
  return report ? res.json(reviewPayload(req, report)) : undefined;
});

app.put('/api/reports/:id/review/ticks', (req, res) => {
  const report = requireWorkflowReport(req, res);
  if (!report) {
    return undefined;
  }
  const gate = reviewGate(store, req.user, 'ticks', report, { overrideReason: req.body && req.body.overrideReason });
  if (!gate.allowed) {
    return refuse(res, gate);
  }
  try {
    store.setReviewTicks(req.params.id, req.body && req.body.ticks);
    return res.json(reviewPayload(req, loadReport(req.params.id)));
  } catch (error) {
    return reviewError(res, error);
  }
});

const REVIEW_ACTIONS = { submit: 'in_review', approve: 'approved', return: 'draft' };

app.post('/api/reports/:id/review/:action', (req, res) => {
  const report = requireWorkflowReport(req, res);
  if (!report) {
    return undefined;
  }
  const { action } = req.params;
  const body = req.body || {};
  // The step is recorded against the signed-in person; their profile gives the name and designation.
  const who = { actor: req.user.name, designation: req.user.designation, userId: req.user.id };
  const gate = reviewGate(store, req.user, action, report, { overrideReason: body.overrideReason });
  if (!gate.allowed) {
    return refuse(res, gate);
  }
  const note = gate.overridden ? [`Admin override by ${req.user.name}: ${gate.reason}`, String(body.note || '')].filter(Boolean).join(' | ') : body.note;
  try {
    let warnings = [];
    if (action === 'check') {
      store.markChecked(req.params.id, { ...who, note });
    } else if (REVIEW_ACTIONS[action]) {
      const result = store.transition(req.params.id, REVIEW_ACTIONS[action], { ...who, note, acknowledgements: body.acknowledgements || {} });
      warnings = result.warnings;
    } else {
      return res.status(404).json({ message: 'Unknown review step.' });
    }
    if (gate.overridden) {
      auth.audit(req.user, 'review.override', 'report', req.params.id, { step: action, reason: gate.reason });
    }
    if (action === 'approve') {
      peopleSync.afterTransition(action, loadReport(req.params.id)).catch((error) => console.error('Site Work Log sync failed', error));
    }
    return res.json(reviewPayload(req, loadReport(req.params.id), { stepWarnings: warnings }));
  } catch (error) {
    return reviewError(res, error);
  }
});

app.post('/api/reports/:id/reissue', (req, res) => {
  if (!requireWorkflowReport(req, res)) {
    return undefined;
  }
  const body = req.body || {};
  try {
    const copy = store.reissueReport(req.params.id, { actor: req.user.name, userId: req.user.id, reason: body.reason });
    return res.status(201).json(presentReport(copy));
  } catch (error) {
    return reviewError(res, error);
  }
});


// ---- Project workspaces (Session 6C.4) -------------------------------------------------------
const ISO_DAY = /^\d{4}-\d{2}-\d{2}$/;

app.get('/api/projects', (_req, res) => {
  res.json(listProjectSummaries(store));
});

// One workspace per Zoho project: finds it, or creates it from the project details given.
app.post('/api/projects', (req, res) => {
  try {
    const { project, created } = ensureProject(store, (req.body && req.body.project) || req.body);
    return res.status(created ? 201 : 200).json({ projectId: project.id, created });
  } catch (error) {
    return res.status(400).json({ message: error instanceof Error ? error.message : 'Invalid project.', code: error.code });
  }
});

app.get('/api/projects/:id', (req, res) => {
  const workspace = buildWorkspace(store, req.params.id);
  return workspace ? res.json(workspace) : res.status(404).json({ message: 'Project not found.' });
});

// Project-level details: planned quantities (the denominator of progress), unit, engineer...
app.put('/api/projects/:id', (req, res) => {
  const project = store.getProject(req.params.id);
  if (!project) {
    return res.status(404).json({ message: 'Project not found.' });
  }
  const patch = pickProjectFields((req.body && req.body.project) || req.body);
  delete patch.zohoProjectId;
  const sent = (req.body && req.body.project) || req.body || {};
  if ('workOrder' in sent) {
    patch.workOrder = String(sent.workOrder || '').trim(); // optional: an empty value clears it
  }
  if (patch.dataCollectionMode && !['On-site entry', 'Printed sheets'].includes(patch.dataCollectionMode)) {
    return res.status(400).json({ message: 'Data collection mode is On-site entry or Printed sheets.' });
  }
  if (patch.plannedQuantities) {
    const planned = {};
    Object.entries(patch.plannedQuantities).forEach(([testId, value]) => {
      const number = Number(value);
      if (String(value).trim() !== '' && Number.isFinite(number) && number >= 0) {
        planned[testId] = number;
      }
    });
    patch.plannedQuantities = planned;
  }
  store.updateProject(req.params.id, patch);
  return res.json(buildWorkspace(store, req.params.id));
});

// Records a site day: the daily report of a work date, one per date.
app.post('/api/projects/:id/days', (req, res) => {
  const project = store.getProject(req.params.id);
  if (!project) {
    return res.status(404).json({ message: 'Project not found.' });
  }
  const body = req.body || {};
  const workDate = String(body.workDate || '').trim();
  if (!ISO_DAY.test(workDate)) {
    return res.status(400).json({ message: 'Choose the work date of this site day.' });
  }
  const existing = findDay(store, project.id, workDate);
  if (existing) {
    return res.status(409).json({ message: `A daily report for ${workDate} already exists: reopen it instead.`, code: 'DAY_EXISTS', reportId: existing.id });
  }
  try {
    const draft = body.report || {};
    const normalized = normalizeReportInput({ ...draft, project: { ...(draft.project || {}), reportDate: workDate } });
    const created = store.addReport({ ...stampReport(normalized, null, req.user), kind: 'daily', projectId: project.id, collectionDate: workDate, createdBy: req.user.id });
    if (!project.plannedQuantities && normalized.project.plannedQuantities) {
      store.updateProject(project.id, { plannedQuantities: normalized.project.plannedQuantities });
    }
    return res.status(201).json(presentReport(created));
  } catch (error) {
    return res.status(400).json({
      message: error instanceof Error ? error.message : 'Invalid report payload.',
      errors: error instanceof ReportValidationError ? error.errors : undefined
    });
  }
});

// Stitches every daily report of the project into a NEW draft final report; an earlier final
// report is kept as it was. Days not yet approved are included and flagged.
app.post('/api/projects/:id/final/assemble', (req, res) => {
  if (!store.getProject(req.params.id)) {
    return res.status(404).json({ message: 'Project not found.' });
  }
  try {
    const result = assembleFor(store, req.params.id, { actor: req.user.name, userId: req.user.id });
    return res.status(201).json({ ...result, report: presentReport(result.report) });
  } catch (error) {
    return res.status(400).json({ message: error instanceof Error ? error.message : 'Could not assemble the final report.' });
  }
});

// Saves changes to a daily report that is still a draft (a day being entered over several visits).
app.put('/api/reports/:id', (req, res) => {
  const current = store.getReport(req.params.id);
  if (!current) {
    return res.status(404).json({ message: 'Report not found.' });
  }
  if (getReportKind(current) !== 'daily') {
    return res.status(400).json({ message: 'Only a daily report can be edited. A final report is assembled from the daily reports.' });
  }
  if (!mayEditReport(req.user, current)) {
    return res.status(403).json({ message: current.createdBy ? 'Only the engineer who created this daily report (or an admin) can edit it.' : 'This report has no creator: only an admin can edit it.', code: 'NOT_OWNER' });
  }
  if (getReviewStatus(current) !== 'draft') {
    return res.status(409).json({ message: 'Only a draft daily report can be edited. Return it to draft from the review screen first.', code: 'NOT_DRAFT' });
  }
  const body = req.body || {};
  if (body.project && body.project.reportDate && body.project.reportDate !== current.collectionDate) {
    return res.status(400).json({ message: `The work date of this day is ${current.collectionDate} and cannot be changed here.` });
  }
  try {
    const normalized = normalizeReportInput({ ...body, project: { ...(body.project || {}), reportDate: current.collectionDate } });
    const updated = store.updateReport(req.params.id, () => ({ ...stampReport({ ...normalized, collectionDate: current.collectionDate }, current, req.user) }));
    return res.json(presentReport(updated));
  } catch (error) {
    return res.status(400).json({
      message: error instanceof Error ? error.message : 'Invalid report payload.',
      errors: error instanceof ReportValidationError ? error.errors : undefined
    });
  }
});

// ---- New-project notices (Session 6D) --------------------------------------------------------
// Who is asking is the "I am ..." choice on Home for now (no logins yet). This is the ONE place the
// server reads it; per-engineer logins will replace it.
function noticeRecipientFrom(req) {
  return { keys: [req.user.email.toLowerCase(), req.user.name.toLowerCase()], label: req.user.name };
}

// A notice is for its recipient (an admin may answer on someone's behalf).
function mayAnswerNotice(req, notice) {
  return hasRole(req.user, 'admin') || noticeRecipientFrom(req).keys.includes(notice.recipientKey);
}

app.get('/api/notices', (req, res) => {
  const { keys } = noticeRecipientFrom(req);
  const seen = new Map();
  keys.forEach((key) => store.listNotices({ recipientKey: key, status: 'open' }).forEach((notice) => seen.set(notice.id, notice)));
  return res.json([...seen.values()]);
});

app.post('/api/notices/:id/answer', (req, res) => {
  const notice = store.getNotice(req.params.id);
  if (notice && !mayAnswerNotice(req, notice)) {
    return res.status(403).json({ message: 'This notice is addressed to someone else.', code: 'NOT_RECIPIENT' });
  }
  try {
    const body = req.body || {};
    return res.json(answerNotice(store, req.params.id, { phonesAllowed: body.phonesAllowed, mobileData: body.mobileData, by: req.user.name, userId: req.user.id }));
  } catch (error) {
    return res.status(error.code === 'NOT_FOUND' ? 404 : error.code === 'ANSWERS_REQUIRED' ? 400 : 409).json({ message: error.message, code: error.code });
  }
});

app.post('/api/notices/:id/dismiss', (req, res) => {
  const notice = store.getNotice(req.params.id);
  if (notice && !mayAnswerNotice(req, notice)) {
    return res.status(403).json({ message: 'This notice is addressed to someone else.', code: 'NOT_RECIPIENT' });
  }
  try {
    dismissNotice(store, req.params.id, req.user.id);
    return res.json({ dismissed: true });
  } catch (error) {
    return res.status(error.code === 'NOT_FOUND' ? 404 : 409).json({ message: error.message, code: error.code });
  }
});

// Runs the new-project check now (the timer runs it every ZOHO_POLL_MINUTES).
app.post('/api/zoho/poll', async (_req, res) => {
  try {
    return res.json(await pollNewProjects({ store }));
  } catch (error) {
    console.error('Zoho project check failed', error);
    return res.status(500).json({ message: 'The Zoho project check failed.' });
  }
});

// Fills the workspace's planned quantities from the Zoho Planned Quantity of the Onsite tasks. Only
// tests with one clear Zoho number are set; every other value is left as it is.
app.post('/api/projects/:id/zoho-refresh', async (req, res) => {
  const project = store.getProject(req.params.id);
  if (!project) {
    return res.status(404).json({ message: 'Project not found.' });
  }
  if (!project.zohoProjectId) {
    return res.status(400).json({ message: 'This project is not linked to a Zoho project.' });
  }
  try {
    const { plan } = await loadTaskPlan(project.zohoProjectId);
    const planned = { ...(project.plannedQuantities || {}) };
    const filled = [];
    Object.values(plan.tests).forEach((test) => {
      if (test.planned !== null) {
        planned[test.testId] = test.planned;
        filled.push(test.testId);
      }
    });
    store.updateProject(project.id, { plannedQuantities: planned });
    return res.json({ workspace: buildWorkspace(store, project.id), filled, warnings: plan.warnings, tests: plan.tests });
  } catch (error) {
    console.error('Failed to refresh from Zoho', error);
    return res.status(500).json({ message: 'Failed to read the Zoho project tasks.' });
  }
});

// ---- Schedule (6F) -----------------------------------------------------------------------------
// Everyone signed in sees the whole team, their own row first. Nothing here writes to Zoho.
app.get('/api/schedule', async (req, res) => {
  try {
    return res.json(await schedule.snapshot({ from: String(req.query.from || ''), to: String(req.query.to || ''), user: req.user }));
  } catch (error) {
    console.error('Failed to build the schedule', error);
    return res.status(500).json({ message: 'Failed to build the schedule.' });
  }
});

app.post('/api/schedule/refresh', async (req, res) => {
  try {
    const body = req.body || {};
    return res.json(await schedule.refresh({ from: String(body.from || ''), to: String(body.to || ''), user: req.user }));
  } catch (error) {
    console.error('Failed to refresh the schedule', error);
    return res.status(500).json({ message: 'Failed to refresh the schedule.' });
  }
});

// Admin only (the /api/admin rule): working days and the rate per test.
app.put('/api/admin/schedule-settings', (req, res) => {
  const result = schedule.saveSettings(req.body || {}, req.user.email);
  if (result.error) {
    return res.status(400).json({ message: result.error });
  }
  auth.audit(req.user, 'schedule.settings', 'schedule', 'settings', { workingDays: result.settings.workingDays, rates: Object.fromEntries(Object.entries(result.settings.rates).map(([id, rate]) => [id, rate.perDay])) });
  return res.json(result.settings);
});

// ---- Site Work Log in Zoho People (6G): admin only (the /api/admin rule) ---------------------------
app.get('/api/admin/people-sync', async (_req, res) => {
  try {
    return res.json(await peopleSync.status());
  } catch (error) {
    console.error('Failed to read the Site Work Log status', error);
    return res.status(500).json({ message: 'Failed to read the Site Work Log status.' });
  }
});

app.get('/api/admin/people-sync/totals', (req, res) => {
  try {
    return res.json(peopleSync.totals({ from: String(req.query.from || ''), to: String(req.query.to || '') }));
  } catch (error) {
    return res.status(error.code === 'BAD_RANGE' ? 400 : 500).json({ message: error.code === 'BAD_RANGE' ? error.message : 'Failed to total the Site Work Log.' });
  }
});

app.post('/api/admin/people-sync/backfill', async (req, res) => {
  const body = req.body || {};
  try {
    const summary = await peopleSync.backfill({ from: String(body.from || ''), to: String(body.to || '') });
    auth.audit(req.user, 'people.backfill', 'people_sync', `${summary.from}..${summary.to}`, summary);
    return res.json(summary);
  } catch (error) {
    if (error.code === 'BAD_RANGE') return res.status(400).json({ message: error.message, code: error.code });
    console.error('Site Work Log backfill failed', error);
    return res.status(500).json({ message: 'The Site Work Log backfill failed.' });
  }
});

app.post('/api/admin/people-sync/retry', async (req, res) => {
  try {
    const tally = await peopleSync.retry();
    auth.audit(req.user, 'people.retry', 'people_sync', 'retry', tally);
    return res.json(tally);
  } catch (error) {
    console.error('Site Work Log retry failed', error);
    return res.status(500).json({ message: 'The Site Work Log retry failed.' });
  }
});

// ---- Client logos ------------------------------------------------------------------------------
app.get('/api/client-logo', (req, res) => {
  const keys = clientKeys({ clientId: req.query.clientId, clientName: req.query.clientName });
  const found = keys.map((key) => store.getClientLogo(key)).find(Boolean);
  if (!found) {
    return res.status(404).json({ message: 'No logo stored for this client.' });
  }
  res.setHeader('Content-Type', found.mime);
  return res.send(found.data);
});

app.put('/api/client-logo', (req, res) => {
  const body = req.body || {};
  // Stored under the Zoho client id and the client name, so a report that only knows the name
  // (or only the id) still finds it.
  const keys = clientKeys({ clientId: body.clientId, clientName: body.clientName });
  if (!keys.length) {
    return res.status(400).json({ message: 'A client id or client name is required.' });
  }
  try {
    const { mime, buffer } = parseLogoDataUrl(body.dataUrl);
    keys.forEach((key) => store.setClientLogo({ key, clientName: body.clientName || '', mime, data: buffer }));
    return res.json({ keys, mime, bytes: buffer.length });
  } catch (error) {
    return res.status(error instanceof LogoError ? 400 : 500).json({ message: error.message });
  }
});

app.delete('/api/client-logo', (req, res) => {
  const deleted = clientKeys({ clientId: req.query.clientId, clientName: req.query.clientName }).map((key) => store.deleteClientLogo(key)).some(Boolean);
  return deleted ? res.status(204).send() : res.status(404).json({ message: 'No logo stored for this client.' });
});

// A single report may override the client's logo (or clear the override with dataUrl: null).
app.put('/api/reports/:id/client-logo', (req, res) => {
  if (!requireWorkflowReport(req, res)) {
    return undefined;
  }
  try {
    const dataUrl = req.body && req.body.dataUrl;
    if (dataUrl) {
      parseLogoDataUrl(dataUrl);
    }
    const updated = store.updateReport(req.params.id, (current) => {
      const next = { ...current };
      next.clientLogoOverride = dataUrl || '';
      return next;
    });
    return res.json({ hasOverride: Boolean(updated.clientLogoOverride) });
  } catch (error) {
    return res.status(error instanceof LogoError ? 400 : 500).json({ message: error.message });
  }
});

app.post('/api/ocr/soil/preview', upload.single('document'), createOcrPreviewHandler(
  previewSoilSheetWithGemini,
  mapSoilOcrPreviewToDraft,
  'Failed to scan soil resistivity sheet.'
));

app.post('/api/ocr/electrode/preview', upload.single('document'), createOcrPreviewHandler(
  previewElectrodeSheetWithGemini,
  mapElectrodeOcrPreviewToDraft,
  'Failed to scan earth electrode sheet.'
));

app.post('/api/ocr/continuity/preview', upload.single('document'), createOcrPreviewHandler(
  previewContinuitySheetWithGemini,
  mapContinuityOcrPreviewToDraft,
  'Failed to scan continuity sheet.'
));

app.post('/api/ocr/loop/preview', upload.single('document'), createOcrPreviewHandler(
  previewLoopImpedanceSheetWithGemini,
  mapLoopImpedanceOcrPreviewToDraft,
  'Failed to scan loop impedance sheet.'
));

app.post('/api/ocr/fault/preview', upload.single('document'), createOcrPreviewHandler(
  previewProspectiveFaultSheetWithGemini,
  mapProspectiveFaultOcrPreviewToDraft,
  'Failed to scan prospective fault current sheet.'
));

app.post('/api/ocr/riser/preview', upload.single('document'), createOcrPreviewHandler(
  previewRiserIntegritySheetWithGemini,
  mapRiserIntegrityOcrPreviewToDraft,
  'Failed to scan riser integrity sheet.'
));

app.post('/api/ocr/earth-continuity/preview', upload.single('document'), createOcrPreviewHandler(
  previewEarthContinuitySheetWithGemini,
  mapEarthContinuityOcrPreviewToDraft,
  'Failed to scan earth continuity sheet.'
));

app.post('/api/ocr/tower-footing/preview', upload.single('document'), createOcrPreviewHandler(
  previewTowerFootingSheetWithGemini,
  mapTowerFootingOcrPreviewToDraft,
  'Failed to scan tower footing sheet.'
));

app.get('*', (_req, res) => {
  res.sendFile(path.join(publicDir, 'index.html'));
});

app.use((error, _req, res, next) => {
  if (error instanceof multer.MulterError) {
    return res.status(400).json({
      message:
        error.code === 'LIMIT_FILE_SIZE'
          ? `Uploaded file is too large. Limit is ${Math.round(config.gemini.maxUploadBytes / (1024 * 1024))} MB.`
          : error.message
    });
  }
  return next(error);
});

app.listen(port, host, () => {
  console.log(`ElectroReports running on http://${host}:${port}`);
  startProjectWatch({ store });
  schedule.start();
});
