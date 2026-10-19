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
const { getReportKind } = require('./reportCollectionModel');
const { buildReviewState } = require('./review');
const { exportReportPdf } = require('./finalReport/service');
const { ConsistencyError } = require('./finalReport/data');
const { LogoError, parseLogoDataUrl, clientKeys } = require('./finalReport/logos');
const { getProjects, getProjectUsers, getPortalUsers, filterProjectsForPicker, getProjectClientUsers, getProjectTasks } = require('./zohoClient');
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
app.use('/generated-pdfs', express.static(generatedDir));
app.use(
  express.static(publicDir, {
    setHeaders(res, filePath) {
      if (filePath.endsWith('index.html') || filePath.endsWith('app.js') || filePath.endsWith('styles.css')) {
        res.setHeader('Cache-Control', 'no-store');
      }
    }
  })
);

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
    const created = store.addReport(normalized);
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

app.get('/api/reports/:id/review', (req, res) => {
  const report = requireWorkflowReport(req, res);
  return report ? res.json(buildReviewState(store, report)) : undefined;
});

app.put('/api/reports/:id/review/ticks', (req, res) => {
  if (!requireWorkflowReport(req, res)) {
    return undefined;
  }
  try {
    store.setReviewTicks(req.params.id, req.body && req.body.ticks);
    return res.json(buildReviewState(store, loadReport(req.params.id)));
  } catch (error) {
    return reviewError(res, error);
  }
});

const REVIEW_ACTIONS = { submit: 'in_review', approve: 'approved', return: 'draft' };

app.post('/api/reports/:id/review/:action', (req, res) => {
  if (!requireWorkflowReport(req, res)) {
    return undefined;
  }
  const { action } = req.params;
  const body = req.body || {};
  const actor = String(body.actor || '').trim();
  if (!actor) {
    return res.status(400).json({ message: 'Enter your name: it is recorded with this step and gives the initials on the cover.' });
  }
  try {
    let warnings = [];
    if (action === 'check') {
      store.markChecked(req.params.id, { actor, note: body.note });
    } else if (REVIEW_ACTIONS[action]) {
      const result = store.transition(req.params.id, REVIEW_ACTIONS[action], { actor, note: body.note, acknowledgements: body.acknowledgements || {} });
      warnings = result.warnings;
    } else {
      return res.status(404).json({ message: 'Unknown review step.' });
    }
    return res.json({ ...buildReviewState(store, loadReport(req.params.id)), stepWarnings: warnings });
  } catch (error) {
    return reviewError(res, error);
  }
});

app.post('/api/reports/:id/reissue', (req, res) => {
  if (!requireWorkflowReport(req, res)) {
    return undefined;
  }
  const body = req.body || {};
  if (!String(body.actor || '').trim()) {
    return res.status(400).json({ message: 'Enter your name to re-issue.' });
  }
  try {
    const copy = store.reissueReport(req.params.id, { actor: body.actor, reason: body.reason });
    return res.status(201).json(presentReport(copy));
  } catch (error) {
    return reviewError(res, error);
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
});
