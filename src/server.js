const fs = require('fs');
const path = require('path');
const express = require('express');
const multer = require('multer');

const config = require('./config');
const { createReportStore } = require('./reportStore');
const { normalizeReportInput, TEST_LIBRARY } = require('./reportModel');
const { generateElectroReportPdf } = require('./pdfService');
const { getAiStatus, generateReportNarrative } = require('./aiService');
const { getProjects, getProjectUsers, getPortalUsers } = require('./zohoClient');
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

app.get('/api/zoho/projects', async (req, res) => {
  try {
    const query = String(req.query.q || '').trim();
    const projects = await getProjects(query);
    return res.json(projects);
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

app.get('/api/reports', (req, res) => {
  const query = String(req.query.q || '').trim();
  res.json(store.listReports(query));
});

app.get('/api/reports/:id', (req, res) => {
  const report = store.getReport(req.params.id);
  if (!report) {
    return res.status(404).json({ message: 'Report not found.' });
  }
  return res.json(report);
});

app.post('/api/reports', (req, res) => {
  try {
    const normalized = normalizeReportInput(req.body);
    const created = store.addReport(normalized);
    return res.status(201).json(created);
  } catch (error) {
    return res.status(400).json({
      message: error instanceof Error ? error.message : 'Invalid report payload.'
    });
  }
});

app.post('/api/reports/:id/ai/generate', async (req, res) => {
  const report = store.getReport(req.params.id);
  if (!report) {
    return res.status(404).json({ message: 'Report not found.' });
  }

  try {
    const narrative = await generateReportNarrative(report);
    const updated = store.updateReport(req.params.id, (current) => ({
      ...current,
      aiNarrative: narrative
    }));
    return res.json(updated);
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
  const report = store.getReport(req.params.id);
  if (!report) {
    return res.status(404).json({ message: 'Report not found.' });
  }

  try {
    const pdf = await generateElectroReportPdf(report, {
      generatedDir,
      publicPathPrefix: '/generated-pdfs'
    });
    return res.json(pdf);
  } catch (error) {
    console.error('Failed to generate ElectroReports PDF', error);
    return res.status(500).json({ message: 'Failed to generate PDF.' });
  }
});

app.post('/api/ocr/soil/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewSoilSheetWithGemini(config, req.file);
    const draftPatch = mapSoilOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Soil OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan soil resistivity sheet.'
    });
  }
});

app.post('/api/ocr/electrode/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewElectrodeSheetWithGemini(config, req.file);
    const draftPatch = mapElectrodeOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Electrode OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan earth electrode sheet.'
    });
  }
});

app.post('/api/ocr/continuity/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewContinuitySheetWithGemini(config, req.file);
    const draftPatch = mapContinuityOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Continuity OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan continuity sheet.'
    });
  }
});

app.post('/api/ocr/loop/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewLoopImpedanceSheetWithGemini(config, req.file);
    const draftPatch = mapLoopImpedanceOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Loop Impedance OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan loop impedance sheet.'
    });
  }
});

app.post('/api/ocr/fault/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewProspectiveFaultSheetWithGemini(config, req.file);
    const draftPatch = mapProspectiveFaultOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Prospective Fault Current OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan prospective fault current sheet.'
    });
  }
});

app.post('/api/ocr/riser/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewRiserIntegritySheetWithGemini(config, req.file);
    const draftPatch = mapRiserIntegrityOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Riser Integrity OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan riser integrity sheet.'
    });
  }
});

app.post('/api/ocr/earth-continuity/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewEarthContinuitySheetWithGemini(config, req.file);
    const draftPatch = mapEarthContinuityOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Earth Continuity OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan earth continuity sheet.'
    });
  }
});

app.post('/api/ocr/tower-footing/preview', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'Please choose a sheet image or PDF to scan.' });
    }

    const preview = await previewTowerFootingSheetWithGemini(config, req.file);
    const draftPatch = mapTowerFootingOcrPreviewToDraft(preview);

    return res.json({
      success: true,
      preview,
      draftPatch,
      fileName: req.file.originalname || 'uploaded-sheet'
    });
  } catch (error) {
    console.error('Failed to preview Tower Footing OCR with Gemini', error);
    return res.status(500).json({
      message: error instanceof Error ? error.message : 'Failed to scan tower footing sheet.'
    });
  }
});

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
