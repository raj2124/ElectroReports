const fs = require('fs');
const path = require('path');
const express = require('express');

const config = require('./config');
const { createReportStore } = require('./reportStore');
const { normalizeReportInput, TEST_LIBRARY } = require('./reportModel');
const { generateElectroReportPdf } = require('./pdfService');
const { getProjects, getProjectUsers, getPortalUsers } = require('../../src/zohoClient');

const app = express();
const rootDir = config.app.rootDir;
const publicDir = config.app.publicDir;
const generatedDir = config.app.generatedDir;
const dataFilePath = config.app.dataFilePath;
const store = createReportStore({ filePath: dataFilePath });
const port = config.app.port;
const host = config.app.host;

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

app.get('*', (_req, res) => {
  res.sendFile(path.join(publicDir, 'index.html'));
});

app.listen(port, host, () => {
  console.log(`ElectroReports running on http://${host}:${port}`);
});
