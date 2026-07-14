const path = require('path');
const dotenv = require('dotenv');

const rootDir = path.resolve(__dirname, '..');
const envPath = path.join(rootDir, '.env');

// Force ElectroReports and any shared modules it imports to resolve dotenv from
// the ElectroReports folder instead of the workspace root.
if (process.cwd() !== rootDir) {
  process.chdir(rootDir);
}

dotenv.config({ path: envPath });

function toBool(value, defaultValue = false) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }
  return String(value).toLowerCase() === 'true';
}

const host = process.env.ELECTROREPORTS_HOST || '127.0.0.1';
const port = Number.parseInt(process.env.ELECTROREPORTS_PORT || '3011', 10);
const referencePdfPaths = String(process.env.OPENAI_REFERENCE_PDFS || '')
  .split(path.delimiter)
  .map((entry) => entry.trim())
  .filter(Boolean);

// AI provider roles (Session 3 rewire):
//   Claude  = PRIMARY   -> report narrative (overallAssessment, executiveSummary,
//                          closingSummary, keyFindings, moduleSummaries)
//   OpenAI  = SECONDARY  -> measurement-sheet observations/recommendations/remarks
//                          + sanity-check of Claude's narrative (no numbers)
//   Gemini  = OCR only   (unchanged)
const claude = {
  apiKey: String(process.env.ANTHROPIC_API_KEY || '').trim(),
  model: String(process.env.ANTHROPIC_MODEL || 'claude-opus-4-8').trim(),
  role: String(process.env.CLAUDE_ROLE || 'report-narrative').trim(),
  referencePdfPaths
};

const openai = {
  apiKey: String(process.env.OPENAI_API_KEY || '').trim(),
  model: String(process.env.OPENAI_MODEL || 'gpt-5.6-sol').trim(),
  role: String(process.env.OPENAI_ROLE || 'observations-and-sanity-check').trim(),
  referencePdfPaths
};

module.exports = {
  app: {
    rootDir,
    publicDir: path.join(rootDir, 'public'),
    generatedDir: path.join(rootDir, 'generated-pdfs'),
    dataFilePath: path.join(rootDir, 'data', 'reports.json'),
    host,
    port
  },
  zoho: {
    useMock: toBool(process.env.ZOHO_USE_MOCK, true),
    enrichProjectStage: toBool(process.env.ZOHO_ENRICH_PROJECT_STAGE, true),
    baseUrl: process.env.ZOHO_BASE_URL || 'https://projectsapi.zoho.in/restapi',
    accountsBaseUrl: process.env.ZOHO_ACCOUNTS_BASE_URL || 'https://accounts.zoho.in',
    organizationUserEmailDomain: process.env.ORG_USER_EMAIL_DOMAIN || 'elegrow.com',
    portalId: process.env.ZOHO_PORTAL_ID || '',
    accessToken: process.env.ZOHO_ACCESS_TOKEN || '',
    refreshToken: process.env.ZOHO_REFRESH_TOKEN || '',
    clientId: process.env.ZOHO_CLIENT_ID || '',
    clientSecret: process.env.ZOHO_CLIENT_SECRET || '',
    projectsEndpoint: process.env.ZOHO_PROJECTS_ENDPOINT || ''
  },
  claude,
  openai,
  // `ai` kept as a backward-compatible alias to the OpenAI config so existing
  // references (getAiStatus, fallback path) keep working after the role split.
  ai: openai,
  gemini: {
    enabled: toBool(process.env.GEMINI_ENABLED, true),
    apiKey: String(process.env.GEMINI_API_KEY || '').trim(),
    model: String(process.env.GEMINI_MODEL || 'gemini-2.5-pro').trim(),
    maxUploadBytes: Number.parseInt(process.env.GEMINI_MAX_UPLOAD_BYTES || '15728640', 10)
  },
  envPath
};
