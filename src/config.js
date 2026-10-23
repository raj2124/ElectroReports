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
  model: String(process.env.ANTHROPIC_MODEL || 'claude-opus-5-5').trim(),
  // Opus 5.5 defaults to medium effort and medium is intended for the report writer: it only
  // phrases findings from computed results (6C). Raise ANTHROPIC_EFFORT only if text quality needs it.
  effort: String(process.env.ANTHROPIC_EFFORT || 'medium').trim(),
  role: String(process.env.CLAUDE_ROLE || 'report-narrative').trim(),
  referencePdfPaths
};

const openaiSanityModel = String(process.env.OPENAI_SANITY_MODEL || process.env.OPENAI_MODEL || 'gpt-6-sol').trim();
const openai = {
  apiKey: String(process.env.OPENAI_API_KEY || '').trim(),
  // Per-row observations/remarks: the cheaper model. Once-per-report sanity check: the stronger one.
  observationModel: String(process.env.OPENAI_OBSERVATION_MODEL || 'gpt-6-luna').trim(),
  sanityModel: openaiSanityModel,
  // Also used for the OpenAI narrative fallback and for any photo/visual check.
  model: openaiSanityModel,
  role: String(process.env.OPENAI_ROLE || 'observations-and-sanity-check').trim(),
  referencePdfPaths
};

// Instruments that may appear in a report, with the range and accuracy printed in its Instruments
// table. `match` is matched (case-insensitively, spaces ignored) against the instrument name
// recorded on the sheet. Edit the range / accuracy here to match the instrument datasheets.
const instrumentCatalogue = [
  { match: 'mi3290', name: 'Metrel MI 3290 GX Earth Analyser', range: '0.01 Ω – 19.9 kΩ (earth resistance); 0.1 – 9999 Ω·m (soil)', accuracy: '±(2 % of reading + 3 digits)' },
  { match: 'mi3152', name: 'Metrel MI 3152 EurotestXC', range: '0.00 Ω – 1999 Ω (earth); 0.00 Ω – 9.99 kΩ (loop)', accuracy: '±(5 % of reading + 5 digits)' },
  { match: 'mi3295', name: 'Metrel MI 3295 Step & Contact Voltage Meter', range: '0.01 – 199.9 V (step / touch voltage)', accuracy: '±(3 % of reading + 3 digits)' },
  { match: 'kyoritsu4105a', name: 'Kyoritsu 4105A Digital Earth Tester', range: '0 – 20 / 200 / 2000 Ω', accuracy: '±(2 % of reading + 0.1 Ω)' },
  { match: 'kyoritsu4118a', name: 'Kyoritsu 4118A Loop / PSC Tester', range: '0 – 20 / 200 / 2000 Ω (loop); 0 – 20 kA (PSC)', accuracy: '±(2.5 % of reading + 3 digits)' }
];

// Microsoft Graph (Session 6H): Outlook calendar sync (write, only the app's own events, only mailboxes in the
// Exchange-scoped group) and the SharePoint knowledge base (STRICTLY read-only: Sites.Selected with "read").
const kbSitesDefault = [
  { name: 'ETPL_OP_26_27', url: 'https://eletechnology.sharepoint.com/sites/ETPL_OP_26_27', folders: [], scope: 'client' },
  { name: 'FY23-24', url: 'https://eletechnology.sharepoint.com/sites/FY23-24', folders: [], scope: 'client' },
  { name: 'FY23-24-StandardsGuideline', url: 'https://eletechnology.sharepoint.com/sites/FY23-24-StandardsGuideline', folders: ['Shared Documents/ETPL Knowledge Hub/02 Technical Knowledge'], scope: 'general' }
];
function parseKbSites(value) {
  if (!String(value || '').trim()) return kbSitesDefault;
  try {
    const list = JSON.parse(value);
    return Array.isArray(list) ? list.map((site) => ({ name: String(site.name || '').trim(), url: String(site.url || '').trim().replace(/\/+$/, ''), folders: (Array.isArray(site.folders) ? site.folders : []).map((entry) => String(entry).replace(/^\/+|\/+$/g, '')), scope: site.scope === 'general' ? 'general' : 'client' })).filter((site) => site.name && site.url) : kbSitesDefault;
  } catch (_error) {
    return kbSitesDefault;
  }
}
const graph = {
  useMock: toBool(process.env.GRAPH_USE_MOCK, true),
  tenantId: String(process.env.GRAPH_TENANT_ID || '').trim(),
  clientId: String(process.env.GRAPH_CLIENT_ID || '').trim(),
  certPath: String(process.env.GRAPH_CERT_PATH || '').trim(),
  certThumbprintSha256: String(process.env.GRAPH_CERT_THUMBPRINT_SHA256 || '').replace(/[^0-9a-fA-F]/g, ''),
  clientSecret: String(process.env.GRAPH_CLIENT_SECRET || '').trim(),
  baseUrl: 'https://graph.microsoft.com/v1.0',
  authority: 'https://login.microsoftonline.com',
  appBaseUrl: String(process.env.APP_BASE_URL || '').trim().replace(/\/+$/, ''),
  timezone: String(process.env.GRAPH_CALENDAR_TIMEZONE || 'India Standard Time').trim(),
  kb: {
    sites: parseKbSites(process.env.GRAPH_KB_SITES),
    refreshMinutes: Number.parseInt(process.env.GRAPH_KB_REFRESH_MINUTES || '60', 10),
    maxMb: Number.parseFloat(process.env.GRAPH_KB_MAX_MB || '40'),
    dbPath: process.env.GRAPH_KB_DB || path.join(path.dirname(process.env.ELECTROREPORTS_DATA_FILE || path.join(rootDir, 'data', 'reports.json')), 'knowledge.db')
  }
};

// Scanned sheets (Session 6I): OCR_USE_MOCK / AI_CHECKS_USE_MOCK default to the mocks, so nothing is sent to a provider
// until they are set to false on purpose (the provider billing is not ready).
const scanDataDir = path.join(path.dirname(process.env.ELECTROREPORTS_DATA_FILE || path.join(rootDir, 'data', 'reports.json')), 'scans');
const scan = {
  useMock: toBool(process.env.OCR_USE_MOCK, true),
  dir: process.env.ELECTROREPORTS_SCANS_DIR || scanDataDir,
  minConfidence: Math.min(1, Math.max(0, Number.parseFloat(process.env.OCR_MIN_CONFIDENCE || '0.8'))),
  maxPages: Number.parseInt(process.env.OCR_MAX_PAGES || '30', 10),
  maxUploadBytes: Math.round(Number.parseFloat(process.env.OCR_MAX_UPLOAD_MB || '40') * 1024 * 1024),
  aiChecksUseMock: toBool(process.env.AI_CHECKS_USE_MOCK, true)
};

// Sign-in (Session 6E). mode: 'zoho' (Sign in with Zoho, OpenID Connect) or 'dev' (a user picker for
// development and tests, never the default). With neither configured the API stays closed.
const authMode = String(process.env.AUTH_MODE || '').trim().toLowerCase();
const auth = {
  mode: authMode,
  zoho: {
    clientId: String(process.env.AUTH_ZOHO_CLIENT_ID || '').trim(),
    clientSecret: String(process.env.AUTH_ZOHO_CLIENT_SECRET || '').trim(),
    redirectUri: String(process.env.AUTH_ZOHO_REDIRECT_URI || `http://127.0.0.1:${port}/auth/callback`).trim(),
    accountsUrl: String(process.env.AUTH_ZOHO_ACCOUNTS_URL || 'https://accounts.zoho.in').trim().replace(/\/+$/, ''),
    // OpenID Connect first; set AUTH_ZOHO_SCOPE=AaaServer.profile.READ if the e-mail is not in the id_token
    scope: String(process.env.AUTH_ZOHO_SCOPE || 'openid email profile').trim()
  },
  allowedDomain: String(process.env.AUTH_ALLOWED_DOMAIN || 'elegrow.com').trim().toLowerCase(),
  idleHours: Number.parseFloat(process.env.AUTH_SESSION_IDLE_HOURS || '12'),
  recheckMinutes: Number.parseFloat(process.env.AUTH_RECHECK_MINUTES || '60'),
  sessionSecret: String(process.env.AUTH_SESSION_SECRET || '').trim(),
  seedAdmins: String(process.env.AUTH_SEED_ADMINS !== undefined ? process.env.AUTH_SEED_ADMINS : 'shivam@elegrow.com').split(',').map((entry) => entry.trim().toLowerCase()).filter(Boolean),
  // dev mode only: "email:role+role:Name:Designation,..."
  devUsers: String(process.env.AUTH_DEV_USERS || '').split(',').map((entry) => entry.trim()).filter(Boolean)
};

module.exports = {
  auth,
  graph,
  scan,
  instrumentCatalogue,
  app: {
    rootDir,
    publicDir: path.join(rootDir, 'public'),
    generatedDir: process.env.ELECTROREPORTS_GENERATED_DIR || path.join(rootDir, 'generated-pdfs'),
    dataFilePath: process.env.ELECTROREPORTS_DATA_FILE || path.join(rootDir, 'data', 'reports.json'),
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
    projectsEndpoint: process.env.ZOHO_PROJECTS_ENDPOINT || '',
    // Zoho People (Session 6F, read only): leave, holidays and the employee list. India data centre.
    peopleBaseUrl: String(process.env.ZOHO_PEOPLE_BASE_URL || 'https://people.zoho.in').trim().replace(/\/+$/, ''),
    // Site Work Log (Session 6G): the People form that gets one record per engineer, work date and test. Its
    // fields are found by label when the app writes; `fields` only holds overrides by API name, keyed by the
    // app's field key (e.g. {"completedQty":"Completed_Qty1"}). The form has to be created in People first;
    // until it exists nothing is lost: the writes wait as "pending" and the admin Backfill / Retry sends them.
    workLog: {
      form: String(process.env.ZOHO_PEOPLE_WORKLOG_FORM || 'site_work_log').trim(),
      fields: (() => { try { return JSON.parse(process.env.ZOHO_PEOPLE_WORKLOG_FIELDS || '{}'); } catch (_error) { return {}; } })()
    },
    // New-project watch (Session 6D): how often Zoho is polled (0 turns it off) and which project
    // statuses get a notice. D03 projects only, created on or after 2026-01-01, with no workspace yet.
    pollMinutes: Number.parseInt(process.env.ZOHO_POLL_MINUTES || '15', 10),
    notifyStatuses: String(process.env.ZOHO_NOTIFY_STATUSES || 'Planning,Active').split(',').map((entry) => entry.trim()).filter(Boolean),
    // Task custom fields are found in the task payload by API name first, then by label. These
    // overrides (API names) are used when no task of a project has a value yet.
    taskFields: {
      eshaTest: String(process.env.ZOHO_TASK_FIELD_ESHA_TEST || '').trim(),
      plannedQuantity: String(process.env.ZOHO_TASK_FIELD_PLANNED_QTY || '').trim(),
      category: String(process.env.ZOHO_TASK_FIELD_CATEGORY || 'category').trim()
    }
  },
  // Scheduling (Session 6F). Zoho is read only; what is fetched is cached with its time. The working
  // days and the rates per test are defaults: an admin edits them in the app (stored in the database).
  schedule: {
    refreshMinutes: Number.parseInt(process.env.SCHEDULE_REFRESH_MINUTES || '30', 10),
    manualRefreshSeconds: Number.parseInt(process.env.SCHEDULE_MANUAL_REFRESH_SECONDS || '60', 10),
    // Onsite tasks in these statuses are left off the calendar (everything else shows, whatever the project status).
    hiddenTaskStatuses: String(process.env.SCHEDULE_HIDDEN_STATUSES || 'Closed').split(',').map((entry) => entry.trim()).filter(Boolean),
    workingDays: String(process.env.SCHEDULE_WORKING_DAYS || 'Mon,Tue,Wed,Thu,Fri,Sat').split(',').map((entry) => entry.trim()).filter(Boolean)
  },
  claude,
  openai,
  // `ai` kept as a backward-compatible alias to the OpenAI config so existing
  // references (getAiStatus, fallback path) keep working after the role split.
  ai: openai,
  gemini: {
    enabled: toBool(process.env.GEMINI_ENABLED, true),
    apiKey: String(process.env.GEMINI_API_KEY || '').trim(),
    model: String(process.env.GEMINI_MODEL || 'gemini-3.8-flash').trim(),
    // Re-reads a page once when the primary read has low-confidence or illegible cells.
    escalationModel: String(process.env.GEMINI_ESCALATION_MODEL || 'gemini-4-argon').trim(),
    maxUploadBytes: Number.parseInt(process.env.GEMINI_MAX_UPLOAD_BYTES || '15728640', 10)
  },
  envPath
};
