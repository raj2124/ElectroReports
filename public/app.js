const app = document.getElementById('app');

const LOCAL_TEST_LIBRARY = [
  {
    id: "soilResistivity",
    sheetCode: "MS-01",
    label: "Soil Resistivity Test",
    shortLabel: "Soil",
    description: "Direction-wise soil resistivity readings with automatic average and IS 3043 corrosivity class."
  },
  {
    id: "electrodeResistance",
    sheetCode: "MS-02",
    label: "Earth Electrode Resistance Test",
    shortLabel: "Electrode",
    description: "Pit-wise resistance against the IS 3043:2018 Cl. 14.2 limit calculated from site soil resistivity."
  },
  {
    id: "continuityTest",
    sheetCode: "MS-03",
    label: "Earth Continuity & Bonding Test",
    shortLabel: "Continuity",
    description: "Point-to-point bonding resistance with test-lead resistance subtracted."
  },
  {
    id: "loopImpedanceTest",
    sheetCode: "MS-04",
    label: "Loop Impedance Test",
    shortLabel: "Loop",
    description: "Measured Zs values for protection loop verification."
  },
  {
    id: "prospectiveFaultCurrent",
    sheetCode: "MS-05",
    label: "Prospective Fault Current",
    shortLabel: "PFC",
    description: "PFC (L-PE) and PSC (L-N) against the device breaking capacity."
  },
  {
    id: "riserIntegrityTest",
    sheetCode: "MS-06",
    label: "Riser / Grid Integrity Test",
    shortLabel: "Riser",
    description: "Resistance verification towards equipment and earth grid."
  },
  {
    id: "groundResistance",
    sheetCode: "MS-07",
    label: "Ground (Grid) Resistance Test",
    shortLabel: "Ground",
    description: "In-service grid resistance against the limit of the associated pit."
  },
  {
    id: "stepTouchVoltage",
    sheetCode: "MS-08",
    label: "Step & Touch Voltage Test",
    shortLabel: "Step & Touch",
    description: "Injected-current step and touch voltages scaled by If/It against the tolerable limits."
  },
  {
    id: "towerFootingResistance",
    sheetCode: "MS-09",
    label: "Tower Footing Resistance Measurement & Analysis",
    shortLabel: "Tower Footing",
    description: "Grouped tower footing impedance and current analysis with per-tower totals."
  },
  {
    id: "earthContinuityTest",
    sheetCode: "legacy",
    retired: true,
    label: "Earth Continuity Test",
    shortLabel: "Earth Continuity",
    description: "Retired: folded into Earth Continuity & Bonding (MS-03). Kept so older reports still open."
  }
];

const TOWER_FOOT_POINTS = ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'];
const PHASE_MEASURED_POINTS = ['R-E', 'Y-E', 'B-E'];
const DRAFT_STORAGE_KEY = 'electroreports-builder-draft-v1';
const DRAFT_AUTOSAVE_DELAY_MS = 450;

function buildRowId(prefix = 'row') {
  const timestamp = Date.now().toString(36);
  const suffix = Math.random().toString(36).slice(2, 8);
  return `${prefix}-${timestamp}-${suffix}`;
}

function createOcrSheetState() {
  return {
    mode: 'manual',
    scanning: false,
    selectedFile: null,
    preview: null,
    previewMeta: null,
    draftPatch: null,
    warnings: [],
    uncertainFields: [],
    error: ''
  };
}

const state = {
  catalog: [...LOCAL_TEST_LIBRARY],
  reports: [],
  view: 'dashboard',
  search: '',
  loadingReports: false,
  ai: {
    configured: false,
    model: '',
    referenceDocuments: [],
    loadingStatus: false,
    generating: false
  },
  ocr: {
    configured: false,
    enabled: false,
    model: '',
    supportedSheets: [],
    loadingStatus: false,
    sheets: {
      soilResistivity: createOcrSheetState(),
      electrodeResistance: createOcrSheetState(),
      continuityTest: createOcrSheetState(),
      loopImpedanceTest: createOcrSheetState(),
      prospectiveFaultCurrent: createOcrSheetState(),
      riserIntegrityTest: createOcrSheetState(),
      earthContinuityTest: createOcrSheetState(),
      towerFootingResistance: createOcrSheetState()
    }
  },
  zoho: {
    projects: [],
    users: [],
    clientUsers: [],
    loadingProjects: false,
    loadingUsers: false,
    loadingClientUsers: false
  },
  draft: createDraft(),
  soilPage: 0,
  stepIndex: 0,
  activeReport: null,
  editing: null, // { projectId, workDate, reportId }: the builder holds a site day of a project workspace
  saving: false,
  exporting: false,
  observationEditor: null,
  observationUploading: false,
  toast: null,
  restoredDraftNoticePending: false
};

function getLocalStorage() {
  try {
    return window.localStorage;
  } catch (_error) {
    return null;
  }
}

function cloneSerializable(value) {
  return JSON.parse(JSON.stringify(value));
}

function normalizeRestoredOcrImports(input) {
  const normalized = {};
  const source = input && typeof input === 'object' ? input : {};
  LOCAL_TEST_LIBRARY.forEach((test) => {
    const record = source[test.id];
    if (!record || typeof record !== 'object') {
      return;
    }
    normalized[test.id] = {
      sheetId: test.id,
      sheetLabel: safeText(record.sheetLabel, test.label),
      fileName: safeText(record.fileName, ''),
      mimeType: safeText(record.mimeType, ''),
      fileSize: Number.isFinite(Number(record.fileSize)) ? Number(record.fileSize) : null,
      model: safeText(record.model, ''),
      scannedAt: safeText(record.scannedAt, ''),
      appliedAt: safeText(record.appliedAt, ''),
      extractedCount: Number.isFinite(Number(record.extractedCount)) ? Number(record.extractedCount) : null,
      warnings: (Array.isArray(record.warnings) ? record.warnings : []).map((value) => safeText(value, '')).filter(Boolean),
      uncertainFields: (Array.isArray(record.uncertainFields) ? record.uncertainFields : [])
        .map((item) => ({
          path: safeText(item?.path, ''),
          reason: safeText(item?.reason, '')
        }))
        .filter((item) => item.path || item.reason)
    };
  });
  return normalized;
}

function projectHasMeaningfulData(source) {
  const project = source?.project || {};
  return Boolean(
    safeText(project.projectNo, '') ||
      safeText(project.clientName, '') ||
      safeText(project.siteLocation, '') ||
      safeText(project.workOrder, '') ||
      safeText(project.engineerName, '') ||
      safeText(project.clientRepresentative, '') ||
      safeText(project.groundCondition, '') ||
      safeText(project.assumedSoilResistivity, '') ||
      safeText(project.zohoProjectId, '') ||
      safeText(project.zohoProjectName, '') ||
      safeText(project.zohoProjectOwner, '') ||
      safeText(project.zohoProjectStage, '')
  );
}

function testsDifferFromDefault(source) {
  const defaultTests = createDraft().tests;
  return Object.keys(defaultTests).some((key) => Boolean(source?.tests?.[key]) !== Boolean(defaultTests[key]));
}

function draftHasMeaningfulProgress(source) {
  return Boolean(
    projectHasMeaningfulData(source) ||
      testsDifferFromDefault(source) ||
      draftHasMeaningfulSoilData(source) ||
      draftHasMeaningfulElectrodeData(source) ||
      draftHasMeaningfulContinuityData(source) ||
      draftHasMeaningfulLoopData(source) ||
      draftHasMeaningfulFaultData(source) ||
      draftHasMeaningfulRiserData(source) ||
      draftHasMeaningfulEarthContinuityData(source) ||
      draftHasMeaningfulGroundData(source) ||
      draftHasMeaningfulStepTouchData(source) ||
      draftHasMeaningfulSheetHeaders(source) ||
      draftHasMeaningfulTowerData(source) ||
      Object.keys(normalizeRestoredOcrImports(source?.ocrImports)).length
  );
}

function persistDraftSnapshotNow() {
  const storage = getLocalStorage();
  if (!storage) {
    return;
  }

  if (state.view !== 'builder' || !draftHasMeaningfulProgress(state.draft)) {
    storage.removeItem(DRAFT_STORAGE_KEY);
    return;
  }

  storage.setItem(
    DRAFT_STORAGE_KEY,
    JSON.stringify({
      version: 1,
      savedAt: new Date().toISOString(),
      stepIndex: state.stepIndex,
      editing: state.editing ? { ...state.editing } : null,
      draft: cloneSerializable(state.draft)
    })
  );
}

function scheduleDraftAutosave() {
  window.clearTimeout(scheduleDraftAutosave.timeoutId);
  scheduleDraftAutosave.timeoutId = window.setTimeout(() => {
    persistDraftSnapshotNow();
  }, DRAFT_AUTOSAVE_DELAY_MS);
}

function clearDraftSnapshot() {
  window.clearTimeout(scheduleDraftAutosave.timeoutId);
  const storage = getLocalStorage();
  if (storage) {
    storage.removeItem(DRAFT_STORAGE_KEY);
  }
}

function restoreDraftSnapshot() {
  const storage = getLocalStorage();
  if (!storage) {
    return false;
  }

  const raw = storage.getItem(DRAFT_STORAGE_KEY);
  if (!raw) {
    return false;
  }

  try {
    const parsed = JSON.parse(raw);
    const restoredDraft = restoreDraftShape(parsed?.draft);
    if (!draftHasMeaningfulProgress(restoredDraft)) {
      storage.removeItem(DRAFT_STORAGE_KEY);
      return false;
    }

    state.draft = restoredDraft;
    state.stepIndex = Number.isInteger(parsed?.stepIndex) ? parsed.stepIndex : 0;
    state.editing = parsed && parsed.editing && parsed.editing.projectId ? parsed.editing : null;
    state.view = 'builder';
    state.restoredDraftNoticePending = true;
    return true;
  } catch (_error) {
    storage.removeItem(DRAFT_STORAGE_KEY);
    return false;
  }
}

function safeText(value, fallback = '-') {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || fallback;
}

function escapeHtml(value) {
  return safeText(value, '').replace(/[&<>"']/g, (char) => {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[char];
  });
}

function cloneRowPhotos(photos) {
  return (Array.isArray(photos) ? photos : []).map((photo) => ({
    id: safeText(photo?.id, buildRowId('photo')),
    name: safeText(photo?.name, 'Observation image'),
    dataUrl: safeText(photo?.dataUrl, '')
  }));
}

function rowHasObservationData(row) {
  return Boolean(safeText(row?.rowObservation, '')) || cloneRowPhotos(row?.rowPhotos).some((photo) => photo.dataUrl);
}

function getElectrodeMeasuredValue(row) {
  return toNumber(row?.resistanceWithGrid) ?? toNumber(row?.resistanceWithoutGrid) ?? toNumber(row?.measuredResistance);
}

function getSectionRows(source, section, direction) {
  if (section === 'soilResistivity') {
    const location = source.soilResistivity?.locations?.[0];
    return location?.[direction] || [];
  }
  return source[section] || [];
}

function getSoilLocations(source) {
  if (Array.isArray(source?.soilResistivity?.locations) && source.soilResistivity.locations.length) {
    return source.soilResistivity.locations;
  }
  return [
    {
      locationId: buildRowId('soil-location'),
      name: 'Location 1',
      direction1: source?.soilResistivity?.direction1 || [],
      direction2: source?.soilResistivity?.direction2 || [],
      drivenElectrodeDiameter: source?.soilResistivity?.drivenElectrodeDiameter || '',
      drivenElectrodeLength: source?.soilResistivity?.drivenElectrodeLength || '',
      notes: source?.soilResistivity?.notes || ''
    }
  ];
}

function soilLocationHasData(location) {
  if (!location) {
    return false;
  }

  const rowHasData = (row) => safeText(row?.resistivity, '') || safeText(row?.rowObservation, '') || (Array.isArray(row?.rowPhotos) && row.rowPhotos.length);
  return Boolean(
    safeText(location.name, '') ||
      safeText(location.drivenElectrodeDiameter, '') ||
      safeText(location.drivenElectrodeLength, '') ||
      safeText(location.notes, '') ||
      SOIL_DIRECTION_KEYS.some((key) => Array.isArray(location[key]) && location[key].some(rowHasData))
  );
}

function draftHasMeaningfulSoilData(source) {
  return getSoilLocations(source).some((location, index) => {
    if (index === 0 && safeText(location.name, '') === 'Location 1' && !soilLocationHasData({ ...location, name: '' })) {
      return false;
    }
    return soilLocationHasData(location);
  });
}

function getOcrSheetState(sheetId) {
  if (!state.ocr.sheets[sheetId]) {
    state.ocr.sheets[sheetId] = createOcrSheetState();
  }
  return state.ocr.sheets[sheetId];
}

function resetOcrSheetState(sheetId, keepMode = true) {
  const sheetState = getOcrSheetState(sheetId);
  sheetState.selectedFile = null;
  sheetState.preview = null;
  sheetState.previewMeta = null;
  sheetState.draftPatch = null;
  sheetState.warnings = [];
  sheetState.uncertainFields = [];
  sheetState.error = '';
  sheetState.scanning = false;
  if (!keepMode) {
    sheetState.mode = 'manual';
  }
}

function resetAllOcrStates(keepMode = false) {
  [
    'soilResistivity',
    'electrodeResistance',
    'continuityTest',
    'loopImpedanceTest',
    'prospectiveFaultCurrent',
    'riserIntegrityTest',
    'earthContinuityTest',
    'towerFootingResistance'
  ].forEach((sheetId) => {
    resetOcrSheetState(sheetId, keepMode);
  });
}

function buildOcrImportRecord(sheetId, sheetState) {
  const preview = sheetState.preview || {};
  const extractedCount = Array.isArray(preview.locations)
    ? preview.locations.length
    : Array.isArray(preview.rows)
      ? preview.rows.length
      : Array.isArray(preview.groups)
        ? preview.groups.length
        : null;

  return {
    sheetId,
    sheetLabel: getOcrSheetLabel(sheetId),
    fileName: safeText(sheetState.previewMeta?.fileName, ''),
    mimeType: safeText(sheetState.previewMeta?.mimeType, ''),
    fileSize: Number.isFinite(Number(sheetState.previewMeta?.fileSize)) ? Number(sheetState.previewMeta.fileSize) : null,
    model: safeText(state.ocr.model, ''),
    scannedAt: safeText(sheetState.previewMeta?.scannedAt, new Date().toISOString()),
    appliedAt: new Date().toISOString(),
    extractedCount,
    warnings: (Array.isArray(sheetState.warnings) ? sheetState.warnings : []).map((value) => safeText(value, '')).filter(Boolean),
    uncertainFields: (Array.isArray(sheetState.uncertainFields) ? sheetState.uncertainFields : [])
      .map((item) => ({
        path: safeText(item?.path, ''),
        reason: safeText(item?.reason, '')
      }))
      .filter((item) => item.path || item.reason)
  };
}

function getSectionRow(source, section, index, direction, groupIndex = null, locationIndex = null) {
  if (section === 'towerFootingResistance' && Number.isInteger(groupIndex)) {
    return source.towerFootingResistance?.[groupIndex]?.readings?.[index] || null;
  }
  if (section === 'soilResistivity' && Number.isInteger(locationIndex)) {
    return getSoilLocations(source)?.[locationIndex]?.[direction]?.[index] || null;
  }
  return getSectionRows(source, section, direction)[index] || null;
}

function getTestLabel(section) {
  const match = LOCAL_TEST_LIBRARY.find((test) => test.id === section);
  return match ? match.label : 'Measurement Row';
}

function getDraftReportTitle() {
  const selected = LOCAL_TEST_LIBRARY.filter((test) => state.draft.tests[test.id]);
  if (!selected.length) {
    return 'ElectroReports Assessment';
  }
  if (selected.length === 1) {
    return selected[0].label;
  }
  return 'Earthing System Health Assessment';
}

function getRowObservationTitle(section, index, direction, groupIndex = null, locationIndex = null) {
  if (section === 'towerFootingResistance' && Number.isInteger(groupIndex)) {
    const group = state.draft.towerFootingResistance?.[groupIndex];
    const foot = group?.readings?.[index]?.measurementPointLocation || TOWER_FOOT_POINTS[index] || `Foot-${index + 1}`;
    const tower = safeText(group?.mainLocationTower, `Tower ${groupIndex + 1}`);
    return `Tower Footing Resistance Measurement & Analysis | ${tower} | ${foot}`;
  }
  if (section === 'soilResistivity' && Number.isInteger(locationIndex)) {
    const location = getSoilLocations(state.draft)?.[locationIndex];
    const locationName = safeText(location?.name, `Location ${locationIndex + 1}`);
    return `${getTestLabel(section)} | ${locationName} | Row ${index + 1} | ${String(direction || 'direction1').replace('direction', 'Direction ')}`;
  }
  if (section === 'loopImpedanceTest') {
    const row = state.draft.loopImpedanceTest?.[index];
    return `${getTestLabel(section)} | ${safeText(row?.location || row?.feederTag, `Group ${Math.floor(index / 3) + 1}`)} | ${safeText(row?.measuredPoints, 'Point')}`;
  }
  if (section === 'prospectiveFaultCurrent') {
    const row = state.draft.prospectiveFaultCurrent?.[index];
    return `${getTestLabel(section)} | ${safeText(row?.location || row?.feederTag, `Group ${Math.floor(index / 3) + 1}`)} | ${safeText(row?.measuredPoints, 'Point')}`;
  }
  const base = `${getTestLabel(section)} | Row ${index + 1}`;
  if (section !== 'soilResistivity') {
    return base;
  }
  return `${base} | ${direction === 'direction2' ? 'Direction 2' : 'Direction 1'}`;
}

function toNumber(value) {
  const numeric = Number.parseFloat(String(value === undefined || value === null ? '' : value).trim());
  return Number.isFinite(numeric) ? numeric : null;
}

function round(value, digits = 2) {
  if (!Number.isFinite(value)) {
    return null;
  }
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function loadFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error(`Failed to read ${file.name || 'image'}.`));
    reader.readAsDataURL(file);
  });
}

function compressImageDataUrl(dataUrl, maxDimension = 1600, quality = 0.82) {
  return new Promise((resolve) => {
    const image = new window.Image();
    image.onload = () => {
      const largestSide = Math.max(image.width, image.height, 1);
      const scale = Math.min(1, maxDimension / largestSide);
      const width = Math.max(1, Math.round(image.width * scale));
      const height = Math.max(1, Math.round(image.height * scale));
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      const context = canvas.getContext('2d');
      if (!context) {
        resolve(dataUrl);
        return;
      }
      context.drawImage(image, 0, 0, width, height);
      resolve(canvas.toDataURL('image/jpeg', quality));
    };
    image.onerror = () => resolve(dataUrl);
    image.src = dataUrl;
  });
}

async function buildObservationPhoto(file) {
  const sourceDataUrl = await loadFileAsDataUrl(file);
  const optimizedDataUrl = await compressImageDataUrl(sourceDataUrl);
  return {
    id: buildRowId('photo'),
    name: safeText(file?.name, 'Observation image'),
    dataUrl: optimizedDataUrl
  };
}

function average(values) {
  const filtered = values.filter((value) => Number.isFinite(value));
  if (!filtered.length) {
    return null;
  }
  return filtered.reduce((sum, value) => sum + value, 0) / filtered.length;
}

// ---- Assessments --------------------------------------------------------------
// Every status, priority, limit, basis, remark and soil class shown in this file is
// computed by the server (src/assessmentService.js, rules from the acceptance
// register). The browser holds no rules: it sends the report or draft, caches what
// comes back, and renders it. Until a result arrives rows read as Pending.

const ASSESSMENT_SECTIONS = [
  'electrodeResistance',
  'continuityTest',
  'loopImpedanceTest',
  'prospectiveFaultCurrent',
  'riserIntegrityTest',
  'groundResistance',
  'stepTouchVoltage',
  'earthContinuityTest',
  'towerFootingResistance'
];

const PENDING_ROW_ASSESSMENT = Object.freeze({
  status: Object.freeze({ label: 'Pending', tone: 'neutral' }),
  standard: Object.freeze({ reference: '', limitLabel: '' }),
  priority: null,
  limit: null,
  totalImpedanceZt: null,
  totalCurrentItotal: null,
  zsat: null,
  comment: '',
  remark: ''
});

const EMPTY_ASSESSMENT = Object.freeze({
  profile: null,
  basis: { rho: null, rhoAssumed: false, rhoSource: '' },
  soil: {
    direction1Average: null,
    direction2Average: null,
    overallAverage: null,
    category: { label: 'Insufficient Data', tone: 'neutral' },
    locations: []
  },
  tests: {},
  modules: [],
  actions: { electrodeP1: 0, electrodeP2: 0, criticalTotal: 0, total: 0 },
  measurementCoverage: { total: 0, completed: 0, score: 0 },
  healthySignals: 0,
  tone: 'neutral',
  label: 'Calculating...',
  banners: { electrode: '' },
  validation: { errors: [], warnings: [] },
  reviewItems: [],
  references: {}
});

const assessmentCache = {
  byObject: new WeakMap(),
  byKey: new Map(),
  bySource: new WeakMap(),
  draft: { signature: null, data: null, timer: null, requested: null }
};

function assessmentRowKey(section, row) {
  const id = row && (row.rowId || row.groupId);
  return id ? `${section}:${id}` : null;
}

// Bind a server result to the rows it describes (by object and by id, since detail
// tables render copies of rows). A result whose row counts no longer match the
// source is not bound at all: showing another row's status is worse than Pending.
function registerAssessment(source, data) {
  if (!source || !data || data.error) {
    return;
  }
  assessmentCache.bySource.set(source, data);
  ASSESSMENT_SECTIONS.forEach((section) => {
    const rows = Array.isArray(source[section]) ? source[section] : [];
    const results = data.tests && data.tests[section] ? data.tests[section].rows : [];
    if (rows.length !== results.length) {
      return;
    }
    rows.forEach((row, index) => {
      if (!row || typeof row !== 'object') {
        return;
      }
      assessmentCache.byObject.set(row, results[index]);
      const key = assessmentRowKey(section, row);
      if (key) {
        assessmentCache.byKey.set(key, results[index]);
      }
    });
  });
}

function rowAssessment(section, row) {
  if (!row || typeof row !== 'object') {
    return PENDING_ROW_ASSESSMENT;
  }
  const key = assessmentRowKey(section, row);
  return assessmentCache.byObject.get(row) || (key && assessmentCache.byKey.get(key)) || PENDING_ROW_ASSESSMENT;
}

function sourceAssessment(source) {
  if (!source) {
    return EMPTY_ASSESSMENT;
  }
  if (!assessmentCache.bySource.has(source) && source.assessment && !source.assessment.error) {
    registerAssessment(source, source.assessment);
  }
  return assessmentCache.bySource.get(source) || EMPTY_ASSESSMENT;
}

// The request body for a draft: rule inputs only, no photos or OCR bookkeeping.
function draftAssessmentBody(draft) {
  return JSON.stringify(draft, (key, value) => (key === 'rowPhotos' || key === 'ocrImports' ? undefined : value));
}

// Called on every render of the builder. Rebinds the latest result to the current row
// objects, and asks the server for a fresh one (debounced) when the draft has changed.
function syncDraftAssessment() {
  const slot = assessmentCache.draft;
  const body = draftAssessmentBody(state.draft);
  if (slot.data && slot.signature === body) {
    registerAssessment(state.draft, slot.data);
    return;
  }
  if (slot.data) {
    registerAssessment(state.draft, slot.data);
  }
  if (slot.requested === body) {
    return;
  }
  slot.requested = body;
  window.clearTimeout(slot.timer);
  slot.timer = window.setTimeout(async () => {
    try {
      const data = await api('/api/assessments', { method: 'POST', body });
      if (slot.requested !== body) {
        return;
      }
      slot.signature = body;
      slot.data = data;
      if (state.view === 'builder' && draftAssessmentBody(state.draft) === body) {
        rerenderPreservingFocus();
      }
    } catch (error) {
      slot.requested = null;
    }
  }, 200);
}

function selectedTests(source) {
  return state.catalog.filter((test) => Boolean(source.tests[test.id]));
}

function selectedTestCount(source) {
  return selectedTests(source).length;
}

// Summary for the dashboard, builder and detail views, computed by the server.
function assessmentSummary(source) {
  return { ...sourceAssessment(source), selected: selectedTests(source) };
}

function formatDisplayDate(value) {
  const text = safeText(value, '');
  if (!text) {
    return '';
  }

  const candidate = /^\d+$/.test(text) ? Number(text) : text;
  const date = new Date(candidate);
  if (!Number.isFinite(date.getTime())) {
    return text;
  }

  return new Intl.DateTimeFormat('en-GB', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  }).format(date);
}

function isLikelyZohoReference(value) {
  const text = safeText(value, '');
  return /^\d{10,}$/.test(text);
}

function getZohoProjectDisplayName(project) {
  return safeText(project.name || project.projectNumber, 'Untitled Project');
}

function getZohoProjectVisibleCode(project) {
  const projectNumber = safeText(project.projectNumber, '');
  const projectName = safeText(project.name, '');
  if (!projectNumber || isLikelyZohoReference(projectNumber) || projectNumber === projectName) {
    return '';
  }
  return projectNumber;
}

function getProjectStageMeta(value) {
  const label = safeText(value, 'Not Synced');
  const normalized = label.toLowerCase().replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();

  if (!normalized) {
    return { label: 'Not Synced', tone: 'neutral' };
  }

  if (
    normalized.includes('review') ||
    normalized.includes('active') ||
    normalized.includes('execution') ||
    normalized.includes('progress') ||
    normalized.includes('ongoing')
  ) {
    return { label, tone: 'brand' };
  }

  if (
    normalized.includes('plan') ||
    normalized.includes('pending') ||
    normalized.includes('hold')
  ) {
    return { label, tone: 'warning' };
  }

  if (
    normalized.includes('complete') ||
    normalized.includes('closed') ||
    normalized.includes('done')
  ) {
    return { label, tone: 'healthy' };
  }

  if (
    normalized.includes('cancel') ||
    normalized.includes('drop')
  ) {
    return { label, tone: 'critical' };
  }

  return { label, tone: 'neutral' };
}

function builderReadiness(assessment = assessmentSummary(state.draft)) {
  const project = state.draft.project;
  const projectComplete = ['projectNo', 'clientName', 'siteLocation', 'reportDate', 'engineerName'].filter((field) => safeText(project[field], '')).length;
  const projectScore = Math.round((projectComplete / 5) * 100);
  const selectionComplete = assessment.selected.length ? 1 : 0;
  const selectionScore = selectionComplete ? 100 : 0;
  const measurementScore = assessment.measurementCoverage.score;
  const score = Math.round(projectScore * 0.45 + selectionScore * 0.1 + measurementScore * 0.45);
  return {
    score,
    projectComplete,
    selectionComplete,
    projectScore,
    selectionScore,
    measurementScore
  };
}

function stepGuidance(stepId) {
  const map = {
    project: 'Capture the core job references first so the saved report and PDF stay traceable in the field.',
    selection: 'Select only the sections actually tested on site. The rest will stay out of the report.',
    soilResistivity: 'Tick the reading unit (R or rho), record each direction at each spacing, and the app calculates rho, the direction means and the IS 3043 corrosivity class.',
    electrodeResistance: 'Each electrode status is calculated from the site soil resistivity and the electrode dimensions.',
    continuityTest: 'Record the test-lead resistance once in the header; it is subtracted from every reading.',
    loopImpedanceTest: 'Record Zs values for protection verification at the tested points.',
    prospectiveFaultCurrent: 'Record the breaking capacity from the nameplate, and both PFC (L-PE) and PSC (L-N) for each phase.',
    groundResistance: 'Read at 50%, 61.8% and 70%; each reading is compared with the limit of the pit sharing its tag.',
    stepTouchVoltage: 'Record If, ts and the surface layer in the header; each voltage is scaled by If/It before comparison.',
    riserIntegrityTest: 'Use this for resistance checks towards equipment and towards the grid.',
    earthContinuityTest: 'Retired sheet: record these points on Earth Continuity & Bonding (MS-03).',
    towerFootingResistance: 'Each tower location keeps 4 fixed footing rows with grouped totals; Zsat comes from the project acceptance profile.'
  };
  return map[stepId] || 'Continue entering the field measurements for the selected scope.';
}

function currentStepNote(stepId) {
  const map = {
    project: 'Keep the project references complete so saved reports and PDFs stay traceable.',
    selection: 'Choose only the measurement sheets actually executed on site.',
    soilResistivity: 'Complete both directions so the soil mean and corrosivity class can calculate.',
    electrodeResistance: 'Enter the without-grid (isolated electrode) reading; it is compared with the calculated limit.',
    continuityTest: 'Measured R is the decision field; distance and test current are context.',
    loopImpedanceTest: 'Use measured Zs to flag circuits that need protection review.',
    prospectiveFaultCurrent: 'Without a recorded breaking capacity the result is Not Assessable.',
    groundResistance: 'Use the same tag as the pit so its calculated limit can be applied.',
    stepTouchVoltage: 'Test current It is needed on every row for the If/It scaling.',
    riserIntegrityTest: 'Capture both equipment-side and grid-side resistance for each point.',
    earthContinuityTest: 'Use measured value to identify earth path continuity issues quickly.',
    towerFootingResistance: 'Each tower location needs all 4 footing rows to complete grouped totals.'
  };
  return map[stepId] || '';
}

function renderSelectedModuleTags(source) {
  const selected = selectedTests(source);
  if (!selected.length) {
    return '<p class="muted">No measurement sections selected yet.</p>';
  }
  return `<div class="tag-cloud">${selected.map((test) => `<span class="soft-chip">${escapeHtml(test.shortLabel || test.label)}</span>`).join('')}</div>`;
}

function renderStandardsSummary(source) {
  const selected = selectedTests(source);
  if (!selected.length) {
    return '<p class="muted">Select a measurement section to show its reference standards.</p>';
  }

  return `
    <div class="standards-list">
      ${selected
        .map((test) => {
          const references = sourceAssessment(source).references[test.id] || [];
          return `
            <article class="standard-entry">
              <strong>${escapeHtml(test.shortLabel || test.label)}</strong>
              <span>${escapeHtml(references.join(' • ') || 'Reference loads with the assessment')}</span>
            </article>
          `;
        })
        .join('')}
    </div>
  `;
}

function renderProgressBar(label, value, tone = 'healthy') {
  return `
    <div class="progress-block">
      <div class="progress-head">
        <span>${escapeHtml(label)}</span>
        <strong>${value}%</strong>
      </div>
      <div class="progress-track">
        <span class="progress-fill progress-fill-${tone}" style="width:${Math.max(0, Math.min(100, value))}%"></span>
      </div>
    </div>
  `;
}

function renderAssessmentHighlights(assessment) {
  const towerModule = assessment.modules.find((module) => module.id === 'towerFootingResistance');

  const cards = [
    {
      label: 'Selected Scope',
      value: assessment.selected.length ? `${assessment.selected.length}` : '-',
      tone: assessment.selected.length ? 'brand' : 'neutral'
    },
    {
      label: 'Healthy Modules',
      value: assessment.selected.length ? `${assessment.healthySignals}/${assessment.selected.length}` : '-',
      tone: assessment.healthySignals === assessment.selected.length && assessment.selected.length ? 'healthy' : assessment.healthySignals ? 'warning' : 'neutral'
    },
    {
      label: 'Action Points',
      value: String(assessment.actions.total),
      tone: assessment.actions.total ? 'warning' : 'healthy'
    },
    {
      label: 'Critical Findings',
      value: String(assessment.actions.criticalTotal),
      tone: assessment.actions.criticalTotal ? 'critical' : 'healthy'
    },
    {
      label: 'Mean Soil',
      value: assessment.soil.overallAverage === null ? '-' : `${assessment.soil.overallAverage} ohm-m`,
      tone: assessment.soil.category.tone
    },
    {
      label: 'Tower Groups',
      value: towerModule ? `${towerModule.completeGroups || 0}/${towerModule.groupCount || 0}` : '-',
      tone: towerModule ? towerModule.tone : 'neutral'
    }
  ];

  return `
    <div class="aside-stats aside-stats-rich">
      ${cards
        .map(
          (card) => `
            <div class="aside-stat-card aside-stat-card-${card.tone}">
              <span>${escapeHtml(card.label)}</span>
              <strong>${escapeHtml(card.value)}</strong>
            </div>
          `
        )
        .join('')}
    </div>
  `;
}

function renderModuleHealthList(assessment) {
  if (!assessment.modules.length) {
    return '<p class="muted">Select a measurement section to see module-specific health signals.</p>';
  }

  return `
    <div class="module-health-list">
      ${assessment.modules
        .map(
          (module) => `
            <article class="module-health-item module-health-item-${module.tone}">
              <div class="module-health-head">
                <strong>${escapeHtml(module.shortLabel)}</strong>
                ${pill(module.tone, module.badge)}
              </div>
              <p>${escapeHtml(module.detail)}</p>
            </article>
          `
        )
        .join('')}
    </div>
  `;
}

function renderScopeStandardsPanel(source) {
  return `
    <div class="builder-info-stack">
      <div>
        <h4>Selected Scope</h4>
        ${renderSelectedModuleTags(source)}
      </div>
      <div>
        <h4>Reference Standards</h4>
        ${renderStandardsSummary(source)}
      </div>
    </div>
  `;
}

function getSteps() {
  return [
    { id: 'project', label: 'Project' },
    { id: 'selection', label: 'Selection' },
    ...selectedTests(state.draft).map((test) => ({
      id: test.id,
      label: test.shortLabel || test.label
    }))
  ];
}

// Entry screens for the measurement sheets use the full page width; the project and
// selection steps are forms and keep a comfortable width.
function isSheetStep(stepId) {
  return stepId !== 'project' && stepId !== 'selection';
}

function currentStepId() {
  const steps = getSteps();
  return steps[state.stepIndex] ? steps[state.stepIndex].id : 'project';
}

function showToast(message, tone = 'neutral') {
  state.toast = { message, tone };
  render();
  window.clearTimeout(showToast.timeoutId);
  showToast.timeoutId = window.setTimeout(() => {
    state.toast = null;
    render();
  }, 3200);
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    headers: { 'Content-Type': 'application/json' },
    ...options
  });

  if (response.status === 401) {
    // the session ended (idle timeout, signed out elsewhere): back to the sign-in screen
    authUi.user = null;
    state.view = 'signin';
    render();
  }
  if (!response.ok) {
    const data = await response.json().catch(() => ({ message: 'Request failed.' }));
    const failure = new Error(data.message || 'Request failed.');
    failure.errors = data.errors;
    failure.code = data.code;
    failure.reportId = data.reportId;
    throw failure;
  }

  if (response.status === 204) {
    return null;
  }

  return response.json();
}

async function apiFormData(path, formData) {
  const response = await fetch(path, {
    method: 'POST',
    body: formData
  });

  if (!response.ok) {
    const data = await response.json().catch(() => ({ message: 'Request failed.' }));
    const failure = new Error(data.message || 'Request failed.');
    failure.errors = data.errors;
    throw failure;
  }

  return response.json();
}

async function loadCatalog() {
  try {
    const data = await api('/api/catalog');
    if (Array.isArray(data.tests) && data.tests.length) {
      state.catalog = data.tests;
    }
  } catch (_error) {
    state.catalog = [...LOCAL_TEST_LIBRARY];
  }
}

async function loadAiStatus() {
  state.ai.loadingStatus = true;
  render();
  try {
    const data = await api('/api/ai/status');
    state.ai.configured = Boolean(data.configured);
    state.ai.model = safeText(data.model, '');
    state.ai.referenceDocuments = Array.isArray(data.referenceDocuments) ? data.referenceDocuments : [];
  } catch (_error) {
    state.ai.configured = false;
    state.ai.model = '';
    state.ai.referenceDocuments = [];
  } finally {
    state.ai.loadingStatus = false;
    render();
  }
}

async function loadOcrStatus() {
  state.ocr.loadingStatus = true;
  render();
  try {
    const data = await api('/api/ocr/status');
    state.ocr.configured = Boolean(data.configured);
    state.ocr.enabled = Boolean(data.enabled);
    state.ocr.model = safeText(data.model, '');
    state.ocr.supportedSheets = Array.isArray(data.supportedSheets) ? data.supportedSheets : [];
  } catch (_error) {
    state.ocr.configured = false;
    state.ocr.enabled = false;
    state.ocr.model = '';
    state.ocr.supportedSheets = [];
  } finally {
    state.ocr.loadingStatus = false;
    render();
  }
}

function getOcrRoute(sheetId) {
  const routes = {
    soilResistivity: '/api/ocr/soil/preview',
    electrodeResistance: '/api/ocr/electrode/preview',
    continuityTest: '/api/ocr/continuity/preview',
    loopImpedanceTest: '/api/ocr/loop/preview',
    prospectiveFaultCurrent: '/api/ocr/fault/preview',
    riserIntegrityTest: '/api/ocr/riser/preview',
    earthContinuityTest: '/api/ocr/earth-continuity/preview',
    towerFootingResistance: '/api/ocr/tower-footing/preview'
  };
  return routes[sheetId] || '';
}

function getOcrSheetLabel(sheetId) {
  const test = LOCAL_TEST_LIBRARY.find((item) => item.id === sheetId);
  return test ? test.label : 'Measurement Sheet';
}

async function previewSheetOcr(sheetId) {
  const sheetState = getOcrSheetState(sheetId);
  const route = getOcrRoute(sheetId);

  if (!state.ocr.enabled || !state.ocr.configured) {
    sheetState.error = 'Gemini OCR is not configured yet.';
    render();
    return;
  }

  if (!route || !state.ocr.supportedSheets.includes(sheetId)) {
    sheetState.error = 'OCR preview is not available for this measurement sheet yet.';
    render();
    return;
  }

  if (!sheetState.selectedFile) {
    sheetState.error = `Choose a ${getOcrSheetLabel(sheetId)} image or PDF first.`;
    render();
    return;
  }

  sheetState.scanning = true;
  sheetState.error = '';
  sheetState.preview = null;
  sheetState.previewMeta = null;
  sheetState.draftPatch = null;
  sheetState.warnings = [];
  sheetState.uncertainFields = [];
  render();

  try {
    const formData = new FormData();
    formData.append('document', sheetState.selectedFile);
    const result = await apiFormData(route, formData);
    sheetState.preview = result.preview || null;
    sheetState.previewMeta = {
      fileName: safeText(result.fileName, sheetState.selectedFile?.name || 'uploaded-sheet'),
      mimeType: safeText(result.mimeType, sheetState.selectedFile?.type || ''),
      fileSize: Number.isFinite(Number(result.fileSize)) ? Number(result.fileSize) : Number(sheetState.selectedFile?.size || 0),
      scannedAt: safeText(result.scannedAt, new Date().toISOString())
    };
    sheetState.draftPatch = result.draftPatch || null;
    sheetState.warnings = Array.isArray(result.preview?.warnings) ? result.preview.warnings : [];
    sheetState.uncertainFields = Array.isArray(result.preview?.uncertainFields) ? result.preview.uncertainFields : [];

    const hasData =
      (sheetId === 'soilResistivity' && Array.isArray(sheetState.preview?.locations) && sheetState.preview.locations.length) ||
      ((sheetId === 'electrodeResistance' ||
        sheetId === 'continuityTest' ||
        sheetId === 'riserIntegrityTest' ||
        sheetId === 'earthContinuityTest') &&
        Array.isArray(sheetState.preview?.rows) &&
        sheetState.preview.rows.length) ||
      ((sheetId === 'loopImpedanceTest' ||
        sheetId === 'prospectiveFaultCurrent' ||
        sheetId === 'towerFootingResistance') &&
        Array.isArray(sheetState.preview?.groups) &&
        sheetState.preview.groups.length);

    if (!hasData) {
      sheetState.error = `No ${getOcrSheetLabel(sheetId)} rows were extracted from the uploaded sheet.`;
    } else {
      showToast(`${getOcrSheetLabel(sheetId)} scanned. Review the preview before applying it.`, 'healthy');
    }
  } catch (error) {
    sheetState.error = error.message || `${getOcrSheetLabel(sheetId)} OCR preview failed.`;
  } finally {
    sheetState.scanning = false;
    render();
  }
}

function applySheetOcrPreview(sheetId) {
  const sheetState = getOcrSheetState(sheetId);

  if (sheetId === 'soilResistivity') {
    const previewLocations = sheetState.draftPatch?.locations;
    if (!Array.isArray(previewLocations) || !previewLocations.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }

    if (draftHasMeaningfulSoilData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Soil Resistivity entries. Continue?');
      if (!confirmed) {
        return;
      }
    }

    state.draft.soilResistivity.locations = previewLocations.map((location, index) => ({
      ...location,
      name: safeText(location.name, `Location ${index + 1}`)
    }));
  } else if (sheetId === 'electrodeResistance') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulElectrodeData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Earth Electrode Resistance entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.electrodeResistance = rows.map((row) => upgradeRow('electrodeResistance', row));
  } else if (sheetId === 'continuityTest') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulContinuityData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Continuity Test entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.continuityTest = rows.map((row, index) => ({
      ...row,
      srNo: safeText(row.srNo, String(index + 1))
    }));
  } else if (sheetId === 'loopImpedanceTest') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulLoopData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Loop Impedance entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.loopImpedanceTest = rows.map((row, index) => ({
      ...row,
      srNo: safeText(row.srNo, String(index + 1)),
      measuredPoints: safeText(row.measuredPoints, PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length])
    }));
  } else if (sheetId === 'prospectiveFaultCurrent') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulFaultData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Prospective Fault Current entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.prospectiveFaultCurrent = rows.map((row, index) => ({
      ...row,
      srNo: safeText(row.srNo, String(index + 1)),
      measuredPoints: safeText(row.measuredPoints, PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length])
    }));
  } else if (sheetId === 'riserIntegrityTest') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulRiserData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Riser / Grid Integrity entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.riserIntegrityTest = rows.map((row, index) => ({
      ...row,
      srNo: safeText(row.srNo, String(index + 1))
    }));
  } else if (sheetId === 'earthContinuityTest') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulEarthContinuityData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Earth Continuity entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.earthContinuityTest = rows.map((row, index) => ({
      ...row,
      srNo: safeText(row.srNo, String(index + 1))
    }));
  } else if (sheetId === 'towerFootingResistance') {
    const rows = sheetState.draftPatch?.rows;
    if (!Array.isArray(rows) || !rows.length) {
      sheetState.error = 'No OCR preview is ready to apply.';
      render();
      return;
    }
    if (draftHasMeaningfulTowerData(state.draft)) {
      const confirmed = window.confirm('Applying this OCR preview will replace the current Tower Footing entries. Continue?');
      if (!confirmed) {
        return;
      }
    }
    state.draft.towerFootingResistance = rows.map((row, index) => ({
      ...row,
      srNo: safeText(row.srNo, String(index + 1))
    }));
  } else {
    return;
  }

  state.draft.ocrImports = {
    ...(state.draft.ocrImports || {}),
    [sheetId]: buildOcrImportRecord(sheetId, sheetState)
  };
  scheduleDraftAutosave();
  resetOcrSheetState(sheetId, false);
  getOcrSheetState(sheetId).mode = 'manual';
  showToast(`OCR preview applied to the ${getOcrSheetLabel(sheetId)} sheet.`, 'healthy');
  render();
}

async function loadReports() {
  state.loadingReports = true;
  render();
  try {
    state.reports = await api(`/api/reports${state.search ? `?q=${encodeURIComponent(state.search)}` : ''}`);
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    state.loadingReports = false;
    render();
  }
}

async function loadZohoProjects(query = '') {
  state.zoho.loadingProjects = true;
  render();
  try {
    state.zoho.projects = await api(`/api/zoho/projects${query ? `?q=${encodeURIComponent(query)}` : ''}`);
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    state.zoho.loadingProjects = false;
    render();
  }
}

async function loadZohoUsers(projectId) {
  if (!projectId) {
    state.zoho.users = [];
    render();
    return;
  }
  state.zoho.loadingUsers = true;
  render();
  try {
    const [projectUsersResult, portalUsersResult] = await Promise.allSettled([
      api(`/api/zoho/projects/${encodeURIComponent(projectId)}/users`),
      api('/api/zoho/users')
    ]);

    if (projectUsersResult.status === 'rejected' && portalUsersResult.status === 'rejected') {
      throw projectUsersResult.reason || portalUsersResult.reason;
    }

    const mergedUsers = [
      ...(projectUsersResult.status === 'fulfilled' ? projectUsersResult.value : []),
      ...(portalUsersResult.status === 'fulfilled' ? portalUsersResult.value : [])
    ].reduce((map, user) => {
      const displayName = safeText(user.displayName || user.name, '');
      if (!displayName) {
        return map;
      }
      const key = `${safeText(user.id, '')}::${safeText(user.email, '').toLowerCase()}::${displayName.toLowerCase()}`;
      if (!map.has(key)) {
        map.set(key, { ...user, displayName });
      }
      return map;
    }, new Map());

    state.zoho.users = Array.from(mergedUsers.values());
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    state.zoho.loadingUsers = false;
    render();
  }
}

async function loadZohoClientUsers(projectId) {
  if (!projectId) {
    state.zoho.clientUsers = [];
    return;
  }
  state.zoho.loadingClientUsers = true;
  render();
  try {
    const clientUsers = await api(`/api/zoho/projects/${encodeURIComponent(projectId)}/client-users`);
    state.zoho.clientUsers = Array.isArray(clientUsers) ? clientUsers : [];
    // Company name wins: prefer the Zoho client company over the name-parsed abbreviation.
    const withCompany = state.zoho.clientUsers.find((user) => safeText(user.companyName, ''));
    if (withCompany) {
      state.draft.project.clientName = safeText(withCompany.companyName, state.draft.project.clientName);
    }
  } catch (error) {
    // Client contacts are supplementary (e.g. portal returns 403 without the
    // client-user scope); keep the parsed client name and fail silently in the UI.
    state.zoho.clientUsers = [];
    console.warn('Zoho client contacts unavailable:', error.message);
  } finally {
    state.zoho.loadingClientUsers = false;
    render();
  }
}

function applyZohoProject(projectId) {
  const project = state.zoho.projects.find((item) => String(item.id) === String(projectId));
  if (!project) {
    state.draft.project.zohoProjectId = '';
    state.draft.project.zohoProjectName = '';
    state.draft.project.zohoProjectOwner = '';
    state.draft.project.zohoProjectStage = '';
    state.zoho.users = [];
    state.zoho.clientUsers = [];
    resetPlanUi();
    render();
    return;
  }

  state.draft.project.zohoProjectId = safeText(project.id, '');
  state.draft.project.zohoProjectName = safeText(project.name, '');
  state.draft.project.zohoProjectOwner = safeText(project.ownerName, '');
  state.draft.project.zohoProjectStage = safeText(project.stage, '');
  state.draft.project.projectNo = safeText(
    project.projectNumber || getZohoProjectVisibleCode(project) || project.name,
    state.draft.project.projectNo
  );
  state.draft.project.clientName = safeText(project.clientName, state.draft.project.clientName);
  state.draft.project.workOrder = safeText(project.workOrderNo, state.draft.project.workOrder);
  // The engineer is the owner of the Onsite tasks (set from the Zoho task plan), not the project owner.
  loadZohoPlan(project.id);
  loadZohoUsers(project.id);
  loadZohoClientUsers(project.id);
  showToast(`Synced ${safeText(project.projectNumber || project.name)} from Zoho Projects.`, 'healthy');
  render();
}

function pill(tone, text) {
  return `<span class="pill pill-${tone}">${escapeHtml(text)}</span>`;
}

function renderFloatingShapes(count = 18, className = 'hero-shapes') {
  const shapes = Array.from({ length: count }, (_, index) => {
    const size = 5 + ((index * 2) % 6);
    const left = 6 + ((index * 7) % 88);
    const delay = -((index * 1.37) % 11.5);
    const duration = 9.4 + ((index * 0.83) % 5.2);
    const driftOne = -9 + ((index * 5) % 19);
    const driftTwo = -12 + ((index * 7) % 25);
    const driftThree = -8 + ((index * 11) % 17);
    const radius = index % 3 === 0 ? '999px' : '8px';

    return `
      <span
        style="
          --shape-left:${left}%;
          --shape-size:${size}px;
          --shape-delay:${delay}s;
          --shape-duration:${duration}s;
          --shape-drift-one:${driftOne}px;
          --shape-drift-two:${driftTwo}px;
          --shape-drift-three:${driftThree}px;
          --shape-radius:${radius};
        "
      ></span>
    `;
  }).join('');

  return `<div class="${className}" aria-hidden="true">${shapes}</div>`;
}

function brandHeader() {
  return `
    <header class="brand-bar">
      <div class="brand-shell">
        <div class="brand-lockup">
          <img class="brand-logo" src="/assets/elegrow-logo-full.png" alt="Elegrow Technology" />
        </div>
        <div class="brand-support">
          ${authUi.user ? renderUserBar() : '<span>Need help?</span><strong>info@elegrow.com</strong>'}
        </div>
      </div>
    </header>
  `;
}

function dashboardView() {
  const uniqueClients = new Set(state.reports.map((report) => report.project.clientName)).size;
  const reportsNeedingAction = state.reports.filter((report) => assessmentSummary(report).tone !== 'healthy').length;
  const zohoCount = state.zoho.projects.length;

  return `
    <section class="hero-panel hero-panel-dashboard">
      <div class="hero-banner">
        <div class="hero-copy">
          <p class="eyebrow">ElectroReports</p>
          <h2>Electrical Reporting for Field Engineers</h2>
          <p>
            Capture site measurements, sync live Zoho Projects, and prepare clear client-ready reports with a faster,
            more dependable workflow.
          </p>
          <div class="button-row hero-actions">
            <button class="button button-primary" data-action="new-report">Start New Report</button>
            <button class="button button-secondary" data-action="search-focus">Search Reports</button>
          </div>
          <div class="hero-inline-notes hero-facts">
            <span>Zoho Projects ${zohoCount ? 'connected' : 'ready to sync'}</span>
            <span>PDF export enabled</span>
            <span>Built for electrical audit teams</span>
          </div>
        </div>
        <div class="hero-side" aria-hidden="true">
          <div class="hero-mark"></div>
        </div>
        ${renderFloatingShapes(58)}
      </div>
    </section>

    <section class="summary-orb-row" aria-label="Dashboard summary">
      <article class="summary-orb summary-orb-blue reveal-card">
        <span>Total Reports</span>
        <strong>${state.reports.length}</strong>
      </article>
      <article class="summary-orb summary-orb-silver reveal-card">
        <span>Clients</span>
        <strong>${uniqueClients}</strong>
      </article>
      <article class="summary-orb summary-orb-gold reveal-card">
        <span>Needs Review</span>
        <strong>${reportsNeedingAction}</strong>
      </article>
    </section>

    ${renderHomeNotices()}

    ${renderHomeProjects()}

    <section class="surface toolbar">
      <div class="dashboard-section-copy">
        <p class="section-kicker">Dashboard</p>
        <h3>Saved Reports</h3>
        <p class="muted">Search saved reports by project number, client, location, or engineer.</p>
      </div>
      <div class="toolbar-actions">
        <input
          id="reportSearchInput"
          class="search-input"
          type="search"
          placeholder="Search project number, client, location, or engineer"
          value="${escapeHtml(state.search)}"
        />
      </div>
    </section>

    ${
      state.loadingReports
        ? '<section class="surface empty-state"><h3>Loading reports...</h3></section>'
        : state.reports.length
          ? `<section class="card-grid">
              ${state.reports
                .map((report) => {
                  const summary = assessmentSummary(report);
                  return `
                    <article class="report-card report-card-${summary.tone}">
                      <div class="report-card-head">
                        <span class="chip">${escapeHtml(report.project.projectNo)}</span>
                        <span class="report-card-date">${escapeHtml(formatDisplayDate(report.project.reportDate) || report.project.reportDate)}</span>
                      </div>
                      <h4>${escapeHtml(report.project.clientName)}</h4>
                      <p class="report-card-location">${escapeHtml(report.project.siteLocation)}</p>
                      <div class="report-card-meta">
                        <div>
                          <span>Engineer</span>
                          <strong>${escapeHtml(report.project.engineerName)}</strong>
                        </div>
                        <div>
                          <span>Tests</span>
                          <strong>${selectedTestCount(report)}</strong>
                        </div>
                      </div>
                      <div class="report-card-tags">
                        ${renderSelectedModuleTags(report)}
                      </div>
                      <div class="report-card-summary">
                        <div>
                          <span>Mean Soil</span>
                          <strong>${summary.soil.overallAverage === null ? '-' : `${summary.soil.overallAverage} ohm-m`}</strong>
                        </div>
                        <div>
                          ${pill(summary.tone, summary.label)}
                        </div>
                      </div>
                      <div class="report-card-alerts">
                        <span>P1 electrodes: <strong>${summary.actions.electrodeP1}</strong></span>
                        <span>Total action points: <strong>${summary.actions.total}</strong></span>
                      </div>
                      <div class="card-actions">
                        <button class="button button-secondary" data-action="open-report" data-id="${report.id}">View Details</button>
                        ${isReviewable(report) ? `<button class="button button-primary" data-action="open-review" data-id="${report.id}">Review</button>` : ''}
                        <button class="button button-ghost" data-action="delete-report" data-id="${report.id}">Delete</button>
                      </div>
                    </article>
                  `;
                })
                .join('')}
            </section>`
          : `<section class="surface empty-state">
              <div class="empty-state-copy">
                <h3>No Reports Yet</h3>
                <p>Create your first ElectroReports report to start building your live measurement archive.</p>
                <button class="button button-primary" data-action="new-report">Create First Report</button>
              </div>
            </section>`
    }

    <section class="surface zoho-overview">
      <div class="section-head">
        <div class="dashboard-section-copy">
          <p class="section-kicker">Zoho Projects</p>
          <h3>Live Zoho Projects</h3>
          <p class="muted zoho-overview-copy">Start a report from a live Zoho project and keep engineer, status, and project references aligned.</p>
        </div>
        <div class="button-row">
          <button class="button button-secondary" data-action="refresh-zoho-projects">Refresh Zoho</button>
        </div>
      </div>
      ${renderZohoProjectCards()}
    </section>
  `;
}

function sectionCard(title, subtitle, content) {
  return `
    <section class="surface section-card">
      <div class="section-head">
        <div>
          <p class="section-kicker">${escapeHtml(subtitle)}</p>
          <h3>${escapeHtml(title)}</h3>
        </div>
      </div>
      ${content}
    </section>
  `;
}

function renderField(label, value, bind, type = 'text', placeholder = '') {
  const input = `
      <input
        type="${type}"
        class="text-input"
        placeholder="${escapeHtml(placeholder)}"
        value="${escapeHtml(value || '')}"
        data-bind="${bind}"
      />`;
  return `
    <label class="field">
      <span>${escapeHtml(label)}</span>
      ${type === 'date' && typeof dateControl === 'function' ? dateControl(input, value) : input}
    </label>
  `;
}

function renderZohoProjectCards() {
  if (state.zoho.loadingProjects) {
    return '<div class="zoho-empty">Loading Zoho Projects...</div>';
  }
  if (!state.zoho.projects.length) {
    return '<div class="zoho-empty">No Zoho Projects available. Check your Zoho configuration or mock mode.</div>';
  }
  return `
    <div class="zoho-grid">
      ${state.zoho.projects.slice(0, 6).map((project) => {
        const stage = getProjectStageMeta(project.stage);
        const displayName = getZohoProjectDisplayName(project);
        const visibleCode = getZohoProjectVisibleCode(project);
        const visibleDate = formatDisplayDate(project.updatedAt || project.createdAt);
        const projectHeading = visibleCode || displayName;
        return `
          <article class="zoho-card zoho-card-${stage.tone} reveal-card">
            <div class="zoho-card-topline">
              ${visibleDate ? `<span class="meta-date-badge">${escapeHtml(visibleDate)}</span>` : ''}
            </div>
            <h4>${escapeHtml(projectHeading)}</h4>
            <div class="zoho-card-meta zoho-card-meta-compact">
              <div class="zoho-meta-block zoho-meta-panel">
                <span>Owner</span>
                <strong>${escapeHtml(project.ownerName || '-')}</strong>
              </div>
              <div class="zoho-meta-block zoho-meta-panel zoho-meta-panel-status">
                <span>Status</span>
                <strong>${escapeHtml(stage.label)}</strong>
              </div>
            </div>
            <div class="zoho-card-actions">
              <button class="button button-primary" data-action="open-workspace-zoho" data-project-id="${escapeHtml(project.id)}">Open workspace</button>
              <button class="button button-secondary" data-action="use-zoho-project" data-project-id="${escapeHtml(project.id)}">Single report</button>
            </div>
          </article>
        `;
      }).join('')}
    </div>
  `;
}

// Planned quantity per selected test (the denominator of "planned vs completed" in the
// reports). Entered by hand for now.
function renderPlannedQuantities() {
  const tests = state.catalog.filter((test) => state.draft.tests[test.id]);
  if (!tests.length) {
    return '';
  }
  const planned = state.draft.project.plannedQuantities || {};
  return `
    <div class="planned-qty">
      <p class="section-kicker">Planned quantities</p>
      <div class="field-grid">
        ${tests
          .map(
            (test) =>
              `<div class="planned-cell">${renderField(
                `${test.shortLabel || test.label} - planned ${test.unitLabel || 'qty'}`,
                planned[test.id],
                `project.plannedQuantities.${test.id}`,
                'number'
              )}${plannedNote(test.id)}</div>`
          )
          .join('')}
      </div>
    </div>`;
}

function renderProjectStep() {
  const zohoProjectOptions = state.zoho.projects
    .map((project) => {
      const selected = state.draft.project.zohoProjectId === String(project.id) ? 'selected' : '';
      const displayName = getZohoProjectDisplayName(project);
      const visibleCode = getZohoProjectVisibleCode(project);
      const label = visibleCode ? `${displayName} (${visibleCode})` : displayName;
      return `<option value="${escapeHtml(project.id)}" ${selected}>${escapeHtml(label)}</option>`;
    })
    .join('');

  return sectionCard(
    'Project Information',
    'Step 1',
    `
      <div class="zoho-sync-panel${state.editing ? ' is-hidden' : ''}">
        <div class="zoho-sync-copy">
          <p class="section-kicker">Zoho Projects Sync</p>
          <h4>Pick a live Zoho project to prefill this report</h4>
          <p>Project number, client, work order, owner, and project reference can be pulled directly from Zoho.</p>
        </div>
        <div class="zoho-sync-controls">
          <select class="text-input" id="zohoProjectPicker">
            <option value="">Select a Zoho project</option>
            ${zohoProjectOptions}
          </select>
          <button class="button button-secondary" data-action="refresh-zoho-projects">Refresh Zoho Projects</button>
        </div>
      </div>
      <div class="print-sheets-row">
        <p class="muted">Going to site? Print the measurement sheets for this project, pre-filled and ready to write on.</p>
        ${renderPrintButton()}
      </div>
      ${
        state.draft.project.zohoProjectId
          ? `
            <div class="zoho-selected-bar">
              <div class="zoho-selected-primary">
                <span>Zoho project</span>
                <strong>${escapeHtml(state.draft.project.zohoProjectName || '-')}</strong>
              </div>
              <div><span>Owner</span><strong>${escapeHtml(state.draft.project.zohoProjectOwner || '-')}</strong></div>
              <div><span>Stage</span><strong>${escapeHtml(state.draft.project.zohoProjectStage || '-')}</strong></div>
            </div>
          `
          : ''
      }
      ${
        state.zoho.users.length
          ? `
            <div class="zoho-user-strip">
              <span>Zoho users available for assignment</span>
              <p class="zoho-user-help">Project users and other Zoho portal users can be selected here for the technician field.</p>
              <div class="tag-cloud">
                ${state.zoho.users
                  .map((user) => {
                    const rawName = safeText(user.displayName || user.name, 'User');
                    const isActive =
                      safeText(state.draft.project.engineerName, '').toLowerCase() === rawName.toLowerCase();
                    return `<button class="soft-chip soft-chip-button${isActive ? ' soft-chip-button-active' : ''}" data-action="use-zoho-user" data-user-name="${escapeHtml(rawName)}">${escapeHtml(rawName)}</button>`;
                  })
                  .join('')}
              </div>
            </div>
          `
          : state.zoho.loadingUsers
            ? '<div class="zoho-user-strip"><span>Loading Zoho users...</span></div>'
            : ''
      }
      ${
        state.zoho.clientUsers && state.zoho.clientUsers.length
          ? `
            <div class="zoho-user-strip">
              <span>Client contact (from Zoho)</span>
              <p class="zoho-user-help">Client-side contacts on this Zoho project. Shown for reference; not saved to the report.</p>
              <div class="tag-cloud">
                ${state.zoho.clientUsers
                  .map((user) => {
                    const nm = safeText(user.displayName || user.name, 'Client contact');
                    const em = safeText(user.email, '');
                    const co = safeText(user.companyName, '');
                    const label = nm + (em ? ` <${em}>` : '') + (co ? ` — ${co}` : '');
                    return `<span class="soft-chip">${escapeHtml(label)}</span>`;
                  })
                  .join('')}
              </div>
            </div>
          `
          : state.zoho.loadingClientUsers
            ? '<div class="zoho-user-strip"><span>Loading client contacts...</span></div>'
            : ''
      }
      ${renderZohoTaskPanel()}
      <div class="field-grid">
        ${renderField('Project Number', state.draft.project.projectNo, 'project.projectNo', 'text', 'PRJ-2026-001')}
        ${renderField('Client Name', state.draft.project.clientName, 'project.clientName', 'text', 'Client Name')}
        ${renderField('Site Location', state.draft.project.siteLocation, 'project.siteLocation', 'text', 'Main Plant, Surat')}
        ${state.editing ? `<div class="field"><span>Work date</span><strong>${escapeHtml(formatDisplayDate(state.editing.workDate) || state.editing.workDate)}</strong><small class="muted">Fixed for this site day of the project.</small></div>` : renderField('Work date', state.draft.project.reportDate, 'project.reportDate', 'date')}
        ${renderField('Testing Engineer', state.draft.project.engineerName, 'project.engineerName', 'text', 'Engineer Name')}
        ${renderField('Unit / plant', state.draft.project.unit, 'project.unit', 'text', 'FI Plant')}
        ${renderField('Unit code (required for the final report)', state.draft.project.unitCode, 'project.unitCode', 'text', 'FI')}
        ${state.editing ? renderField('Plan for next day', state.draft.project.nextDayPlan, 'project.nextDayPlan', 'text', 'e.g. Earth pits EP-07 to EP-10') : ''}
        ${state.editing ? `<label class="field"><span>Data collection mode</span><select class="text-input" data-bind="project.dataCollectionMode">${['On-site entry', 'Printed sheets'].map((mode) => `<option value="${mode}" ${state.draft.project.dataCollectionMode === mode ? 'selected' : ''}>${mode}</option>`).join('')}</select></label>` : ''}
      </div>
      ${renderPlannedQuantities()}
      ${renderJobHeaderFields()}
    `
  );
}

function renderSelectionStep() {
  return sectionCard(
    'Select Measurement Sections',
    'Step 2',
    `
      <div class="selection-grid">
        ${state.catalog
          .filter((test) => !test.retired || state.draft.tests[test.id])
          .map((test) => {
            const active = state.draft.tests[test.id];
            return `
              <button
                class="selection-tile ${active ? 'selection-tile-active' : ''}"
                type="button"
                data-action="toggle-test"
                data-test="${test.id}"
                aria-pressed="${active ? 'true' : 'false'}"
              >
                <div>
                  ${test.sheetCode && test.sheetCode !== 'legacy' ? `<span class="soft-chip">${escapeHtml(test.sheetCode)}</span>` : ''}
                  <strong>${escapeHtml(test.label)}</strong>
                  <p>${escapeHtml(test.description)}</p>
                </div>
                <span class="toggle-mark">${active ? 'Selected' : 'Select'}</span>
              </button>
            `;
          })
          .join('')}
      </div>
      <div class="selection-support-grid">
        <article class="glass-note-card">
          <p class="section-kicker">Report Scope</p>
          <h4>${escapeHtml(getDraftReportTitle())}</h4>
          <p>The final PDF title follows the selected scope automatically. One test keeps its own name; multiple tests use Earthing System Health Assessment.</p>
        </article>
        <article class="glass-note-card">
          <p class="section-kicker">Testing Equipment</p>
          <h4>Instruments are recorded on each sheet</h4>
          <p>Tick the instrument, serial number and calibration due date in the header of each measurement sheet, as on the printed sheets.</p>
        </article>
      </div>
    `
  );
}

function actionButtonsCell(section, index, row, direction, extraAttributes = '') {
  const hasObservation = rowHasObservationData(row);
  return `
    <td class="table-action-cell">
      <div class="table-action-buttons">
        <button
          class="icon-button icon-button-danger"
          type="button"
          data-action="remove-row"
          data-section="${section}"
          data-index="${index}"
          ${direction ? `data-direction="${direction}"` : ''}
          ${extraAttributes}
          aria-label="Remove row"
          title="Remove row"
        >
          <span class="icon-button-symbol">&times;</span>
        </button>
        <button
          class="icon-button icon-button-observation ${hasObservation ? 'icon-button-observation-active' : ''}"
          type="button"
          data-action="open-row-observation"
          data-section="${section}"
          data-index="${index}"
          ${direction ? `data-direction="${direction}"` : ''}
          ${extraAttributes}
          aria-label="${hasObservation ? 'Edit row observation' : 'Add row observation'}"
          title="${hasObservation ? 'Edit row observation' : 'Add row observation'}"
        >
          <span class="icon-button-symbol">+</span>
        </button>
      </div>
    </td>
  `;
}

function renderObservationCard(row, title = 'Row Observation') {
  const text = safeText(row?.rowObservation, '');
  const photos = cloneRowPhotos(row?.rowPhotos).filter((photo) => photo.dataUrl);
  if (!text && !photos.length) {
    return '';
  }

  return `
    <article class="row-observation-card">
      <div class="row-observation-head">
        <strong>${escapeHtml(title)}</strong>
        ${row?.rowId ? `<span class="row-observation-badge">${escapeHtml(row.rowId)}</span>` : ''}
      </div>
      ${text ? `<p class="row-observation-copy">${escapeHtml(text)}</p>` : ''}
      ${
        photos.length
          ? `<div class="row-observation-gallery">
              ${photos
                .map(
                  (photo, photoIndex) => `
                    <figure class="row-observation-photo">
                      <img src="${escapeHtml(photo.dataUrl)}" alt="${escapeHtml(photo.name || `Observation photo ${photoIndex + 1}`)}" />
                    </figure>
                  `
                )
                .join('')}
            </div>`
          : ''
      }
    </article>
  `;
}

function renderObservationDrawer() {
  if (state.view !== 'builder' || !state.observationEditor) {
    return '';
  }

  const { section, index, direction, remark, photos, rowId, groupIndex, locationIndex } = state.observationEditor;
  const title = getRowObservationTitle(section, index, direction, groupIndex, locationIndex);

  return `
    <div class="observation-drawer-layer">
      <button class="observation-drawer-backdrop" type="button" data-action="close-row-observation" aria-label="Close observation panel"></button>
      <aside class="observation-drawer" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <div class="observation-drawer-head">
          <div>
            <p class="section-kicker">Row Observation</p>
            <h3>${escapeHtml(title)}</h3>
            <p class="muted">Attach optional field remarks and photo evidence only for this measurement row.</p>
          </div>
          <button class="icon-button" type="button" data-action="close-row-observation" aria-label="Close observation panel">
            <span class="icon-button-symbol">&times;</span>
          </button>
        </div>

        <div class="observation-meta-row">
          <span class="row-observation-badge">${escapeHtml(rowId)}</span>
          <span class="muted">Mobile friendly: tap the upload area to open camera or gallery.</span>
        </div>

        <div class="observation-drawer-body">
          <label class="field">
            <span>Observation / Remark</span>
            <textarea class="text-area" rows="5" placeholder="Add optional row-specific observation..." data-observation-field="remark">${escapeHtml(
              remark
            )}</textarea>
          </label>

          <div class="field">
            <span>Photos / Evidence</span>
            <input id="rowObservationFiles" class="observation-upload-input" type="file" accept="image/*" multiple />
            <label class="observation-upload-tile" for="rowObservationFiles">
              <strong>${state.observationUploading ? 'Preparing images...' : 'Tap to add photos'}</strong>
              <p>${state.observationUploading ? 'Optimizing images for report storage...' : 'Camera or gallery works on mobile devices.'}</p>
            </label>
            ${
              photos.length
                ? `<div class="observation-gallery">
                    ${photos
                      .map(
                        (photo, photoIndex) => `
                          <figure class="observation-thumb">
                            <img src="${escapeHtml(photo.dataUrl)}" alt="${escapeHtml(photo.name || `Observation photo ${photoIndex + 1}`)}" />
                            <button
                              class="observation-remove-photo"
                              type="button"
                              data-action="remove-observation-photo"
                              data-photo-index="${photoIndex}"
                              aria-label="Remove photo"
                            >
                              &times;
                            </button>
                          </figure>
                        `
                      )
                      .join('')}
                  </div>`
                : '<p class="observation-empty">No photos added for this row yet.</p>'
            }
          </div>
        </div>

        <div class="wizard-nav observation-drawer-actions">
          <button class="button button-secondary" type="button" data-action="close-row-observation">Cancel</button>
          <button class="button button-primary" type="button" data-action="save-row-observation" ${state.observationUploading ? 'disabled' : ''}>
            Save Observation
          </button>
        </div>
      </aside>
    </div>
  `;
}

function autoTextCell(text) {
  return `<div class="auto-eval-text">${escapeHtml(text)}</div>`;
}

function standardCell(reference, limitLabel) {
  return `
    <div class="standard-cell">
      <strong>${escapeHtml(limitLabel)}</strong>
      <span>${escapeHtml(reference)}</span>
    </div>
  `;
}

function tableSurface(title, subtitle, header, body, footer = '', tableOptions = {}) {
  return `
    <div class="mini-surface">
      <div class="mini-surface-head">
        <div>
          <h4>${escapeHtml(title)}</h4>
          <p>${escapeHtml(subtitle)}</p>
        </div>
      </div>
      <div class="table-wrap ${tableOptions.className ? `sheet-table-wrap ${tableOptions.className.includes('sheet-dense') ? 'sheet-dense-wrap' : ''}` : ''}">
        <table class="data-table ${tableOptions.className || ''}" ${tableOptions.tableStyle ? `style="${tableOptions.tableStyle}"` : ''}>
          ${tableOptions.colgroup || ''}
          <thead>${header}</thead>
          <tbody>${body}</tbody>
        </table>
      </div>
      ${footer}
    </div>
  `;
}

function soilOcrPreviewLocationTable(location) {
  const maxRows = Math.max(location.direction1.length, location.direction2.length, 1);
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Spacing of Probes (m)</th>
            <th>Direction 1 Resistivity (ohm-m)</th>
            <th>Direction 2 Resistivity (ohm-m)</th>
          </tr>
        </thead>
        <tbody>
          ${Array.from({ length: maxRows }, (_, index) => {
            const direction1 = location.direction1[index] || {};
            const direction2 = location.direction2[index] || {};
            return `
              <tr>
                <td>${escapeHtml(direction1.spacing || direction2.spacing || '-')}</td>
                <td>${escapeHtml(direction1.resistivity || '-')}</td>
                <td>${escapeHtml(direction2.resistivity || '-')}</td>
              </tr>
            `;
          }).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderOcrWarningsAndErrors(sheetState) {
  return `
    ${
      sheetState.error
        ? `<p class="detail-note detail-note-critical">${escapeHtml(sheetState.error)}</p>`
        : ''
    }
    ${
      sheetState.warnings.length
        ? `<div class="soil-ocr-notes"><strong>Warnings</strong><ul>${sheetState.warnings
            .map((warning) => `<li>${escapeHtml(warning)}</li>`)
            .join('')}</ul></div>`
        : ''
    }
    ${
      sheetState.uncertainFields.length
        ? `<div class="soil-ocr-notes"><strong>Review Carefully</strong><ul>${sheetState.uncertainFields
            .map((item) => `<li><code>${escapeHtml(item.path)}</code> - ${escapeHtml(item.reason || 'Low OCR confidence')}</li>`)
            .join('')}</ul></div>`
        : ''
    }
  `;
}

function renderSoilOcrAssist() {
  const sheetId = 'soilResistivity';
  const sheetState = getOcrSheetState(sheetId);
  const preview = sheetState.preview;
  const selectedFileName = sheetState.selectedFile ? sheetState.selectedFile.name : 'No sheet selected yet.';
  const previewLocations = Array.isArray(preview?.locations) ? preview.locations : [];

  return `
    <section class="mini-surface soil-ocr-assist">
      <div class="mini-surface-head mini-surface-head-location">
        <div>
          <h4>Soil Sheet Autofill</h4>
          <p>Choose whether to record manually or import a physical soil resistivity sheet for OCR preview.</p>
        </div>
        <div class="mini-surface-head-actions">
          <button class="soft-chip soft-chip-button${sheetState.mode === 'manual' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="manual">Fill Manually</button>
          <button class="soft-chip soft-chip-button${sheetState.mode === 'upload' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="upload">Autofill From Upload</button>
        </div>
      </div>
      ${
        sheetState.mode === 'upload'
          ? `
            <div class="soil-ocr-panel">
              <div class="soil-ocr-upload">
                <input id="soilOcrFileInput" class="observation-upload-input" type="file" accept="image/*,application/pdf" />
                <label class="observation-upload-tile soil-ocr-upload-tile" for="soilOcrFileInput">
                  <strong>${sheetState.scanning ? 'Scanning sheet...' : 'Tap to choose soil sheet image or PDF'}</strong>
                  <p>${escapeHtml(selectedFileName)}</p>
                  <span class="muted">${
                    state.ocr.configured
                      ? `Gemini OCR ready${state.ocr.model ? ` (${escapeHtml(state.ocr.model)})` : ''}.`
                      : 'Gemini OCR is not configured yet.'
                  }</span>
                </label>
                <div class="soil-ocr-actions">
                  <button class="button button-primary" type="button" data-action="scan-ocr" data-sheet="${sheetId}" ${
                    !state.ocr.configured || !sheetState.selectedFile || sheetState.scanning ? 'disabled' : ''
                  }>${sheetState.scanning ? 'Scanning...' : 'Preview Extracted Data'}</button>
                  <button class="button button-ghost" type="button" data-action="clear-ocr" data-sheet="${sheetId}" ${
                    !sheetState.selectedFile && !sheetState.preview ? 'disabled' : ''
                  }>Clear</button>
                </div>
              </div>
              ${renderOcrWarningsAndErrors(sheetState)}
              ${
                previewLocations.length
                  ? `
                    <div class="soil-ocr-preview">
                      <div class="detail-subsection-head">
                        <h4>Preview Before Applying</h4>
                        <div class="button-row">
                          <button class="button button-secondary" type="button" data-action="apply-ocr-preview" data-sheet="${sheetId}">Apply To Soil Sheet</button>
                        </div>
                      </div>
                      <div class="stack-grid">
                        ${previewLocations
                          .map(
                            (location, index) => `
                              <section class="mini-surface mini-surface-location">
                                <div class="mini-surface-head mini-surface-head-location">
                                  <div>
                                    <h4>${escapeHtml(location.name || `Location ${index + 1}`)}</h4>
                                    <p>Preview extracted values before writing into the Soil Resistivity measurement sheet.</p>
                                  </div>
                                </div>
                                <div class="field-grid">
                                  <div class="field"><span>Driven Electrode Diameter (mm)</span><input class="table-input" value="${escapeHtml(
                                    safeText(location.drivenElectrodeDiameter, '')
                                  )}" disabled /></div>
                                  <div class="field"><span>Driven Electrode Length (m)</span><input class="table-input" value="${escapeHtml(
                                    safeText(location.drivenElectrodeLength, '')
                                  )}" disabled /></div>
                                </div>
                                ${soilOcrPreviewLocationTable(location)}
                                ${
                                  safeText(location.notes, '')
                                    ? `<p class="detail-note"><strong>Notes:</strong> ${escapeHtml(location.notes)}</p>`
                                    : ''
                                }
                              </section>
                            `
                          )
                          .join('')}
                      </div>
                    </div>
                  `
                  : ''
              }
            </div>
          `
          : '<p class="detail-note">Manual data entry remains active. Switch to upload mode anytime if the engineer wants Gemini to preview a physical soil sheet.</p>'
      }
    </section>
  `;
}

function electrodeOcrPreviewTable(rows) {
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Pit Tag</th>
            <th>Location</th>
            <th>Electrode Type</th>
            <th>Material</th>
            <th>Length (m)</th>
            <th>Dia (mm)</th>
            <th>Resistance Without Grid</th>
            <th>Resistance With Grid</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row) => `
                <tr>
                  <td>${escapeHtml(row.tag || '-')}</td>
                  <td>${escapeHtml(row.location || '-')}</td>
                  <td>${escapeHtml(row.electrodeType || '-')}</td>
                  <td>${escapeHtml(row.materialType || '-')}</td>
                  <td>${escapeHtml(row.length || '-')}</td>
                  <td>${escapeHtml(row.diameter || '-')}</td>
                  <td>${escapeHtml(row.resistanceWithoutGrid || '-')}</td>
                  <td>${escapeHtml(row.resistanceWithGrid || '-')}</td>
                </tr>
              `
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderElectrodeOcrAssist() {
  const sheetId = 'electrodeResistance';
  const sheetState = getOcrSheetState(sheetId);
  const previewRows = Array.isArray(sheetState.preview?.rows) ? sheetState.preview.rows : [];
  const selectedFileName = sheetState.selectedFile ? sheetState.selectedFile.name : 'No sheet selected yet.';

  return `
    <section class="mini-surface soil-ocr-assist">
      <div class="mini-surface-head mini-surface-head-location">
        <div>
          <h4>Electrode Sheet Autofill</h4>
          <p>Choose whether to record manually or import a physical earth electrode sheet for OCR preview.</p>
        </div>
        <div class="mini-surface-head-actions">
          <button class="soft-chip soft-chip-button${sheetState.mode === 'manual' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="manual">Fill Manually</button>
          <button class="soft-chip soft-chip-button${sheetState.mode === 'upload' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="upload">Autofill From Upload</button>
        </div>
      </div>
      ${
        sheetState.mode === 'upload'
          ? `
            <div class="soil-ocr-panel">
              <div class="soil-ocr-upload">
                <input id="electrodeOcrFileInput" class="observation-upload-input" type="file" accept="image/*,application/pdf" />
                <label class="observation-upload-tile soil-ocr-upload-tile" for="electrodeOcrFileInput">
                  <strong>${sheetState.scanning ? 'Scanning sheet...' : 'Tap to choose electrode sheet image or PDF'}</strong>
                  <p>${escapeHtml(selectedFileName)}</p>
                  <span class="muted">${
                    state.ocr.configured
                      ? `Gemini OCR ready${state.ocr.model ? ` (${escapeHtml(state.ocr.model)})` : ''}.`
                      : 'Gemini OCR is not configured yet.'
                  }</span>
                </label>
                <div class="soil-ocr-actions">
                  <button class="button button-primary" type="button" data-action="scan-ocr" data-sheet="${sheetId}" ${
                    !state.ocr.configured || !sheetState.selectedFile || sheetState.scanning ? 'disabled' : ''
                  }>${sheetState.scanning ? 'Scanning...' : 'Preview Extracted Data'}</button>
                  <button class="button button-ghost" type="button" data-action="clear-ocr" data-sheet="${sheetId}" ${
                    !sheetState.selectedFile && !sheetState.preview ? 'disabled' : ''
                  }>Clear</button>
                </div>
              </div>
              ${renderOcrWarningsAndErrors(sheetState)}
              ${
                previewRows.length
                  ? `
                    <div class="soil-ocr-preview">
                      <div class="detail-subsection-head">
                        <h4>Preview Before Applying</h4>
                        <div class="button-row">
                          <button class="button button-secondary" type="button" data-action="apply-ocr-preview" data-sheet="${sheetId}">Apply To Electrode Sheet</button>
                        </div>
                      </div>
                      ${electrodeOcrPreviewTable(previewRows)}
                    </div>
                  `
                  : ''
              }
            </div>
          `
          : '<p class="detail-note">Manual data entry remains active. Switch to upload mode anytime if the engineer wants Gemini to preview a physical electrode sheet.</p>'
      }
    </section>
  `;
}

function continuityOcrPreviewTable(rows) {
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Sr. No.</th>
            <th>Main Location</th>
            <th>Measurement Point</th>
            <th>Resistance (ohm)</th>
            <th>Impedance (ohm)</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row) => `
                <tr>
                  <td>${escapeHtml(row.srNo || '-')}</td>
                  <td>${escapeHtml(row.mainLocation || '-')}</td>
                  <td>${escapeHtml(row.measurementPoint || '-')}</td>
                  <td>${escapeHtml(row.resistance || '-')}</td>
                  <td>${escapeHtml(row.impedance || '-')}</td>
                </tr>
              `
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderContinuityOcrAssist() {
  const sheetId = 'continuityTest';
  const sheetState = getOcrSheetState(sheetId);
  const previewRows = Array.isArray(sheetState.preview?.rows) ? sheetState.preview.rows : [];
  const selectedFileName = sheetState.selectedFile ? sheetState.selectedFile.name : 'No sheet selected yet.';

  return `
    <section class="mini-surface soil-ocr-assist">
      <div class="mini-surface-head mini-surface-head-location">
        <div>
          <h4>Continuity Sheet Autofill</h4>
          <p>Choose whether to record manually or import a physical continuity sheet for OCR preview.</p>
        </div>
        <div class="mini-surface-head-actions">
          <button class="soft-chip soft-chip-button${sheetState.mode === 'manual' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="manual">Fill Manually</button>
          <button class="soft-chip soft-chip-button${sheetState.mode === 'upload' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="upload">Autofill From Upload</button>
        </div>
      </div>
      ${
        sheetState.mode === 'upload'
          ? `
            <div class="soil-ocr-panel">
              <div class="soil-ocr-upload">
                <input id="continuityOcrFileInput" class="observation-upload-input" type="file" accept="image/*,application/pdf" />
                <label class="observation-upload-tile soil-ocr-upload-tile" for="continuityOcrFileInput">
                  <strong>${sheetState.scanning ? 'Scanning sheet...' : 'Tap to choose continuity sheet image or PDF'}</strong>
                  <p>${escapeHtml(selectedFileName)}</p>
                  <span class="muted">${
                    state.ocr.configured
                      ? `Gemini OCR ready${state.ocr.model ? ` (${escapeHtml(state.ocr.model)})` : ''}.`
                      : 'Gemini OCR is not configured yet.'
                  }</span>
                </label>
                <div class="soil-ocr-actions">
                  <button class="button button-primary" type="button" data-action="scan-ocr" data-sheet="${sheetId}" ${
                    !state.ocr.configured || !sheetState.selectedFile || sheetState.scanning ? 'disabled' : ''
                  }>${sheetState.scanning ? 'Scanning...' : 'Preview Extracted Data'}</button>
                  <button class="button button-ghost" type="button" data-action="clear-ocr" data-sheet="${sheetId}" ${
                    !sheetState.selectedFile && !sheetState.preview ? 'disabled' : ''
                  }>Clear</button>
                </div>
              </div>
              ${renderOcrWarningsAndErrors(sheetState)}
              ${
                previewRows.length
                  ? `
                    <div class="soil-ocr-preview">
                      <div class="detail-subsection-head">
                        <h4>Preview Before Applying</h4>
                        <div class="button-row">
                          <button class="button button-secondary" type="button" data-action="apply-ocr-preview" data-sheet="${sheetId}">Apply To Continuity Sheet</button>
                        </div>
                      </div>
                      ${continuityOcrPreviewTable(previewRows)}
                    </div>
                  `
                  : ''
              }
            </div>
          `
          : '<p class="detail-note">Manual data entry remains active. Switch to upload mode anytime if the engineer wants Gemini to preview a physical continuity sheet.</p>'
      }
    </section>
  `;
}

function renderSharedOcrAssist({ sheetId, title, description, inputId, chooseLabel, previewHtml, applyLabel }) {
  const sheetState = getOcrSheetState(sheetId);
  const selectedFileName = sheetState.selectedFile ? sheetState.selectedFile.name : 'No sheet selected yet.';

  return `
    <section class="mini-surface soil-ocr-assist">
      <div class="mini-surface-head mini-surface-head-location">
        <div>
          <h4>${escapeHtml(title)}</h4>
          <p>${escapeHtml(description)}</p>
        </div>
        <div class="mini-surface-head-actions">
          <button class="soft-chip soft-chip-button${sheetState.mode === 'manual' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="manual">Fill Manually</button>
          <button class="soft-chip soft-chip-button${sheetState.mode === 'upload' ? ' soft-chip-button-active' : ''}" type="button" data-action="set-ocr-entry-mode" data-sheet="${sheetId}" data-mode="upload">Autofill From Upload</button>
        </div>
      </div>
      ${
        sheetState.mode === 'upload'
          ? `
            <div class="soil-ocr-panel">
              <div class="soil-ocr-upload">
                <input id="${escapeHtml(inputId)}" class="observation-upload-input" type="file" accept="image/*,application/pdf" />
                <label class="observation-upload-tile soil-ocr-upload-tile" for="${escapeHtml(inputId)}">
                  <strong>${sheetState.scanning ? 'Scanning sheet...' : escapeHtml(chooseLabel)}</strong>
                  <p>${escapeHtml(selectedFileName)}</p>
                  <span class="muted">${
                    state.ocr.configured
                      ? `Gemini OCR ready${state.ocr.model ? ` (${escapeHtml(state.ocr.model)})` : ''}.`
                      : 'Gemini OCR is not configured yet.'
                  }</span>
                </label>
                <div class="soil-ocr-actions">
                  <button class="button button-primary" type="button" data-action="scan-ocr" data-sheet="${sheetId}" ${
                    !state.ocr.configured || !sheetState.selectedFile || sheetState.scanning ? 'disabled' : ''
                  }>${sheetState.scanning ? 'Scanning...' : 'Preview Extracted Data'}</button>
                  <button class="button button-ghost" type="button" data-action="clear-ocr" data-sheet="${sheetId}" ${
                    !sheetState.selectedFile && !sheetState.preview ? 'disabled' : ''
                  }>Clear</button>
                </div>
              </div>
              ${renderOcrWarningsAndErrors(sheetState)}
              ${
                previewHtml
                  ? `
                    <div class="soil-ocr-preview">
                      <div class="detail-subsection-head">
                        <h4>Preview Before Applying</h4>
                        <div class="button-row">
                          <button class="button button-secondary" type="button" data-action="apply-ocr-preview" data-sheet="${sheetId}">${escapeHtml(applyLabel)}</button>
                        </div>
                      </div>
                      ${previewHtml}
                    </div>
                  `
                  : ''
              }
            </div>
          `
          : '<p class="detail-note">Manual data entry remains active. Switch to upload mode anytime if the engineer wants Gemini to preview a physical sheet.</p>'
      }
    </section>
  `;
}

function phaseOcrPreviewTable(groups, includeFaultCurrent = false) {
  const rows = [];
  (Array.isArray(groups) ? groups : []).forEach((group) => {
    (Array.isArray(group.rows) ? group.rows : []).forEach((row) => {
      rows.push({
        location: group.location,
        feederTag: group.feederTag,
        deviceType: group.deviceType,
        deviceRating: group.deviceRating,
        breakingCapacity: group.breakingCapacity,
        measuredPoints: row.measuredPoints,
        loopImpedance: row.loopImpedance,
        prospectiveFaultCurrent: row.prospectiveFaultCurrent,
        voltage: row.voltage
      });
    });
  });

  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Location of Panel / Equipment</th>
            <th>Name of Feeder & Tag No.</th>
            <th>Type</th>
            <th>Rating (A)</th>
            <th>Breaking (kA)</th>
            <th>Measured Points</th>
            <th>Loop Impedance (Z)</th>
            ${includeFaultCurrent ? '<th>Prospective Fault Current</th>' : ''}
            <th>Voltage</th>
          </tr>
        </thead>
        <tbody>
          ${rows
            .map(
              (row) => `
                <tr>
                  <td>${escapeHtml(row.location || '-')}</td>
                  <td>${escapeHtml(row.feederTag || '-')}</td>
                  <td>${escapeHtml(row.deviceType || '-')}</td>
                  <td>${escapeHtml(row.deviceRating || '-')}</td>
                  <td>${escapeHtml(row.breakingCapacity || '-')}</td>
                  <td>${escapeHtml(row.measuredPoints || '-')}</td>
                  <td>${escapeHtml(row.loopImpedance || '-')}</td>
                  ${includeFaultCurrent ? `<td>${escapeHtml(row.prospectiveFaultCurrent || '-')}</td>` : ''}
                  <td>${escapeHtml(row.voltage || '-')}</td>
                </tr>
              `
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function riserOcrPreviewTable(rows) {
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Sr. No.</th>
            <th>Main Location</th>
            <th>Measurement Point</th>
            <th>Towards Equipment</th>
            <th>Towards Grid</th>
          </tr>
        </thead>
        <tbody>
          ${(Array.isArray(rows) ? rows : [])
            .map(
              (row) => `
                <tr>
                  <td>${escapeHtml(row.srNo || '-')}</td>
                  <td>${escapeHtml(row.mainLocation || '-')}</td>
                  <td>${escapeHtml(row.measurementPoint || '-')}</td>
                  <td>${escapeHtml(row.resistanceTowardsEquipment || '-')}</td>
                  <td>${escapeHtml(row.resistanceTowardsGrid || '-')}</td>
                </tr>
              `
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function earthContinuityOcrPreviewTable(rows) {
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Sr. No.</th>
            <th>Tag</th>
            <th>Location / Building</th>
            <th>Distance</th>
            <th>Measured Value</th>
          </tr>
        </thead>
        <tbody>
          ${(Array.isArray(rows) ? rows : [])
            .map(
              (row) => `
                <tr>
                  <td>${escapeHtml(row.srNo || '-')}</td>
                  <td>${escapeHtml(row.tag || '-')}</td>
                  <td>${escapeHtml(row.locationBuildingName || '-')}</td>
                  <td>${escapeHtml(row.distance || '-')}</td>
                  <td>${escapeHtml(row.measuredValue || '-')}</td>
                </tr>
              `
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function towerOcrPreviewTable(groups) {
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Sr. No.</th>
            <th>Main Location – Tower</th>
            <th>Measurement Point Location</th>
            <th>Foot to Earthing Connection Status</th>
            <th>Measured Current I (mA)</th>
            <th>Measured Impedance (ohm)</th>
          </tr>
        </thead>
        <tbody>
          ${(Array.isArray(groups) ? groups : [])
            .map((group) =>
              (Array.isArray(group.readings) ? group.readings : [])
                .map(
                  (reading) => `
                    <tr>
                      <td>${escapeHtml(group.srNo || '-')}</td>
                      <td>${escapeHtml(group.mainLocationTower || '-')}</td>
                      <td>${escapeHtml(reading.measurementPointLocation || '-')}</td>
                      <td>${escapeHtml(reading.footToEarthingConnectionStatus || '-')}</td>
                      <td>${escapeHtml(reading.measuredCurrentMa || '-')}</td>
                      <td>${escapeHtml(reading.measuredImpedance || '-')}</td>
                    </tr>
                  `
                )
                .join('')
            )
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderLoopOcrAssist() {
  const previewGroups = Array.isArray(getOcrSheetState('loopImpedanceTest').preview?.groups)
    ? getOcrSheetState('loopImpedanceTest').preview.groups
    : [];
  return renderSharedOcrAssist({
    sheetId: 'loopImpedanceTest',
    title: 'Loop Sheet Autofill',
    description: 'Choose whether to record manually or import a physical loop impedance sheet for OCR preview.',
    inputId: 'loopOcrFileInput',
    chooseLabel: 'Tap to choose loop impedance sheet image or PDF',
    previewHtml: previewGroups.length ? phaseOcrPreviewTable(previewGroups, false) : '',
    applyLabel: 'Apply To Loop Sheet'
  });
}

function renderFaultOcrAssist() {
  const previewGroups = Array.isArray(getOcrSheetState('prospectiveFaultCurrent').preview?.groups)
    ? getOcrSheetState('prospectiveFaultCurrent').preview.groups
    : [];
  return renderSharedOcrAssist({
    sheetId: 'prospectiveFaultCurrent',
    title: 'PFC Sheet Autofill',
    description: 'Choose whether to record manually or import a physical prospective fault current sheet for OCR preview.',
    inputId: 'faultOcrFileInput',
    chooseLabel: 'Tap to choose PFC sheet image or PDF',
    previewHtml: previewGroups.length ? phaseOcrPreviewTable(previewGroups, true) : '',
    applyLabel: 'Apply To PFC Sheet'
  });
}

function renderRiserOcrAssist() {
  const previewRows = Array.isArray(getOcrSheetState('riserIntegrityTest').preview?.rows)
    ? getOcrSheetState('riserIntegrityTest').preview.rows
    : [];
  return renderSharedOcrAssist({
    sheetId: 'riserIntegrityTest',
    title: 'Riser Sheet Autofill',
    description: 'Choose whether to record manually or import a physical riser / grid integrity sheet for OCR preview.',
    inputId: 'riserOcrFileInput',
    chooseLabel: 'Tap to choose riser integrity sheet image or PDF',
    previewHtml: previewRows.length ? riserOcrPreviewTable(previewRows) : '',
    applyLabel: 'Apply To Riser Sheet'
  });
}

function renderEarthContinuityOcrAssist() {
  const previewRows = Array.isArray(getOcrSheetState('earthContinuityTest').preview?.rows)
    ? getOcrSheetState('earthContinuityTest').preview.rows
    : [];
  return renderSharedOcrAssist({
    sheetId: 'earthContinuityTest',
    title: 'Earth Continuity Sheet Autofill',
    description: 'Choose whether to record manually or import a physical earth continuity sheet for OCR preview.',
    inputId: 'earthContinuityOcrFileInput',
    chooseLabel: 'Tap to choose earth continuity sheet image or PDF',
    previewHtml: previewRows.length ? earthContinuityOcrPreviewTable(previewRows) : '',
    applyLabel: 'Apply To Earth Continuity Sheet'
  });
}

function renderTowerOcrAssist() {
  const previewGroups = Array.isArray(getOcrSheetState('towerFootingResistance').preview?.groups)
    ? getOcrSheetState('towerFootingResistance').preview.groups
    : [];
  return renderSharedOcrAssist({
    sheetId: 'towerFootingResistance',
    title: 'Tower Footing Sheet Autofill',
    description: 'Choose whether to record manually or import a physical tower footing sheet for OCR preview.',
    inputId: 'towerFootingOcrFileInput',
    chooseLabel: 'Tap to choose tower footing sheet image or PDF',
    previewHtml: previewGroups.length ? towerOcrPreviewTable(previewGroups) : '',
    applyLabel: 'Apply To Tower Footing Sheet'
  });
}

function renderBuilderStep() {
  const step = currentStepId();
  if (step === 'project') {
    return renderProjectStep();
  }
  if (step === 'selection') {
    return renderSelectionStep();
  }
  return renderSheetStep(step);
}

function builderView() {
  const steps = getSteps();
  const assessment = assessmentSummary(state.draft);
  const readiness = builderReadiness(assessment);
  const currentStep = currentStepId();
  return `
    <section class="surface builder-shell">
      <div class="builder-header">
        <div>
          <p class="section-kicker">Report Builder</p>
          <h2>${state.editing ? `Site day ${escapeHtml(formatDisplayDate(state.editing.workDate) || state.editing.workDate)}` : 'New ElectroReports Project'}</h2>
          <p>Only the sections you select will appear in the report flow and in the final PDF.</p>
        </div>
        <div class="button-row">
          ${state.editing ? '<button class="button button-secondary" data-action="workspace-back">Back to project</button>' : '<button class="button button-secondary" data-action="go-dashboard">Cancel</button>'}
          <button class="button button-primary" data-action="save-report" ${state.saving ? 'disabled' : ''}>${state.saving ? 'Saving...' : state.editing ? 'Save site day' : 'Save Report'}</button>
        </div>
        ${renderFloatingShapes(18, 'hero-shapes hero-shapes-subtle')}
      </div>

      <div class="stepper">
        ${steps
          .map((step, index) => {
            const active = index === state.stepIndex;
            const complete = index < state.stepIndex;
            return `
              <button
                class="step-pill ${active ? 'step-pill-active' : ''} ${complete ? 'step-pill-complete' : ''}"
                data-action="go-step"
                data-step-index="${index}"
                type="button"
              >
                <span>${index + 1}</span>
                <strong>${escapeHtml(step.label)}</strong>
                ${typeof stepStatus === 'function' ? (() => { const status = stepStatus(step.id); return `<i class="step-dot step-dot-${status}" role="img" aria-label="${STEP_STATUS_LABEL[status]}" title="${STEP_STATUS_LABEL[status]}"></i>`; })() : ''}
              </button>
            `;
          })
          .join('')}
      </div>

      <div class="builder-layout ${isSheetStep(currentStep) ? 'builder-layout-sheet' : 'builder-layout-form'}">
        <div class="builder-main">
          ${renderBuilderStep()}
        </div>
        <aside class="builder-aside">
          <section class="mini-surface aside-panel">
            <p class="section-kicker">Live Readiness</p>
            <h4>${assessment.label}</h4>
            <p>${stepGuidance(currentStep)} ${currentStepNote(currentStep)}</p>
            ${renderProgressBar('Project completeness', readiness.projectScore, readiness.projectScore === 100 ? 'healthy' : readiness.projectScore > 60 ? 'warning' : 'critical')}
            ${renderProgressBar('Measurement coverage', readiness.measurementScore, readiness.measurementScore > 84 ? 'healthy' : readiness.measurementScore > 40 ? 'warning' : 'critical')}
            ${renderProgressBar('Overall readiness', readiness.score, readiness.score > 84 ? 'healthy' : readiness.score > 50 ? 'warning' : 'critical')}
          </section>

          <section class="mini-surface aside-panel">
            <p class="section-kicker">Live Assessment</p>
            ${renderAssessmentHighlights(assessment)}
            <p class="aside-compact-note">Current step: <strong>${escapeHtml(steps[state.stepIndex].label)}</strong></p>
          </section>
        </aside>
      </div>

      <div class="builder-followup">
        <section class="mini-surface aside-panel aside-panel-wide">
          <p class="section-kicker">Scope & Standards</p>
          ${renderScopeStandardsPanel(state.draft)}
        </section>

        <section class="mini-surface aside-panel aside-panel-wide">
          <p class="section-kicker">Module Health</p>
          ${renderModuleHealthList(assessment)}
        </section>
      </div>

      <div class="wizard-nav">
        <button class="button button-ghost" data-action="step-prev" ${state.stepIndex === 0 ? 'disabled' : ''}>Back</button>
        ${
          state.stepIndex < steps.length - 1
            ? '<button class="button button-primary" data-action="step-next">Next Step</button>'
            : `<button class="button button-success" data-action="save-report" ${state.saving ? 'disabled' : ''}>${state.saving ? 'Saving...' : 'Save & Finish'}</button>`
        }
      </div>
    </section>
  `;
}

function reportSectionBlock(title, content) {
  return `
    <section class="surface detail-block">
      <div class="section-head">
        <h3>${escapeHtml(title)}</h3>
      </div>
      ${content}
    </section>
  `;
}

function simpleDetailTable(columns, rows, options = {}) {
  const observationRenderer = options.observationRenderer || null;
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join('')}</tr>
        </thead>
        <tbody>
          ${rows
            .map((row, index) => {
              const observationHtml = observationRenderer ? observationRenderer(row, index) : renderObservationCard(row);
              return `
                <tr>${columns.map((column) => `<td>${column.render ? column.render(row) : escapeHtml(row[column.key])}</td>`).join('')}</tr>
                ${
                  observationHtml
                    ? `<tr class="detail-table-observation"><td colspan="${columns.length}"><div class="row-observation-stack">${observationHtml}</div></td></tr>`
                    : ''
                }
              `;
            })
            .join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderAiNarrativeSection(report) {
  const narrative = report?.aiNarrative;
  if (!narrative) {
    return reportSectionBlock(
      'AI Report Assist',
      `
        <div class="detail-grid">
          <div><span>AI Status</span><strong>${state.ai.configured ? 'Ready to generate' : 'Not configured'}</strong></div>
          <div><span>Model</span><strong>${escapeHtml(state.ai.model || '-')}</strong></div>
          <div><span>Reference PDFs</span><strong>${escapeHtml(state.ai.referenceDocuments.length ? state.ai.referenceDocuments.join(', ') : '-')}</strong></div>
        </div>
        <p class="detail-note">${
          state.ai.configured
            ? 'Generate a grounded executive summary, key findings, and module summaries from the saved report.'
            : 'Add an OpenAI API key in the ElectroReports .env file to enable grounded report narrative generation.'
        }</p>
      `
    );
  }

  const findings = Array.isArray(narrative.keyFindings) ? narrative.keyFindings : [];
  const moduleSummaries = Array.isArray(narrative.moduleSummaries) ? narrative.moduleSummaries : [];

  return reportSectionBlock(
    'AI Report Assist',
    `
      <div class="detail-grid">
        <div><span>Generated At</span><strong>${escapeHtml(safeText(narrative.generatedAt, '-'))}</strong></div>
        <div><span>Model</span><strong>${escapeHtml(safeText(narrative.model, '-'))}</strong></div>
        <div><span>Reference PDFs</span><strong>${escapeHtml((narrative.referenceDocuments || []).join(', ') || '-')}</strong></div>
      </div>
      <div class="detail-subsection">
        <div class="detail-subsection-head"><h4>Overall Assessment</h4></div>
        <p class="detail-note">${escapeHtml(safeText(narrative.overallAssessment, '-'))}</p>
      </div>
      <div class="detail-subsection">
        <div class="detail-subsection-head"><h4>Executive Summary</h4></div>
        <p class="detail-note">${escapeHtml(safeText(narrative.executiveSummary, '-'))}</p>
      </div>
      ${
        findings.length
          ? `
            <div class="detail-subsection">
              <div class="detail-subsection-head"><h4>Key Findings & Recommended Actions</h4></div>
              ${simpleDetailTable(
                [
                  { label: 'Test Area', key: 'testArea' },
                  { label: 'Key Finding', key: 'keyFinding' },
                  { label: 'Priority', key: 'priority' },
                  { label: 'Recommended Action', key: 'recommendedAction' },
                  { label: 'Status', key: 'status' }
                ],
                findings
              )}
            </div>
          `
          : ''
      }
      ${
        moduleSummaries.length
          ? `
            <div class="detail-subsection">
              <div class="detail-subsection-head"><h4>Module Summaries</h4></div>
              ${simpleDetailTable(
                [
                  { label: 'Measurement Sheet', key: 'label' },
                  { label: 'Summary', key: 'summary' }
                ],
                moduleSummaries
              )}
            </div>
          `
          : ''
      }
    `
  );
}

function detailView() {
  const report = state.activeReport;
  const assessment = assessmentSummary(report);
  const soil = assessment.soil;

  return `
    <section class="surface detail-shell">
      <div class="detail-hero">
        <div class="detail-hero-copy">
          <p class="section-kicker">${escapeHtml(report.project.projectNo)}</p>
          <h2>${escapeHtml(report.project.clientName)}</h2>
          <p>${escapeHtml(report.project.siteLocation)} | Engineer: ${escapeHtml(report.project.engineerName)}</p>
          <div class="detail-hero-tags">
            ${pill(assessment.tone, assessment.label)}
            ${renderSelectedModuleTags(report)}
          </div>
        </div>
        <div class="button-row detail-hero-actions">
          <button class="button button-secondary" data-action="go-dashboard">Back to Dashboard</button>
          <button class="button button-primary" data-action="new-report">New Report</button>
          ${isReviewable(report) ? `<button class="button button-primary" data-action="open-review" data-id="${report.id}">Review &amp; approval</button>` : `<button class="button button-secondary" data-action="generate-ai-summary" data-id="${report.id}" ${!state.ai.configured || state.ai.generating ? 'disabled' : ''}>${
            state.ai.generating ? 'Generating AI Summary...' : report.aiNarrative ? 'Refresh AI Summary' : 'Generate AI Summary'
          }</button>`}
          <button class="button button-success" data-action="export-pdf" data-id="${report.id}" ${state.exporting ? 'disabled' : ''}>${state.exporting ? 'Generating PDF...' : 'Export PDF'}</button>
          ${renderExportNote(report)}
        </div>
        ${renderFloatingShapes(16, 'hero-shapes hero-shapes-subtle')}
      </div>

      <div class="metric-strip">
        <article class="metric-box">
          <span>Selected Tests</span>
          <strong>${selectedTestCount(report)}</strong>
        </article>
        <article class="metric-box">
          <span>Mean Soil Resistivity</span>
          <strong>${soil.overallAverage === null ? '-' : `${soil.overallAverage} ohm-m`}</strong>
          ${pill(soil.category.tone, soil.category.label)}
        </article>
        <article class="metric-box">
          <span>Date of Testing</span>
          <strong>${escapeHtml(report.project.reportDate)}</strong>
        </article>
        <article class="metric-box">
          <span>Action Points</span>
          <strong>${assessment.actions.total}</strong>
        </article>
      </div>

      ${renderReviewChecklist(report)}

      ${reportSectionBlock(
        'Project Details',
        `
          <div class="detail-grid">
            <div><span>Project Number</span><strong>${escapeHtml(report.project.projectNo)}</strong></div>
            <div><span>Client Name</span><strong>${escapeHtml(report.project.clientName)}</strong></div>
            <div><span>Site Location</span><strong>${escapeHtml(report.project.siteLocation)}</strong></div>
            <div><span>Work Order / Ref</span><strong>${escapeHtml(report.project.workOrder)}</strong></div>
            <div><span>Work date</span><strong>${escapeHtml(report.project.reportDate)}</strong></div>
            <div><span>Testing Engineer</span><strong>${escapeHtml(report.project.engineerName)}</strong></div>
            ${report.project.clientRepresentative ? `<div><span>Client representative</span><strong>${escapeHtml(report.project.clientRepresentative)}</strong></div>` : ''}
            ${report.project.groundCondition ? `<div><span>Ground condition</span><strong>${escapeHtml(report.project.groundCondition)}</strong></div>` : ''}
            <div><span>Acceptance profile</span><strong>${escapeHtml((sourceAssessment(report).profile && sourceAssessment(report).profile.id) || report.project.acceptanceProfile || 'ETPL-DEFAULT')}</strong></div>
            ${report.project.assumedSoilResistivity ? `<div><span>Assumed soil resistivity</span><strong>${escapeHtml(report.project.assumedSoilResistivity)} ohm-m (${escapeHtml(report.project.assumedSoilResistivitySource || 'source not stated')})</strong></div>` : ''}
            ${report.project.zohoProjectName ? `<div><span>Zoho Project</span><strong>${escapeHtml(report.project.zohoProjectName)}</strong></div>` : ''}
            ${report.project.zohoProjectStage ? `<div><span>Zoho Stage</span><strong>${escapeHtml(report.project.zohoProjectStage)}</strong></div>` : ''}
          </div>
        `
      )}

      ${renderAiNarrativeSection(report)}

      ${report.tests.soilResistivity ? reportSectionBlock('MS-01 Soil Resistivity', renderSoilDetail(report)) : ''}

      ${allSheets()
        .filter((sheet) => sheet.id !== 'soilResistivity' && report.tests && report.tests[sheet.id])
        .map((sheet) => reportSectionBlock(sheet.code === 'legacy' ? sheet.title : `${sheet.code} ${sheet.title}`, renderSheetDetail(report, sheet)))
        .join('')}
    </section>
  `;
}

function toastHtml() {
  if (!state.toast) {
    return '';
  }
  return `<aside class="toast toast-${state.toast.tone}">${escapeHtml(state.toast.message)}</aside>`;
}

let motionBound = false;

function updateMotionEffects() {
  const root = document.documentElement;
  const scrollRange = Math.max(document.body.scrollHeight - window.innerHeight, 1);
  const scrollProgress = Math.max(0, Math.min(1, window.scrollY / scrollRange));
  root.style.setProperty('--scroll-progress', scrollProgress.toFixed(4));
}

function bindMotionEffects() {
  if (motionBound) {
    return;
  }
  motionBound = true;

  window.addEventListener('scroll', updateMotionEffects, { passive: true });
  window.addEventListener(
    'pointermove',
    (event) => {
      const x = event.clientX / Math.max(window.innerWidth, 1);
      const y = event.clientY / Math.max(window.innerHeight, 1);
      document.documentElement.style.setProperty('--pointer-x', x.toFixed(4));
      document.documentElement.style.setProperty('--pointer-y', y.toFixed(4));
    },
    { passive: true }
  );

  updateMotionEffects();
}

function render() {
  if (state.view === 'builder') {
    syncDraftAssessment();
  }
  app.dataset.view = state.view;
  app.innerHTML = `
    ${brandHeader()}
    <main class="app-shell">
      ${state.view === 'dashboard' ? dashboardView() : ''}
      ${state.view === 'builder' ? builderView() : ''}
      ${state.view === 'detail' && state.activeReport ? detailView() : ''}
      ${state.view === 'review' ? reviewView() : ''}
      ${state.view === 'workspace' ? workspaceView() : ''}
      ${state.view === 'schedule' && authUi.user ? scheduleView() : ''}
      ${state.view === 'signin' ? signInView() : ''}
      ${state.view === 'profile' && authUi.user ? profileView() : ''}
      ${state.view === 'admin' && authUi.user ? adminView() : ''}
      ${state.view === 'kb' && authUi.user ? kbView() : ''}
      ${state.view === 'references' && authUi.user ? referencesView() : ''}
      ${state.view === 'scans' && authUi.user ? scansView() : ''}
      ${state.view === 'scan-review' && authUi.user ? scanReviewView() : ''}
    </main>
    ${renderObservationDrawer()}
    ${renderPrintDialog()}
    ${toastHtml()}
  `;
  if (state.view === 'builder') {
    scheduleDraftAutosave();
  }
  bindMotionEffects();
  updateMotionEffects();
  if (typeof restoreTickPicker === 'function') {
    restoreTickPicker();
  }
  if (typeof hideExplain === 'function') {
    hideExplain();
  }
}

function buildFocusSelector(element) {
  if (!element) {
    return null;
  }
  if (element.dataset.bind) {
    return `[data-bind="${element.dataset.bind}"]`;
  }
  if (element.dataset.observationField) {
    return `[data-observation-field="${element.dataset.observationField}"]`;
  }
  if (element.dataset.section) {
    const parts = [
      `[data-section="${element.dataset.section}"]`,
      `[data-index="${element.dataset.index}"]`,
      `[data-field="${element.dataset.field}"]`
    ];
    if (element.dataset.readingIndex) {
      parts.push(`[data-reading-index="${element.dataset.readingIndex}"]`);
    }
    if (element.dataset.groupIndex) {
      parts.push(`[data-group-index="${element.dataset.groupIndex}"]`);
    }
    if (element.dataset.locationIndex) {
      parts.push(`[data-location-index="${element.dataset.locationIndex}"]`);
    }
    if (element.dataset.direction) {
      parts.push(`[data-direction="${element.dataset.direction}"]`);
    }
    return parts.join('');
  }
  return element.id ? `#${element.id}` : null;
}

function rerenderPreservingFocus() {
  const activeElement = document.activeElement;
  const selector = buildFocusSelector(activeElement);
  const selectionStart = activeElement && typeof activeElement.selectionStart === 'number' ? activeElement.selectionStart : null;
  const selectionEnd = activeElement && typeof activeElement.selectionEnd === 'number' ? activeElement.selectionEnd : null;
  render();
  if (!selector) {
    return;
  }
  const nextElement = document.querySelector(selector);
  if (!nextElement) {
    return;
  }
  nextElement.focus();
  if (selectionStart !== null && typeof nextElement.setSelectionRange === 'function') {
    nextElement.setSelectionRange(selectionStart, selectionEnd === null ? selectionStart : selectionEnd);
  }
}

function getPathSegments(bind) {
  return String(bind || '').split('.');
}

function setByPath(target, bind, value) {
  const segments = getPathSegments(bind);
  let cursor = target;
  for (let i = 0; i < segments.length - 1; i += 1) {
    if (cursor[segments[i]] === undefined || cursor[segments[i]] === null) {
      cursor[segments[i]] = {};
    }
    cursor = cursor[segments[i]];
  }
  cursor[segments[segments.length - 1]] = value;
}

function openObservationEditor(section, index, direction, groupIndex = null, locationIndex = null) {
  const row = getSectionRow(state.draft, section, index, direction, groupIndex, locationIndex);
  if (!row) {
    return;
  }
  state.observationEditor = {
    section,
    index,
    direction,
    groupIndex,
    locationIndex,
    rowId: safeText(row.rowId, buildRowId('row')),
    remark: safeText(row.rowObservation, ''),
    photos: cloneRowPhotos(row.rowPhotos)
  };
}

function closeObservationEditor() {
  state.observationEditor = null;
  state.observationUploading = false;
}

function saveObservationEditor() {
  if (!state.observationEditor) {
    return;
  }
  const row = getSectionRow(
    state.draft,
    state.observationEditor.section,
    state.observationEditor.index,
    state.observationEditor.direction,
    state.observationEditor.groupIndex,
    state.observationEditor.locationIndex
  );
  if (!row) {
    closeObservationEditor();
    return;
  }
  row.rowId = state.observationEditor.rowId;
  row.rowObservation = String(state.observationEditor.remark || '').trim();
  row.rowPhotos = cloneRowPhotos(state.observationEditor.photos).filter((photo) => photo.dataUrl);
  scheduleDraftAutosave();
  closeObservationEditor();
  showToast('Row observation saved.', 'healthy');
}

async function appendObservationFiles(fileList) {
  if (!state.observationEditor || !fileList?.length) {
    return;
  }
  state.observationUploading = true;
  render();
  try {
    const uploaded = await Promise.all(Array.from(fileList).map((file) => buildObservationPhoto(file)));
    state.observationEditor.photos.push(...uploaded.filter((photo) => photo.dataUrl));
  } catch (error) {
    showToast(error.message || 'Failed to add row photos.', 'critical');
  } finally {
    state.observationUploading = false;
    render();
  }
}

function validateProjectStep() {
  const project = state.draft.project;
  const fields = [
    ['projectNo', 'Project number is required.'],
    ['clientName', 'Client name is required.'],
    ['siteLocation', 'Site location is required.'],
    ['reportDate', 'Date of testing is required.'],
    ['engineerName', 'Testing engineer is required.']
  ];
  for (const [field, message] of fields) {
    if (!safeText(project[field], '')) {
      showToast(message, 'critical');
      return false;
    }
  }
  return true;
}

function validateSelectionStep() {
  if (!selectedTestCount(state.draft)) {
    showToast('Select at least one measurement section.', 'critical');
    return false;
  }
  return true;
}

function validateCurrentStep() {
  const step = currentStepId();
  if (step === 'project') {
    return validateProjectStep();
  }
  if (step === 'selection') {
    return validateSelectionStep();
  }
  return true;
}

async function saveReport() {
  if (state.editing) {
    return saveWorkspaceDay();
  }
  state.saving = true;
  closeObservationEditor();
  resetAllOcrStates(false);
  render();
  try {
    const created = await api('/api/reports', {
      method: 'POST',
      body: JSON.stringify(state.draft)
    });
    state.activeReport = created;
    state.view = 'detail';
    state.draft = createDraft();
    state.stepIndex = 0;
    clearDraftSnapshot();
    const warnings = ((created.assessment && created.assessment.validation && created.assessment.validation.warnings) || []).length;
    showToast(warnings ? `Report saved with ${warnings} warning${warnings === 1 ? '' : 's'} to check.` : 'Report saved successfully.', warnings ? 'warning' : 'healthy');
    await loadReports();
  } catch (error) {
    const extra = Array.isArray(error.errors) && error.errors.length > 1 ? ` (${error.errors.length} problems; see the sheets)` : '';
    showToast(`${error.message}${extra}`, 'critical');
  } finally {
    state.saving = false;
    render();
  }
}

async function openReport(id) {
  try {
    resetAllOcrStates(false);
    state.activeReport = await api(`/api/reports/${encodeURIComponent(id)}`);
    state.view = 'detail';
    render();
  } catch (error) {
    showToast(error.message, 'critical');
  }
}

async function deleteReport(id) {
  if (!window.confirm('Delete this report permanently?')) {
    return;
  }
  try {
    await api(`/api/reports/${encodeURIComponent(id)}`, { method: 'DELETE' });
    if (state.activeReport && state.activeReport.id === id) {
      state.activeReport = null;
      state.view = 'dashboard';
    }
    showToast('Report deleted.', 'healthy');
    await loadReports();
  } catch (error) {
    showToast(error.message, 'critical');
  }
}

function shouldUseMobilePdfFlow() {
  const touchPoints = navigator.maxTouchPoints || 0;
  const isCompactViewport = window.matchMedia('(max-width: 920px)').matches;
  return touchPoints > 0 && isCompactViewport;
}

async function presentGeneratedPdf(result) {
  const absoluteUrl = new URL(result.pdfUrl, window.location.origin).toString();
  const fileName = result.fileName || absoluteUrl.split('/').pop() || 'ElectroReports.pdf';

  if (!shouldUseMobilePdfFlow()) {
    window.open(absoluteUrl, '_blank', 'noopener');
    return;
  }

  if (typeof navigator.share === 'function') {
    try {
      const response = await fetch(absoluteUrl);
      if (response.ok) {
        const blob = await response.blob();
        const file = new File([blob], fileName, { type: 'application/pdf' });
        if (!navigator.canShare || navigator.canShare({ files: [file] })) {
          await navigator.share({
            title: 'ElectroReports PDF',
            files: [file]
          });
          return;
        }
      }
    } catch (error) {
      if (error?.name === 'AbortError') {
        return;
      }
    }

    try {
      await navigator.share({
        title: 'ElectroReports PDF',
        url: absoluteUrl
      });
      return;
    } catch (error) {
      if (error?.name === 'AbortError') {
        return;
      }
    }
  }

  window.location.href = absoluteUrl;
}

async function exportPdf(id) {
  state.exporting = true;
  render();
  try {
    const result = await api(`/api/reports/${encodeURIComponent(id)}/pdf`, { method: 'POST' });
    await presentGeneratedPdf(result);
    showToast('PDF generated successfully.', 'healthy');
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    state.exporting = false;
    render();
  }
}

async function generateAiNarrative(id) {
  state.ai.generating = true;
  render();
  try {
    const updated = await api(`/api/reports/${encodeURIComponent(id)}/ai/generate`, { method: 'POST' });
    state.activeReport = updated;
    state.reports = state.reports.map((report) => (report.id === updated.id ? updated : report));
    showToast('AI summary generated successfully.', 'healthy');
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    state.ai.generating = false;
    render();
  }
}

document.addEventListener('click', async (event) => {
  const button = event.target.closest('[data-action]');
  if (!button) {
    return;
  }

  const action = button.dataset.action;

  if (handlePrintAction(action, button)) {
    return;
  }

  if (handleReviewAction(action, button)) {
    return;
  }

  if (handleAuthAction(action, button)) {
    return;
  }

  if (typeof handleKbAction === 'function' && handleKbAction(action, button)) {
    return;
  }

  if (typeof handleScanAction === 'function' && (handleScanAction(action, button) || handleScanAdminAction(action, button))) {
    return;
  }

  if (handleNoticeAction(action, button)) {
    return;
  }

  if (handleScheduleAction(action, button)) {
    return;
  }

  if (handleWorkspaceAction(action, button)) {
    return;
  }

  if (handleSheetAction(action, button)) {
    render();
    return;
  }

  if (action === 'new-report') {
    closeObservationEditor();
    state.editing = null;
    resetPlanUi();
    state.draft = createDraft();
    state.stepIndex = 0;
    state.view = 'builder';
    resetAllOcrStates(false);
    clearDraftSnapshot();
    render();
    return;
  }

  if (action === 'go-dashboard') {
    closeObservationEditor();
    state.editing = null;
    state.view = 'dashboard';
    state.activeReport = null;
    resetAllOcrStates(false);
    persistDraftSnapshotNow();
    render();
    return;
  }

  if (action === 'set-ocr-entry-mode') {
    const sheetId = button.dataset.sheet;
    const sheetState = getOcrSheetState(sheetId);
    const nextMode = button.dataset.mode === 'upload' ? 'upload' : 'manual';
    if (nextMode === 'manual') {
      resetOcrSheetState(sheetId, false);
    } else {
      sheetState.mode = 'upload';
      sheetState.error = '';
    }
    render();
    return;
  }

  if (action === 'scan-ocr') {
    await previewSheetOcr(button.dataset.sheet);
    return;
  }

  if (action === 'clear-ocr') {
    const sheetId = button.dataset.sheet;
    resetOcrSheetState(sheetId, getOcrSheetState(sheetId).mode === 'upload');
    render();
    return;
  }

  if (action === 'apply-ocr-preview') {
    applySheetOcrPreview(button.dataset.sheet);
    return;
  }

  if (action === 'search-focus') {
    const input = document.getElementById('reportSearchInput');
    if (input) {
      input.focus();
    }
    return;
  }

  if (action === 'toggle-test') {
    const testId = button.dataset.test;
    state.draft.tests[testId] = !state.draft.tests[testId];
    const steps = getSteps();
    state.stepIndex = Math.min(state.stepIndex, steps.length - 1);
    render();
    return;
  }

  if (action === 'step-next') {
    if (!validateCurrentStep()) {
      return;
    }
    state.stepIndex = Math.min(state.stepIndex + 1, getSteps().length - 1);
    render();
    return;
  }

  if (action === 'step-prev') {
    state.stepIndex = Math.max(state.stepIndex - 1, 0);
    render();
    return;
  }

  if (action === 'go-step') {
    const nextIndex = Number(button.dataset.stepIndex);
    if (nextIndex > state.stepIndex && !validateCurrentStep()) {
      return;
    }
    state.stepIndex = nextIndex;
    render();
    return;
  }

  if (action === 'open-row-observation') {
    openObservationEditor(
      button.dataset.section,
      Number(button.dataset.index),
      button.dataset.direction,
      button.dataset.groupIndex === undefined ? null : Number(button.dataset.groupIndex),
      button.dataset.locationIndex === undefined ? null : Number(button.dataset.locationIndex)
    );
    render();
    return;
  }

  if (action === 'close-row-observation') {
    closeObservationEditor();
    render();
    return;
  }

  if (action === 'save-row-observation') {
    saveObservationEditor();
    render();
    return;
  }

  if (action === 'remove-observation-photo') {
    if (state.observationEditor) {
      state.observationEditor.photos.splice(Number(button.dataset.photoIndex), 1);
      render();
    }
    return;
  }

  if (action === 'save-report') {
    if (!validateProjectStep() || !validateSelectionStep()) {
      return;
    }
    await saveReport();
    return;
  }

  if (action === 'open-report') {
    await openReport(button.dataset.id);
    return;
  }

  if (action === 'delete-report') {
    await deleteReport(button.dataset.id);
    return;
  }

  if (action === 'export-pdf') {
    await exportPdf(button.dataset.id);
    return;
  }

  if (action === 'generate-ai-summary') {
    await generateAiNarrative(button.dataset.id);
    return;
  }

  if (action === 'refresh-zoho-projects') {
    await loadZohoProjects();
    return;
  }

  if (action === 'use-zoho-project') {
    state.editing = null;
    applyZohoProject(button.dataset.projectId);
    state.view = 'builder';
    state.stepIndex = 0;
    render();
    return;
  }

  if (action === 'use-zoho-user') {
    state.draft.project.engineerName = button.dataset.userName || state.draft.project.engineerName;
    showToast(`Assigned ${state.draft.project.engineerName} from Zoho users.`, 'healthy');
    render();
  }
});

document.addEventListener('input', (event) => {
  const target = event.target;

  if (handlePrintInput(target)) {
    return;
  }

  if (handleReviewInput(target)) {
    return;
  }

  if (typeof handleKbInput === 'function' && handleKbInput(target)) {
    return;
  }

  if (typeof handleScanInput === 'function' && handleScanInput(target)) {
    return;
  }

  if (handleWorkspaceInput(target)) {
    return;
  }

  // Ticks are handled on `change` (one event per tap); typing and selects are handled here.
  if (target.type === 'checkbox' || target.type === 'radio') {
    return;
  }

  if (target.matches('[data-bind]')) {
    if (target.dataset.bind === 'project.unitCode') {
      target.value = target.value.replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(0, 8);
    }
    setByPath(state.draft, target.dataset.bind, target.value);
    planOnPlannedEdit(target.dataset.bind);
    scheduleDraftAutosave();
    if (state.view === 'builder') {
      syncDraftAssessment();
    }
    return;
  }

  if (target.matches('[data-observation-field]')) {
    if (state.observationEditor) {
      state.observationEditor[target.dataset.observationField] = target.value;
    }
    return;
  }

  if (target.id === 'reportSearchInput') {
    state.search = target.value;
    window.clearTimeout(document.searchDebounce);
    document.searchDebounce = window.setTimeout(() => {
      loadReports();
    }, 220);
    return;
  }

  if (target.matches('[data-section], [data-soil-cell], [data-soil-location], [data-header-sheet]')) {
    if (applyFieldInput(target)) {
      if (typeof autoAddRow === 'function') {
        autoAddRow(target);
      }
      scheduleDraftAutosave();
      rerenderPreservingFocus();
    }
    return;
  }
});

document.addEventListener('change', async (event) => {
  const target = event.target;
  if (handleReviewChange(target)) {
    return;
  }
  if (handleAuthChange(target)) {
    return;
  }
  if (typeof handleKbChange === 'function' && handleKbChange(target)) {
    return;
  }
  if (typeof handleScanChange === 'function' && handleScanChange(target)) {
    return;
  }
  if (handlePlanChange(target)) {
    return;
  }
  if (handleNoticeChange(target)) {
    return;
  }
  if (handleScheduleChange(target)) {
    return;
  }
  if (target.id === 'zohoProjectPicker') {
    applyZohoProject(target.value);
    return;
  }
  if (
    target.id === 'soilOcrFileInput' ||
    target.id === 'electrodeOcrFileInput' ||
    target.id === 'continuityOcrFileInput' ||
    target.id === 'loopOcrFileInput' ||
    target.id === 'faultOcrFileInput' ||
    target.id === 'riserOcrFileInput' ||
    target.id === 'earthContinuityOcrFileInput' ||
    target.id === 'towerFootingOcrFileInput'
  ) {
    const sheetIdMap = {
      soilOcrFileInput: 'soilResistivity',
      electrodeOcrFileInput: 'electrodeResistance',
      continuityOcrFileInput: 'continuityTest',
      loopOcrFileInput: 'loopImpedanceTest',
      faultOcrFileInput: 'prospectiveFaultCurrent',
      riserOcrFileInput: 'riserIntegrityTest',
      earthContinuityOcrFileInput: 'earthContinuityTest',
      towerFootingOcrFileInput: 'towerFootingResistance'
    };
    const sheetId = sheetIdMap[target.id];
    const sheetState = getOcrSheetState(sheetId);
    const [file] = Array.from(target.files || []);
    sheetState.selectedFile = file || null;
    sheetState.preview = null;
    sheetState.previewMeta = null;
    sheetState.draftPatch = null;
    sheetState.warnings = [];
    sheetState.uncertainFields = [];
    sheetState.error = '';
    render();
    return;
  }
  if (target.id === 'rowObservationFiles') {
    await appendObservationFiles(target.files);
    target.value = '';
    return;
  }
  if (target.dataset.bind === 'project.soilSpacingPreset') {
    applySoilPreset();
    scheduleDraftAutosave();
    render();
    return;
  }
  if (target.dataset.bind && (target.type === 'radio' || target.type === 'checkbox')) {
    setByPath(state.draft, target.dataset.bind, target.value);
    scheduleDraftAutosave();
    render();
    return;
  }
  if (target.matches('[data-section], [data-soil-cell], [data-soil-location], [data-header-sheet]')) {
    if (applyFieldInput(target)) {
      if (typeof autoAddRow === 'function') {
        autoAddRow(target);
      }
      scheduleDraftAutosave();
      rerenderPreservingFocus();
    }
    return;
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && state.observationEditor) {
    closeObservationEditor();
    render();
  }
});

// The sheet definitions (what each entry screen contains) and the acceptance profiles come
// from the server; the draft cannot be built or restored without them.
async function loadSheetDefinitions() {
  try {
    const [definitions, profiles] = await Promise.all([api('/api/sheets'), api('/api/acceptance/profiles')]);
    setSheetDefinitions(definitions);
    setAcceptanceProfiles(profiles);
  } catch (error) {
    showToast(`Could not load the measurement sheets: ${error.message}`, 'critical');
  }
}

async function init() {
  await loadAuth();
  if (!authUi.user) {
    state.view = 'signin';
    render();
    return;
  }
  await loadSheetDefinitions();
  if (!restoreDraftSnapshot()) {
    state.draft = createDraft();
  }
  render();
  if (state.restoredDraftNoticePending) {
    state.restoredDraftNoticePending = false;
    showToast('Recovered your in-progress draft.', 'warning');
  }
  await loadAiStatus();
  await loadOcrStatus();
  await loadCatalog();
  await loadZohoProjects();
  await loadReports();
  await loadProjects();
  await loadNotices();
  if (typeof consumeDeepLink === 'function') {
    await consumeDeepLink();
  }
}

window.addEventListener('beforeunload', () => {
  persistDraftSnapshotNow();
});

init();
