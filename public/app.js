const app = document.getElementById('app');

const LOCAL_TEST_LIBRARY = [
  {
    id: 'soilResistivity',
    label: 'Soil Resistivity Test',
    shortLabel: 'Soil',
    description: 'Direction-wise soil resistivity readings with automatic average and category.'
  },
  {
    id: 'electrodeResistance',
    label: 'Earth Electrode Resistance Test',
    shortLabel: 'Electrode',
    description: 'Pit-wise resistance against the 4.60 ohm permissible limit.'
  },
  {
    id: 'continuityTest',
    label: 'Continuity Test',
    shortLabel: 'Continuity',
    description: 'Point-to-point resistance and impedance verification.'
  },
  {
    id: 'loopImpedanceTest',
    label: 'Loop Impedance Test',
    shortLabel: 'Loop',
    description: 'Measured Zs values for protection loop verification.'
  },
  {
    id: 'prospectiveFaultCurrent',
    label: 'Prospective Fault Current',
    shortLabel: 'PFC',
    description: 'Feeder details with loop impedance and fault current.'
  },
  {
    id: 'riserIntegrityTest',
    label: 'Riser / Grid Integrity Test',
    shortLabel: 'Riser',
    description: 'Resistance verification towards equipment and earth grid.'
  },
  {
    id: 'earthContinuityTest',
    label: 'Earth Continuity Test',
    shortLabel: 'Earth Continuity',
    description: 'Earth path continuity by tag, location, and measured value.'
  },
  {
    id: 'towerFootingResistance',
    label: 'Tower Footing Resistance Measurement & Analysis',
    shortLabel: 'Tower Footing',
    description: 'Grouped tower footing impedance and current analysis with per-tower totals.'
  }
];

const TEST_STANDARD_REFERENCES = {
  soilResistivity: ['IEEE 81-2012', 'IS 3043:2018 Clause 9.2'],
  electrodeResistance: ['IS 3043:2018', 'IS 3043:2018 Clause 8'],
  continuityTest: ['IEC 60364-6', 'IS 732 continuity verification'],
  loopImpedanceTest: ['IEC 60364-6', 'IS/IEC protective loop verification'],
  prospectiveFaultCurrent: ['IEC 60909', 'IS/IEC fault level assessment'],
  riserIntegrityTest: ['IS 3043:2018', 'IEC 60364 earth continuity guidance'],
  earthContinuityTest: ['IS 3043:2018', 'IEC 60364 continuity verification'],
  towerFootingResistance: ['Fixed Zsat 10 ohm', '4 fixed footing rows per tower location']
};

const SOIL_SPACING_PRESETS = ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0', '4.5', '5.0'];
const TOWER_FOOT_POINTS = ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'];
const PHASE_MEASURED_POINTS = ['R-E', 'Y-E', 'B-E'];
const EQUIPMENT_LIBRARY = [
  { id: 'mi3152', label: 'MI 3152 EurotestXC' },
  { id: 'mi3290', label: 'MI 3290 GF Earth Analyser' },
  { id: 'kyoritsu4118a', label: 'Kyoritsu Digital PSC Loop Tester 4118A' }
];
const DRAFT_STORAGE_KEY = 'electroreports-builder-draft-v1';
const DRAFT_AUTOSAVE_DELAY_MS = 450;

function buildRowId(prefix = 'row') {
  const timestamp = Date.now().toString(36);
  const suffix = Math.random().toString(36).slice(2, 8);
  return `${prefix}-${timestamp}-${suffix}`;
}

function defaultSoilRow(spacing = '') {
  return {
    rowId: buildRowId('soil'),
    spacing,
    resistivity: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultSoilLocation(name = 'Location 1') {
  return {
    locationId: buildRowId('soil-location'),
    name,
    direction1: SOIL_SPACING_PRESETS.slice(0, 6).map((spacing) => defaultSoilRow(spacing)),
    direction2: SOIL_SPACING_PRESETS.slice(0, 6).map((spacing) => defaultSoilRow(spacing)),
    drivenElectrodeDiameter: '',
    drivenElectrodeLength: '',
    notes: ''
  };
}

function defaultElectrodeRow() {
  return {
    rowId: buildRowId('electrode'),
    tag: '',
    location: '',
    electrodeType: 'Rod',
    materialType: 'Copper',
    length: '',
    diameter: '',
    resistanceWithoutGrid: '',
    resistanceWithGrid: '',
    observation: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultContinuityRow(srNo = '') {
  return {
    rowId: buildRowId('continuity'),
    srNo,
    mainLocation: '',
    measurementPoint: '',
    resistance: '',
    impedance: '',
    comment: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultLoopRow(srNo = '') {
  return {
    rowId: buildRowId('loop'),
    srNo,
    location: '',
    feederTag: '',
    deviceType: 'MCB',
    deviceRating: '',
    breakingCapacity: '',
    measuredPoints: 'R-E',
    loopImpedance: '',
    voltage: '230',
    remarks: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultFaultRow(srNo = '') {
  return {
    rowId: buildRowId('fault'),
    srNo,
    location: '',
    feederTag: '',
    deviceType: 'MCB',
    deviceRating: '',
    breakingCapacity: '',
    measuredPoints: 'R-E',
    loopImpedance: '',
    prospectiveFaultCurrent: '',
    voltage: '230',
    comment: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultLoopGroup(startIndex = 1) {
  return PHASE_MEASURED_POINTS.map((point, index) => ({
    ...defaultLoopRow(String(startIndex + index)),
    measuredPoints: point
  }));
}

function defaultFaultGroup(startIndex = 1) {
  return PHASE_MEASURED_POINTS.map((point, index) => ({
    ...defaultFaultRow(String(startIndex + index)),
    measuredPoints: point
  }));
}

function defaultRiserRow(srNo = '') {
  return {
    rowId: buildRowId('riser'),
    srNo,
    mainLocation: '',
    measurementPoint: '',
    resistanceTowardsEquipment: '',
    resistanceTowardsGrid: '',
    comment: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultEarthContinuityRow(srNo = '') {
  return {
    rowId: buildRowId('earth'),
    srNo,
    tag: '',
    locationBuildingName: '',
    distance: '',
    measuredValue: '',
    remark: '',
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultTowerFootingReading(foot) {
  return {
    measurementPointLocation: '',
    footLabel: foot,
    footToEarthingConnectionStatus: 'Given',
    measuredCurrentMa: '',
    measuredImpedance: '',
    rowId: buildRowId('tower'),
    rowObservation: '',
    rowPhotos: []
  };
}

function defaultTowerFootingGroup(srNo = '') {
  return {
    groupId: buildRowId('tower-group'),
    srNo,
    mainLocationTower: '',
    readings: TOWER_FOOT_POINTS.map((foot) => {
      const reading = defaultTowerFootingReading(foot);
      reading.measurementPointLocation = foot;
      return reading;
    }),
    totalImpedanceZt: '',
    totalCurrentItotal: '',
    standardTolerableImpedanceZsat: '10',
    remarks: ''
  };
}

function createDraft() {
  return {
    project: {
      projectNo: '',
      clientName: '',
      siteLocation: '',
      workOrder: '',
      reportDate: new Date().toISOString().slice(0, 10),
      engineerName: '',
      equipmentSelections: EQUIPMENT_LIBRARY.map((equipment) => equipment.id),
      zohoProjectId: '',
      zohoProjectName: '',
      zohoProjectOwner: '',
      zohoProjectStage: ''
    },
    tests: {
      soilResistivity: true,
      electrodeResistance: true,
      continuityTest: false,
      loopImpedanceTest: false,
      prospectiveFaultCurrent: false,
      riserIntegrityTest: false,
      earthContinuityTest: false,
      towerFootingResistance: false
    },
    soilResistivity: {
      locations: [defaultSoilLocation('Location 1')]
    },
    electrodeResistance: [defaultElectrodeRow()],
    continuityTest: [defaultContinuityRow('1')],
    loopImpedanceTest: defaultLoopGroup(1),
    prospectiveFaultCurrent: defaultFaultGroup(1),
    riserIntegrityTest: [defaultRiserRow('1')],
    earthContinuityTest: [defaultEarthContinuityRow('1')],
    towerFootingResistance: [defaultTowerFootingGroup('1')],
    ocrImports: {}
  };
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
    loadingProjects: false,
    loadingUsers: false
  },
  draft: createDraft(),
  stepIndex: 0,
  activeReport: null,
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

function restoreDraftShape(payload) {
  const draft = createDraft();
  const source = payload && typeof payload === 'object' ? payload : {};

  if (source.project && typeof source.project === 'object') {
    draft.project = { ...draft.project, ...cloneSerializable(source.project) };
  }

  if (source.tests && typeof source.tests === 'object') {
    draft.tests = { ...draft.tests, ...cloneSerializable(source.tests) };
  }

  if (source.soilResistivity && typeof source.soilResistivity === 'object') {
    draft.soilResistivity = cloneSerializable(source.soilResistivity);
  }

  [
    'electrodeResistance',
    'continuityTest',
    'loopImpedanceTest',
    'prospectiveFaultCurrent',
    'riserIntegrityTest',
    'earthContinuityTest',
    'towerFootingResistance'
  ].forEach((section) => {
    if (Array.isArray(source[section])) {
      draft[section] = cloneSerializable(source[section]);
    }
  });

  draft.ocrImports = normalizeRestoredOcrImports(source.ocrImports);
  return draft;
}

function projectHasMeaningfulData(source) {
  const project = source?.project || {};
  return Boolean(
    safeText(project.projectNo, '') ||
      safeText(project.clientName, '') ||
      safeText(project.siteLocation, '') ||
      safeText(project.workOrder, '') ||
      safeText(project.engineerName, '') ||
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
      (Array.isArray(location.direction1) && location.direction1.some(rowHasData)) ||
      (Array.isArray(location.direction2) && location.direction2.some(rowHasData))
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

function draftHasMeaningfulElectrodeData(source) {
  return (Array.isArray(source?.electrodeResistance) ? source.electrodeResistance : []).some((row) => {
    return Boolean(
      safeText(row?.tag, '') ||
        safeText(row?.location, '') ||
        safeText(row?.electrodeType, '') !== 'Rod' ||
        safeText(row?.materialType, '') !== 'Copper' ||
        safeText(row?.length, '') ||
        safeText(row?.diameter, '') ||
        safeText(row?.resistanceWithoutGrid, '') ||
        safeText(row?.resistanceWithGrid, '') ||
        rowHasObservationData(row)
    );
  });
}

function draftHasMeaningfulContinuityData(source) {
  return (Array.isArray(source?.continuityTest) ? source.continuityTest : []).some((row, index) => {
    return Boolean(
      (safeText(row?.srNo, '') && safeText(row?.srNo, '') !== String(index + 1)) ||
        safeText(row?.mainLocation, '') ||
        safeText(row?.measurementPoint, '') ||
        safeText(row?.resistance, '') ||
        safeText(row?.impedance, '') ||
        rowHasObservationData(row)
    );
  });
}

function draftHasMeaningfulLoopData(source) {
  return (Array.isArray(source?.loopImpedanceTest) ? source.loopImpedanceTest : []).some((row, index) => {
    return Boolean(
      (safeText(row?.srNo, '') && safeText(row?.srNo, '') !== String(index + 1)) ||
        safeText(row?.location, '') ||
        safeText(row?.feederTag, '') ||
        safeText(row?.deviceType, '') !== 'MCB' ||
        safeText(row?.deviceRating, '') ||
        safeText(row?.breakingCapacity, '') ||
        safeText(row?.loopImpedance, '') ||
        safeText(row?.voltage, '') !== '230' ||
        rowHasObservationData(row)
    );
  });
}

function draftHasMeaningfulFaultData(source) {
  return (Array.isArray(source?.prospectiveFaultCurrent) ? source.prospectiveFaultCurrent : []).some((row, index) => {
    return Boolean(
      (safeText(row?.srNo, '') && safeText(row?.srNo, '') !== String(index + 1)) ||
        safeText(row?.location, '') ||
        safeText(row?.feederTag, '') ||
        safeText(row?.deviceType, '') !== 'MCB' ||
        safeText(row?.deviceRating, '') ||
        safeText(row?.breakingCapacity, '') ||
        safeText(row?.loopImpedance, '') ||
        safeText(row?.prospectiveFaultCurrent, '') ||
        safeText(row?.voltage, '') !== '230' ||
        rowHasObservationData(row)
    );
  });
}

function draftHasMeaningfulRiserData(source) {
  return (Array.isArray(source?.riserIntegrityTest) ? source.riserIntegrityTest : []).some((row, index) => {
    return Boolean(
      (safeText(row?.srNo, '') && safeText(row?.srNo, '') !== String(index + 1)) ||
        safeText(row?.mainLocation, '') ||
        safeText(row?.measurementPoint, '') ||
        safeText(row?.resistanceTowardsEquipment, '') ||
        safeText(row?.resistanceTowardsGrid, '') ||
        rowHasObservationData(row)
    );
  });
}

function draftHasMeaningfulEarthContinuityData(source) {
  return (Array.isArray(source?.earthContinuityTest) ? source.earthContinuityTest : []).some((row, index) => {
    return Boolean(
      (safeText(row?.srNo, '') && safeText(row?.srNo, '') !== String(index + 1)) ||
        safeText(row?.tag, '') ||
        safeText(row?.locationBuildingName, '') ||
        safeText(row?.distance, '') ||
        safeText(row?.measuredValue, '') ||
        rowHasObservationData(row)
    );
  });
}

function towerGroupHasMeaningfulData(group, index) {
  return Boolean(
    (safeText(group?.srNo, '') && safeText(group?.srNo, '') !== String(index + 1)) ||
      safeText(group?.mainLocationTower, '') ||
      (Array.isArray(group?.readings) &&
        group.readings.some((reading, readingIndex) => {
          return Boolean(
            safeText(reading?.measurementPointLocation, '') !== TOWER_FOOT_POINTS[readingIndex] ||
              safeText(reading?.footToEarthingConnectionStatus, '') !== 'Given' ||
              safeText(reading?.measuredCurrentMa, '') ||
              safeText(reading?.measuredImpedance, '') ||
              rowHasObservationData(reading)
          );
        }))
  );
}

function draftHasMeaningfulTowerData(source) {
  return (Array.isArray(source?.towerFootingResistance) ? source.towerFootingResistance : []).some((group, index) =>
    towerGroupHasMeaningfulData(group, index)
  );
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
    return `${getTestLabel(section)} | ${locationName} | Row ${index + 1} | ${direction === 'direction2' ? 'Direction 2' : 'Direction 1'}`;
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

function getSoilCategory(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Insufficient Data', tone: 'neutral' };
  }
  if (value < 100) {
    return { label: 'Low', tone: 'healthy' };
  }
  if (value <= 500) {
    return { label: 'Medium', tone: 'warning' };
  }
  return { label: 'High', tone: 'critical' };
}

function getElectrodeStatus(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (value <= 4.6) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  return { label: 'Exceeds Permissible Limit', tone: 'critical' };
}

function getContinuityStatus(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (value <= 0.5) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (value <= 1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getLoopStatus(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (value <= 1) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (value <= 1.5) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getRiserStatus(equipment, grid) {
  const values = [equipment, grid].filter((value) => Number.isFinite(value));
  if (!values.length) {
    return { label: 'Pending', tone: 'neutral' };
  }
  const highest = Math.max(...values);
  if (highest <= 0.5) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (highest <= 1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getEarthContinuityStatus(value) {
  const numeric = toNumber(value);
  if (!Number.isFinite(numeric)) {
    return { label: 'Pending', tone: 'neutral' };
  }
  if (numeric <= 0.5) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (numeric <= 1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

const STANDARD_GUIDANCE = {
  electrodeResistance: {
    reference: 'IS 3043',
    limitLabel: '4.60 ohm project limit'
  },
  continuityTest: {
    reference: 'Continuity guidance',
    limitLabel: '0.50 ohm continuity reference'
  },
  loopImpedanceTest: {
    reference: 'IEC 60364 / device-specific max Zs',
    limitLabel: 'Circuit-specific maximum Zs'
  },
  prospectiveFaultCurrent: {
    reference: 'IEC breaking-capacity check',
    limitLabel: 'PFC must not exceed device breaking capacity'
  },
  riserIntegrityTest: {
    reference: 'Continuity / bonding guidance',
    limitLabel: '0.50 ohm continuity reference'
  },
  earthContinuityTest: {
    reference: 'Earth continuity guidance',
    limitLabel: '0.50 ohm continuity reference'
  },
  towerFootingResistance: {
    reference: 'Project default tower footing limit',
    limitLabel: '10.00 ohm Zsat default'
  }
};

function toReportBandLabel(status) {
  if (status?.tone === 'healthy') {
    return 'Healthy';
  }
  if (status?.tone === 'warning') {
    return 'Warning';
  }
  if (status?.tone === 'critical') {
    return '> Permissible Limit';
  }
  return 'Pending';
}

function toRiserCommentLabel(status) {
  if (status?.tone === 'healthy') {
    return 'Healthy';
  }
  if (status?.tone === 'neutral') {
    return 'Pending';
  }
  return 'Un-Healthy';
}

function deriveElectrodeAssessment(row) {
  const measured = getElectrodeMeasuredValue(row);
  const withoutGrid = toNumber(row?.resistanceWithoutGrid);
  const withGrid = toNumber(row?.resistanceWithGrid) ?? toNumber(row?.measuredResistance);
  const status = getElectrodeStatus(measured);
  const standard = STANDARD_GUIDANCE.electrodeResistance;

  let comment = `Enter the resistance value with grid to compare against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Resistance with grid is within the ${standard.limitLabel} under ${standard.reference}.`;
  } else if (status.tone === 'critical') {
    comment = `Resistance with grid exceeds the ${standard.limitLabel}; review earthing improvement as per ${standard.reference}.`;
  }

  if (Number.isFinite(withoutGrid) && Number.isFinite(withGrid)) {
    if (withGrid > withoutGrid) {
      comment += ' The with-grid value is higher than the without-grid value, so verify the test setup and bonding path.';
    } else if (withGrid < withoutGrid) {
      comment += ' The with-grid value improved after bonding to the grid, which is the expected trend.';
    }
  }

  return {
    status,
    standard,
    comment
  };
}

function deriveContinuityAssessment(row) {
  const resistance = toNumber(row?.resistance);
  const impedance = toNumber(row?.impedance);
  const status = getContinuityStatus(resistance);
  const standard = STANDARD_GUIDANCE.continuityTest;

  let comment = `Enter a resistance reading to assess continuity against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Continuity resistance is within the ${standard.limitLabel}; bond/joint continuity is acceptable.`;
  } else if (status.tone === 'warning') {
    comment = `Continuity resistance is above the ${standard.limitLabel}; inspect joints, terminations, and bonding integrity.`;
  } else if (status.tone === 'critical') {
    comment = `Continuity resistance is well above the ${standard.limitLabel}; urgent investigation of the continuity path is recommended.`;
  }

  if (Number.isFinite(impedance)) {
    comment += ` Recorded impedance: ${round(impedance, 2)} ohm.`;
  }

  return {
    status,
    standard,
    comment
  };
}

function deriveLoopAssessment(row) {
  const measured = toNumber(row?.loopImpedance) ?? toNumber(row?.measuredZs);
  const status = getLoopStatus(measured);
  const standard = STANDARD_GUIDANCE.loopImpedanceTest;

  let comment = `Enter measured Zs and verify it against the protective device's maximum permitted Zs under ${standard.reference}.`;
  if (status.tone === 'healthy') {
    comment = `Measured Zs is in the healthy range, but final compliance still depends on the circuit/device-specific maximum Zs under ${standard.reference}.`;
  } else if (status.tone === 'warning') {
    comment = `Measured Zs is elevated; compare it carefully against the circuit/device-specific maximum Zs and confirm disconnection performance.`;
  } else if (status.tone === 'critical') {
    comment = `Measured Zs is high and likely to challenge disconnection performance; verify immediately against the protective device's maximum Zs.`;
  }

  return {
    status,
    standard,
    comment
  };
}

function deriveFaultAssessment(row) {
  const pfc = toNumber(row?.prospectiveFaultCurrent);
  const breakingCapacity = toNumber(row?.breakingCapacity);
  const standard = STANDARD_GUIDANCE.prospectiveFaultCurrent;
  let status = { label: 'Pending', tone: 'neutral' };
  let comment = `Enter both prospective fault current and device breaking capacity to verify that ${standard.limitLabel}.`;

  if (Number.isFinite(pfc) && Number.isFinite(breakingCapacity)) {
    if (pfc <= breakingCapacity * 0.9) {
      status = { label: 'Healthy', tone: 'healthy' };
      comment = 'Prospective fault current is comfortably within the device breaking capacity.';
    } else if (pfc <= breakingCapacity) {
      status = { label: 'Needs Attention', tone: 'warning' };
      comment = 'Prospective fault current is within the device breaking capacity, but the safety margin is small.';
    } else {
      status = { label: 'Critical', tone: 'critical' };
      comment = 'Prospective fault current exceeds the device breaking capacity; review protective device selection immediately.';
    }
  }

  return {
    status,
    standard,
    comment
  };
}

function deriveRiserAssessment(row) {
  const equipment = toNumber(row?.resistanceTowardsEquipment);
  const grid = toNumber(row?.resistanceTowardsGrid);
  const status = getRiserStatus(equipment, grid);
  const standard = STANDARD_GUIDANCE.riserIntegrityTest;

  let comment = `Enter both resistance readings to assess continuity against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Riser continuity readings are within the ${standard.limitLabel} toward equipment and grid.`;
  } else if (status.tone === 'warning') {
    comment = `One or both riser continuity readings are above the ${standard.limitLabel}; inspect joints, lugs, and bonding interfaces.`;
  } else if (status.tone === 'critical') {
    comment = `Riser continuity readings are high; investigate the riser path, bonding terminations, and earth grid connection urgently.`;
  }

  return {
    status,
    standard,
    comment
  };
}

function deriveEarthContinuityAssessment(row) {
  const measured = toNumber(row?.measuredValue);
  const status = getEarthContinuityStatus(measured);
  const standard = STANDARD_GUIDANCE.earthContinuityTest;

  let comment = `Enter a measured value to assess earth continuity against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Earth continuity is within the ${standard.limitLabel}; the earth path is acceptable for this point.`;
  } else if (status.tone === 'warning') {
    comment = `Earth continuity is above the ${standard.limitLabel}; inspect the earth path, joints, and terminations.`;
  } else if (status.tone === 'critical') {
    comment = `Earth continuity is well above the ${standard.limitLabel}; urgent investigation of the earth path is recommended.`;
  }

  return {
    status,
    standard,
    comment
  };
}

function buildTowerGroupKey(group) {
  const groupId = safeText(group?.groupId, '');
  const location = safeText(group?.mainLocationTower, '');
  return groupId || location || `group:${safeText(group?.srNo, 'tower')}`;
}

function summarizeTowerGroups(groups) {
  const summaries = new Map();

  (Array.isArray(groups) ? groups : []).forEach((group) => {
    const readings = Array.isArray(group?.readings) ? group.readings : [];
    const impedanceValues = readings.map((reading) => toNumber(reading?.measuredImpedance)).filter((value) => Number.isFinite(value));
    const currentValues = readings.map((reading) => toNumber(reading?.measuredCurrentMa)).filter((value) => Number.isFinite(value));
    const key = buildTowerGroupKey(group);
    const hasAnyInput = Boolean(safeText(group?.mainLocationTower, '')) || readings.some((reading) => {
      return (
        safeText(reading?.footToEarthingConnectionStatus, 'Given') !== 'Given' ||
        safeText(reading?.measuredCurrentMa, '') ||
        safeText(reading?.measuredImpedance, '') ||
        rowHasObservationData(reading)
      );
    });

    summaries.set(key, {
      impedanceCount: impedanceValues.length,
      currentCount: currentValues.length,
      totalImpedanceZt: impedanceValues.length === TOWER_FOOT_POINTS.length ? round(impedanceValues.reduce((sum, value) => sum + value, 0), 2) : null,
      totalCurrentItotal: currentValues.length === TOWER_FOOT_POINTS.length ? round(currentValues.reduce((sum, value) => sum + value, 0), 2) : null,
      hasAnyInput
    });
  });

  return summaries;
}

function deriveTowerFootingAssessment(group, groupSummary) {
  const standard = STANDARD_GUIDANCE.towerFootingResistance;
  const zsat = 10;
  const totalImpedanceZt = groupSummary?.totalImpedanceZt ?? null;
  const totalCurrentItotal = groupSummary?.totalCurrentItotal ?? null;
  let status = { label: 'Pending', tone: 'neutral' };
  let comment = '-';

  if (Number.isFinite(totalImpedanceZt)) {
    if (totalImpedanceZt <= zsat) {
      status = { label: 'Healthy', tone: 'healthy' };
      comment = 'Healthy';
    } else if (totalImpedanceZt <= zsat * 1.2) {
      status = { label: 'Marginal', tone: 'warning' };
      comment = 'Marginal';
    } else {
      status = { label: 'Not Acceptable', tone: 'critical' };
      comment = 'Not Acceptable';
    }
  } else if (groupSummary?.hasAnyInput) {
    comment = '-';
  }

  return {
    status,
    standard,
    totalImpedanceZt,
    totalCurrentItotal,
    zsat: round(zsat, 2),
    comment
  };
}

function calculateSoilSummary(source) {
  const locations = getSoilLocations(source).map((location, index) => {
    const direction1Average = average((location.direction1 || []).map((row) => toNumber(row.resistivity)));
    const direction2Average = average((location.direction2 || []).map((row) => toNumber(row.resistivity)));
    const overallAverage = average([direction1Average, direction2Average].filter((value) => Number.isFinite(value)));
    return {
      locationId: location.locationId || `soil-location-${index + 1}`,
      name: safeText(location.name, `Location ${index + 1}`),
      direction1Average: round(direction1Average, 2),
      direction2Average: round(direction2Average, 2),
      overallAverage: round(overallAverage, 2),
      category: getSoilCategory(overallAverage),
      drivenElectrodeDiameter: safeText(location.drivenElectrodeDiameter, ''),
      drivenElectrodeLength: safeText(location.drivenElectrodeLength, ''),
      notes: safeText(location.notes, '')
    };
  });

  const direction1Average = average(locations.map((location) => location.direction1Average).filter((value) => Number.isFinite(value)));
  const direction2Average = average(locations.map((location) => location.direction2Average).filter((value) => Number.isFinite(value)));
  const overallAverage = average(locations.map((location) => location.overallAverage).filter((value) => Number.isFinite(value)));
  return {
    direction1Average: round(direction1Average, 2),
    direction2Average: round(direction2Average, 2),
    overallAverage: round(overallAverage, 2),
    category: getSoilCategory(overallAverage),
    locations
  };
}

function selectedTests(source) {
  return state.catalog.filter((test) => Boolean(source.tests[test.id]));
}

function selectedTestCount(source) {
  return selectedTests(source).length;
}

function countActionRows(source) {
  const towerSummaries = summarizeTowerGroups(source.towerFootingResistance || []);
  const electrodeCritical = (source.electrodeResistance || []).filter((row) => getElectrodeStatus(getElectrodeMeasuredValue(row)).tone === 'critical').length;
  const continuityAttention = (source.continuityTest || []).filter((row) => !['healthy', 'neutral'].includes(deriveContinuityAssessment(row).status.tone)).length;
  const loopAttention = (source.loopImpedanceTest || []).filter((row) => !['healthy', 'neutral'].includes(deriveLoopAssessment(row).status.tone)).length;
  const faultAttention = (source.prospectiveFaultCurrent || []).filter((row) => deriveFaultAssessment(row).status.tone !== 'healthy' && deriveFaultAssessment(row).status.tone !== 'neutral').length;
  const riserAttention = (source.riserIntegrityTest || []).filter((row) => !['healthy', 'neutral'].includes(deriveRiserAssessment(row).status.tone)).length;
  const earthContinuityAttention = (source.earthContinuityTest || []).filter((row) => !['healthy', 'neutral'].includes(deriveEarthContinuityAssessment(row).status.tone)).length;
  const towerGroups = Array.isArray(source.towerFootingResistance) ? source.towerFootingResistance : [];
  const towerFootingAttention = towerGroups.filter((group) => {
    return !['healthy', 'neutral'].includes(deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group))).status.tone);
  }).length;
  const criticalTotal =
    electrodeCritical +
    (source.continuityTest || []).filter((row) => deriveContinuityAssessment(row).status.tone === 'critical').length +
    (source.loopImpedanceTest || []).filter((row) => deriveLoopAssessment(row).status.tone === 'critical').length +
    (source.prospectiveFaultCurrent || []).filter((row) => deriveFaultAssessment(row).status.tone === 'critical').length +
    (source.riserIntegrityTest || []).filter((row) => deriveRiserAssessment(row).status.tone === 'critical').length +
    (source.earthContinuityTest || []).filter((row) => deriveEarthContinuityAssessment(row).status.tone === 'critical').length +
    towerGroups.filter((group) => deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group))).status.tone === 'critical').length;

  return {
    electrodeCritical,
    continuityAttention,
    loopAttention,
    faultAttention,
    riserAttention,
    earthContinuityAttention,
    towerFootingAttention,
    criticalTotal,
    total:
      electrodeCritical +
      continuityAttention +
      loopAttention +
      faultAttention +
      riserAttention +
      earthContinuityAttention +
      towerFootingAttention
  };
}

function summarizeStatusRows(rows, deriveFn) {
  const statuses = (Array.isArray(rows) ? rows : [])
    .map((row) => deriveFn(row).status)
    .filter((status) => status && status.tone !== 'neutral');

  const healthy = statuses.filter((status) => status.tone === 'healthy').length;
  const warning = statuses.filter((status) => status.tone === 'warning').length;
  const critical = statuses.filter((status) => status.tone === 'critical').length;

  return {
    tested: statuses.length,
    healthy,
    warning,
    critical,
    attention: warning + critical
  };
}

function toneFromSummary(summary) {
  if (summary.critical > 0) {
    return 'critical';
  }
  if (summary.attention > 0) {
    return 'warning';
  }
  if (summary.healthy > 0) {
    return 'healthy';
  }
  return 'neutral';
}

function buildModuleInsights(source, soilSummary = calculateSoilSummary(source)) {
  const selected = selectedTests(source);
  const towerGroups = Array.isArray(source.towerFootingResistance) ? source.towerFootingResistance : [];
  const towerSummaries = summarizeTowerGroups(towerGroups);

  return selected.map((test) => {
    if (test.id === 'soilResistivity') {
      const locationCount = soilSummary.locations.length;
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: soilSummary.category.tone,
        badge: soilSummary.category.label,
        detail:
          soilSummary.overallAverage === null
            ? `Awaiting numeric readings across ${locationCount} soil location${locationCount === 1 ? '' : 's'}.`
            : `${soilSummary.overallAverage} ohm-m mean across ${locationCount} location${locationCount === 1 ? '' : 's'}.`,
        complete: soilSummary.overallAverage !== null
      };
    }

    if (test.id === 'electrodeResistance') {
      const summary = summarizeStatusRows(source.electrodeResistance, (row) => deriveElectrodeAssessment(row));
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: toneFromSummary(summary),
        badge: summary.critical ? `${summary.critical} critical` : summary.healthy ? `${summary.healthy} healthy` : 'Pending',
        detail: summary.tested ? `${summary.healthy}/${summary.tested} within limit • ${summary.attention} need review.` : 'Awaiting with-grid resistance values.',
        complete: summary.tested > 0
      };
    }

    if (test.id === 'continuityTest') {
      const summary = summarizeStatusRows(source.continuityTest, (row) => deriveContinuityAssessment(row));
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: toneFromSummary(summary),
        badge: summary.critical ? `${summary.critical} critical` : summary.attention ? `${summary.attention} review` : summary.healthy ? `${summary.healthy} healthy` : 'Pending',
        detail: summary.tested ? `${summary.tested} points checked • ${summary.attention} above reference.` : 'Awaiting continuity resistance readings.',
        complete: summary.tested > 0
      };
    }

    if (test.id === 'loopImpedanceTest') {
      const summary = summarizeStatusRows(source.loopImpedanceTest, (row) => deriveLoopAssessment(row));
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: toneFromSummary(summary),
        badge: summary.critical ? `${summary.critical} critical` : summary.attention ? `${summary.attention} review` : summary.healthy ? `${summary.healthy} healthy` : 'Pending',
        detail: summary.tested ? `${summary.tested} Zs points assessed • ${summary.attention} elevated.` : 'Awaiting Zs entries for protective loop review.',
        complete: summary.tested > 0
      };
    }

    if (test.id === 'prospectiveFaultCurrent') {
      const summary = summarizeStatusRows(source.prospectiveFaultCurrent, (row) => deriveFaultAssessment(row));
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: toneFromSummary(summary),
        badge: summary.critical ? `${summary.critical} critical` : summary.attention ? `${summary.attention} review` : summary.healthy ? `${summary.healthy} healthy` : 'Pending',
        detail: summary.tested ? `${summary.tested} fault levels compared • ${summary.attention} near or above breaking capacity.` : 'Awaiting PFC and breaking-capacity entries.',
        complete: summary.tested > 0
      };
    }

    if (test.id === 'riserIntegrityTest') {
      const summary = summarizeStatusRows(source.riserIntegrityTest, (row) => deriveRiserAssessment(row));
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: toneFromSummary(summary),
        badge: summary.critical ? `${summary.critical} critical` : summary.attention ? `${summary.attention} review` : summary.healthy ? `${summary.healthy} healthy` : 'Pending',
        detail: summary.tested ? `${summary.tested} riser paths assessed • ${summary.attention} need follow-up.` : 'Awaiting riser resistance inputs.',
        complete: summary.tested > 0
      };
    }

    if (test.id === 'earthContinuityTest') {
      const summary = summarizeStatusRows(source.earthContinuityTest, (row) => deriveEarthContinuityAssessment(row));
      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: toneFromSummary(summary),
        badge: summary.critical ? `${summary.critical} critical` : summary.attention ? `${summary.attention} review` : summary.healthy ? `${summary.healthy} healthy` : 'Pending',
        detail: summary.tested ? `${summary.tested} earth points checked • ${summary.attention} above threshold.` : 'Awaiting earth continuity values.',
        complete: summary.tested > 0
      };
    }

    if (test.id === 'towerFootingResistance') {
      const assessments = towerGroups.map((group) => deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group))));
      const completeGroups = assessments.filter((assessment) => Number.isFinite(assessment.totalImpedanceZt) && Number.isFinite(assessment.totalCurrentItotal)).length;
      const healthy = assessments.filter((assessment) => assessment.status.tone === 'healthy').length;
      const warning = assessments.filter((assessment) => assessment.status.tone === 'warning').length;
      const critical = assessments.filter((assessment) => assessment.status.tone === 'critical').length;

      return {
        id: test.id,
        label: test.label,
        shortLabel: test.shortLabel || test.label,
        tone: critical ? 'critical' : warning ? 'warning' : healthy ? 'healthy' : 'neutral',
        badge: critical ? `${critical} critical` : warning ? `${warning} marginal` : healthy ? `${healthy} healthy` : 'Pending',
        detail: towerGroups.length ? `${completeGroups}/${towerGroups.length} tower groups calculated.` : 'Awaiting tower location measurements.',
        complete: completeGroups > 0,
        groupCount: towerGroups.length,
        completeGroups
      };
    }

    return {
      id: test.id,
      label: test.label,
      shortLabel: test.shortLabel || test.label,
      tone: 'neutral',
      badge: 'Pending',
      detail: 'Awaiting measurement values.',
      complete: false
    };
  });
}

function measurementCoverageFromModules(modules) {
  const total = Array.isArray(modules) ? modules.length : 0;
  const completed = (Array.isArray(modules) ? modules : []).filter((module) => module.complete).length;
  return {
    total,
    completed,
    score: total ? Math.round((completed / total) * 100) : 0
  };
}

function buildAssessmentSummary(source) {
  const soil = calculateSoilSummary(source);
  const actions = countActionRows(source);
  const selected = selectedTests(source);
  const modules = buildModuleInsights(source, soil);
  const measurementCoverage = measurementCoverageFromModules(modules);
  const criticalSignals = actions.criticalTotal + (soil.category.tone === 'critical' ? 1 : 0);
  const warningSignals = actions.total - actions.criticalTotal + (soil.category.tone === 'warning' ? 1 : 0);
  const healthySignals = [
    source.tests.soilResistivity && soil.category.tone === 'healthy' ? 1 : 0,
    source.tests.electrodeResistance && actions.electrodeCritical === 0 && source.electrodeResistance.length ? 1 : 0,
    source.tests.continuityTest && actions.continuityAttention === 0 && source.continuityTest.length ? 1 : 0,
    source.tests.loopImpedanceTest && actions.loopAttention === 0 && source.loopImpedanceTest.length ? 1 : 0,
    source.tests.prospectiveFaultCurrent && actions.faultAttention === 0 && source.prospectiveFaultCurrent.length ? 1 : 0,
    source.tests.riserIntegrityTest && actions.riserAttention === 0 && source.riserIntegrityTest.length ? 1 : 0,
    source.tests.earthContinuityTest && actions.earthContinuityAttention === 0 && source.earthContinuityTest.length ? 1 : 0,
    source.tests.towerFootingResistance && actions.towerFootingAttention === 0 && source.towerFootingResistance.length ? 1 : 0
  ].reduce((sum, value) => sum + value, 0);

  let tone = 'healthy';
  let label = 'Healthy Snapshot';
  if (criticalSignals > 0) {
    tone = 'critical';
    label = 'Action Required';
  } else if (warningSignals > 0 || soil.category.tone === 'warning') {
    tone = 'warning';
    label = 'Review Needed';
  } else if (!selected.length) {
    tone = 'neutral';
    label = 'Awaiting Selection';
  }

  return {
    soil,
    actions,
    selected,
    modules,
    measurementCoverage,
    healthySignals,
    tone,
    label
  };
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

function builderReadiness(assessment = buildAssessmentSummary(state.draft)) {
  const project = state.draft.project;
  const projectComplete = ['projectNo', 'clientName', 'siteLocation', 'workOrder', 'reportDate', 'engineerName'].filter((field) => safeText(project[field], '')).length;
  const projectScore = Math.round((projectComplete / 6) * 100);
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
    soilResistivity: 'Enter measured resistivity values only. The app calculates direction averages and the final category.',
    electrodeResistance: 'Each electrode status updates from the 4.60 ohm limit automatically.',
    continuityTest: 'Use this for point-to-point resistance and impedance checks.',
    loopImpedanceTest: 'Record Zs values for protection verification at the tested points.',
    prospectiveFaultCurrent: 'Capture breaker and feeder context with the measured fault current values.',
    riserIntegrityTest: 'Use this for resistance checks towards equipment and towards the grid.',
    earthContinuityTest: 'Record tagged earth continuity points with distance and measured value.',
    towerFootingResistance: 'Each tower location keeps 4 fixed footing rows with grouped totals and a fixed Zsat 10 standard.'
  };
  return map[stepId] || 'Continue entering the field measurements for the selected scope.';
}

function currentStepNote(stepId) {
  const map = {
    project: 'Keep the project references complete so saved reports and PDFs stay traceable.',
    selection: 'Choose only the measurement sheets actually executed on site.',
    soilResistivity: 'Complete both directions so the soil mean and category can calculate.',
    electrodeResistance: 'Focus on the with-grid value because that drives the health result.',
    continuityTest: 'Resistance is the main decision field; impedance is supporting context.',
    loopImpedanceTest: 'Use measured Zs to flag circuits that need protection review.',
    prospectiveFaultCurrent: 'Breaking capacity and PFC must both be filled for a real check.',
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
          const references = TEST_STANDARD_REFERENCES[test.id] || [];
          return `
            <article class="standard-entry">
              <strong>${escapeHtml(test.shortLabel || test.label)}</strong>
              <span>${escapeHtml(references.join(' • ') || 'Reference standard to be confirmed')}</span>
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

  if (!response.ok) {
    const data = await response.json().catch(() => ({ message: 'Request failed.' }));
    throw new Error(data.message || 'Request failed.');
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
    throw new Error(data.message || 'Request failed.');
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
    state.draft.electrodeResistance = rows;
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

function applyZohoProject(projectId) {
  const project = state.zoho.projects.find((item) => String(item.id) === String(projectId));
  if (!project) {
    state.draft.project.zohoProjectId = '';
    state.draft.project.zohoProjectName = '';
    state.draft.project.zohoProjectOwner = '';
    state.draft.project.zohoProjectStage = '';
    state.zoho.users = [];
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
  if (!safeText(state.draft.project.engineerName, '')) {
    state.draft.project.engineerName = safeText(project.ownerName, '');
  }
  loadZohoUsers(project.id);
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
          <span>Need help?</span>
          <strong>info@elegrow.com</strong>
        </div>
      </div>
    </header>
  `;
}

function dashboardView() {
  const uniqueClients = new Set(state.reports.map((report) => report.project.clientName)).size;
  const reportsNeedingAction = state.reports.filter((report) => buildAssessmentSummary(report).tone !== 'healthy').length;
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
                  const summary = buildAssessmentSummary(report);
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
                        <span>Critical electrodes: <strong>${summary.actions.electrodeCritical}</strong></span>
                        <span>Total action points: <strong>${summary.actions.total}</strong></span>
                      </div>
                      <div class="card-actions">
                        <button class="button button-secondary" data-action="open-report" data-id="${report.id}">View Details</button>
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
  return `
    <label class="field">
      <span>${escapeHtml(label)}</span>
      <input
        type="${type}"
        class="text-input"
        placeholder="${escapeHtml(placeholder)}"
        value="${escapeHtml(value || '')}"
        data-bind="${bind}"
      />
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
              <button class="button button-secondary" data-action="use-zoho-project" data-project-id="${escapeHtml(project.id)}">Use In ElectroReports</button>
            </div>
          </article>
        `;
      }).join('')}
    </div>
  `;
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
      <div class="zoho-sync-panel">
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
      ${
        state.draft.project.zohoProjectId
          ? `
            <div class="zoho-selected-bar">
              <div class="zoho-selected-primary">
                <span>Project Number</span>
                <strong>${escapeHtml(state.draft.project.projectNo || state.draft.project.zohoProjectName || '-')}</strong>
                ${
                  state.draft.project.zohoProjectName &&
                  state.draft.project.zohoProjectName !== state.draft.project.projectNo
                    ? `<small>${escapeHtml(state.draft.project.zohoProjectName)}</small>`
                    : ''
                }
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
      <div class="field-grid">
        ${renderField('Project Number', state.draft.project.projectNo, 'project.projectNo', 'text', 'PRJ-2026-001')}
        ${renderField('Client Name', state.draft.project.clientName, 'project.clientName', 'text', 'Client Name')}
        ${renderField('Site Location', state.draft.project.siteLocation, 'project.siteLocation', 'text', 'Main Plant, Surat')}
        ${renderField('Work Order / Ref', state.draft.project.workOrder, 'project.workOrder', 'text', 'WO-628')}
        ${renderField('Date of Testing', state.draft.project.reportDate, 'project.reportDate', 'date')}
        ${renderField('Testing Engineer', state.draft.project.engineerName, 'project.engineerName', 'text', 'Engineer Name')}
      </div>
    `
  );
}

function renderSelectionStep() {
  const selectedEquipment = Array.isArray(state.draft.project.equipmentSelections) ? state.draft.project.equipmentSelections : [];
  return sectionCard(
    'Select Measurement Sections',
    'Step 2',
    `
      <div class="selection-grid">
        ${state.catalog
          .map((test) => {
            const active = state.draft.tests[test.id];
            return `
              <button
                class="selection-tile ${active ? 'selection-tile-active' : ''}"
                type="button"
                data-action="toggle-test"
                data-test="${test.id}"
              >
                <div>
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
          <h4>Select Instruments Used On Site</h4>
          <p>These selections will flow into the Equipment Required section of the final PDF methodology chapter.</p>
          <div class="tag-cloud">
            ${EQUIPMENT_LIBRARY.map((equipment) => {
              const active = selectedEquipment.includes(equipment.id);
              return `<button class="soft-chip soft-chip-button${active ? ' soft-chip-button-active' : ''}" type="button" data-action="toggle-equipment" data-equipment-id="${equipment.id}">${escapeHtml(equipment.label)}</button>`;
            }).join('')}
          </div>
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

function soilRowHtml(direction, row, index, locationIndex) {
  return `
    <tr>
      <td><input class="table-input" value="${escapeHtml(row.spacing)}" data-section="soilResistivity" data-location-index="${locationIndex}" data-direction="${direction}" data-index="${index}" data-field="spacing" /></td>
      <td><input class="table-input" value="${escapeHtml(row.resistivity)}" data-section="soilResistivity" data-location-index="${locationIndex}" data-direction="${direction}" data-index="${index}" data-field="resistivity" /></td>
      ${actionButtonsCell('soilResistivity', index, row, direction, `data-location-index="${locationIndex}"`)}
    </tr>
  `;
}

function tableSurface(title, subtitle, header, body, footer = '') {
  return `
    <div class="mini-surface">
      <div class="mini-surface-head">
        <div>
          <h4>${escapeHtml(title)}</h4>
          <p>${escapeHtml(subtitle)}</p>
        </div>
      </div>
      <div class="table-wrap">
        <table class="data-table">
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

function renderSoilStep() {
  const summary = calculateSoilSummary(state.draft);
  const locations = getSoilLocations(state.draft);
  return sectionCard(
    'Soil Resistivity Test',
    'Measurement Sheet',
    `
      ${renderSoilOcrAssist()}
      <div class="metric-strip">
        <article class="metric-box">
          <span>Direction 1 Average</span>
          <strong>${summary.direction1Average === null ? '-' : `${summary.direction1Average} ohm-m`}</strong>
        </article>
        <article class="metric-box">
          <span>Direction 2 Average</span>
          <strong>${summary.direction2Average === null ? '-' : `${summary.direction2Average} ohm-m`}</strong>
        </article>
        <article class="metric-box">
          <span>Mean Soil Resistivity</span>
          <strong>${summary.overallAverage === null ? '-' : `${summary.overallAverage} ohm-m`}</strong>
          ${pill(summary.category.tone, summary.category.label)}
        </article>
      </div>
      <div class="stack-grid">
        ${locations
          .map((location, locationIndex) => {
            const locationSummary = summary.locations[locationIndex] || {};
            return `
              <section class="mini-surface mini-surface-location">
                <div class="mini-surface-head mini-surface-head-location">
                  <div>
                    <h4>${escapeHtml(safeText(location.name, `Location ${locationIndex + 1}`))}</h4>
                    <p>Record both directions separately for this soil test location.</p>
                  </div>
                  <div class="mini-surface-head-actions">
                    ${pill(locationSummary.category?.tone || 'neutral', locationSummary.category?.label || 'Insufficient Data')}
                    ${
                      locations.length > 1
                        ? `<button class="button button-ghost" type="button" data-action="remove-soil-location" data-location-index="${locationIndex}">Remove Location</button>`
                        : ''
                    }
                  </div>
                </div>
                <div class="field-grid">
                  ${renderField('Location Name', location.name, `soilResistivity.locations.${locationIndex}.name`, 'text', `Location ${locationIndex + 1}`)}
                  ${renderField(
                    'Driven Earth Electrode - Diameter (mm)',
                    location.drivenElectrodeDiameter,
                    `soilResistivity.locations.${locationIndex}.drivenElectrodeDiameter`,
                    'number',
                    '40'
                  )}
                  ${renderField(
                    'Driven Earth Electrode - Length (m)',
                    location.drivenElectrodeLength,
                    `soilResistivity.locations.${locationIndex}.drivenElectrodeLength`,
                    'number',
                    '3'
                  )}
                </div>
                <div class="metric-strip">
                  <article class="metric-box">
                    <span>Direction 1 Average</span>
                    <strong>${locationSummary.direction1Average === null || locationSummary.direction1Average === undefined ? '-' : `${locationSummary.direction1Average} ohm-m`}</strong>
                  </article>
                  <article class="metric-box">
                    <span>Direction 2 Average</span>
                    <strong>${locationSummary.direction2Average === null || locationSummary.direction2Average === undefined ? '-' : `${locationSummary.direction2Average} ohm-m`}</strong>
                  </article>
                  <article class="metric-box">
                    <span>Location Mean</span>
                    <strong>${locationSummary.overallAverage === null || locationSummary.overallAverage === undefined ? '-' : `${locationSummary.overallAverage} ohm-m`}</strong>
                  </article>
                </div>
                <div class="split-grid">
                  ${tableSurface(
                    'Direction 1',
                    'Keep the measured columns exactly as per the sheet.',
                    '<tr><th>Spacing of Probes (m)</th><th>Resistivity rho (ohm-m)</th><th>Action</th></tr>',
                    location.direction1.map((row, index) => soilRowHtml('direction1', row, index, locationIndex)).join(''),
                    `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="soilResistivity" data-location-index="${locationIndex}" data-direction="direction1">Add Measurement Row</button></div>`
                  )}
                  ${tableSurface(
                    'Direction 2',
                    'Second direction for the mean calculation.',
                    '<tr><th>Spacing of Probes (m)</th><th>Resistivity rho (ohm-m)</th><th>Action</th></tr>',
                    location.direction2.map((row, index) => soilRowHtml('direction2', row, index, locationIndex)).join(''),
                    `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="soilResistivity" data-location-index="${locationIndex}" data-direction="direction2">Add Measurement Row</button></div>`
                  )}
                </div>
                <label class="field">
                  <span>Site Notes</span>
                  <textarea class="text-area" rows="3" data-bind="soilResistivity.locations.${locationIndex}.notes" placeholder="Ground condition, nearby structures, or observations">${escapeHtml(location.notes)}</textarea>
                </label>
              </section>
            `;
          })
          .join('')}
      </div>
      <div class="mini-surface-foot">
        <button class="button button-secondary" type="button" data-action="add-soil-location">Add Another Location</button>
      </div>
    `
  );
}

function electrodeRowHtml(row, index) {
  const assessment = deriveElectrodeAssessment(row);
  const status = assessment.status;
  return `
    <tr>
      <td><input class="table-input" value="${escapeHtml(row.tag)}" data-section="electrodeResistance" data-index="${index}" data-field="tag" /></td>
      <td><input class="table-input" value="${escapeHtml(row.location)}" data-section="electrodeResistance" data-index="${index}" data-field="location" /></td>
      <td>
        <select class="table-input" data-section="electrodeResistance" data-index="${index}" data-field="electrodeType">
          ${['Rod', 'Pipe', 'Plate', 'Strip'].map((option) => `<option ${row.electrodeType === option ? 'selected' : ''}>${option}</option>`).join('')}
        </select>
      </td>
      <td>
        <select class="table-input" data-section="electrodeResistance" data-index="${index}" data-field="materialType">
          ${['Copper', 'GI', 'Copper Bonded'].map((option) => `<option ${row.materialType === option ? 'selected' : ''}>${option}</option>`).join('')}
        </select>
      </td>
      <td><input class="table-input" value="${escapeHtml(row.length)}" data-section="electrodeResistance" data-index="${index}" data-field="length" /></td>
      <td><input class="table-input" value="${escapeHtml(row.diameter)}" data-section="electrodeResistance" data-index="${index}" data-field="diameter" /></td>
      <td><input class="table-input" value="${escapeHtml(safeText(row.resistanceWithoutGrid, ''))}" data-section="electrodeResistance" data-index="${index}" data-field="resistanceWithoutGrid" /></td>
      <td><input class="table-input" value="${escapeHtml(safeText(row.resistanceWithGrid || row.measuredResistance, ''))}" data-section="electrodeResistance" data-index="${index}" data-field="resistanceWithGrid" /></td>
      <td>${standardCell(assessment.standard.reference, assessment.standard.limitLabel)}</td>
      <td>${pill(status.tone, status.label)}</td>
      <td>${autoTextCell(assessment.comment)}</td>
      ${actionButtonsCell('electrodeResistance', index, row)}
    </tr>
  `;
}

function renderElectrodeStep() {
  return sectionCard(
    'Earth Electrode Resistance Test',
    'Measurement Sheet',
    `
      ${renderElectrodeOcrAssist()}
      <div class="limit-banner">Permissible limit: 4.60 ohm as per IS 3043</div>
      ${tableSurface(
        'Electrode Measurements',
        'Status is calculated automatically from the resistance value with grid.',
        '<tr><th>Pit Tag</th><th>Location</th><th>Electrode Type</th><th>Material</th><th>Length (m)</th><th>Dia (mm)</th><th>Resistance Value Without Grid (ohm)</th><th>Resistance Value With Grid (ohm)</th><th>Standard</th><th>Status</th><th>Comment / Observation</th><th>Action</th></tr>',
        state.draft.electrodeResistance.map((row, index) => electrodeRowHtml(row, index)).join(''),
        '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="electrodeResistance">Add Electrode</button></div>'
      )}
    `
  );
}

function towerReadingActionCell(groupIndex, readingIndex, reading, rowSpan = 1) {
  const hasObservation = rowHasObservationData(reading);
  return `
    <td class="table-action-cell tower-group-action-cell" ${rowSpan > 1 ? `rowspan="${rowSpan}"` : ''}>
      <div class="table-action-buttons">
        <button
          class="icon-button icon-button-observation ${hasObservation ? 'icon-button-observation-active' : ''}"
          type="button"
          data-action="open-row-observation"
          data-section="towerFootingResistance"
          data-index="${readingIndex}"
          data-group-index="${groupIndex}"
          aria-label="${hasObservation ? 'Edit row observation' : 'Add row observation'}"
          title="${hasObservation ? 'Edit row observation' : 'Add row observation'}"
        >
          <span class="icon-button-symbol">+</span>
        </button>
      </div>
    </td>
  `;
}

function towerGroupRowsHtml(group, groupIndex, summaries) {
  const assessment = deriveTowerFootingAssessment(group, summaries.get(buildTowerGroupKey(group)));
  const readings = Array.isArray(group.readings) ? group.readings : [];

  return readings
    .map((reading, readingIndex) => {
      const rowClass = [
        readingIndex === 0 ? 'tower-group-row-start' : '',
        readingIndex === readings.length - 1 ? 'tower-group-row-end' : ''
      ]
        .filter(Boolean)
        .join(' ');

      const sharedCells =
        readingIndex === 0
          ? `
              <td class="tower-group-cell tower-group-number" rowspan="${readings.length}">${escapeHtml(group.srNo)}</td>
              <td class="tower-group-cell tower-group-main" rowspan="${readings.length}">
                <input
                  class="table-input"
                  value="${escapeHtml(group.mainLocationTower)}"
                  data-section="towerFootingResistance"
                  data-index="${groupIndex}"
                  data-field="mainLocationTower"
                />
                ${
                  state.draft.towerFootingResistance.length > 1
                    ? `<button class="button button-ghost tower-group-remove" type="button" data-action="remove-tower-group" data-group-index="${groupIndex}">Remove Tower</button>`
                    : ''
                }
              </td>
            `
          : '';

      const sharedSummaryCells =
        readingIndex === 0
          ? `
              <td class="tower-group-cell tower-group-merged" rowspan="${readings.length}">${autoTextCell(assessment.totalImpedanceZt === null ? '-' : String(assessment.totalImpedanceZt))}</td>
              <td class="tower-group-cell tower-group-merged" rowspan="${readings.length}">${autoTextCell(assessment.totalCurrentItotal === null ? '-' : String(assessment.totalCurrentItotal))}</td>
              <td class="tower-group-cell tower-group-merged" rowspan="${readings.length}">${autoTextCell('10')}</td>
            `
          : '';

      const sharedRemarkCell =
        readingIndex === 0
          ? `
              <td class="tower-group-cell tower-group-merged" rowspan="${readings.length}">
                ${autoTextCell(assessment.comment)}
              </td>
            `
          : '';

      return `
        <tr class="${rowClass}">
          ${sharedCells}
          <td class="tower-foot-label">${escapeHtml(reading.measurementPointLocation || TOWER_FOOT_POINTS[readingIndex])}</td>
          <td>
            <select
              class="table-input"
              data-section="towerFootingResistance"
              data-index="${groupIndex}"
              data-reading-index="${readingIndex}"
              data-field="footToEarthingConnectionStatus"
            >
              ${['Given', 'Not Given']
                .map((option) => `<option ${safeText(reading.footToEarthingConnectionStatus, 'Given') === option ? 'selected' : ''}>${option}</option>`)
                .join('')}
            </select>
          </td>
          <td><input class="table-input" value="${escapeHtml(reading.measuredCurrentMa)}" data-section="towerFootingResistance" data-index="${groupIndex}" data-reading-index="${readingIndex}" data-field="measuredCurrentMa" /></td>
          <td><input class="table-input" value="${escapeHtml(reading.measuredImpedance)}" data-section="towerFootingResistance" data-index="${groupIndex}" data-reading-index="${readingIndex}" data-field="measuredImpedance" /></td>
          ${sharedSummaryCells}
          ${sharedRemarkCell}
          ${towerReadingActionCell(groupIndex, readingIndex, reading)}
        </tr>
      `;
    })
    .join('');
}

function renderTowerFootingStep() {
  const summaries = summarizeTowerGroups(state.draft.towerFootingResistance);
  return sectionCard(
    'Tower Footing Resistance Measurement & Analysis',
    'Measurement Sheet',
    `
      ${renderTowerOcrAssist()}
      ${tableSurface(
        'Tower Footing Resistance Measurement & Analysis',
        'Each tower location contains 4 fixed footing rows: Foot-1, Foot-2, Foot-3, and Foot-4.',
        '<tr><th>Sr. No.</th><th>Main Location – Tower</th><th>Measurement Point Location</th><th>Foot to Earthing Connection Status</th><th>Measured Current I (mA)</th><th>Measured Impedance (ohm)</th><th>Total Impedance Zt (ohm)</th><th>Total Current | Total (A)</th><th>Standard Tolerable Impedance Value Zsat</th><th>Remarks</th><th>Action</th></tr>',
        state.draft.towerFootingResistance.map((group, groupIndex) => towerGroupRowsHtml(group, groupIndex, summaries)).join(''),
        '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="towerFootingResistance">Add New Tower Location</button></div>'
      )}
    `
  );
}

function phaseGroupInputCell(section, index, field, value, rowSpan = 3, type = 'text', options = []) {
  if (type === 'select') {
    return `
      <td class="phase-group-cell" rowspan="${rowSpan}">
        <select class="table-input" data-section="${section}" data-index="${index}" data-field="${field}" data-group-sync="true">
          ${options.map((option) => `<option ${value === option ? 'selected' : ''}>${option}</option>`).join('')}
        </select>
      </td>
    `;
  }
  return `<td class="phase-group-cell" rowspan="${rowSpan}"><input class="table-input" value="${escapeHtml(value)}" data-section="${section}" data-index="${index}" data-field="${field}" data-group-sync="true" /></td>`;
}

function renderLoopGroupRows(groupRows, groupIndex) {
  const baseIndex = groupIndex * PHASE_MEASURED_POINTS.length;
  const first = groupRows[0];
  return groupRows
    .map((row, offset) => {
      const index = baseIndex + offset;
      const assessment = deriveLoopAssessment(row);
      return `
        <tr>
          <td>${escapeHtml(row.srNo)}</td>
          ${
            offset === 0
              ? `
                ${phaseGroupInputCell('loopImpedanceTest', index, 'location', first.location)}
                ${phaseGroupInputCell('loopImpedanceTest', index, 'feederTag', first.feederTag)}
                ${phaseGroupInputCell('loopImpedanceTest', index, 'deviceType', first.deviceType, 3, 'select', ['ACB', 'MCCB', 'MCB', 'SFU'])}
                ${phaseGroupInputCell('loopImpedanceTest', index, 'deviceRating', first.deviceRating)}
                ${phaseGroupInputCell('loopImpedanceTest', index, 'breakingCapacity', first.breakingCapacity)}
              `
              : ''
          }
          <td class="table-input-static">${escapeHtml(row.measuredPoints)}</td>
          <td><input class="table-input" value="${escapeHtml(row.loopImpedance)}" data-section="loopImpedanceTest" data-index="${index}" data-field="loopImpedance" /></td>
          <td><input class="table-input" value="${escapeHtml(row.voltage)}" data-section="loopImpedanceTest" data-index="${index}" data-field="voltage" /></td>
          <td>${autoTextCell(assessment.comment)}</td>
          ${actionButtonsCell('loopImpedanceTest', index, row)}
        </tr>
      `;
    })
    .join('');
}

function renderLoopStep() {
  const groups = [];
  for (let index = 0; index < state.draft.loopImpedanceTest.length; index += PHASE_MEASURED_POINTS.length) {
    groups.push(state.draft.loopImpedanceTest.slice(index, index + PHASE_MEASURED_POINTS.length));
  }
  return sectionCard(
    'Loop Impedance Test',
    'Measurement Sheet',
    `
      ${renderLoopOcrAssist()}
      ${tableSurface(
        'Loop Impedance Test',
        'Add feeder groups as needed for the executed field scope.',
        `
          <tr>
            <th rowspan="2">Sr. No.</th>
            <th rowspan="2">Location of Panel / Equipment</th>
            <th rowspan="2">Name of Feeder & Tag No.</th>
            <th colspan="3">Protective Device Details</th>
            <th rowspan="2">Measured Points</th>
            <th rowspan="2">Loop Impedance (Z)</th>
            <th rowspan="2">Voltage</th>
            <th rowspan="2">Comment</th>
            <th rowspan="2">Action</th>
          </tr>
          <tr>
            <th>Type (ACB / MCCB / MCB / SFU)</th>
            <th>Rating (A)</th>
            <th>Breaking (kA)</th>
          </tr>
        `,
        groups.map((groupRows, groupIndex) => renderLoopGroupRows(groupRows, groupIndex)).join(''),
        '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="loopImpedanceTest">Add Feeder Group</button></div>'
      )}
    `
  );
}

function renderFaultGroupRows(groupRows, groupIndex) {
  const baseIndex = groupIndex * PHASE_MEASURED_POINTS.length;
  const first = groupRows[0];
  return groupRows
    .map((row, offset) => {
      const index = baseIndex + offset;
      const assessment = deriveFaultAssessment(row);
      return `
        <tr>
          <td>${escapeHtml(row.srNo)}</td>
          ${
            offset === 0
              ? `
                ${phaseGroupInputCell('prospectiveFaultCurrent', index, 'location', first.location)}
                ${phaseGroupInputCell('prospectiveFaultCurrent', index, 'feederTag', first.feederTag)}
                ${phaseGroupInputCell('prospectiveFaultCurrent', index, 'deviceType', first.deviceType, 3, 'select', ['ACB', 'MCCB', 'MCB', 'SFU'])}
                ${phaseGroupInputCell('prospectiveFaultCurrent', index, 'deviceRating', first.deviceRating)}
                ${phaseGroupInputCell('prospectiveFaultCurrent', index, 'breakingCapacity', first.breakingCapacity)}
              `
              : ''
          }
          <td class="table-input-static">${escapeHtml(row.measuredPoints)}</td>
          <td><input class="table-input" value="${escapeHtml(row.loopImpedance)}" data-section="prospectiveFaultCurrent" data-index="${index}" data-field="loopImpedance" /></td>
          <td><input class="table-input" value="${escapeHtml(row.prospectiveFaultCurrent)}" data-section="prospectiveFaultCurrent" data-index="${index}" data-field="prospectiveFaultCurrent" /></td>
          <td><input class="table-input" value="${escapeHtml(row.voltage)}" data-section="prospectiveFaultCurrent" data-index="${index}" data-field="voltage" /></td>
          <td>${autoTextCell(assessment.comment)}</td>
          ${actionButtonsCell('prospectiveFaultCurrent', index, row)}
        </tr>
      `;
    })
    .join('');
}

function renderFaultStep() {
  const groups = [];
  for (let index = 0; index < state.draft.prospectiveFaultCurrent.length; index += PHASE_MEASURED_POINTS.length) {
    groups.push(state.draft.prospectiveFaultCurrent.slice(index, index + PHASE_MEASURED_POINTS.length));
  }
  return sectionCard(
    'Prospective Fault Current',
    'Measurement Sheet',
    `
      ${renderFaultOcrAssist()}
      ${tableSurface(
        'Prospective Fault Current',
        'Add feeder groups as needed for the executed field scope.',
        `
          <tr>
            <th rowspan="2">Sr. No.</th>
            <th rowspan="2">Location of Panel / Equipment</th>
            <th rowspan="2">Name of Feeder & Tag No.</th>
            <th colspan="3">Protective Device Details</th>
            <th rowspan="2">Measured Points</th>
            <th rowspan="2">Loop Impedance (Z)</th>
            <th rowspan="2">Prospective Fault Current</th>
            <th rowspan="2">Voltage</th>
            <th rowspan="2">Comment</th>
            <th rowspan="2">Action</th>
          </tr>
          <tr>
            <th>Type (ACB / MCCB / MCB / SFU)</th>
            <th>Rating (A)</th>
            <th>Breaking (kA)</th>
          </tr>
        `,
        groups.map((groupRows, groupIndex) => renderFaultGroupRows(groupRows, groupIndex)).join(''),
        '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="prospectiveFaultCurrent">Add Feeder Group</button></div>'
      )}
    `
  );
}

function renderContinuityStep() {
  return sectionCard(
    'Continuity Test',
    'Measurement Sheet',
    `
      ${renderContinuityOcrAssist()}
      ${tableSurface(
        'Continuity Test',
        'Add rows as needed for the executed field scope.',
        `<tr><th>Sr. No.</th><th>Main Location</th><th>Measurement Point</th><th>Resistance (ohm)</th><th>Impedance (ohm)</th><th>Status</th><th>Comment</th><th>Action</th></tr>`,
        state.draft.continuityTest
          .map((row, index) => {
            return `
              <tr>
                <td><input class="table-input" value="${escapeHtml(row.srNo)}" data-section="continuityTest" data-index="${index}" data-field="srNo" /></td>
                <td><input class="table-input" value="${escapeHtml(row.mainLocation)}" data-section="continuityTest" data-index="${index}" data-field="mainLocation" /></td>
                <td><input class="table-input" value="${escapeHtml(row.measurementPoint)}" data-section="continuityTest" data-index="${index}" data-field="measurementPoint" /></td>
                <td><input class="table-input" value="${escapeHtml(row.resistance)}" data-section="continuityTest" data-index="${index}" data-field="resistance" /></td>
                <td><input class="table-input" value="${escapeHtml(row.impedance)}" data-section="continuityTest" data-index="${index}" data-field="impedance" /></td>
                <td>${pill(deriveContinuityAssessment(row).status.tone, deriveContinuityAssessment(row).status.label)}</td>
                <td>${autoTextCell(deriveContinuityAssessment(row).comment)}</td>
                ${actionButtonsCell('continuityTest', index, row)}
              </tr>
            `;
          })
          .join(''),
        `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="continuityTest">Add Row</button></div>`
      )}
    `
  );
}

function genericStep(title, subtitle, section, columns, statusRenderer, prefix = '') {
  return sectionCard(
    title,
    subtitle,
    `
      ${prefix}
      ${tableSurface(
        title,
        'Add rows as needed for the executed field scope.',
        `<tr>${columns.map((column) => `<th>${escapeHtml(column.label)}</th>`).join('')}<th>Action</th></tr>`,
        state.draft[section]
          .map((row, index) => {
            return `
              <tr>
                ${columns
                  .map((column) => {
                    if (column.type === 'select') {
                      return `
                        <td>
                          <select class="table-input" data-section="${section}" data-index="${index}" data-field="${column.key}">
                            ${column.options.map((option) => `<option ${row[column.key] === option ? 'selected' : ''}>${option}</option>`).join('')}
                          </select>
                        </td>
                      `;
                    }
                    if (column.render) {
                      return `<td>${column.render(row, index)}</td>`;
                    }
                    return `<td><input class="table-input" value="${escapeHtml(row[column.key])}" data-section="${section}" data-index="${index}" data-field="${column.key}" /></td>`;
                  })
                  .join('')}
                ${actionButtonsCell(section, index, row)}
              </tr>
            `;
          })
          .join(''),
        `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${section}">Add Row</button></div>`
      )}
    `
  );
}

function renderBuilderStep() {
  const step = currentStepId();
  if (step === 'project') {
    return renderProjectStep();
  }
  if (step === 'selection') {
    return renderSelectionStep();
  }
  if (step === 'soilResistivity') {
    return renderSoilStep();
  }
  if (step === 'electrodeResistance') {
    return renderElectrodeStep();
  }
  if (step === 'continuityTest') {
    return renderContinuityStep();
  }
  if (step === 'loopImpedanceTest') {
    return renderLoopStep();
  }
  if (step === 'prospectiveFaultCurrent') {
    return renderFaultStep();
  }
  if (step === 'riserIntegrityTest') {
    return genericStep('Riser / Grid Integrity Test', 'Measurement Sheet', 'riserIntegrityTest', [
      { key: 'srNo', label: 'Sr. No.' },
      { key: 'mainLocation', label: 'Main Location' },
      { key: 'measurementPoint', label: 'Measurement Point' },
      { key: 'resistanceTowardsEquipment', label: 'Towards Equipment' },
      { key: 'resistanceTowardsGrid', label: 'Towards Grid' },
      {
        key: 'status',
        label: 'Status',
        render: (row) => {
          const assessment = deriveRiserAssessment(row);
          return pill(assessment.status.tone, assessment.status.label);
        }
      },
      {
        key: 'comment',
        label: 'Comment',
        render: (row) => autoTextCell(deriveRiserAssessment(row).comment)
      }
    ], null, renderRiserOcrAssist());
  }
  if (step === 'earthContinuityTest') {
    return genericStep('Earth Continuity Test', 'Measurement Sheet', 'earthContinuityTest', [
      { key: 'srNo', label: 'Sr. No.' },
      { key: 'tag', label: 'Tag' },
      { key: 'locationBuildingName', label: 'Location / Building' },
      { key: 'distance', label: 'Distance' },
      { key: 'measuredValue', label: 'Measured Value' },
      {
        key: 'status',
        label: 'Status',
        render: (row) => {
          const assessment = deriveEarthContinuityAssessment(row);
          return pill(assessment.status.tone, assessment.status.label);
        }
      },
      {
        key: 'remark',
        label: 'Remark',
        render: (row) => autoTextCell(deriveEarthContinuityAssessment(row).comment)
      }
    ], null, renderEarthContinuityOcrAssist());
  }
  if (step === 'towerFootingResistance') {
    return renderTowerFootingStep();
  }
  return '';
}

function builderView() {
  const steps = getSteps();
  const assessment = buildAssessmentSummary(state.draft);
  const readiness = builderReadiness(assessment);
  const currentStep = currentStepId();
  return `
    <section class="surface builder-shell">
      <div class="builder-header">
        <div>
          <p class="section-kicker">Report Builder</p>
          <h2>New ElectroReports Project</h2>
          <p>Only the sections you select will appear in the report flow and in the final PDF.</p>
        </div>
        <div class="button-row">
          <button class="button button-secondary" data-action="go-dashboard">Cancel</button>
          <button class="button button-primary" data-action="save-report" ${state.saving ? 'disabled' : ''}>${state.saving ? 'Saving...' : 'Save Report'}</button>
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
              </button>
            `;
          })
          .join('')}
      </div>

      <div class="builder-layout">
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

function towerFootingDetailTable(groups) {
  const summaries = summarizeTowerGroups(groups);
  return `
    <div class="table-wrap">
      <table class="data-table tower-data-table">
        <thead>
          <tr>
            <th>Sr. No.</th>
            <th>Main Location – Tower</th>
            <th>Measurement Point Location</th>
            <th>Foot to Earthing Connection Status</th>
            <th>Measured Current I (mA)</th>
            <th>Measured Impedance (ohm)</th>
            <th>Total Impedance Zt (ohm)</th>
            <th>Total Current | Total (A)</th>
            <th>Standard Tolerable Impedance Value Zsat</th>
            <th>Remarks</th>
            <th>Observation</th>
          </tr>
        </thead>
        <tbody>
          ${(Array.isArray(groups) ? groups : [])
            .map((group) => {
              const assessment = deriveTowerFootingAssessment(group, summaries.get(buildTowerGroupKey(group)));
              const readings = Array.isArray(group.readings) ? group.readings : [];
              return readings
                .map((reading, readingIndex) => {
                  return `
                    <tr class="${readingIndex === 0 ? 'tower-group-row-start' : ''} ${readingIndex === readings.length - 1 ? 'tower-group-row-end' : ''}">
                      ${
                        readingIndex === 0
                          ? `<td class="tower-group-cell" rowspan="${readings.length}">${escapeHtml(group.srNo)}</td>
                             <td class="tower-group-cell" rowspan="${readings.length}">${escapeHtml(group.mainLocationTower)}</td>`
                          : ''
                      }
                      <td class="tower-foot-label">${escapeHtml(reading.measurementPointLocation)}</td>
                      <td>${escapeHtml(safeText(reading.footToEarthingConnectionStatus, 'Given'))}</td>
                      <td>${escapeHtml(safeText(reading.measuredCurrentMa, ''))}</td>
                      <td>${escapeHtml(safeText(reading.measuredImpedance, ''))}</td>
                      ${
                        readingIndex === 0
                          ? `<td class="tower-group-cell" rowspan="${readings.length}">${escapeHtml(
                              assessment.totalImpedanceZt === null ? '-' : String(assessment.totalImpedanceZt)
                            )}</td>
                             <td class="tower-group-cell" rowspan="${readings.length}">${escapeHtml(
                               assessment.totalCurrentItotal === null ? '-' : String(assessment.totalCurrentItotal)
                             )}</td>
                             <td class="tower-group-cell" rowspan="${readings.length}">10</td>`
                          : ''
                      }
                      ${
                        readingIndex === 0
                          ? `<td class="tower-group-cell" rowspan="${readings.length}">${autoTextCell(assessment.comment)}</td>`
                          : ''
                      }
                      <td>${renderObservationCard(reading, safeText(reading.measurementPointLocation, 'Foot')) || '-'}</td>
                    </tr>
                  `;
                })
                .join('');
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
  const assessment = buildAssessmentSummary(report);
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
          <button class="button button-secondary" data-action="generate-ai-summary" data-id="${report.id}" ${!state.ai.configured || state.ai.generating ? 'disabled' : ''}>${
            state.ai.generating ? 'Generating AI Summary...' : report.aiNarrative ? 'Refresh AI Summary' : 'Generate AI Summary'
          }</button>
          <button class="button button-success" data-action="export-pdf" data-id="${report.id}" ${state.exporting ? 'disabled' : ''}>${state.exporting ? 'Generating PDF...' : 'Export PDF'}</button>
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

      ${reportSectionBlock(
        'Project Details',
        `
          <div class="detail-grid">
            <div><span>Project Number</span><strong>${escapeHtml(report.project.projectNo)}</strong></div>
            <div><span>Client Name</span><strong>${escapeHtml(report.project.clientName)}</strong></div>
            <div><span>Site Location</span><strong>${escapeHtml(report.project.siteLocation)}</strong></div>
            <div><span>Work Order / Ref</span><strong>${escapeHtml(report.project.workOrder)}</strong></div>
            <div><span>Date of Testing</span><strong>${escapeHtml(report.project.reportDate)}</strong></div>
            <div><span>Testing Engineer</span><strong>${escapeHtml(report.project.engineerName)}</strong></div>
            ${report.project.zohoProjectName ? `<div><span>Zoho Project</span><strong>${escapeHtml(report.project.zohoProjectName)}</strong></div>` : ''}
            ${report.project.zohoProjectStage ? `<div><span>Zoho Stage</span><strong>${escapeHtml(report.project.zohoProjectStage)}</strong></div>` : ''}
          </div>
        `
      )}

      ${renderAiNarrativeSection(report)}

      ${report.tests.soilResistivity
        ? reportSectionBlock(
            'Soil Resistivity Test',
            `
              <div class="metric-strip">
                <article class="metric-box"><span>Direction 1 Average</span><strong>${soil.direction1Average === null ? '-' : `${soil.direction1Average} ohm-m`}</strong></article>
                <article class="metric-box"><span>Direction 2 Average</span><strong>${soil.direction2Average === null ? '-' : `${soil.direction2Average} ohm-m`}</strong></article>
                <article class="metric-box"><span>Classification</span><strong>${escapeHtml(soil.category.label)}</strong>${pill(soil.category.tone, soil.category.label)}</article>
              </div>
              ${soil.locations
                .map((location, locationIndex) => {
                  const sourceLocation = getSoilLocations(report)[locationIndex];
                  return `
                    <div class="detail-subsection">
                      <div class="detail-subsection-head">
                        <h4>${escapeHtml(location.name)}</h4>
                        ${pill(location.category.tone, location.category.label)}
                      </div>
                      <div class="detail-grid">
                        <div><span>Direction 1 Average</span><strong>${location.direction1Average === null ? '-' : `${location.direction1Average} ohm-m`}</strong></div>
                        <div><span>Direction 2 Average</span><strong>${location.direction2Average === null ? '-' : `${location.direction2Average} ohm-m`}</strong></div>
                        <div><span>Location Mean</span><strong>${location.overallAverage === null ? '-' : `${location.overallAverage} ohm-m`}</strong></div>
                        <div><span>Driven Electrode Diameter</span><strong>${escapeHtml(location.drivenElectrodeDiameter || '-')}</strong></div>
                        <div><span>Driven Electrode Length</span><strong>${escapeHtml(location.drivenElectrodeLength || '-')}</strong></div>
                      </div>
                      ${simpleDetailTable(
                        [
                          { label: 'Spacing of Probes (m)', key: 'spacing' },
                          { label: 'Direction 1 Resistivity (ohm-m)', key: 'direction1' },
                          { label: 'Direction 2 Resistivity (ohm-m)', key: 'direction2' }
                        ],
                        (sourceLocation?.direction1 || []).map((row, index) => ({
                          spacing: row.spacing,
                          direction1: row.resistivity,
                          direction2: sourceLocation?.direction2[index] ? sourceLocation.direction2[index].resistivity : '-'
                        })),
                        {
                          observationRenderer: (_row, index) => {
                            return [
                              renderObservationCard(sourceLocation?.direction1[index], `${location.name} | Direction 1 Observation`),
                              renderObservationCard(sourceLocation?.direction2[index], `${location.name} | Direction 2 Observation`)
                            ]
                              .filter(Boolean)
                              .join('');
                          }
                        }
                      )}
                      ${
                        sourceLocation?.notes
                          ? `<p class="detail-note"><strong>Site Notes:</strong> ${escapeHtml(sourceLocation.notes)}</p>`
                          : ''
                      }
                    </div>
                  `;
                })
                .join('')}
            `
          )
        : ''}

      ${report.tests.electrodeResistance
        ? reportSectionBlock(
            'Earth Electrode Resistance Test',
            simpleDetailTable(
              [
                { label: 'Pit Tag', key: 'tag' },
                { label: 'Location', key: 'location' },
                { label: 'Type', key: 'electrodeType' },
                { label: 'Material', key: 'materialType' },
                { label: 'Length (m)', key: 'length' },
                { label: 'Dia (mm)', key: 'diameter' },
                { label: 'Resistance Without Grid (ohm)', key: 'resistanceWithoutGrid' },
                { label: 'Resistance With Grid (ohm)', key: 'resistanceWithGrid' },
                {
                  label: 'Standard',
                  render: () => standardCell(STANDARD_GUIDANCE.electrodeResistance.reference, STANDARD_GUIDANCE.electrodeResistance.limitLabel)
                },
                {
                  label: 'Comment',
                  render: (row) => {
                    const assessment = deriveElectrodeAssessment(row);
                    return autoTextCell(toReportBandLabel(assessment.status));
                  }
                },
                { label: 'Observation', render: () => autoTextCell('') },
                { label: 'Recommendation', render: () => autoTextCell('') },
                { label: 'Priority of Action', render: () => autoTextCell('') }
              ],
              report.electrodeResistance.map((row) => ({
                ...row,
                resistanceWithoutGrid: safeText(row.resistanceWithoutGrid, ''),
                resistanceWithGrid: safeText(row.resistanceWithGrid || row.measuredResistance, '')
              })),
              {
                observationRenderer: (row) => renderObservationCard(row)
              }
            )
          )
        : ''}

      ${report.tests.continuityTest
        ? reportSectionBlock(
            'Continuity Test',
            simpleDetailTable(
              [
                { label: 'Sr. No.', key: 'srNo' },
                { label: 'Main Location', key: 'mainLocation' },
                { label: 'Measurement Point', key: 'measurementPoint' },
                { label: 'Resistance (ohm)', key: 'resistance' },
                { label: 'Impedance (ohm)', key: 'impedance' },
                {
                  label: 'Status',
                  render: (row) => {
                    const assessment = deriveContinuityAssessment(row);
                    return pill(assessment.status.tone, assessment.status.label);
                  }
                },
                {
                  label: 'Comment',
                  render: (row) => autoTextCell(safeText(row.comment, deriveContinuityAssessment(row).comment))
                }
              ],
              report.continuityTest,
              {
                observationRenderer: (row) => renderObservationCard(row)
              }
            )
          )
        : ''}

      ${report.tests.loopImpedanceTest
        ? reportSectionBlock(
            'Loop Impedance Test',
            simpleDetailTable(
              [
                { label: 'Sr. No.', key: 'srNo' },
                { label: 'Location of Panel / Equipment', key: 'location' },
                { label: 'Name of Feeder & Tag No.', key: 'feederTag' },
                { label: 'Type', key: 'deviceType' },
                { label: 'Rating (A)', key: 'deviceRating' },
                { label: 'Breaking (kA)', key: 'breakingCapacity' },
                { label: 'Measured Points', key: 'measuredPoints' },
                { label: 'Loop Impedance (ohm)', key: 'loopImpedance' },
                { label: 'Voltage (V)', key: 'voltage' },
                {
                  label: 'Remark',
                  render: (row) => autoTextCell(safeText(row.remarks, toReportBandLabel(deriveLoopAssessment(row).status)))
                }
              ],
              report.loopImpedanceTest,
              {
                observationRenderer: (row) => renderObservationCard(row)
              }
            )
          )
        : ''}

      ${report.tests.prospectiveFaultCurrent
        ? reportSectionBlock(
            'Prospective Fault Current',
            simpleDetailTable(
              [
                { label: 'Sr. No.', key: 'srNo' },
                { label: 'Location of Panel / Equipment', key: 'location' },
                { label: 'Name of Feeder & Tag No.', key: 'feederTag' },
                { label: 'Device', key: 'deviceType' },
                { label: 'Rating (A)', key: 'deviceRating' },
                { label: 'Breaking (kA)', key: 'breakingCapacity' },
                { label: 'Measured Points', key: 'measuredPoints' },
                { label: 'Loop Z (ohm)', key: 'loopImpedance' },
                { label: 'PFC', key: 'prospectiveFaultCurrent' },
                { label: 'Voltage', key: 'voltage' },
                {
                  label: 'Remark',
                  render: (row) => autoTextCell(safeText(row.comment, toReportBandLabel(deriveFaultAssessment(row).status)))
                }
              ],
              report.prospectiveFaultCurrent,
              {
                observationRenderer: (row) => renderObservationCard(row)
              }
            )
          )
        : ''}

      ${report.tests.riserIntegrityTest
        ? reportSectionBlock(
            'Riser / Grid Integrity Test',
            simpleDetailTable(
              [
                { label: 'Sr. No.', key: 'srNo' },
                { label: 'Main Location', key: 'mainLocation' },
                { label: 'Measurement Point', key: 'measurementPoint' },
                { label: 'Towards Equipment', key: 'resistanceTowardsEquipment' },
                { label: 'Towards Grid', key: 'resistanceTowardsGrid' },
                {
                  label: 'Comment',
                  render: (row) => autoTextCell(toRiserCommentLabel(deriveRiserAssessment(row).status))
                },
                { label: 'Observation', render: () => autoTextCell('') },
                { label: 'Recommendation', render: () => autoTextCell('') }
              ],
              report.riserIntegrityTest,
              {
                observationRenderer: (row) => renderObservationCard(row)
              }
            )
          )
        : ''}

      ${report.tests.earthContinuityTest
        ? reportSectionBlock(
            'Earth Continuity Test',
            simpleDetailTable(
              [
                { label: 'Sr. No.', key: 'srNo' },
                { label: 'Tag', key: 'tag' },
                { label: 'Location / Building', key: 'locationBuildingName' },
                { label: 'Distance', key: 'distance' },
                { label: 'Measured Value', key: 'measuredValue' },
                {
                  label: 'Status',
                  render: (row) => {
                    const assessment = deriveEarthContinuityAssessment(row);
                    return pill(assessment.status.tone, assessment.status.label);
                  }
                },
                {
                  label: 'Remark',
                  render: (row) => autoTextCell(safeText(row.remark, deriveEarthContinuityAssessment(row).comment))
                }
              ],
              report.earthContinuityTest,
              {
                observationRenderer: (row) => renderObservationCard(row)
              }
            )
          )
        : ''}

      ${report.tests.towerFootingResistance
        ? reportSectionBlock(
            'Tower Footing Resistance Measurement & Analysis',
            towerFootingDetailTable(report.towerFootingResistance)
          )
        : ''}
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
  app.innerHTML = `
    ${brandHeader()}
    <main class="app-shell">
      ${state.view === 'dashboard' ? dashboardView() : ''}
      ${state.view === 'builder' ? builderView() : ''}
      ${state.view === 'detail' && state.activeReport ? detailView() : ''}
    </main>
    ${renderObservationDrawer()}
    ${toastHtml()}
  `;
  if (state.view === 'builder') {
    scheduleDraftAutosave();
  }
  bindMotionEffects();
  updateMotionEffects();
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
    ['workOrder', 'Work order / reference is required.'],
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

function renumberSection(section) {
  const rows = state.draft[section];
  if (!Array.isArray(rows)) {
    return;
  }
  rows.forEach((row, index) => {
    if (Object.prototype.hasOwnProperty.call(row, 'srNo')) {
      row.srNo = String(index + 1);
    }
    if ((section === 'loopImpedanceTest' || section === 'prospectiveFaultCurrent') && Object.prototype.hasOwnProperty.call(row, 'measuredPoints')) {
      row.measuredPoints = PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length];
    }
  });
}

function syncPhaseGroupField(section, index, field, value) {
  const rows = state.draft[section];
  if (!Array.isArray(rows)) {
    return false;
  }
  const groupStart = Math.floor(index / PHASE_MEASURED_POINTS.length) * PHASE_MEASURED_POINTS.length;
  for (let offset = 0; offset < PHASE_MEASURED_POINTS.length; offset += 1) {
    if (rows[groupStart + offset]) {
      rows[groupStart + offset][field] = value;
    }
  }
  return true;
}

function addRow(section, direction) {
  if (section === 'soilResistivity') {
    return;
  }
  if (section === 'soilResistivity-location') {
    state.draft.soilResistivity.locations.push(defaultSoilLocation(`Location ${state.draft.soilResistivity.locations.length + 1}`));
    return;
  }
  if (section === 'electrodeResistance') {
    state.draft.electrodeResistance.push(defaultElectrodeRow());
    return;
  }
  if (section === 'continuityTest') {
    state.draft.continuityTest.push(defaultContinuityRow(String(state.draft.continuityTest.length + 1)));
    return;
  }
  if (section === 'loopImpedanceTest') {
    state.draft.loopImpedanceTest.push(...defaultLoopGroup(state.draft.loopImpedanceTest.length + 1));
    return;
  }
  if (section === 'prospectiveFaultCurrent') {
    state.draft.prospectiveFaultCurrent.push(...defaultFaultGroup(state.draft.prospectiveFaultCurrent.length + 1));
    return;
  }
  if (section === 'riserIntegrityTest') {
    state.draft.riserIntegrityTest.push(defaultRiserRow(String(state.draft.riserIntegrityTest.length + 1)));
    return;
  }
  if (section === 'earthContinuityTest') {
    state.draft.earthContinuityTest.push(defaultEarthContinuityRow(String(state.draft.earthContinuityTest.length + 1)));
    return;
  }
  if (section === 'towerFootingResistance') {
    state.draft.towerFootingResistance.push(defaultTowerFootingGroup(String(state.draft.towerFootingResistance.length + 1)));
    return;
  }
}

function removeRow(section, index, direction, locationIndex = null) {
  if (
    state.observationEditor &&
    state.observationEditor.section === section &&
    safeText(state.observationEditor.direction, '') === safeText(direction, '') &&
    (section !== 'soilResistivity' || state.observationEditor.locationIndex === locationIndex)
  ) {
    closeObservationEditor();
  }
  if (section === 'soilResistivity') {
    const location = state.draft.soilResistivity.locations?.[locationIndex];
    if (!location) {
      return;
    }
    if (location[direction].length === 1) {
      showToast('Keep at least one row in each soil direction.', 'warning');
      return;
    }
    location[direction].splice(index, 1);
    return;
  }
  if (section === 'loopImpedanceTest' || section === 'prospectiveFaultCurrent') {
    if (state.draft[section].length <= PHASE_MEASURED_POINTS.length) {
      showToast('Keep at least one 3-point feeder group in the selected section.', 'warning');
      return;
    }
    const groupStart = Math.floor(index / PHASE_MEASURED_POINTS.length) * PHASE_MEASURED_POINTS.length;
    state.draft[section].splice(groupStart, PHASE_MEASURED_POINTS.length);
    renumberSection(section);
    return;
  }
  if (state.draft[section].length === 1) {
    showToast('Keep at least one row in the selected section.', 'warning');
    return;
  }
  state.draft[section].splice(index, 1);
  renumberSection(section);
}

function addSoilLocation() {
  state.draft.soilResistivity.locations.push(defaultSoilLocation(`Location ${state.draft.soilResistivity.locations.length + 1}`));
}

function removeSoilLocation(locationIndex) {
  if (state.observationEditor && state.observationEditor.section === 'soilResistivity' && state.observationEditor.locationIndex === locationIndex) {
    closeObservationEditor();
  }
  if (state.draft.soilResistivity.locations.length === 1) {
    showToast('Keep at least one soil location in the report.', 'warning');
    return;
  }
  state.draft.soilResistivity.locations.splice(locationIndex, 1);
  state.draft.soilResistivity.locations.forEach((location, index) => {
    if (!safeText(location.name, '')) {
      location.name = `Location ${index + 1}`;
    }
  });
}

function removeTowerGroup(groupIndex) {
  if (state.observationEditor && state.observationEditor.section === 'towerFootingResistance' && state.observationEditor.groupIndex === groupIndex) {
    closeObservationEditor();
  }
  if (state.draft.towerFootingResistance.length === 1) {
    showToast('Keep at least one tower location group in the selected section.', 'warning');
    return;
  }
  state.draft.towerFootingResistance.splice(groupIndex, 1);
  state.draft.towerFootingResistance.forEach((group, index) => {
    group.srNo = String(index + 1);
  });
}

async function saveReport() {
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
    showToast('Report saved successfully.', 'healthy');
    await loadReports();
  } catch (error) {
    showToast(error.message, 'critical');
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

  if (action === 'new-report') {
    closeObservationEditor();
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

  if (action === 'toggle-equipment') {
    const equipmentId = button.dataset.equipmentId;
    const selections = new Set(Array.isArray(state.draft.project.equipmentSelections) ? state.draft.project.equipmentSelections : []);
    if (selections.has(equipmentId)) {
      selections.delete(equipmentId);
    } else {
      selections.add(equipmentId);
    }
    state.draft.project.equipmentSelections = EQUIPMENT_LIBRARY.filter((equipment) => selections.has(equipment.id)).map((equipment) => equipment.id);
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

  if (action === 'add-row') {
    if (button.dataset.section === 'soilResistivity') {
      const locationIndex = Number(button.dataset.locationIndex);
      const location = state.draft.soilResistivity.locations?.[locationIndex];
      if (location) {
        const nextSpacing = SOIL_SPACING_PRESETS[location[button.dataset.direction].length] || '';
        location[button.dataset.direction].push(defaultSoilRow(nextSpacing));
      }
    } else {
      addRow(button.dataset.section, button.dataset.direction);
    }
    render();
    return;
  }

  if (action === 'add-soil-location') {
    addSoilLocation();
    render();
    return;
  }

  if (action === 'remove-row') {
    removeRow(button.dataset.section, Number(button.dataset.index), button.dataset.direction, button.dataset.locationIndex === undefined ? null : Number(button.dataset.locationIndex));
    render();
    return;
  }

  if (action === 'remove-soil-location') {
    removeSoilLocation(Number(button.dataset.locationIndex));
    render();
    return;
  }

  if (action === 'remove-tower-group') {
    removeTowerGroup(Number(button.dataset.groupIndex));
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

  if (target.matches('[data-bind]')) {
    setByPath(state.draft, target.dataset.bind, target.value);
    scheduleDraftAutosave();
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

  if (target.matches('[data-section]')) {
    const section = target.dataset.section;
    const index = Number(target.dataset.index);
    const readingIndex = target.dataset.readingIndex === undefined ? null : Number(target.dataset.readingIndex);
    const locationIndex = target.dataset.locationIndex === undefined ? null : Number(target.dataset.locationIndex);
    const field = target.dataset.field;
    const direction = target.dataset.direction;

    if (section === 'soilResistivity') {
      state.draft.soilResistivity.locations[locationIndex][direction][index][field] = target.value;
    } else if (section === 'towerFootingResistance') {
      if (Number.isInteger(readingIndex)) {
        state.draft.towerFootingResistance[index].readings[readingIndex][field] = target.value;
      } else {
        state.draft.towerFootingResistance[index][field] = target.value;
      }
    } else if ((section === 'loopImpedanceTest' || section === 'prospectiveFaultCurrent') && target.dataset.groupSync === 'true') {
      syncPhaseGroupField(section, index, field, target.value);
    } else {
      state.draft[section][index][field] = target.value;
    }
    scheduleDraftAutosave();
    rerenderPreservingFocus();
  }
});

document.addEventListener('change', async (event) => {
  const target = event.target;
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
  if (target.matches('[data-section]')) {
    const section = target.dataset.section;
    const index = Number(target.dataset.index);
    const readingIndex = target.dataset.readingIndex === undefined ? null : Number(target.dataset.readingIndex);
    const locationIndex = target.dataset.locationIndex === undefined ? null : Number(target.dataset.locationIndex);
    const field = target.dataset.field;
    const direction = target.dataset.direction;
    if (section === 'soilResistivity') {
      state.draft.soilResistivity.locations[locationIndex][direction][index][field] = target.value;
    } else if (section === 'towerFootingResistance') {
      if (Number.isInteger(readingIndex)) {
        state.draft.towerFootingResistance[index].readings[readingIndex][field] = target.value;
      } else {
        state.draft.towerFootingResistance[index][field] = target.value;
      }
    } else if ((section === 'loopImpedanceTest' || section === 'prospectiveFaultCurrent') && target.dataset.groupSync === 'true') {
      syncPhaseGroupField(section, index, field, target.value);
    } else {
      state.draft[section][index][field] = target.value;
    }
    scheduleDraftAutosave();
    rerenderPreservingFocus();
  }
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && state.observationEditor) {
    closeObservationEditor();
    render();
  }
});

async function init() {
  restoreDraftSnapshot();
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
}

window.addEventListener('beforeunload', () => {
  persistDraftSnapshotNow();
});

init();
