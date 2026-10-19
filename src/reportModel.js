// All limits, formulas, factors and classifications come from the project's acceptance
// profile (ET-D03-REG-002 R1+, see ./acceptance.js): values on the criterion, factors,
// bands and labels in its `params`. The only literals left here are the condition-code
// sets and the formulas themselves, which are fixed by the standard's text.
const { resolveProfile, BASIS_LABELS, DEFAULT_PROFILE_ID } = require('./acceptance');
const { engineerRemark } = require('./remarks');
const { validateReport } = require('./sheetValidation');
const { SOIL_SPACING_PRESETS: SOIL_PRESETS, DEFAULT_SOIL_PRESET, SHEETS, INSTRUMENTS, GROUND_CONDITIONS } = require('./sheetDefinitions');

const CONDITION_CODES = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'OTH'];
const PRIORITY_P1_CODES = new Set(['C1', 'C2', 'C4', 'C5']);
const PRIORITY_P2_CODES = new Set(['C3', 'C6', 'C7', 'C8']);

const TEST_LIBRARY = [
  {
    id: 'soilResistivity',
    unitLabel: 'locations',
    sheetCode: 'MS-01',
    label: 'Soil Resistivity Test',
    shortLabel: 'Soil',
    description: 'Direction-wise soil resistivity readings with automatic average and IS 3043 corrosivity class.'
  },
  {
    id: 'electrodeResistance',
    unitLabel: 'pits',
    sheetCode: 'MS-02',
    label: 'Earth Electrode Resistance Test',
    shortLabel: 'Electrode',
    description: 'Pit-wise resistance against the IS 3043:2018 Cl. 14.2 limit calculated from site soil resistivity.'
  },
  {
    id: 'continuityTest',
    unitLabel: 'points',
    sheetCode: 'MS-03',
    label: 'Earth Continuity & Bonding Test',
    shortLabel: 'Continuity',
    description: 'Point-to-point bonding resistance with test-lead resistance subtracted.'
  },
  {
    id: 'loopImpedanceTest',
    unitLabel: 'panels',
    sheetCode: 'MS-04',
    label: 'Loop Impedance Test',
    shortLabel: 'Loop',
    description: 'Measured Zs values for protection loop verification.'
  },
  {
    id: 'prospectiveFaultCurrent',
    unitLabel: 'panels',
    sheetCode: 'MS-05',
    label: 'Prospective Fault Current',
    shortLabel: 'PFC',
    description: 'PFC (L-PE) and PSC (L-N) against the device breaking capacity.'
  },
  {
    id: 'riserIntegrityTest',
    unitLabel: 'points',
    sheetCode: 'MS-06',
    label: 'Riser / Grid Integrity Test',
    shortLabel: 'Riser',
    description: 'Resistance verification towards equipment and earth grid.'
  },
  {
    id: 'groundResistance',
    unitLabel: 'points',
    sheetCode: 'MS-07',
    label: 'Ground (Grid) Resistance Test',
    shortLabel: 'Ground',
    description: 'In-service grid resistance against the limit of the associated pit.'
  },
  {
    id: 'stepTouchVoltage',
    unitLabel: 'points',
    sheetCode: 'MS-08',
    label: 'Step & Touch Voltage Test',
    shortLabel: 'Step & Touch',
    description: 'Injected-current step and touch voltages scaled by If/It against the tolerable limits.'
  },
  {
    id: 'towerFootingResistance',
    unitLabel: 'towers',
    sheetCode: 'MS-09',
    label: 'Tower Footing Resistance Measurement & Analysis',
    shortLabel: 'Tower Footing',
    description: 'Grouped tower footing impedance and current analysis with per-tower totals.'
  },
  {
    id: 'earthContinuityTest',
    unitLabel: 'points',
    sheetCode: 'legacy',
    retired: true,
    label: 'Earth Continuity Test',
    shortLabel: 'Earth Continuity',
    description: 'Retired: folded into Earth Continuity & Bonding (MS-03). Kept so older reports still open.'
  }
];

const SOIL_DIRECTION_KEYS = ['direction1', 'direction2', 'direction3', 'direction4', 'direction5', 'direction6', 'direction7', 'direction8'];
const TOWER_FOOT_POINTS = ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'];
const PHASE_MEASURED_POINTS = ['R-E', 'Y-E', 'B-E'];
const EQUIPMENT_LIBRARY = [
  {
    id: 'mi3152',
    label: 'MI 3152 EurotestXC'
  },
  {
    id: 'mi3290',
    label: 'MI 3290 GF Earth Analyser'
  },
  {
    id: 'kyoritsu4118a',
    label: 'Kyoritsu Digital PSC Loop Tester 4118A'
  }
];

function buildRowId(prefix = 'row') {
  const timestamp = Date.now().toString(36);
  const suffix = Math.random().toString(36).slice(2, 8);
  return `${prefix}-${timestamp}-${suffix}`;
}

function createRowMeta(prefix = 'row') {
  return {
    rowId: buildRowId(prefix),
    rowObservation: '',
    rowPhotos: []
  };
}

function createDefaultSoilRow(spacing = '', resistivity = '') {
  return {
    ...createRowMeta('soil'),
    spacing,
    resistivity,
    remark: ''
  };
}

function soilPreset(presetId) {
  return SOIL_PRESETS[presetId] || SOIL_PRESETS[DEFAULT_SOIL_PRESET];
}

// A soil location for the project's spacing preset: one row per spacing in every
// direction (the printed sheet has 4 direction columns per page; the substation
// preset has 8 directions).
function createDefaultSoilLocation(name = 'Location 1', presetId = DEFAULT_SOIL_PRESET) {
  const preset = soilPreset(presetId);
  const location = {
    locationId: buildRowId('soil-location'),
    name,
    // 'ohm-m': values are rho as read. 'ohm': values are R as read, rho = 2*pi*a*R.
    inputUnit: '',
    probeDepthCm: '',
    weather: '',
    bearings: Object.fromEntries(SOIL_DIRECTION_KEYS.map((key) => [key, ''])),
    drivenElectrodeDiameter: '',
    drivenElectrodeLength: '',
    notes: ''
  };
  SOIL_DIRECTION_KEYS.forEach((key, index) => {
    location[key] = index < preset.directions ? preset.spacings.map((spacing) => createDefaultSoilRow(spacing, '')) : [];
  });
  return location;
}

function createDefaultElectrodeRow() {
  return {
    ...createRowMeta('electrode'),
    tag: '',
    location: '',
    electrodeType: '',
    materialType: '',
    length: '',
    diameter: '',
    plateLengthM: '',
    plateWidthM: '',
    plateArea: '',
    resistanceWithoutGrid: '',
    checkR50: '',
    checkR70: '',
    resistanceWithGrid: '',
    method: '',
    conditionCodes: [],
    photoRef: '',
    observation: ''
  };
}

function createDefaultContinuityRow() {
  return {
    ...createRowMeta('continuity'),
    srNo: '',
    tag: '',
    mainLocation: '',
    referencePoint: '',
    measurementPoint: '',
    distanceM: '',
    testCurrentA: '',
    resistance: '',
    impedance: '',
    testLeadResistance: '',
    conditionCodes: [],
    photoRef: '',
    comment: ''
  };
}

function createDefaultLoopImpedanceRow() {
  return {
    ...createRowMeta('loop'),
    srNo: '',
    location: '',
    feederTag: '',
    deviceType: '',
    deviceRating: '',
    tripCurve: '',
    breakingCapacity: '',
    measuredPoints: 'R-E',
    loopImpedance: '',
    voltage: '',
    photoRef: '',
    remarks: ''
  };
}

function createDefaultProspectiveFaultRow() {
  return {
    ...createRowMeta('fault'),
    srNo: '',
    location: '',
    feederTag: '',
    deviceType: '',
    deviceRating: '',
    breakingCapacity: '',
    measuredPoints: 'R-E',
    loopImpedance: '',
    prospectiveFaultCurrent: '',
    shortCircuitCurrentLN: '',
    voltage: '',
    photoRef: '',
    comment: ''
  };
}

function createDefaultLoopGroupRows(startIndex = 1) {
  return PHASE_MEASURED_POINTS.map((point, index) => ({
    ...createDefaultLoopImpedanceRow(),
    srNo: String(startIndex + index),
    measuredPoints: point
  }));
}

function createDefaultFaultGroupRows(startIndex = 1) {
  return PHASE_MEASURED_POINTS.map((point, index) => ({
    ...createDefaultProspectiveFaultRow(),
    srNo: String(startIndex + index),
    measuredPoints: point
  }));
}

function createDefaultRiserIntegrityRow() {
  return {
    ...createRowMeta('riser'),
    srNo: '',
    mainLocation: '',
    measurementPoint: '',
    resistanceTowardsEquipment: '',
    resistanceTowardsGrid: '',
    testLeadResistance: '',
    conditionCodes: [],
    photoRef: '',
    comment: ''
  };
}

function createDefaultGroundResistanceRow() {
  return {
    ...createRowMeta('ground'),
    srNo: '',
    tag: '',
    location: '',
    method: '',
    currentProbeDistanceM: '',
    r50: '',
    r618: '',
    r70: '',
    conditionCodes: [],
    remark: ''
  };
}

function createDefaultStepTouchRow() {
  return {
    ...createRowMeta('steptouch'),
    srNo: '',
    location: '',
    measurementType: '',
    surface: '',
    distanceM: '',
    testCurrentA: '',
    v1: '',
    v2: '',
    v3: '',
    vmax: '',
    photoRef: '',
    remark: ''
  };
}

// Retired (folded into MS-03). Kept so the shape of older reports stays defined.
function createDefaultEarthContinuityRow() {
  return {
    ...createRowMeta('earth'),
    srNo: '',
    tag: '',
    locationBuildingName: '',
    distance: '',
    measuredValue: '',
    testLeadResistance: '',
    remark: ''
  };
}

function createDefaultTowerFootingReading(foot) {
  return {
    measurementPointLocation: '',
    footLabel: foot,
    footToEarthingConnectionStatus: '',
    measuredCurrentMa: '',
    measuredImpedance: '',
    conditionCodes: [],
    photoRef: '',
    remark: '',
    ...createRowMeta('tower')
  };
}

function createDefaultTowerFootingGroup() {
  return {
    groupId: buildRowId('tower-group'),
    srNo: '',
    mainLocationTower: '',
    readings: TOWER_FOOT_POINTS.map((foot) => {
      const reading = createDefaultTowerFootingReading(foot);
      reading.measurementPointLocation = foot;
      return reading;
    }),
    totalImpedanceZt: '',
    totalCurrentItotal: '',
    standardTolerableImpedanceZsat: String(resolveProfile().value('TOWER-ZSAT')),
    remarks: ''
  };
}

// Header block of one sheet (instrument, serial, calibration, times, plus the sheet's
// own printed header cells).
function createDefaultSheetHeader(sheetId) {
  const header = { instrument: '', serialNo: '', calibrationDue: '', timeStart: '', timeFinish: '' };
  const sheet = SHEETS.find((entry) => entry.id === sheetId);
  (sheet ? sheet.headerFields : []).forEach((field) => {
    const fields = field.kind === 'group' ? field.fields : [field];
    fields.forEach((entry) => {
      if (entry.readOnly) {
        return;
      }
      header[entry.key] = entry.kind === 'tick' ? false : '';
    });
  });
  return header;
}

function buildDefaultDraft() {
  return {
    project: {
      projectNo: '',
      clientName: '',
      siteLocation: '',
      workOrder: '',
      reportDate: new Date().toISOString().slice(0, 10),
      engineerName: '',
      clientRepresentative: '',
      groundCondition: '',
      soilSpacingPreset: DEFAULT_SOIL_PRESET,
      equipmentSelections: EQUIPMENT_LIBRARY.map((equipment) => equipment.id),
      acceptanceProfile: 'ETPL-DEFAULT',
      assumedSoilResistivity: '',
      assumedSoilResistivitySource: '',
      zohoProjectId: '',
      zohoProjectName: '',
      zohoProjectOwner: '',
      zohoProjectStage: '',
      unit: '',
      unitCode: '',
      plannedQuantities: {}
    },
    tests: {
      soilResistivity: true,
      electrodeResistance: true,
      continuityTest: false,
      loopImpedanceTest: false,
      prospectiveFaultCurrent: false,
      riserIntegrityTest: false,
      groundResistance: false,
      stepTouchVoltage: false,
      towerFootingResistance: false,
      earthContinuityTest: false
    },
    sheetHeaders: {},
    soilResistivity: {
      locations: [createDefaultSoilLocation('Location 1')]
    },
    electrodeResistance: [createDefaultElectrodeRow()],
    continuityTest: [createDefaultContinuityRow()],
    loopImpedanceTest: createDefaultLoopGroupRows(1),
    prospectiveFaultCurrent: createDefaultFaultGroupRows(1),
    riserIntegrityTest: [createDefaultRiserIntegrityRow()],
    groundResistance: [createDefaultGroundResistanceRow()],
    stepTouchVoltage: [createDefaultStepTouchRow()],
    earthContinuityTest: [createDefaultEarthContinuityRow()],
    towerFootingResistance: [{ ...createDefaultTowerFootingGroup(), srNo: '1' }],
    ocrImports: {}
  };
}

function asTrimmedString(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

function asLooseNumber(value) {
  const numeric = Number.parseFloat(String(value === undefined || value === null ? '' : value).trim());
  return Number.isFinite(numeric) ? numeric : null;
}

function asTextNumber(value) {
  const numeric = asLooseNumber(value);
  return numeric === null ? '' : String(numeric);
}

function average(values) {
  const filtered = values.filter((value) => Number.isFinite(value));
  if (!filtered.length) {
    return null;
  }
  const total = filtered.reduce((sum, value) => sum + value, 0);
  return total / filtered.length;
}

function round(value, digits = 2) {
  if (!Number.isFinite(value)) {
    return null;
  }
  const factor = 10 ** digits;
  return Math.round(value * factor) / factor;
}

function fmt2(value) {
  return Number.isFinite(value) ? value.toFixed(2) : '-';
}

function pendingStatus() {
  return { label: 'Pending', tone: 'neutral' };
}

function healthyStatus() {
  return { label: 'Healthy', tone: 'healthy' };
}

function overLimitStatus() {
  return { label: '>Permissible Limit', tone: 'critical' };
}

function normalizeConditionCodes(value) {
  const list = Array.isArray(value) ? value : typeof value === 'string' ? value.split(/[\s,;]+/) : [];
  const upper = list.map((item) => asTrimmedString(item).toUpperCase());
  return CONDITION_CODES.filter((code) => upper.includes(code));
}

// ---- Assessment context -----------------------------------------------------
// Everything a status depends on besides the row itself: the project's resolved
// acceptance profile and the soil resistivity used for electrode limits.
// Assessment functions take either a context or the report it is built from.

function buildAssessmentContext(report) {
  const profile = resolveProfile(report?.project?.acceptanceProfile);
  const measured = calculateSoilSummary(report).overallAverage;
  const assumed = asLooseNumber(report?.project?.assumedSoilResistivity);
  let rho = null;
  let rhoAssumed = false;
  let rhoSource = '';
  if (Number.isFinite(measured) && measured > 0) {
    rho = measured;
    rhoSource = 'Measured site average (Wenner 4-probe)';
  } else if (Number.isFinite(assumed) && assumed > 0) {
    rho = assumed;
    rhoAssumed = true;
    rhoSource = asTrimmedString(report?.project?.assumedSoilResistivitySource) || 'Source not stated';
  }
  // Limits of the pits by tag, for MS-07 (a ground reading is compared with the limit of
  // the pit it belongs to).
  const pitLimits = new Map();
  (Array.isArray(report?.electrodeResistance) ? report.electrodeResistance : []).forEach((row) => {
    const tag = asTrimmedString(row?.tag).toLowerCase();
    const result = tag
      ? calculateElectrodeLimit({ type: row?.electrodeType, lengthM: asLooseNumber(row?.length), diameterMm: asLooseNumber(row?.diameter), areaM2: plateAreaOf(row), rho }, profile)
      : null;
    if (result) {
      pitLimits.set(tag, result.limit);
    }
  });
  return { profile, rho, rhoAssumed, rhoSource, headers: report?.sheetHeaders || {}, pitLimits };
}

function resolveContext(contextOrReport) {
  if (contextOrReport && contextOrReport.profile && typeof contextOrReport.profile.get === 'function') {
    return contextOrReport;
  }
  return buildAssessmentContext(contextOrReport || null);
}

function getStandardGuidance(profile = resolveProfile()) {
  const cite = (id) => {
    const criterion = profile.get(id);
    const clause = criterion.clause && criterion.clause !== '-' ? `, Cl. ${criterion.clause}` : '';
    return `${criterion.source}${clause}`;
  };
  const valueLabel = (id, what) => {
    const criterion = profile.get(id);
    return `${fmt2(criterion.value)} ohm ${what} (${BASIS_LABELS[criterion.basis] || criterion.basis})`;
  };
  return {
    electrodeResistance: { reference: cite('PIT-ROD'), limitLabel: 'IS 3043:2018 Cl. 14.2 limit calculated for the electrode' },
    continuityTest: { reference: cite('CONT-LIMIT'), limitLabel: valueLabel('CONT-LIMIT', 'continuity limit') },
    loopImpedanceTest: { reference: cite('LOOP-LIMIT'), limitLabel: valueLabel('LOOP-LIMIT', 'Zs limit') },
    prospectiveFaultCurrent: {
      reference: cite('PFC-RULE'),
      limitLabel: 'device breaking capacity above the higher of PFC (L-PE) and PSC (L-N)'
    },
    riserIntegrityTest: { reference: cite('RISER-LIMIT'), limitLabel: valueLabel('RISER-LIMIT', 'riser bonding limit') },
    earthContinuityTest: { reference: cite('CONT-LIMIT'), limitLabel: valueLabel('CONT-LIMIT', 'continuity limit') },
    towerFootingResistance: { reference: cite('TOWER-ZSAT'), limitLabel: valueLabel('TOWER-ZSAT', 'Zsat') },
    groundResistance: { reference: cite('GND-LIMIT'), limitLabel: 'calculated limit of the associated pit (IS 3043:2018 Cl. 14.2), or the client value' },
    stepTouchVoltage: { reference: cite('ST-LIMIT'), limitLabel: 'tolerable step / touch voltage after scaling by If/It' }
  };
}

// ---- Soil -------------------------------------------------------------------
// IS 3043:2018 Cl. 13.6.1, Table 4 (SOIL-CLASS). Corrosivity only: no pass/fail.
function getSoilCategory(value, profile = resolveProfile()) {
  if (!Number.isFinite(value)) {
    return { label: 'Insufficient Data', tone: 'neutral' };
  }
  const criterion = profile.get('SOIL-CLASS');
  const band = criterion.params.bands.find((entry) => entry.max === null || value < entry.max);
  return { label: band.class, tone: 'neutral', reference: `${criterion.source}, Cl. ${criterion.clause}` };
}

// Steel corrosion allowance (IS 3043:2018 Table 5) for a soil resistivity, in percent.
function getCorrosionAllowancePct(value, profile = resolveProfile()) {
  if (!Number.isFinite(value)) {
    return null;
  }
  const entry = profile.get('SOIL-CLASS').params.corrosion_allowance.find((band) => band.max === null || value < band.max);
  return entry ? entry.steel_pct : null;
}

// ---- Earth electrode (pit) --------------------------------------------------
// Plate area for the Cl. 14.2.1 formula is the area of BOTH faces: 2 x length x width.
// Older reports carried the area itself in `plateArea`.
function plateAreaOf(row) {
  const length = asLooseNumber(row?.plateLengthM);
  const width = asLooseNumber(row?.plateWidthM);
  if (Number.isFinite(length) && Number.isFinite(width) && length > 0 && width > 0) {
    return 2 * length * width;
  }
  return asLooseNumber(row?.plateArea);
}

// 50% / 70% check readings against the 61.8% reading (IS 3043:2018 Cl. 41.1.6). The
// tolerance is an ETPL acceptance criterion held on PIT-METHOD params; the standard
// itself gives no number. Flag only: status and priority are never changed by it.
// Not applicable to the clamp and slope methods, which take no 50/70% readings.
function checkFopAgreement({ r618, r50, r70, method }, profile) {
  const methodText = asTrimmedString(method);
  if (methodText === 'Clamp' || methodText === 'Slope' || !Number.isFinite(r618)) {
    return null;
  }
  const checks = [
    { at: '50%', value: r50 },
    { at: '70%', value: r70 }
  ].filter((check) => Number.isFinite(check.value));
  if (!checks.length) {
    return null;
  }
  const params = profile.params('PIT-METHOD');
  const tolerancePct = params.agreement_tolerance_pct;
  const deviation = (value) => (r618 === 0 ? (value === 0 ? 0 : Infinity) : (Math.abs(value - r618) / Math.abs(r618)) * 100);
  const result = {
    tolerancePct,
    toleranceBasis: params.agreement_tolerance_basis,
    toleranceSource: params.agreement_tolerance_source,
    r50DevPct: null,
    r70DevPct: null,
    // A 61.8% reading of 0 has no percentage to compare against; any non-zero check differs.
    comparable: r618 !== 0
  };
  checks.forEach((check) => {
    result[check.at === '50%' ? 'r50DevPct' : 'r70DevPct'] = Number.isFinite(deviation(check.value)) ? round(deviation(check.value), 1) : null;
  });
  // 1e-9 absorbs floating-point noise: 1.05 against 1.00 is exactly 5%.
  result.agrees = checks.every((check) => deviation(check.value) <= tolerancePct + 1e-9);
  return result;
}

function fopRemark(fop) {
  const parts = [];
  if (fop.r50DevPct !== null) parts.push(`50% reading ${fop.r50DevPct}% from the 61.8% value`);
  if (fop.r70DevPct !== null) parts.push(`70% reading ${fop.r70DevPct}% from the 61.8% value`);
  if (!fop.comparable) parts.push('the 61.8% reading is 0, so the checks cannot agree with it');
  return `re-test – 61.8% value unconfirmed (${parts.join('; ')}; tolerance ${fop.tolerancePct}%)`;
}

// Limit by electrode type, IS 3043:2018 Cl. 14.2 (PIT-ROD / PIT-STRIP / PIT-PLATE).
// l is entered in m and d in mm; the rod and strip formulas take both in cm.
function calculateElectrodeLimit({ type, lengthM, diameterMm, areaM2, rho }, profile = resolveProfile()) {
  if (!Number.isFinite(rho) || rho <= 0) {
    return null;
  }
  const kind = asTrimmedString(type).toLowerCase() || 'rod';
  if (kind === 'plate') {
    if (!Number.isFinite(areaM2) || areaM2 <= 0) {
      return null;
    }
    return { limit: (rho / 4) * Math.sqrt(Math.PI / areaM2), criterion: profile.get('PIT-PLATE') };
  }
  const factor = kind === 'strip' ? 4 : kind === 'rod' || kind === 'pipe' ? 2 : null;
  if (!factor || !Number.isFinite(lengthM) || lengthM <= 0 || !Number.isFinite(diameterMm) || diameterMm <= 0) {
    return null;
  }
  const lengthCm = lengthM * 100;
  const diameterCm = diameterMm / 10;
  const logArgument = (factor * lengthCm) / diameterCm;
  if (logArgument <= 1) {
    return null;
  }
  const limit = ((100 * rho) / (2 * Math.PI * lengthCm)) * Math.log(logArgument);
  return { limit, criterion: profile.get(kind === 'strip' ? 'PIT-STRIP' : 'PIT-ROD') };
}

function calculateDrivenElectrodeResistance(soilAverage, electrodeLengthMeters, electrodeDiameterMillimeters) {
  const result = calculateElectrodeLimit({
    type: 'rod',
    lengthM: electrodeLengthMeters,
    diameterMm: electrodeDiameterMillimeters,
    rho: soilAverage
  });
  return result ? round(result.limit, 2) : null;
}

// `options.notMeasurable` marks an electrode that could not be tested on site; it is
// neutral (neither pass nor fail) and is excluded from the overall-status pass rate.
// `options.limit` is the calculated limit in ohm. Without one the row cannot be judged.
function getElectrodeStatus(value, options = {}) {
  const profile = (options && options.profile) || resolveProfile();
  if (options && options.notMeasurable) {
    return { label: 'Not Measurable', tone: 'neutral' };
  }
  const limit = options ? options.limit : null;
  if (!Number.isFinite(value) || !Number.isFinite(limit)) {
    return pendingStatus();
  }
  if (value > limit) {
    return overLimitStatus();
  }
  if (value > profile.params('PIT-BAND').warning_factor * limit) {
    // Within the limit, so it counts as a pass; the warning band only raises priority.
    return { label: 'Warning', tone: 'warning', withinLimit: true };
  }
  return healthyStatus();
}

// Legacy reading (with-grid first), kept for display and AI payloads.
function getElectrodeMeasuredValue(row) {
  return asLooseNumber(row?.resistanceWithGrid) ?? asLooseNumber(row?.resistanceWithoutGrid) ?? asLooseNumber(row?.measuredResistance);
}

// The reading compared with the limit. The Cl. 14.2 formulas give the resistance of
// the isolated electrode, so the without-grid reading is used; with-grid is the
// fallback when no isolated reading was taken.
function getElectrodeAssessedReading(row) {
  const withoutGrid = asLooseNumber(row?.resistanceWithoutGrid);
  if (Number.isFinite(withoutGrid)) {
    return { value: withoutGrid, basis: 'without_grid' };
  }
  const withGrid = asLooseNumber(row?.resistanceWithGrid) ?? asLooseNumber(row?.measuredResistance);
  if (Number.isFinite(withGrid)) {
    return { value: withGrid, basis: 'with_grid' };
  }
  return { value: null, basis: null };
}

// P1: > limit, Not Measurable, or C1/C2/C4/C5. P2: Warning band, or C3/C6/C7/C8.
// P3: Healthy with no code. A healthy pit with no code is never P1.
function deriveElectrodePriority(status, conditionCodes) {
  const codes = normalizeConditionCodes(conditionCodes);
  const label = status?.label;
  if (label === '>Permissible Limit' || label === 'Not Measurable' || codes.some((code) => PRIORITY_P1_CODES.has(code))) {
    return 'P1';
  }
  if (label === 'Warning' || codes.some((code) => PRIORITY_P2_CODES.has(code))) {
    return 'P2';
  }
  return label === 'Healthy' ? 'P3' : null;
}

function deriveElectrodeAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).electrodeResistance;
  const reading = getElectrodeAssessedReading(row);
  const codes = normalizeConditionCodes(row?.conditionCodes);
  const notMeasurable = Boolean(row?.notMeasurable) || (!Number.isFinite(reading.value) && codes.includes('C4'));
  const limitResult = calculateElectrodeLimit(
    {
      type: row?.electrodeType,
      lengthM: asLooseNumber(row?.length),
      diameterMm: asLooseNumber(row?.diameter),
      areaM2: plateAreaOf(row),
      rho: context.rho
    },
    context.profile
  );
  const limit = limitResult ? limitResult.limit : null;
  const status = getElectrodeStatus(reading.value, { notMeasurable, limit, profile: context.profile });
  const fop = checkFopAgreement(
    { r618: asLooseNumber(row?.resistanceWithoutGrid), r50: asLooseNumber(row?.checkR50), r70: asLooseNumber(row?.checkR70), method: row?.method },
    context.profile
  );
  const flags = fop && !fop.agrees ? [{ code: 'FOP_DISAGREE', message: fopRemark(fop), review: true }] : [];
  const priority = deriveElectrodePriority(status, codes);

  const withoutGrid = asLooseNumber(row?.resistanceWithoutGrid);
  const withGrid = asLooseNumber(row?.resistanceWithGrid) ?? asLooseNumber(row?.measuredResistance);
  const limitText = limit === null ? null : `${fmt2(limit)} ohm (${limitResult.criterion.source}, Cl. ${limitResult.criterion.clause})`;

  let comment = 'Enter the measured resistance to compare it against the calculated limit.';
  if (notMeasurable) {
    comment = 'Electrode could not be measured on site; excluded from the pass rate.';
  } else if (context.rho === null) {
    comment = 'Soil resistivity (measured, or assumed with its source) is required to calculate the electrode limit.';
  } else if (limit === null) {
    comment = 'Electrode type, length and diameter (or plate area) are required to calculate the electrode limit.';
  } else if (status.label === '>Permissible Limit') {
    comment = `Resistance ${fmt2(reading.value)} ohm exceeds the calculated limit of ${limitText}.`;
  } else if (status.label === 'Warning') {
    comment = `Resistance ${fmt2(reading.value)} ohm is above 90% of the calculated limit of ${limitText} but within it; monitor at the next test.`;
  } else if (status.label === 'Healthy') {
    comment = `Resistance ${fmt2(reading.value)} ohm is within 90% of the calculated limit of ${limitText}.`;
  }
  if (context.rhoAssumed && limit !== null) {
    comment += ` The limit uses an assumed soil resistivity of ${fmt2(context.rho)} ohm-m (${context.rhoSource}).`;
  }
  if (flags.length) {
    comment += ` ${flags[0].message}.`;
  }

  if (!notMeasurable && Number.isFinite(withoutGrid) && Number.isFinite(withGrid)) {
    if (withGrid > withoutGrid) {
      comment += ' The with-grid value is higher than the without-grid value, so verify the test setup and bonding path.';
    } else if (withGrid < withoutGrid) {
      comment += ' The with-grid value improved after bonding to the grid, which is the expected trend.';
    }
  }

  return {
    status,
    standard,
    measured: reading.value,
    measuredBasis: reading.basis,
    limit: limit === null ? null : round(limit, 2),
    limitCriterionId: limitResult ? limitResult.criterion.id : null,
    rho: context.rho,
    rhoAssumed: context.rhoAssumed,
    rhoSource: context.rhoSource,
    conditionCodes: codes,
    priority,
    fop,
    flags,
    comment
  };
}

// ---- Continuity, loop, PFC, riser -------------------------------------------
// Test-lead resistance is subtracted before comparison (IEEE 81-2012 Cl. 10.2).
// The test-lead resistance is recorded once in the sheet header (as printed).
function leadResistanceFor(context, sheetId, row) {
  // A value on the row wins: older reports store it there, and a final report stitched
  // from several days snapshots each day's header value onto the rows of that day.
  const own = asLooseNumber(row?.testLeadResistance);
  if (Number.isFinite(own)) {
    return own;
  }
  const header = asLooseNumber(context && context.headers && context.headers[sheetId] && context.headers[sheetId].testLeadResistance);
  return Number.isFinite(header) ? header : null;
}

function netOfLeadResistance(measured, lead) {
  if (!Number.isFinite(measured)) {
    return null;
  }
  const leadValue = Number.isFinite(lead) && lead > 0 ? lead : 0;
  return Math.max(0, round(measured - leadValue, 4));
}

function getLimitStatus(value, limit, { strict = false } = {}) {
  if (!Number.isFinite(value) || !Number.isFinite(limit)) {
    return pendingStatus();
  }
  return (strict ? value < limit : value <= limit) ? healthyStatus() : overLimitStatus();
}

// CONT-LIMIT is stated as "< 1 ohm", so it is a strict comparison. `value` is the
// reading already net of test-lead resistance; `options.limit` defaults to ETPL-DEFAULT.
function getContinuityStatus(value, options = {}) {
  return getLimitStatus(value, options.limit ?? resolveProfile().value('CONT-LIMIT'), { strict: true });
}

function getEarthContinuityStatus(value, options = {}) {
  return getContinuityStatus(asLooseNumber(value), options);
}

function getLoopImpedanceStatus(value, options = {}) {
  return getLimitStatus(value, options.limit ?? resolveProfile().value('LOOP-LIMIT'));
}

// Healthy only when both readings (net of lead resistance) are within the limit; one
// reading over the limit fails the point, one reading alone cannot pass it.
function getRiserStatus(equipment, grid, options = {}) {
  const limit = options.limit ?? resolveProfile().value('RISER-LIMIT');
  const values = [equipment, grid].filter((value) => Number.isFinite(value));
  if (!values.length) {
    return pendingStatus();
  }
  if (values.some((value) => value > limit)) {
    return overLimitStatus();
  }
  return values.length === 2 ? healthyStatus() : pendingStatus();
}

function deriveContinuityAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).continuityTest;
  const lead = leadResistanceFor(context, 'continuityTest', row);
  const net = netOfLeadResistance(asLooseNumber(row?.resistance), lead);
  const impedance = asLooseNumber(row?.impedance);
  const status = getContinuityStatus(net, { limit: context.profile.value('CONT-LIMIT') });

  let comment = `Enter a resistance reading to assess continuity against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Continuity resistance is within the ${standard.limitLabel}; bond/joint continuity is acceptable.`;
  } else if (status.tone === 'critical') {
    comment = `Continuity resistance is above the ${standard.limitLabel}; inspect joints, terminations, and bonding integrity.`;
  }
  if (Number.isFinite(lead) && lead > 0 && net !== null) {
    comment += ` Assessed value ${fmt2(net)} ohm after subtracting ${fmt2(lead)} ohm test-lead resistance.`;
  }
  if (Number.isFinite(impedance)) {
    comment += ` Recorded impedance: ${round(impedance, 2)} ohm.`;
  }

  return { status, standard, measured: net, conditionCodes: normalizeConditionCodes(row?.conditionCodes), comment };
}

function deriveLoopImpedanceAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).loopImpedanceTest;
  const measured = asLooseNumber(row?.loopImpedance) ?? asLooseNumber(row?.measuredZs);
  const status = getLoopImpedanceStatus(measured, { limit: context.profile.value('LOOP-LIMIT') });

  let comment = `Enter measured Zs to compare it with the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Measured Zs is within the ${standard.limitLabel}; a panel is healthy only if all three phases are.`;
  } else if (status.tone === 'critical') {
    comment = `Measured Zs exceeds the ${standard.limitLabel}; verify the PE path and confirm disconnection performance against the protective device.`;
  }

  return { status, standard, measured, comment };
}

// Device breaking capacity (kA) must exceed the higher of PFC (L-PE) and PSC (L-N).
// Without a recorded breaking capacity the point is Not Assessable, never Healthy.
function deriveFaultCurrentAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).prospectiveFaultCurrent;
  const pfc = asLooseNumber(row?.prospectiveFaultCurrent);
  const psc = asLooseNumber(row?.shortCircuitCurrentLN);
  const breakingCapacity = asLooseNumber(row?.breakingCapacity);
  const faultValues = [pfc, psc].filter((value) => Number.isFinite(value));
  const governing = faultValues.length ? Math.max(...faultValues) : null;
  let status = pendingStatus();
  let comment = `Enter the prospective fault current to verify that the ${standard.limitLabel}.`;

  if (governing !== null && !Number.isFinite(breakingCapacity)) {
    status = { label: 'Not Assessable', tone: 'neutral' };
    comment = 'Device breaking capacity was not recorded, so the fault current cannot be assessed; record the device nameplate kA.';
  } else if (governing !== null) {
    if (breakingCapacity > governing) {
      status = healthyStatus();
      comment = `Device breaking capacity ${fmt2(breakingCapacity)} kA exceeds the governing fault current of ${fmt2(governing)} kA.`;
    } else {
      status = overLimitStatus();
      comment = `Device breaking capacity ${fmt2(breakingCapacity)} kA does not exceed the governing fault current of ${fmt2(governing)} kA; review protective device selection.`;
    }
    if (!Number.isFinite(psc)) {
      comment += ' PSC not recorded; PFC (L-PE) assessed alone.';
    }
  }

  const pscNotRecorded = governing !== null && !Number.isFinite(psc);
  return { status, standard, measured: governing, pscNotRecorded, comment };
}

function deriveRiserAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).riserIntegrityTest;
  const lead = leadResistanceFor(context, 'riserIntegrityTest', row);
  const equipment = netOfLeadResistance(asLooseNumber(row?.resistanceTowardsEquipment), lead);
  const grid = netOfLeadResistance(asLooseNumber(row?.resistanceTowardsGrid), lead);
  const status = getRiserStatus(equipment, grid, { limit: context.profile.value('RISER-LIMIT') });

  let comment = `Enter both resistance readings to assess the riser against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Riser readings are within the ${standard.limitLabel} towards equipment and grid.`;
  } else if (status.tone === 'critical') {
    comment = `One or both riser readings exceed the ${standard.limitLabel}; inspect joints, lugs, and bonding interfaces.`;
  }
  if (Number.isFinite(lead) && lead > 0) {
    comment += ` Readings are net of ${fmt2(lead)} ohm test-lead resistance.`;
  }

  return { status, standard, conditionCodes: normalizeConditionCodes(row?.conditionCodes), comment };
}

function deriveEarthContinuityAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).earthContinuityTest;
  const lead = leadResistanceFor(context, 'continuityTest', row) ?? leadResistanceFor(context, 'earthContinuityTest', row);
  const net = netOfLeadResistance(asLooseNumber(row?.measuredValue), lead);
  const status = getContinuityStatus(net, { limit: context.profile.value('CONT-LIMIT') });

  let comment = `Enter a measured value to assess earth continuity against the ${standard.limitLabel}.`;
  if (status.tone === 'healthy') {
    comment = `Earth continuity is within the ${standard.limitLabel}; the earth path is acceptable for this point.`;
  } else if (status.tone === 'critical') {
    comment = `Earth continuity is above the ${standard.limitLabel}; inspect the earth path, joints, and terminations.`;
  }
  if (Number.isFinite(lead) && lead > 0 && net !== null) {
    comment += ` Assessed value ${fmt2(net)} ohm after subtracting ${fmt2(lead)} ohm test-lead resistance.`;
  }

  return { status, standard, measured: net, comment };
}

function buildTowerGroupKey(group) {
  const groupId = asTrimmedString(group?.groupId);
  const location = asTrimmedString(group?.mainLocationTower);
  return groupId || location || `group:${asTrimmedString(group?.srNo) || 'tower'}`;
}

function summarizeTowerGroups(groups) {
  const summaries = new Map();

  (Array.isArray(groups) ? groups : []).forEach((group) => {
    const readings = Array.isArray(group?.readings) ? group.readings : [];
    const impedanceValues = readings.map((reading) => asLooseNumber(reading?.measuredImpedance)).filter((value) => Number.isFinite(value));
    const currentValues = readings.map((reading) => asLooseNumber(reading?.measuredCurrentMa)).filter((value) => Number.isFinite(value));
    const key = buildTowerGroupKey(group);
    const hasAnyInput = Boolean(asTrimmedString(group?.mainLocationTower)) || readings.some((reading) => {
      return (
        normalizeConnectionStatus(reading?.footToEarthingConnectionStatus) === 'Not given'
      ) || Boolean(asTrimmedString(reading?.measuredCurrentMa)) || Boolean(asTrimmedString(reading?.measuredImpedance)) || Boolean(asTrimmedString(reading?.rowObservation)) || (Array.isArray(reading?.rowPhotos) && reading.rowPhotos.length);
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

// Zt = sum of the 4 feet. Healthy <= Zsat; Marginal <= 1.2 x Zsat; else Not Acceptable.
// Zsat comes from the project's profile (TOWER-ZSAT). Marginal is above Zsat, so it is
// not counted as within limit in the overall pass rate.
function deriveTowerFootingAssessment(group, groupSummary, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).towerFootingResistance;
  const zsat = context.profile.value('TOWER-ZSAT');
  const totalImpedanceZt = groupSummary?.totalImpedanceZt ?? null;
  const totalCurrentItotal = groupSummary?.totalCurrentItotal ?? null;
  let status = pendingStatus();
  let comment = '-';

  if (Number.isFinite(totalImpedanceZt)) {
    if (totalImpedanceZt <= zsat) {
      status = healthyStatus();
      comment = 'Healthy';
    } else if (totalImpedanceZt <= zsat * context.profile.params('TOWER-ZSAT').marginal_factor) {
      status = { label: 'Marginal', tone: 'warning' };
      comment = 'Marginal';
    } else {
      status = { label: 'Not Acceptable', tone: 'critical' };
      comment = 'Not Acceptable';
    }
  }

  // A foot marked "Not given" has no earth connection: reported as a defect, with the
  // remark and a CAP line (6C). It does not change the tower status, which stays on Zt.
  const readings = Array.isArray(group?.readings) ? group.readings : [];
  const notGiven = readings
    .map((reading, index) => ({ reading, index }))
    .filter(({ reading }) => normalizeConnectionStatus(reading?.footToEarthingConnectionStatus) === 'Not given')
    .map(({ index }) => `F-${index + 1}`);
  const flags = notGiven.length
    ? [{ code: 'FOOT_NOT_GIVEN', feet: notGiven, review: false, message: `Foot ${notGiven.join(', ')} has no earth connection (defect)` }]
    : [];
  const conditionCodes = [...new Set(readings.flatMap((reading) => normalizeConditionCodes(reading?.conditionCodes)))];

  return {
    status,
    standard,
    totalImpedanceZt,
    totalCurrentItotal,
    zsat: round(zsat, 2),
    flags,
    conditionCodes,
    defectText: flags.length ? `${flags[0].message}.` : '',
    comment
  };
}

// ---- Ground (grid) resistance, MS-07 -----------------------------------------
// Compared with the calculated limit of the pit sharing its tag (IS 3043:2018 Cl. 14.2),
// or the client's value when the profile gives one on GND-LIMIT. Healthy or
// >Permissible Limit: the pit warning band and P1/P2/P3 are electrode-only.
function deriveGroundResistanceAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).groundResistance;
  const reading = asLooseNumber(row?.r618);
  const tag = asTrimmedString(row?.tag);
  const clientValue = context.profile.get('GND-LIMIT').value;
  const pitLimit = tag ? context.pitLimits && context.pitLimits.get(tag.toLowerCase()) : undefined;
  const limit = Number.isFinite(clientValue) ? clientValue : Number.isFinite(pitLimit) ? pitLimit : null;
  const limitSource = Number.isFinite(clientValue) ? 'client value' : Number.isFinite(pitLimit) ? `limit of pit ${tag}` : null;
  const fop = checkFopAgreement({ r618: reading, r50: asLooseNumber(row?.r50), r70: asLooseNumber(row?.r70), method: row?.method }, context.profile);
  const flags = fop && !fop.agrees ? [{ code: 'FOP_DISAGREE', message: fopRemark(fop), review: true }] : [];

  let status = pendingStatus();
  let comment = 'Enter the resistance at 61.8% to compare it with the limit of the associated pit.';
  if (Number.isFinite(reading) && limit === null) {
    comment = tag
      ? `No pit with tag ${tag} has a calculable limit (soil resistivity and pit dimensions are needed), so this reading cannot be assessed.`
      : 'Enter the tag of the pit this ground belongs to so its limit can be used.';
  } else if (Number.isFinite(reading)) {
    status = getLimitStatus(reading, limit);
    comment =
      status.label === 'Healthy'
        ? `Ground resistance ${fmt2(reading)} ohm is within the ${fmt2(limit)} ohm ${limitSource}.`
        : `Ground resistance ${fmt2(reading)} ohm exceeds the ${fmt2(limit)} ohm ${limitSource}.`;
  }
  if (flags.length) {
    comment += ` ${flags[0].message}.`;
  }

  return {
    status,
    standard,
    measured: Number.isFinite(reading) ? reading : null,
    limit: limit === null ? null : round(limit, 2),
    limitSource,
    fop,
    flags,
    conditionCodes: normalizeConditionCodes(row?.conditionCodes),
    comment
  };
}

// ---- Step & touch (IS 3043:2018 Cl. 10.3, IEEE 81-2012 Cl. 9.3 g) -------------
// Pure calculations only: step & touch is not yet a selectable test in the app.
// Tolerable voltages for a 50 kg or 70 kg body. Cs = 1 and rho_s = rho when there is
// no surface layer. Returns null until rho_s and ts are known.
function calculateTolerableStepTouch({ surfaceResistivity, surfaceFactor = 1, faultClearingTime, bodyWeightKg = 70 }, profile = resolveProfile()) {
  const k = profile.params('ST-LIMIT')[`k_${bodyWeightKg}kg`];
  if (!k || !Number.isFinite(surfaceResistivity) || surfaceResistivity <= 0 || !Number.isFinite(faultClearingTime) || faultClearingTime <= 0) {
    return null;
  }
  const cs = Number.isFinite(surfaceFactor) ? surfaceFactor : 1;
  const root = Math.sqrt(faultClearingTime);
  return {
    k,
    step: ((1000 + 6 * cs * surfaceResistivity) * k) / root,
    touch: ((1000 + 1.5 * cs * surfaceResistivity) * k) / root
  };
}

// The measured voltage is scaled by If/It before comparison. A client profile may
// carry design-calculation limits on ST-LIMIT, in the fields named by its
// params.client_override_fields (step first, then touch); they win.
function assessStepTouchReading({ kind, measuredVolts, testCurrent, faultCurrent, tolerable }, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const criterion = context.profile.get('ST-LIMIT');
  const [stepField, touchField] = context.profile.params('ST-LIMIT').client_override_fields;
  const clientLimit = criterion[kind === 'step' ? stepField : touchField];
  const hasClientLimit = Number.isFinite(clientLimit);
  const limit = hasClientLimit ? clientLimit : tolerable ? tolerable[kind] : null;
  const valid = [measuredVolts, testCurrent, faultCurrent].every((value) => Number.isFinite(value)) && testCurrent > 0 && faultCurrent > 0;
  if (!valid || !Number.isFinite(limit)) {
    return { status: pendingStatus(), scaledVolts: null, limit: Number.isFinite(limit) ? limit : null, limitSource: null };
  }
  const scaledVolts = (measuredVolts * faultCurrent) / testCurrent;
  return {
    status: scaledVolts < limit ? healthyStatus() : overLimitStatus(),
    scaledVolts,
    limit,
    limitSource: hasClientLimit ? `${criterion.source}, Cl. ${criterion.clause}` : 'IS 3043:2018 Cl. 10.3'
  };
}

// Surface layer derating factor Cs. A client design value (header override) wins and is
// printed as client-specified; with no surface layer Cs = 1; otherwise the IEEE 80-2013
// formula named on ST-LIMIT params: Cs = 1 - 0.09 (1 - rho/rho_s) / (2 h_s + 0.09).
function resolveSurfaceFactor(header, siteRho, profile) {
  const params = profile.params('ST-LIMIT');
  const override = asLooseNumber(header && header[params.cs_override_field]);
  const surfaceRho = asLooseNumber(header?.surfaceResistivity);
  const thickness = asLooseNumber(header?.surfaceThicknessM);
  const rhoS = Number.isFinite(surfaceRho) && surfaceRho > 0 ? surfaceRho : Number.isFinite(siteRho) ? siteRho : null;
  if (Number.isFinite(override) && override > 0) {
    const source = asTrimmedString(header?.csSource);
    return { cs: override, rhoS, source: `client-specified${source ? ` (${source})` : ''}`, missing: rhoS === null ? ['surface or site soil resistivity'] : [] };
  }
  if (Number.isFinite(surfaceRho) && surfaceRho > 0 && Number.isFinite(thickness) && thickness > 0) {
    if (!Number.isFinite(siteRho) || siteRho <= 0) {
      return { cs: null, rhoS, source: params.cs_formula_reference, missing: ['site soil resistivity (for Cs)'] };
    }
    return { cs: 1 - (0.09 * (1 - siteRho / surfaceRho)) / (2 * thickness + 0.09), rhoS, source: params.cs_formula_reference, missing: [] };
  }
  return { cs: 1, rhoS, source: 'no surface layer (Cs = 1)', missing: rhoS === null ? ['site soil resistivity (rho_s = rho without a surface layer)'] : [] };
}

// ---- Step & touch voltage, MS-08 ----------------------------------------------
// Each row's Vmax (or the largest of V1-V3) is scaled by If/It and compared with the
// tolerable step or touch voltage for the header's surface layer and fault clearing
// time. With body weight "both" the stricter (50 kg) limit governs.
function deriveStepTouchAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).stepTouchVoltage;
  // `dayHeader` is the header of the day the row was measured on (final reports).
  const header = row?.dayHeader || (context.headers && context.headers.stepTouchVoltage) || {};
  const type = asTrimmedString(row?.measurementType);
  const kind = type === 'Step' ? 'step' : type.startsWith('Touch') ? 'touch' : null;
  const volts = ['v1', 'v2', 'v3'].map((key) => asLooseNumber(row?.[key])).filter((value) => Number.isFinite(value));
  const vmax = asLooseNumber(row?.vmax) ?? (volts.length ? Math.max(...volts) : null);
  const testCurrent = asLooseNumber(row?.testCurrentA);
  const faultCurrent = Number.isFinite(asLooseNumber(header.faultCurrentKA)) ? asLooseNumber(header.faultCurrentKA) * 1000 : null;
  const clearingTime = asLooseNumber(header.faultClearingTimeS);
  const basis = asTrimmedString(header.bodyWeightBasis);
  const weights = basis === 'both' ? [50, 70] : basis === '50' ? [50] : basis === '70' ? [70] : [];

  const surface = resolveSurfaceFactor(header, context.rho, context.profile);
  const missing = [];
  if (!kind) missing.push('measurement type');
  if (!Number.isFinite(vmax)) missing.push('voltage reading');
  if (!Number.isFinite(testCurrent)) missing.push('test current It');
  if (!Number.isFinite(faultCurrent)) missing.push('earth fault current If (header)');
  if (!Number.isFinite(clearingTime)) missing.push('fault clearing time ts (header)');
  if (!weights.length) missing.push('body weight basis (header)');
  missing.push(...surface.missing);

  const base = { standard, vmax: Number.isFinite(vmax) ? vmax : null, cs: surface.cs, csSource: surface.source, rhoS: surface.rhoS, missing, conditionCodes: [] };
  if (missing.length) {
    return {
      ...base,
      status: pendingStatus(),
      scaledVolts: null,
      limit: null,
      limitSource: null,
      bodyWeightKg: null,
      comment: Number.isFinite(vmax) || kind ? `Cannot be assessed yet: ${missing.join(', ')} not recorded.` : 'Enter the measurement type, test current and voltage readings.'
    };
  }

  const tolerables = weights.map((weight) => ({
    weight,
    tolerable: calculateTolerableStepTouch({ surfaceResistivity: surface.rhoS, surfaceFactor: surface.cs, faultClearingTime: clearingTime, bodyWeightKg: weight }, context.profile)
  }));
  const governing = tolerables.reduce((best, entry) => (!best || entry.tolerable[kind] < best.tolerable[kind] ? entry : best), null);
  const result = assessStepTouchReading({ kind, measuredVolts: vmax, testCurrent, faultCurrent, tolerable: governing.tolerable }, context);
  const comment =
    result.status.label === 'Healthy'
      ? `${vmax} V scaled by If/It is ${fmt2(result.scaledVolts)} V, below the ${fmt2(result.limit)} V ${kind} limit (${result.limitSource}; ${governing.weight} kg; Cs ${fmt2(surface.cs)}, ${surface.source}).`
      : `${vmax} V scaled by If/It is ${fmt2(result.scaledVolts)} V, which is not below the ${fmt2(result.limit)} V ${kind} limit (${result.limitSource}; ${governing.weight} kg; Cs ${fmt2(surface.cs)}, ${surface.source}).`;
  return {
    ...base,
    status: result.status,
    scaledVolts: result.scaledVolts,
    limit: round(result.limit, 2),
    limitSource: result.limitSource,
    bodyWeightKg: governing.weight,
    comment
  };
}

function valueHasAnyInput(value, key = '') {
  if (key === 'rowId') {
    return false;
  }
  if (Array.isArray(value)) {
    return value.some((entry) => valueHasAnyInput(entry));
  }
  if (value && typeof value === 'object') {
    return Object.entries(value).some(([childKey, childValue]) => valueHasAnyInput(childValue, childKey));
  }
  return asTrimmedString(value) !== '';
}

function rowHasAnyValue(row) {
  return Object.entries(row || {}).some(([key, value]) => valueHasAnyInput(value, key));
}

function normalizeObservationPhotos(photos) {
  return (Array.isArray(photos) ? photos : [])
    .map((photo) => ({
      id: asTrimmedString(photo?.id) || buildRowId('photo'),
      name: asTrimmedString(photo?.name) || 'Observation image',
      dataUrl: asTrimmedString(photo?.dataUrl)
    }))
    .filter((photo) => photo.dataUrl);
}

function normalizeRowMeta(row, prefix = 'row') {
  return {
    rowId: asTrimmedString(row?.rowId) || buildRowId(prefix),
    rowObservation: asTrimmedString(row?.rowObservation),
    rowPhotos: normalizeObservationPhotos(row?.rowPhotos)
  };
}

function normalizeSoilRows(rows) {
  return (Array.isArray(rows) ? rows : [])
    // The printed sheet pre-fills the spacings; a row only counts once it holds a
    // reading, a remark or an observation.
    .filter((row) => asTrimmedString(row?.resistivity) !== '' || asTrimmedString(row?.remark) !== '' || asTrimmedString(row?.rowObservation) !== '' || (Array.isArray(row?.rowPhotos) && row.rowPhotos.length > 0))
    .map((row) => ({
      ...normalizeRowMeta(row, 'soil'),
      spacing: asTextNumber(row.spacing),
      resistivity: asTextNumber(row.resistivity),
      remark: asTrimmedString(row.remark)
    }));
}

function normalizeSoilBearings(bearings) {
  return Object.fromEntries(SOIL_DIRECTION_KEYS.map((key) => [key, asTextNumber(bearings && bearings[key])]));
}

function soilLocationHasAnyValue(location) {
  return (
    Boolean(asTrimmedString(location?.name)) ||
    Boolean(asTrimmedString(location?.drivenElectrodeDiameter)) ||
    Boolean(asTrimmedString(location?.drivenElectrodeLength)) ||
    Boolean(asTrimmedString(location?.notes)) ||
    SOIL_DIRECTION_KEYS.some((key) => normalizeSoilRows(location?.[key]).length > 0)
  );
}

function normalizeSoilLocations(input) {
  const rawLocations =
    Array.isArray(input?.locations) && input.locations.length
      ? input.locations
      : [input || {}];

  return rawLocations
    .filter((location) => soilLocationHasAnyValue(location))
    .map((location, index) => ({
      locationId: asTrimmedString(location?.locationId) || buildRowId('soil-location'),
      name: asTrimmedString(location?.name) || `Location ${index + 1}`,
      ...Object.fromEntries(SOIL_DIRECTION_KEYS.map((key) => [key, normalizeSoilRows(location?.[key])])),
      inputUnit: asTrimmedString(location?.inputUnit) === 'ohm' ? 'ohm' : asTrimmedString(location?.inputUnit) === 'ohm-m' ? 'ohm-m' : '',
      probeDepthCm: asTextNumber(location?.probeDepthCm),
      weather: asTrimmedString(location?.weather),
      bearings: normalizeSoilBearings(location?.bearings),
      drivenElectrodeDiameter: asTrimmedString(location?.drivenElectrodeDiameter),
      drivenElectrodeLength: asTrimmedString(location?.drivenElectrodeLength),
      notes: asTrimmedString(location?.notes)
    }));
}

function normalizeRows(rows, fieldMap, rowPrefix = 'row', codeFields = []) {
  return (Array.isArray(rows) ? rows : [])
    .filter(rowHasAnyValue)
    .map((row) => {
      const normalized = normalizeRowMeta(row, rowPrefix);
      fieldMap.forEach((field) => {
        normalized[field] = asTrimmedString(row[field]);
      });
      codeFields.forEach((field) => {
        normalized[field] = normalizeConditionCodes(row[field]);
      });
      return normalized;
    });
}

function towerReadingHasAnyValue(reading) {
  return (
    Boolean(asTrimmedString(reading?.measuredCurrentMa)) ||
    Boolean(asTrimmedString(reading?.measuredImpedance)) ||
    Boolean(asTrimmedString(reading?.rowObservation)) ||
    (Array.isArray(reading?.rowPhotos) && reading.rowPhotos.length > 0) ||
    Boolean(asTrimmedString(reading?.photoRef)) ||
    Boolean(asTrimmedString(reading?.remark)) ||
    normalizeConditionCodes(reading?.conditionCodes).length > 0 ||
    normalizeConnectionStatus(reading?.footToEarthingConnectionStatus) !== ''
  );
}

// 'Given' | 'Not given'; older reports wrote 'Not Given'.
function normalizeConnectionStatus(value) {
  const text = asTrimmedString(value).toLowerCase();
  return text === 'not given' ? 'Not given' : text === 'given' ? 'Given' : '';
}

function normalizeTowerReadingFields(source) {
  return {
    footToEarthingConnectionStatus: normalizeConnectionStatus(source?.footToEarthingConnectionStatus),
    measuredCurrentMa: asTrimmedString(source?.measuredCurrentMa),
    measuredImpedance: asTrimmedString(source?.measuredImpedance),
    conditionCodes: normalizeConditionCodes(source?.conditionCodes),
    photoRef: asTrimmedString(source?.photoRef),
    remark: asTrimmedString(source?.remark)
  };
}

function towerGroupHasAnyValue(group) {
  return (
    Boolean(asTrimmedString(group?.mainLocationTower)) ||
    (Array.isArray(group?.readings) && group.readings.some((reading) => towerReadingHasAnyValue(reading)))
  );
}

function normalizeTowerFootingGroups(inputGroups) {
  const groups = Array.isArray(inputGroups) ? inputGroups : [];
  const looksGrouped = groups.some((group) => Array.isArray(group?.readings) || Object.prototype.hasOwnProperty.call(group || {}, 'groupId'));
  const normalizedGroups = groups
    .filter((group) => towerGroupHasAnyValue(group))
    .map((group, groupIndex) => {
      const readings = Array.isArray(group?.readings) ? group.readings : [];
      const normalizedReadings = TOWER_FOOT_POINTS.map((foot, readingIndex) => {
        const source =
          readings.find((reading) => asTrimmedString(reading?.measurementPointLocation) === foot) ||
          readings[readingIndex] ||
          {};
        return {
          ...normalizeRowMeta(source, 'tower'),
          measurementPointLocation: foot,
          ...normalizeTowerReadingFields(source)
        };
      });

      return {
        groupId: asTrimmedString(group?.groupId) || buildRowId('tower-group'),
        srNo: asTrimmedString(group?.srNo) || String(groupIndex + 1),
        mainLocationTower: asTrimmedString(group?.mainLocationTower),
        readings: normalizedReadings,
        totalImpedanceZt: '',
        totalCurrentItotal: '',
        standardTolerableImpedanceZsat: '',
        remarks: asTrimmedString(group?.remarks)
      };
    });

  if (normalizedGroups.length) {
    return normalizedGroups;
  }

  if (looksGrouped) {
    return [];
  }

  const legacyRows = normalizeRows(
    inputGroups,
    [
      'srNo',
      'mainLocationTower',
      'measurementPointLocation',
      'footToEarthingConnectionStatus',
      'measuredCurrentMa',
      'measuredImpedance',
      'totalImpedanceZt',
      'totalCurrentItotal',
      'standardTolerableImpedanceZsat',
      'remarks'
    ],
    'tower'
  );

  const grouped = new Map();
  legacyRows.forEach((row) => {
    const key = asTrimmedString(row?.srNo) || buildTowerGroupKey(row);
    if (!grouped.has(key)) {
      grouped.set(key, {
        groupId: buildRowId('tower-group'),
        srNo: asTrimmedString(row?.srNo) || String(grouped.size + 1),
        mainLocationTower: asTrimmedString(row?.mainLocationTower),
        readings: TOWER_FOOT_POINTS.map((foot) => {
          const reading = createDefaultTowerFootingReading(foot);
          reading.measurementPointLocation = foot;
          return reading;
        }),
        totalImpedanceZt: '',
        totalCurrentItotal: '',
        standardTolerableImpedanceZsat: '',
        remarks: ''
      });
    }
    const group = grouped.get(key);
    const footIndex = Math.max(0, TOWER_FOOT_POINTS.indexOf(asTrimmedString(row?.measurementPointLocation) || 'Foot-1'));
    group.mainLocationTower = group.mainLocationTower || asTrimmedString(row?.mainLocationTower);
    group.readings[footIndex] = {
      ...normalizeRowMeta(row, 'tower'),
      measurementPointLocation: TOWER_FOOT_POINTS[footIndex],
      ...normalizeTowerReadingFields(row)
    };
  });

  return Array.from(grouped.values());
}

function normalizeLoopRows(inputRows) {
  return (Array.isArray(inputRows) ? inputRows : [])
    .filter(rowHasAnyValue)
    .map((row, index) => {
      const normalized = normalizeRowMeta(row, 'loop');
      normalized.srNo = asTrimmedString(row?.srNo) || String(index + 1);
      normalized.location = asTrimmedString(row?.location || row?.mainLocation);
      normalized.feederTag = asTrimmedString(row?.feederTag || row?.panelEquipment);
      normalized.deviceType = asTrimmedString(row?.deviceType);
      normalized.deviceRating = asTrimmedString(row?.deviceRating);
      normalized.breakingCapacity = asTrimmedString(row?.breakingCapacity);
      normalized.measuredPoints = asTrimmedString(row?.measuredPoints) || PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length];
      normalized.tripCurve = asTrimmedString(row?.tripCurve);
      normalized.loopImpedance = asTrimmedString(row?.loopImpedance || row?.measuredZs);
      normalized.voltage = asTrimmedString(row?.voltage);
      normalized.photoRef = asTrimmedString(row?.photoRef);
      normalized.remarks = asTrimmedString(row?.remarks);
      return normalized;
    })
    .map((row, index) => ({
      ...row,
      srNo: String(index + 1),
      measuredPoints: PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length],
      remarks: engineerRemark(row.remarks)
    }));
}

function normalizeFaultRows(inputRows) {
  return (Array.isArray(inputRows) ? inputRows : [])
    .filter(rowHasAnyValue)
    .map((row, index) => {
      const normalized = normalizeRowMeta(row, 'fault');
      normalized.srNo = asTrimmedString(row?.srNo) || String(index + 1);
      normalized.location = asTrimmedString(row?.location || row?.mainLocation);
      normalized.feederTag = asTrimmedString(row?.feederTag || row?.panelEquipment);
      normalized.deviceType = asTrimmedString(row?.deviceType);
      normalized.deviceRating = asTrimmedString(row?.deviceRating);
      normalized.breakingCapacity = asTrimmedString(row?.breakingCapacity);
      normalized.measuredPoints = asTrimmedString(row?.measuredPoints) || PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length];
      normalized.loopImpedance = asTrimmedString(row?.loopImpedance || row?.measuredZs);
      normalized.prospectiveFaultCurrent = asTrimmedString(row?.prospectiveFaultCurrent);
      normalized.shortCircuitCurrentLN = asTrimmedString(row?.shortCircuitCurrentLN);
      normalized.voltage = asTrimmedString(row?.voltage);
      normalized.photoRef = asTrimmedString(row?.photoRef);
      normalized.comment = asTrimmedString(row?.comment);
      return normalized;
    })
    .map((row, index) => ({
      ...row,
      srNo: String(index + 1),
      measuredPoints: PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length],
      comment: engineerRemark(row.comment)
    }));
}

const LEGACY_MATERIALS = { copper: 'Cu', 'copper bonded': 'Cu-bonded' };

function normalizeMaterial(value) {
  const text = asTrimmedString(value);
  return LEGACY_MATERIALS[text.toLowerCase()] || text;
}

// Per-sheet header blocks as printed: instrument, serial, calibration due, times and the
// sheet's own header cells. Soil keeps its header on each location instead.
function normalizeSheetHeaders(input) {
  const headers = {};
  SHEETS.forEach((sheet) => {
    const source = input && typeof input[sheet.id] === 'object' && input[sheet.id] ? input[sheet.id] : null;
    if (!source || sheet.id === 'soilResistivity') {
      return;
    }
    const header = createDefaultSheetHeader(sheet.id);
    let hasInput = false;
    Object.keys(header).forEach((key) => {
      if (typeof header[key] === 'boolean') {
        header[key] = source[key] === true || source[key] === 'true';
      } else {
        header[key] = asTrimmedString(source[key]);
      }
      hasInput = hasInput || Boolean(header[key]);
    });
    if (hasInput) {
      headers[sheet.id] = header;
    }
  });
  return headers;
}

class ReportValidationError extends Error {
  constructor(errors) {
    super(errors.length > 1 ? `${errors[0].message} (${errors.length - 1} more problem${errors.length === 2 ? '' : 's'} on this report.)` : errors[0].message);
    this.name = 'ReportValidationError';
    this.errors = errors;
  }
}

// options.allowRetired: let a report select a retired test (only for reading old data).
function normalizeReportInput(payload, options = {}) {
  const draft = buildDefaultDraft();
  const input = payload || {};

  const project = {
    projectNo: asTrimmedString(input?.project?.projectNo),
    clientName: asTrimmedString(input?.project?.clientName),
    siteLocation: asTrimmedString(input?.project?.siteLocation),
    workOrder: asTrimmedString(input?.project?.workOrder),
    reportDate: asTrimmedString(input?.project?.reportDate),
    engineerName: asTrimmedString(input?.project?.engineerName),
    clientRepresentative: asTrimmedString(input?.project?.clientRepresentative),
    groundCondition: asTrimmedString(input?.project?.groundCondition),
    soilSpacingPreset: asTrimmedString(input?.project?.soilSpacingPreset) || DEFAULT_SOIL_PRESET,
    acceptanceProfile: asTrimmedString(input?.project?.acceptanceProfile) || DEFAULT_PROFILE_ID,
    assumedSoilResistivity: asTextNumber(input?.project?.assumedSoilResistivity),
    assumedSoilResistivitySource: asTrimmedString(input?.project?.assumedSoilResistivitySource),
    equipmentSelections: (Array.isArray(input?.project?.equipmentSelections) ? input.project.equipmentSelections : draft.project.equipmentSelections)
      .map((value) => asTrimmedString(value))
      .filter((value) => EQUIPMENT_LIBRARY.some((equipment) => equipment.id === value)),
    zohoProjectId: asTrimmedString(input?.project?.zohoProjectId),
    zohoProjectName: asTrimmedString(input?.project?.zohoProjectName),
    zohoProjectOwner: asTrimmedString(input?.project?.zohoProjectOwner),
    zohoProjectStage: asTrimmedString(input?.project?.zohoProjectStage)
  };
  // Final-report fields (6C): only kept when given, so older reports normalise exactly as before.
  ['unit', 'unitCode', 'clientId'].forEach((key) => {
    const raw = asTrimmedString(input?.project?.[key]);
    const value = key === 'unitCode' ? raw.replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(0, 8) : raw;
    if (value) {
      project[key] = value;
    }
  });
  const planned = {};
  Object.entries(input?.project?.plannedQuantities || {}).forEach(([testId, value]) => {
    const count = Number.parseInt(value, 10);
    if (TEST_LIBRARY.some((test) => test.id === testId) && Number.isFinite(count) && count >= 0) {
      planned[testId] = count;
    }
  });
  if (Object.keys(planned).length) {
    project.plannedQuantities = planned;
  }

  if (!project.equipmentSelections.length) {
    project.equipmentSelections = [...draft.project.equipmentSelections];
  }

  const tests = {};
  TEST_LIBRARY.forEach((test) => {
    tests[test.id] = Boolean(input?.tests?.[test.id]);
  });

  const selectedCount = TEST_LIBRARY.filter((test) => tests[test.id]).length;
  if (!selectedCount) {
    throw new Error('Select at least one measurement type.');
  }
  TEST_LIBRARY.filter((test) => test.retired && tests[test.id] && !options.allowRetired).forEach((test) => {
    throw new Error(`${test.label} is retired; record these points on Earth Continuity & Bonding (MS-03).`);
  });

  const requiredProjectFields = [
    ['projectNo', 'Project number is required.'],
    ['clientName', 'Client name is required.'],
    ['siteLocation', 'Site location is required.'],
    ['workOrder', 'Work order / reference is required.'],
    ['reportDate', 'Date of testing is required.'],
    ['engineerName', 'Testing engineer is required.']
  ];

  requiredProjectFields.forEach(([field, message]) => {
    if (!project[field]) {
      throw new Error(message);
    }
  });

  // An unknown profile fails loudly rather than printing a source that is not the project's.
  resolveProfile(project.acceptanceProfile);
  if (project.assumedSoilResistivity && !project.assumedSoilResistivitySource) {
    throw new Error('An assumed soil resistivity must state its source.');
  }

  const soilResistivity = { locations: normalizeSoilLocations(input?.soilResistivity) };
  const context = buildAssessmentContext({ project, soilResistivity });

  const normalizedTowerFootingGroups = normalizeTowerFootingGroups(input?.towerFootingResistance).map((group, index) => ({
    ...group,
    srNo: asTrimmedString(group.srNo) || String(index + 1),
    standardTolerableImpedanceZsat: String(context.profile.value('TOWER-ZSAT'))
  }));

  const towerGroupSummaries = summarizeTowerGroups(normalizedTowerFootingGroups);

  const report = {
    project,
    tests,
    sheetHeaders: normalizeSheetHeaders(input?.sheetHeaders),
    soilResistivity,
    electrodeResistance: normalizeRows(
      input?.electrodeResistance,
      [
      'tag',
      'location',
      'electrodeType',
      'materialType',
      'length',
      'diameter',
      'plateLengthM',
      'plateWidthM',
      'plateArea',
      'resistanceWithoutGrid',
      'checkR50',
      'checkR70',
      'resistanceWithGrid',
      'method',
      'photoRef',
      'observation'
      ],
      'electrode',
      ['conditionCodes']
    ).map((row, index) => {
      const source = Array.isArray(input?.electrodeResistance) ? input.electrodeResistance[index] : null;
      const resistanceWithGrid = row.resistanceWithGrid || asTrimmedString(source?.measuredResistance);
      const normalizedRow = {
        ...row,
        materialType: normalizeMaterial(row.materialType),
        resistanceWithGrid,
        observation: asTrimmedString(source?.observation)
      };
      if (source?.notMeasurable === true || source?.notMeasurable === 'true') {
        normalizedRow.notMeasurable = true;
      }
      return normalizedRow;
    }),
    continuityTest: normalizeRows(
      input?.continuityTest,
      [
      'srNo',
      'tag',
      'mainLocation',
      'referencePoint',
      'measurementPoint',
      'distanceM',
      'testCurrentA',
      'resistance',
      'impedance',
      'testLeadResistance',
      'photoRef',
      'comment'
      ],
      'continuity',
      ['conditionCodes']
    ).map((row) => ({
      ...row,
      comment: engineerRemark(row.comment)
    })),
    loopImpedanceTest: normalizeLoopRows(input?.loopImpedanceTest),
    prospectiveFaultCurrent: normalizeFaultRows(input?.prospectiveFaultCurrent),
    riserIntegrityTest: normalizeRows(
      input?.riserIntegrityTest,
      [
      'srNo',
      'mainLocation',
      'measurementPoint',
      'resistanceTowardsEquipment',
      'resistanceTowardsGrid',
      'testLeadResistance',
      'photoRef',
      'comment'
      ],
      'riser',
      ['conditionCodes']
    ).map((row) => ({
      ...row,
      comment: engineerRemark(row.comment)
    })),
    groundResistance: normalizeRows(
      input?.groundResistance,
      ['srNo', 'tag', 'location', 'method', 'currentProbeDistanceM', 'r50', 'r618', 'r70', 'remark'],
      'ground',
      ['conditionCodes']
    ),
    stepTouchVoltage: normalizeRows(
      input?.stepTouchVoltage,
      ['srNo', 'location', 'measurementType', 'surface', 'distanceM', 'testCurrentA', 'v1', 'v2', 'v3', 'vmax', 'photoRef', 'remark'],
      'steptouch'
    ),
    earthContinuityTest: normalizeRows(
      input?.earthContinuityTest,
      [
      'srNo',
      'tag',
      'locationBuildingName',
      'distance',
      'measuredValue',
      'testLeadResistance',
      'remark'
      ],
      'earth'
    ).map((row) => ({
      ...row,
      remark: engineerRemark(row.remark)
    })),
    towerFootingResistance: normalizedTowerFootingGroups.map((group) => {
      const summary = towerGroupSummaries.get(buildTowerGroupKey(group));
      const assessment = deriveTowerFootingAssessment(group, summary, context);
      return {
        ...group,
        totalImpedanceZt: assessment.totalImpedanceZt === null ? '' : String(assessment.totalImpedanceZt),
        totalCurrentItotal: assessment.totalCurrentItotal === null ? '' : String(assessment.totalCurrentItotal),
        standardTolerableImpedanceZsat: String(assessment.zsat),
        remarks: engineerRemark(group.remarks)
      };
    }),
    ocrImports: normalizeOcrImports(input?.ocrImports)
  };

  if (tests.soilResistivity) {
    if (!report.soilResistivity.locations.length) {
      throw new Error('Soil resistivity requires at least one named location.');
    }
    report.soilResistivity.locations.forEach((location, index) => {
      if (!location.direction1.length || !location.direction2.length) {
        throw new Error(`Soil resistivity location ${index + 1} requires readings in both Direction 1 and Direction 2.`);
      }
    });
  }

  if (tests.towerFootingResistance) {
    report.towerFootingResistance.forEach((group, index) => {
      if (!group.mainLocationTower) {
        throw new Error(`Tower footing location ${index + 1} requires Main Location – Tower.`);
      }
      if (!Array.isArray(group.readings) || group.readings.length !== TOWER_FOOT_POINTS.length) {
        throw new Error(`Tower footing location ${index + 1} must contain exactly 4 footing rows.`);
      }
      group.readings.forEach((reading, readingIndex) => {
        if (asTrimmedString(reading.measurementPointLocation) !== TOWER_FOOT_POINTS[readingIndex]) {
          throw new Error(`Tower footing location ${index + 1} must keep fixed footing order Foot-1 to Foot-4.`);
        }
        if (asTrimmedString(reading.measuredCurrentMa) && !Number.isFinite(asLooseNumber(reading.measuredCurrentMa))) {
          throw new Error(`Tower footing location ${index + 1} ${TOWER_FOOT_POINTS[readingIndex]} requires numeric Measured Current I (mA).`);
        }
        if (asTrimmedString(reading.measuredImpedance) && !Number.isFinite(asLooseNumber(reading.measuredImpedance))) {
          throw new Error(`Tower footing location ${index + 1} ${TOWER_FOOT_POINTS[readingIndex]} requires numeric Measured Impedance (ohm).`);
        }
      });
    });
  }

  TEST_LIBRARY.forEach((test) => {
    if (test.id === 'soilResistivity') {
      return;
    }
    if (tests[test.id] && !report[test.id].length) {
      throw new Error(`${test.label} requires at least one row.`);
    }
  });

  // Field-level checks from the sheet definitions. Errors are invalid data and block the
  // save; warnings (incomplete, out of range, unconfirmed) are computed again on every
  // assessment and never stored.
  const { errors } = validateReport(report, buildAssessmentContext(report));
  if (errors.length) {
    throw new ReportValidationError(errors);
  }

  return report;
}

// rho = 2*pi*a*R (IS 3043:2018 Cl. 40.3) when the sheet records R in ohm; otherwise the
// entered value is already rho in ohm-m.
function readSoilRho(row, inputUnit) {
  const value = asLooseNumber(row?.resistivity);
  if (inputUnit !== 'ohm' || !Number.isFinite(value)) {
    return value;
  }
  const spacing = asLooseNumber(row?.spacing);
  return Number.isFinite(spacing) && spacing > 0 ? 2 * Math.PI * spacing * value : null;
}

function calculateSoilSummary(report) {
  const rawLocations =
    Array.isArray(report?.soilResistivity?.locations) && report.soilResistivity.locations.length
      ? report.soilResistivity.locations
      : [
          {
            name: 'Location 1',
            inputUnit: 'ohm-m', // the old single-location shape predates the unit field: values are rho
            direction1: report?.soilResistivity?.direction1 || [],
            direction2: report?.soilResistivity?.direction2 || [],
            drivenElectrodeDiameter: report?.soilResistivity?.drivenElectrodeDiameter || '',
            drivenElectrodeLength: report?.soilResistivity?.drivenElectrodeLength || '',
            notes: report?.soilResistivity?.notes || ''
          }
        ];

  // Site averages are taken from the unrounded location averages, so the printed
  // site rho (and every limit derived from it) does not pick up rounding error.
  const soilProfile = resolveProfile(report?.project?.acceptanceProfile);
  const exact = [];
  const locationSummaries = rawLocations.map((location, index) => {
    const unitText = asTrimmedString(location?.inputUnit);
    const inputUnit = unitText === 'ohm' ? 'ohm' : unitText === 'ohm-m' ? 'ohm-m' : '';
    // Without a ticked reading unit the values cannot be read as R or as rho, so the location
    // is Pending: no rho, no corrosivity class, and it stays out of the site average (and so
    // out of the electrode limits that use it).
    const unitMissing = inputUnit === '';
    // Mean of each direction's readings, then the mean of the direction means
    // (up to 8 directions; 2 perpendicular directions minimum).
    const directionAverages = Object.fromEntries(
      SOIL_DIRECTION_KEYS.map((key) => [
        key,
        unitMissing ? null : average((location[key] || []).map((row) => readSoilRho(row, inputUnit)).filter((value) => Number.isFinite(value)))
      ])
    );
    const direction1Average = directionAverages.direction1;
    const direction2Average = directionAverages.direction2;
    const overallAverage = average(Object.values(directionAverages).filter((value) => Number.isFinite(value)));
    const category = getSoilCategory(overallAverage, soilProfile);
    exact.push({ direction1Average, direction2Average, overallAverage });

    return {
      locationId: asTrimmedString(location?.locationId) || `soil-location-${index + 1}`,
      name: asTrimmedString(location?.name) || `Location ${index + 1}`,
      direction1Average: round(direction1Average, 2),
      direction2Average: round(direction2Average, 2),
      overallAverage: round(overallAverage, 2),
      directionAverages: Object.fromEntries(Object.entries(directionAverages).map(([key, value]) => [key, round(value, 2)])),
      category,
      inputUnit,
      pending: unitMissing,
      pendingReason: unitMissing ? 'reading unit not selected' : '',
      drivenElectrodeDiameter: asTrimmedString(location?.drivenElectrodeDiameter),
      drivenElectrodeLength: asTrimmedString(location?.drivenElectrodeLength),
      notes: asTrimmedString(location?.notes)
    };
  });

  const direction1Average = average(exact.map((location) => location.direction1Average).filter((value) => Number.isFinite(value)));
  const direction2Average = average(exact.map((location) => location.direction2Average).filter((value) => Number.isFinite(value)));
  const overallAverage = average(exact.map((location) => location.overallAverage).filter((value) => Number.isFinite(value)));
  const category = getSoilCategory(overallAverage, soilProfile);

  return {
    direction1Average: round(direction1Average, 2),
    direction2Average: round(direction2Average, 2),
    overallAverage: round(overallAverage, 2),
    category,
    pendingLocations: locationSummaries.filter((location) => location.pending && SOIL_DIRECTION_KEYS.some((key) => ((rawLocations[locationSummaries.indexOf(location)] || {})[key] || []).some((row) => asTrimmedString(row?.resistivity)))).map((location) => location.name),
    locations: locationSummaries
  };
}


// A panel is one location's three phase readings (R-E, Y-E, B-E), stored as
// consecutive rows. It is healthy only if all 3 phases are present and healthy, and
// above limit if any phase is. A panel still being filled in counts against the pass
// rate; one with nothing entered is Pending, and one with no breaking capacity (PFC)
// is Not Assessable and stays out of it.
function summarizePanels(rows, deriveAssessment, contextOrReport) {
  const list = Array.isArray(rows) ? rows : [];
  const context = resolveContext(contextOrReport);
  const panels = [];
  for (let start = 0; start < list.length; start += PHASE_MEASURED_POINTS.length) {
    const phases = list.slice(start, start + PHASE_MEASURED_POINTS.length).map((row) => deriveAssessment(row, context).status);
    const complete = phases.length === PHASE_MEASURED_POINTS.length;
    const allHealthy = complete && phases.every((status) => status.tone === 'healthy');
    const anyOver = phases.some((status) => status.tone === 'critical');
    const allNeutral = phases.length > 0 && phases.every((status) => status.tone === 'neutral');
    const allNotAssessable = phases.length > 0 && phases.every((status) => status.label === 'Not Assessable');
    let status = { label: 'Incomplete', tone: 'warning' };
    if (anyOver) {
      status = overLimitStatus();
    } else if (allHealthy) {
      status = healthyStatus();
    } else if (allNotAssessable) {
      status = { label: 'Not Assessable', tone: 'neutral' };
    } else if (allNeutral) {
      status = pendingStatus();
    }
    panels.push({ startIndex: start, phases, status });
  }
  return panels;
}

// Overall status from a list of pass/fail units ({ tone, withinLimit? }). Neutral units
// (pending, not measurable, not assessable) are listed separately and stay out of the
// pass rate. A unit passes when it is Healthy or a Warning (within limit, `withinLimit`).
//   0 unhealthy            -> "Healthy"
//   pass rate >= 90%       -> "Healthy - Minor Actions"
//   otherwise              -> ">Attention Required"
function summarizeOverallStatus(units, profile = resolveProfile()) {
  const { minor_actions_min_pass: minorActionsMinPass, labels } = profile.params('OVERALL');
  const all = Array.isArray(units) ? units : [];
  const measured = all.filter((unit) => {
    return unit && ['healthy', 'warning', 'critical'].includes(unit.tone);
  });
  const healthy = measured.filter((unit) => unit.tone === 'healthy' || unit.withinLimit === true).length;
  const unhealthy = measured.length - healthy;
  const passRate = measured.length ? healthy / measured.length : null;

  let label = 'Pending';
  let tone = 'neutral';
  if (measured.length) {
    if (unhealthy === 0) {
      label = labels.healthy;
      tone = 'healthy';
    } else if (passRate >= minorActionsMinPass) {
      label = labels.minor;
      tone = 'warning';
    } else {
      label = labels.attention;
      tone = 'critical';
    }
  }
  return {
    label,
    tone,
    total: measured.length,
    healthy,
    unhealthy,
    notAssessed: all.filter(Boolean).length - measured.length,
    passRate: passRate === null ? null : round(passRate, 4)
  };
}

// Pass/fail units per selected test. A loop or PFC unit is a panel (3 phases); soil has
// none because it carries no pass/fail.
function collectTestUnits(report, context = buildAssessmentContext(report)) {
  const units = {};
  const collect = (id, rows, derive) => {
    if (report?.tests?.[id]) {
      units[id] = (Array.isArray(rows) ? rows : []).map((row) => derive(row, context).status);
    }
  };

  collect('electrodeResistance', report?.electrodeResistance, deriveElectrodeAssessment);
  collect('continuityTest', report?.continuityTest, deriveContinuityAssessment);
  if (report?.tests?.loopImpedanceTest) {
    units.loopImpedanceTest = summarizePanels(report.loopImpedanceTest, deriveLoopImpedanceAssessment, context).map((panel) => panel.status);
  }
  if (report?.tests?.prospectiveFaultCurrent) {
    units.prospectiveFaultCurrent = summarizePanels(report.prospectiveFaultCurrent, deriveFaultCurrentAssessment, context).map((panel) => panel.status);
  }
  collect('riserIntegrityTest', report?.riserIntegrityTest, deriveRiserAssessment);
  collect('groundResistance', report?.groundResistance, deriveGroundResistanceAssessment);
  collect('stepTouchVoltage', report?.stepTouchVoltage, deriveStepTouchAssessment);
  collect('earthContinuityTest', report?.earthContinuityTest, deriveEarthContinuityAssessment);
  if (report?.tests?.towerFootingResistance) {
    const summaries = summarizeTowerGroups(report.towerFootingResistance);
    units.towerFootingResistance = (report.towerFootingResistance || []).map((group) => {
      return deriveTowerFootingAssessment(group, summaries.get(buildTowerGroupKey(group)), context).status;
    });
  }
  return units;
}

// Overall status of each selected test (standard 2.10), keyed by test id.
function buildTestSummaries(report, context = buildAssessmentContext(report)) {
  const summaries = {};
  Object.entries(collectTestUnits(report, context)).forEach(([id, units]) => {
    summaries[id] = summarizeOverallStatus(units, context.profile);
  });
  return summaries;
}

function buildOverallStatus(report, context = buildAssessmentContext(report)) {
  return summarizeOverallStatus(Object.values(collectTestUnits(report, context)).flat(), context.profile);
}

function buildExecutiveSnapshot(report) {
  const soil = calculateSoilSummary(report);
  const selectedTests = TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);

  const context = buildAssessmentContext(report);
  const healthyElectrodes = (report?.electrodeResistance || []).filter((row) => {
    return deriveElectrodeAssessment(row, context).status.tone === 'healthy';
  }).length;

  return {
    selectedTests,
    soil,
    healthyElectrodes,
    totalElectrodes: (report?.electrodeResistance || []).length,
    overallStatus: buildOverallStatus(report)
  };
}

function normalizeOcrImports(input) {
  const source = input && typeof input === 'object' ? input : {};
  const normalized = {};

  TEST_LIBRARY.forEach((test) => {
    const record = source[test.id];
    if (!record || typeof record !== 'object') {
      return;
    }

    const warnings = (Array.isArray(record.warnings) ? record.warnings : [])
      .map((value) => asTrimmedString(value))
      .filter(Boolean);

    const uncertainFields = (Array.isArray(record.uncertainFields) ? record.uncertainFields : [])
      .map((item) => ({
        path: asTrimmedString(item?.path),
        reason: asTrimmedString(item?.reason)
      }))
      .filter((item) => item.path || item.reason);

    const fileSize = Number.parseInt(String(record.fileSize ?? ''), 10);
    const extractedCount = Number.parseInt(String(record.extractedCount ?? ''), 10);

    normalized[test.id] = {
      sheetId: test.id,
      sheetLabel: asTrimmedString(record.sheetLabel) || test.label,
      fileName: asTrimmedString(record.fileName),
      mimeType: asTrimmedString(record.mimeType),
      fileSize: Number.isFinite(fileSize) ? fileSize : null,
      model: asTrimmedString(record.model),
      scannedAt: asTrimmedString(record.scannedAt),
      appliedAt: asTrimmedString(record.appliedAt),
      extractedCount: Number.isFinite(extractedCount) ? extractedCount : null,
      warnings,
      uncertainFields
    };
  });

  return normalized;
}

module.exports = {
  ReportValidationError,
  normalizeSheetHeaders,
  createDefaultSheetHeader,
  createDefaultGroundResistanceRow,
  createDefaultStepTouchRow,
  TEST_LIBRARY,
  EQUIPMENT_LIBRARY,
  buildDefaultDraft,
  normalizeReportInput,
  calculateSoilSummary,
  calculateDrivenElectrodeResistance,
  getSoilCategory,
  getCorrosionAllowancePct,
  getElectrodeStatus,
  getContinuityStatus,
  getLoopImpedanceStatus,
  getRiserStatus,
  getEarthContinuityStatus,
  deriveElectrodeAssessment,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveEarthContinuityAssessment,
  deriveTowerFootingAssessment,
  deriveGroundResistanceAssessment,
  deriveStepTouchAssessment,
  checkFopAgreement,
  plateAreaOf,
  resolveSurfaceFactor,
  summarizeTowerGroups,
  buildTowerGroupKey,
  getStandardGuidance,
  buildAssessmentContext,
  calculateElectrodeLimit,
  getElectrodeAssessedReading,
  deriveElectrodePriority,
  normalizeConditionCodes,
  calculateTolerableStepTouch,
  assessStepTouchReading,
  buildExecutiveSnapshot,
  summarizePanels,
  summarizeOverallStatus,
  buildOverallStatus,
  buildTestSummaries,
  normalizeOcrImports,
  createDefaultSoilRow,
  createDefaultSoilLocation,
  createDefaultElectrodeRow,
  createDefaultContinuityRow,
  createDefaultLoopImpedanceRow,
  createDefaultProspectiveFaultRow,
  createDefaultRiserIntegrityRow,
  createDefaultEarthContinuityRow,
  createDefaultTowerFootingGroup,
  asLooseNumber,
  getElectrodeMeasuredValue,
  round
};
