// All limits, formulas and classifications come from the project's acceptance
// profile (ET-D03-REG-002, see ./acceptance.js). The only numbers left in this
// file are ones the register states as prose, not as a value: the 0.9 band edge
// (PIT-BAND), the 1.2 marginal factor (TOWER-ZSAT), the 90% pass rate (OVERALL) and
// the Table 4 corrosivity bands (SOIL-CLASS). Making those machine-readable needs a
// controlled revision of ET-D03-REG-002.
const { resolveProfile, BASIS_LABELS, DEFAULT_PROFILE_ID } = require('./acceptance');

const PIT_BAND_FACTOR = 0.9;
const TOWER_MARGINAL_FACTOR = 1.2;
const OVERALL_PASS_THRESHOLD = 0.9;

const CONDITION_CODES = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'OTH'];
const PRIORITY_P1_CODES = new Set(['C1', 'C2', 'C4', 'C5']);
const PRIORITY_P2_CODES = new Set(['C3', 'C6', 'C7', 'C8']);

const TEST_LIBRARY = [
  {
    id: 'soilResistivity',
    label: 'Soil Resistivity Test',
    shortLabel: 'Soil',
    description: 'Direction-wise soil resistivity readings with automatic average and IS 3043 corrosivity class.'
  },
  {
    id: 'electrodeResistance',
    label: 'Earth Electrode Resistance Test',
    shortLabel: 'Electrode',
    description: 'Pit-wise resistance against the IS 3043:2018 Cl. 14.2 limit calculated from site soil resistivity.'
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
    description: 'Prospective fault current and related feeder details.'
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
    description: 'Earth path continuity by tag, location, and distance.'
  },
  {
    id: 'towerFootingResistance',
    label: 'Tower Footing Resistance Measurement & Analysis',
    shortLabel: 'Tower Footing',
    description: 'Grouped tower footing impedance and current analysis with per-tower totals.'
  }
];

const SOIL_SPACING_PRESETS = ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0', '4.5', '5.0'];
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
    resistivity
  };
}

function createDefaultSoilLocation(name = 'Location 1') {
  return {
    locationId: buildRowId('soil-location'),
    name,
    direction1: SOIL_SPACING_PRESETS.slice(0, 6).map((spacing) => createDefaultSoilRow(spacing, '')),
    direction2: SOIL_SPACING_PRESETS.slice(0, 6).map((spacing) => createDefaultSoilRow(spacing, '')),
    // 'ohm-m': values are rho as read. 'ohm': values are R as read, rho = 2*pi*a*R.
    inputUnit: 'ohm-m',
    drivenElectrodeDiameter: '',
    drivenElectrodeLength: '',
    notes: ''
  };
}

function createDefaultElectrodeRow() {
  return {
    ...createRowMeta('electrode'),
    tag: '',
    location: '',
    electrodeType: 'Rod',
    materialType: 'Copper',
    length: '',
    diameter: '',
    plateArea: '',
    resistanceWithoutGrid: '',
    resistanceWithGrid: '',
    conditionCodes: [],
    observation: ''
  };
}

function createDefaultContinuityRow() {
  return {
    ...createRowMeta('continuity'),
    srNo: '',
    mainLocation: '',
    measurementPoint: '',
    resistance: '',
    impedance: '',
    testLeadResistance: '',
    comment: ''
  };
}

function createDefaultLoopImpedanceRow() {
  return {
    ...createRowMeta('loop'),
    srNo: '',
    location: '',
    feederTag: '',
    deviceType: 'MCB',
    deviceRating: '',
    breakingCapacity: '',
    measuredPoints: 'R-E',
    loopImpedance: '',
    voltage: '230',
    remarks: ''
  };
}

function createDefaultProspectiveFaultRow() {
  return {
    ...createRowMeta('fault'),
    srNo: '',
    location: '',
    feederTag: '',
    deviceType: 'MCB',
    deviceRating: '',
    breakingCapacity: '',
    measuredPoints: 'R-E',
    loopImpedance: '',
    prospectiveFaultCurrent: '',
    shortCircuitCurrentLN: '',
    voltage: '230',
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
    comment: ''
  };
}

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
    footToEarthingConnectionStatus: 'Given',
    measuredCurrentMa: '',
    measuredImpedance: '',
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

function buildDefaultDraft() {
  return {
    project: {
      projectNo: '',
      clientName: '',
      siteLocation: '',
      workOrder: '',
      reportDate: new Date().toISOString().slice(0, 10),
      engineerName: '',
      equipmentSelections: EQUIPMENT_LIBRARY.map((equipment) => equipment.id),
      acceptanceProfile: 'ETPL-DEFAULT',
      assumedSoilResistivity: '',
      assumedSoilResistivitySource: '',
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
      locations: [createDefaultSoilLocation('Location 1')]
    },
    electrodeResistance: [createDefaultElectrodeRow()],
    continuityTest: [createDefaultContinuityRow()],
    loopImpedanceTest: createDefaultLoopGroupRows(1),
    prospectiveFaultCurrent: createDefaultFaultGroupRows(1),
    riserIntegrityTest: [createDefaultRiserIntegrityRow()],
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
  return { profile, rho, rhoAssumed, rhoSource };
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
    towerFootingResistance: { reference: cite('TOWER-ZSAT'), limitLabel: valueLabel('TOWER-ZSAT', 'Zsat') }
  };
}

// ---- Soil -------------------------------------------------------------------
// IS 3043:2018 Cl. 13.6.1, Table 4 (SOIL-CLASS). Corrosivity only: no pass/fail.
function getSoilCategory(value) {
  if (!Number.isFinite(value)) {
    return { label: 'Insufficient Data', tone: 'neutral' };
  }
  const label =
    value < 25 ? 'Severely corrosive' : value < 50 ? 'Moderately corrosive' : value < 100 ? 'Mildly corrosive' : 'Very mildly corrosive';
  return { label, tone: 'neutral', reference: 'IS 3043:2018 Cl. 13.6.1, Table 4' };
}

// ---- Earth electrode (pit) --------------------------------------------------
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
  if (value > PIT_BAND_FACTOR * limit) {
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
      areaM2: asLooseNumber(row?.plateArea),
      rho: context.rho
    },
    context.profile
  );
  const limit = limitResult ? limitResult.limit : null;
  const status = getElectrodeStatus(reading.value, { notMeasurable, limit });
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
    comment
  };
}

// ---- Continuity, loop, PFC, riser -------------------------------------------
// Test-lead resistance is subtracted before comparison (IEEE 81-2012 Cl. 10.2).
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
  const lead = asLooseNumber(row?.testLeadResistance);
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

  return { status, standard, measured: net, comment };
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
  const lead = asLooseNumber(row?.testLeadResistance);
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

  return { status, standard, comment };
}

function deriveEarthContinuityAssessment(row, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const standard = getStandardGuidance(context.profile).earthContinuityTest;
  const lead = asLooseNumber(row?.testLeadResistance);
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
        asTrimmedString(reading?.footToEarthingConnectionStatus) !== '' && asTrimmedString(reading?.footToEarthingConnectionStatus) !== 'Given'
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
    } else if (totalImpedanceZt <= zsat * TOWER_MARGINAL_FACTOR) {
      status = { label: 'Marginal', tone: 'warning' };
      comment = 'Marginal';
    } else {
      status = { label: 'Not Acceptable', tone: 'critical' };
      comment = 'Not Acceptable';
    }
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

// ---- Step & touch (IS 3043:2018 Cl. 10.3, IEEE 81-2012 Cl. 9.3 g) -------------
// Pure calculations only: step & touch is not yet a selectable test in the app.
const STEP_TOUCH_K = { 50: 0.116, 70: 0.157 };

// Tolerable voltages for a 50 kg or 70 kg body. Cs = 1 and rho_s = rho when there is
// no surface layer. Returns null until rho_s and ts are known.
function calculateTolerableStepTouch({ surfaceResistivity, surfaceFactor = 1, faultClearingTime, bodyWeightKg = 70 }) {
  const k = STEP_TOUCH_K[bodyWeightKg];
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
// carry design-calculation limits on ST-LIMIT (stepLimitV / touchLimitV); they win.
function assessStepTouchReading({ kind, measuredVolts, testCurrent, faultCurrent, tolerable }, contextOrReport) {
  const context = resolveContext(contextOrReport);
  const criterion = context.profile.get('ST-LIMIT');
  const clientLimit = kind === 'step' ? criterion.stepLimitV : criterion.touchLimitV;
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
    .filter(rowHasAnyValue)
    .map((row) => ({
      ...normalizeRowMeta(row, 'soil'),
      spacing: asTextNumber(row.spacing),
      resistivity: asTextNumber(row.resistivity)
    }));
}

function soilLocationHasAnyValue(location) {
  return (
    Boolean(asTrimmedString(location?.name)) ||
    Boolean(asTrimmedString(location?.drivenElectrodeDiameter)) ||
    Boolean(asTrimmedString(location?.drivenElectrodeLength)) ||
    Boolean(asTrimmedString(location?.notes)) ||
    normalizeSoilRows(location?.direction1).length > 0 ||
    normalizeSoilRows(location?.direction2).length > 0
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
      direction1: normalizeSoilRows(location?.direction1),
      direction2: normalizeSoilRows(location?.direction2),
      inputUnit: asTrimmedString(location?.inputUnit) === 'ohm' ? 'ohm' : 'ohm-m',
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
    (asTrimmedString(reading?.footToEarthingConnectionStatus) && asTrimmedString(reading?.footToEarthingConnectionStatus) !== 'Given')
  );
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
          footToEarthingConnectionStatus: asTrimmedString(source?.footToEarthingConnectionStatus) || 'Given',
          measuredCurrentMa: asTrimmedString(source?.measuredCurrentMa),
          measuredImpedance: asTrimmedString(source?.measuredImpedance)
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
      footToEarthingConnectionStatus: asTrimmedString(row?.footToEarthingConnectionStatus) || 'Given',
      measuredCurrentMa: asTrimmedString(row?.measuredCurrentMa),
      measuredImpedance: asTrimmedString(row?.measuredImpedance)
    };
  });

  return Array.from(grouped.values());
}

function normalizeLoopRows(inputRows, context) {
  return (Array.isArray(inputRows) ? inputRows : [])
    .filter(rowHasAnyValue)
    .map((row, index) => {
      const normalized = normalizeRowMeta(row, 'loop');
      normalized.srNo = asTrimmedString(row?.srNo) || String(index + 1);
      normalized.location = asTrimmedString(row?.location || row?.mainLocation);
      normalized.feederTag = asTrimmedString(row?.feederTag || row?.panelEquipment);
      normalized.deviceType = asTrimmedString(row?.deviceType) || 'MCB';
      normalized.deviceRating = asTrimmedString(row?.deviceRating);
      normalized.breakingCapacity = asTrimmedString(row?.breakingCapacity);
      normalized.measuredPoints = asTrimmedString(row?.measuredPoints) || PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length];
      normalized.loopImpedance = asTrimmedString(row?.loopImpedance || row?.measuredZs);
      normalized.voltage = asTrimmedString(row?.voltage) || '230';
      normalized.remarks = asTrimmedString(row?.remarks);
      return normalized;
    })
    .map((row, index) => ({
      ...row,
      srNo: String(index + 1),
      measuredPoints: PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length],
      remarks: asTrimmedString(row.remarks) || deriveLoopImpedanceAssessment(row, context).comment
    }));
}

function normalizeFaultRows(inputRows, context) {
  return (Array.isArray(inputRows) ? inputRows : [])
    .filter(rowHasAnyValue)
    .map((row, index) => {
      const normalized = normalizeRowMeta(row, 'fault');
      normalized.srNo = asTrimmedString(row?.srNo) || String(index + 1);
      normalized.location = asTrimmedString(row?.location || row?.mainLocation);
      normalized.feederTag = asTrimmedString(row?.feederTag || row?.panelEquipment);
      normalized.deviceType = asTrimmedString(row?.deviceType) || 'MCB';
      normalized.deviceRating = asTrimmedString(row?.deviceRating);
      normalized.breakingCapacity = asTrimmedString(row?.breakingCapacity);
      normalized.measuredPoints = asTrimmedString(row?.measuredPoints) || PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length];
      normalized.loopImpedance = asTrimmedString(row?.loopImpedance || row?.measuredZs);
      normalized.prospectiveFaultCurrent = asTrimmedString(row?.prospectiveFaultCurrent);
      normalized.shortCircuitCurrentLN = asTrimmedString(row?.shortCircuitCurrentLN);
      normalized.voltage = asTrimmedString(row?.voltage) || '230';
      normalized.comment = asTrimmedString(row?.comment);
      return normalized;
    })
    .map((row, index) => ({
      ...row,
      srNo: String(index + 1),
      measuredPoints: PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length],
      comment: asTrimmedString(row.comment) || deriveFaultCurrentAssessment(row, context).comment
    }));
}

function normalizeReportInput(payload) {
  const draft = buildDefaultDraft();
  const input = payload || {};

  const project = {
    projectNo: asTrimmedString(input?.project?.projectNo),
    clientName: asTrimmedString(input?.project?.clientName),
    siteLocation: asTrimmedString(input?.project?.siteLocation),
    workOrder: asTrimmedString(input?.project?.workOrder),
    reportDate: asTrimmedString(input?.project?.reportDate),
    engineerName: asTrimmedString(input?.project?.engineerName),
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
      'plateArea',
      'resistanceWithoutGrid',
      'resistanceWithGrid',
      'observation'
      ],
      'electrode',
      ['conditionCodes']
    ).map((row, index) => {
      const source = Array.isArray(input?.electrodeResistance) ? input.electrodeResistance[index] : null;
      const resistanceWithGrid = row.resistanceWithGrid || asTrimmedString(source?.measuredResistance);
      const normalizedRow = {
        ...row,
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
      'mainLocation',
      'measurementPoint',
      'resistance',
      'impedance',
      'testLeadResistance',
      'comment'
      ],
      'continuity'
    ).map((row) => ({
      ...row,
      comment: deriveContinuityAssessment(row, context).comment
    })),
    loopImpedanceTest: normalizeLoopRows(input?.loopImpedanceTest, context),
    prospectiveFaultCurrent: normalizeFaultRows(input?.prospectiveFaultCurrent, context),
    riserIntegrityTest: normalizeRows(
      input?.riserIntegrityTest,
      [
      'srNo',
      'mainLocation',
      'measurementPoint',
      'resistanceTowardsEquipment',
      'resistanceTowardsGrid',
      'testLeadResistance',
      'comment'
      ],
      'riser',
      ['conditionCodes']
    ).map((row) => ({
      ...row,
      comment: deriveRiserAssessment(row, context).comment
    })),
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
      remark: deriveEarthContinuityAssessment(row, context).comment
    })),
    towerFootingResistance: normalizedTowerFootingGroups.map((group) => {
      const summary = towerGroupSummaries.get(buildTowerGroupKey(group));
      const assessment = deriveTowerFootingAssessment(group, summary, context);
      return {
        ...group,
        totalImpedanceZt: assessment.totalImpedanceZt === null ? '' : String(assessment.totalImpedanceZt),
        totalCurrentItotal: assessment.totalCurrentItotal === null ? '' : String(assessment.totalCurrentItotal),
        standardTolerableImpedanceZsat: String(assessment.zsat),
        remarks: assessment.comment
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
            direction1: report?.soilResistivity?.direction1 || [],
            direction2: report?.soilResistivity?.direction2 || [],
            drivenElectrodeDiameter: report?.soilResistivity?.drivenElectrodeDiameter || '',
            drivenElectrodeLength: report?.soilResistivity?.drivenElectrodeLength || '',
            notes: report?.soilResistivity?.notes || ''
          }
        ];

  // Site averages are taken from the unrounded location averages, so the printed
  // site rho (and every limit derived from it) does not pick up rounding error.
  const exact = [];
  const locationSummaries = rawLocations.map((location, index) => {
    const inputUnit = asTrimmedString(location?.inputUnit) === 'ohm' ? 'ohm' : 'ohm-m';
    const direction1Values = (location.direction1 || [])
      .map((row) => readSoilRho(row, inputUnit))
      .filter((value) => Number.isFinite(value));
    const direction2Values = (location.direction2 || [])
      .map((row) => readSoilRho(row, inputUnit))
      .filter((value) => Number.isFinite(value));

    const direction1Average = average(direction1Values);
    const direction2Average = average(direction2Values);
    const overallAverage = average([direction1Average, direction2Average].filter((value) => Number.isFinite(value)));
    const category = getSoilCategory(overallAverage);
    exact.push({ direction1Average, direction2Average, overallAverage });

    return {
      locationId: asTrimmedString(location?.locationId) || `soil-location-${index + 1}`,
      name: asTrimmedString(location?.name) || `Location ${index + 1}`,
      direction1Average: round(direction1Average, 2),
      direction2Average: round(direction2Average, 2),
      overallAverage: round(overallAverage, 2),
      category,
      inputUnit,
      drivenElectrodeDiameter: asTrimmedString(location?.drivenElectrodeDiameter),
      drivenElectrodeLength: asTrimmedString(location?.drivenElectrodeLength),
      notes: asTrimmedString(location?.notes)
    };
  });

  const direction1Average = average(exact.map((location) => location.direction1Average).filter((value) => Number.isFinite(value)));
  const direction2Average = average(exact.map((location) => location.direction2Average).filter((value) => Number.isFinite(value)));
  const overallAverage = average(exact.map((location) => location.overallAverage).filter((value) => Number.isFinite(value)));
  const category = getSoilCategory(overallAverage);

  return {
    direction1Average: round(direction1Average, 2),
    direction2Average: round(direction2Average, 2),
    overallAverage: round(overallAverage, 2),
    category,
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
function summarizeOverallStatus(units) {
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
      label = 'Healthy';
      tone = 'healthy';
    } else if (passRate >= OVERALL_PASS_THRESHOLD) {
      label = 'Healthy - Minor Actions';
      tone = 'warning';
    } else {
      label = '>Attention Required';
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
function collectTestUnits(report) {
  const units = {};
  const context = buildAssessmentContext(report);
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
function buildTestSummaries(report) {
  const summaries = {};
  Object.entries(collectTestUnits(report)).forEach(([id, units]) => {
    summaries[id] = summarizeOverallStatus(units);
  });
  return summaries;
}

function buildOverallStatus(report) {
  return summarizeOverallStatus(Object.values(collectTestUnits(report)).flat());
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
  TEST_LIBRARY,
  EQUIPMENT_LIBRARY,
  buildDefaultDraft,
  normalizeReportInput,
  calculateSoilSummary,
  calculateDrivenElectrodeResistance,
  getSoilCategory,
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
