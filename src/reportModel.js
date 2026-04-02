const TEST_LIBRARY = [
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
    resistanceWithoutGrid: '',
    resistanceWithGrid: '',
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
    standardTolerableImpedanceZsat: '10',
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
    towerFootingResistance: [{ ...createDefaultTowerFootingGroup(), srNo: '1' }]
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

function calculateDrivenElectrodeResistance(soilAverage, electrodeLengthMeters, electrodeDiameterMillimeters) {
  const rho = Number.isFinite(soilAverage) ? soilAverage : null;
  const lengthMeters = Number.isFinite(electrodeLengthMeters) ? electrodeLengthMeters : null;
  const diameterMillimeters = Number.isFinite(electrodeDiameterMillimeters) ? electrodeDiameterMillimeters : null;
  if (!Number.isFinite(rho) || !Number.isFinite(lengthMeters) || !Number.isFinite(diameterMillimeters) || lengthMeters <= 0 || diameterMillimeters <= 0) {
    return null;
  }

  const diameterMeters = diameterMillimeters / 1000;
  const result = (rho / (2 * Math.PI * lengthMeters)) * Math.log((2 * lengthMeters) / diameterMeters);
  return round(result, 2);
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

function getElectrodeMeasuredValue(row) {
  return asLooseNumber(row?.resistanceWithGrid) ?? asLooseNumber(row?.resistanceWithoutGrid) ?? asLooseNumber(row?.measuredResistance);
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

function getLoopImpedanceStatus(value) {
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
  const max = Math.max(...values);
  if (max <= 0.5) {
    return { label: 'Healthy', tone: 'healthy' };
  }
  if (max <= 1) {
    return { label: 'Needs Attention', tone: 'warning' };
  }
  return { label: 'Critical', tone: 'critical' };
}

function getEarthContinuityStatus(value) {
  const numeric = asLooseNumber(value);
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
    reference: 'DIO / GOV.UK continuity guidance',
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
    reference: 'IEC / continuity guidance',
    limitLabel: '0.50 ohm continuity reference'
  },
  towerFootingResistance: {
    reference: 'Project default tower footing limit',
    limitLabel: '10.00 ohm Zsat default'
  }
};

function deriveElectrodeAssessment(row) {
  const measured = getElectrodeMeasuredValue(row);
  const withoutGrid = asLooseNumber(row?.resistanceWithoutGrid);
  const withGrid = asLooseNumber(row?.resistanceWithGrid) ?? asLooseNumber(row?.measuredResistance);
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
    measured,
    comment
  };
}

function deriveContinuityAssessment(row) {
  const resistance = asLooseNumber(row?.resistance);
  const impedance = asLooseNumber(row?.impedance);
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

function deriveLoopImpedanceAssessment(row) {
  const measured = asLooseNumber(row?.loopImpedance) ?? asLooseNumber(row?.measuredZs);
  const status = getLoopImpedanceStatus(measured);
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

function deriveFaultCurrentAssessment(row) {
  const pfc = asLooseNumber(row?.prospectiveFaultCurrent);
  const breakingCapacity = asLooseNumber(row?.breakingCapacity);
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
  const equipment = asLooseNumber(row?.resistanceTowardsEquipment);
  const grid = asLooseNumber(row?.resistanceTowardsGrid);
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
  const measured = asLooseNumber(row?.measuredValue);
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
      drivenElectrodeDiameter: asTrimmedString(location?.drivenElectrodeDiameter),
      drivenElectrodeLength: asTrimmedString(location?.drivenElectrodeLength),
      notes: asTrimmedString(location?.notes)
    }));
}

function normalizeRows(rows, fieldMap, rowPrefix = 'row') {
  return (Array.isArray(rows) ? rows : [])
    .filter(rowHasAnyValue)
    .map((row) => {
      const normalized = normalizeRowMeta(row, rowPrefix);
      fieldMap.forEach((field) => {
        normalized[field] = asTrimmedString(row[field]);
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
    (Boolean(asTrimmedString(group?.standardTolerableImpedanceZsat)) && asTrimmedString(group?.standardTolerableImpedanceZsat) !== '10') ||
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
        standardTolerableImpedanceZsat: asTrimmedString(group?.standardTolerableImpedanceZsat) || '10',
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
        standardTolerableImpedanceZsat: asTrimmedString(row?.standardTolerableImpedanceZsat) || '10',
        remarks: ''
      });
    }
    const group = grouped.get(key);
    const footIndex = Math.max(0, TOWER_FOOT_POINTS.indexOf(asTrimmedString(row?.measurementPointLocation) || 'Foot-1'));
    group.mainLocationTower = group.mainLocationTower || asTrimmedString(row?.mainLocationTower);
    group.standardTolerableImpedanceZsat = group.standardTolerableImpedanceZsat || asTrimmedString(row?.standardTolerableImpedanceZsat) || '10';
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

function normalizeLoopRows(inputRows) {
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
      remarks: asTrimmedString(row.remarks) || deriveLoopImpedanceAssessment(row).comment
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
      normalized.deviceType = asTrimmedString(row?.deviceType) || 'MCB';
      normalized.deviceRating = asTrimmedString(row?.deviceRating);
      normalized.breakingCapacity = asTrimmedString(row?.breakingCapacity);
      normalized.measuredPoints = asTrimmedString(row?.measuredPoints) || PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length];
      normalized.loopImpedance = asTrimmedString(row?.loopImpedance || row?.measuredZs);
      normalized.prospectiveFaultCurrent = asTrimmedString(row?.prospectiveFaultCurrent);
      normalized.voltage = asTrimmedString(row?.voltage) || '230';
      normalized.comment = asTrimmedString(row?.comment);
      return normalized;
    })
    .map((row, index) => ({
      ...row,
      srNo: String(index + 1),
      measuredPoints: PHASE_MEASURED_POINTS[index % PHASE_MEASURED_POINTS.length],
      comment: asTrimmedString(row.comment) || deriveFaultCurrentAssessment(row).comment
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

  const normalizedTowerFootingGroups = normalizeTowerFootingGroups(input?.towerFootingResistance).map((group, index) => ({
    ...group,
    srNo: asTrimmedString(group.srNo) || String(index + 1),
    standardTolerableImpedanceZsat: '10'
  }));

  const towerGroupSummaries = summarizeTowerGroups(normalizedTowerFootingGroups);

  const report = {
    project,
    tests,
    soilResistivity: {
      locations: normalizeSoilLocations(input?.soilResistivity)
    },
    electrodeResistance: normalizeRows(
      input?.electrodeResistance,
      [
      'tag',
      'location',
      'electrodeType',
      'materialType',
      'length',
      'diameter',
      'resistanceWithoutGrid',
      'resistanceWithGrid',
      'observation'
      ],
      'electrode'
    ).map((row, index) => {
      const source = Array.isArray(input?.electrodeResistance) ? input.electrodeResistance[index] : null;
      const resistanceWithGrid = row.resistanceWithGrid || asTrimmedString(source?.measuredResistance);
      return {
        ...row,
        resistanceWithGrid,
        observation: asTrimmedString(source?.observation)
      };
    }),
    continuityTest: normalizeRows(
      input?.continuityTest,
      [
      'srNo',
      'mainLocation',
      'measurementPoint',
      'resistance',
      'impedance',
      'comment'
      ],
      'continuity'
    ).map((row) => ({
      ...row,
      comment: deriveContinuityAssessment(row).comment
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
      'comment'
      ],
      'riser'
    ).map((row) => ({
      ...row,
      comment: deriveRiserAssessment(row).comment
    })),
    earthContinuityTest: normalizeRows(
      input?.earthContinuityTest,
      [
      'srNo',
      'tag',
      'locationBuildingName',
      'distance',
      'measuredValue',
      'remark'
      ],
      'earth'
    ).map((row) => ({
      ...row,
      remark: deriveEarthContinuityAssessment(row).comment
    })),
    towerFootingResistance: normalizedTowerFootingGroups.map((group) => {
      const summary = towerGroupSummaries.get(buildTowerGroupKey(group));
      const assessment = deriveTowerFootingAssessment(group, summary);
      return {
        ...group,
        totalImpedanceZt: assessment.totalImpedanceZt === null ? '' : String(assessment.totalImpedanceZt),
        totalCurrentItotal: assessment.totalCurrentItotal === null ? '' : String(assessment.totalCurrentItotal),
        standardTolerableImpedanceZsat: assessment.zsat === null ? '10' : String(assessment.zsat),
        remarks: assessment.comment
      };
    })
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

  const locationSummaries = rawLocations.map((location, index) => {
    const direction1Values = (location.direction1 || [])
      .map((row) => asLooseNumber(row.resistivity))
      .filter((value) => Number.isFinite(value));
    const direction2Values = (location.direction2 || [])
      .map((row) => asLooseNumber(row.resistivity))
      .filter((value) => Number.isFinite(value));

    const direction1Average = average(direction1Values);
    const direction2Average = average(direction2Values);
    const overallAverage = average([direction1Average, direction2Average].filter((value) => Number.isFinite(value)));
    const category = getSoilCategory(overallAverage);

    return {
      locationId: asTrimmedString(location?.locationId) || `soil-location-${index + 1}`,
      name: asTrimmedString(location?.name) || `Location ${index + 1}`,
      direction1Average: round(direction1Average, 2),
      direction2Average: round(direction2Average, 2),
      overallAverage: round(overallAverage, 2),
      category,
      drivenElectrodeDiameter: asTrimmedString(location?.drivenElectrodeDiameter),
      drivenElectrodeLength: asTrimmedString(location?.drivenElectrodeLength),
      notes: asTrimmedString(location?.notes)
    };
  });

  const direction1Average = average(locationSummaries.map((location) => location.direction1Average).filter((value) => Number.isFinite(value)));
  const direction2Average = average(locationSummaries.map((location) => location.direction2Average).filter((value) => Number.isFinite(value)));
  const overallAverage = average(locationSummaries.map((location) => location.overallAverage).filter((value) => Number.isFinite(value)));
  const category = getSoilCategory(overallAverage);

  return {
    direction1Average: round(direction1Average, 2),
    direction2Average: round(direction2Average, 2),
    overallAverage: round(overallAverage, 2),
    category,
    locations: locationSummaries
  };
}

function buildExecutiveSnapshot(report) {
  const soil = calculateSoilSummary(report);
  const selectedTests = TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);

  const healthyElectrodes = (report?.electrodeResistance || []).filter((row) => {
    return getElectrodeStatus(getElectrodeMeasuredValue(row)).tone === 'healthy';
  }).length;

  return {
    selectedTests,
    soil,
    healthyElectrodes,
    totalElectrodes: (report?.electrodeResistance || []).length
  };
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
  STANDARD_GUIDANCE,
  buildExecutiveSnapshot,
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
