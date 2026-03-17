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
  towerFootingResistance: ['Project default Zsat 10 ohm', 'Grouped tower footing analysis']
};

const SOIL_SPACING_PRESETS = ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0', '4.5', '5.0'];
const TOWER_FOOT_POINTS = ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'];

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
    mainLocation: '',
    panelEquipment: '',
    measuredZs: '',
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
    measuredPoints: '',
    loopImpedance: '',
    prospectiveFaultCurrent: '',
    voltage: '230',
    comment: '',
    rowObservation: '',
    rowPhotos: []
  };
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
      direction1: SOIL_SPACING_PRESETS.slice(0, 6).map((spacing) => defaultSoilRow(spacing)),
      direction2: SOIL_SPACING_PRESETS.slice(0, 6).map((spacing) => defaultSoilRow(spacing)),
      notes: ''
    },
    electrodeResistance: [defaultElectrodeRow()],
    continuityTest: [defaultContinuityRow('1')],
    loopImpedanceTest: [defaultLoopRow('1')],
    prospectiveFaultCurrent: [defaultFaultRow('1')],
    riserIntegrityTest: [defaultRiserRow('1')],
    earthContinuityTest: [defaultEarthContinuityRow('1')],
    towerFootingResistance: [defaultTowerFootingGroup('1')]
  };
}

const state = {
  catalog: [...LOCAL_TEST_LIBRARY],
  reports: [],
  view: 'dashboard',
  search: '',
  loadingReports: false,
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
  toast: null
};

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
    return source.soilResistivity[direction] || [];
  }
  return source[section] || [];
}

function getSectionRow(source, section, index, direction, groupIndex = null) {
  if (section === 'towerFootingResistance' && Number.isInteger(groupIndex)) {
    return source.towerFootingResistance?.[groupIndex]?.readings?.[index] || null;
  }
  return getSectionRows(source, section, direction)[index] || null;
}

function getTestLabel(section) {
  const match = LOCAL_TEST_LIBRARY.find((test) => test.id === section);
  return match ? match.label : 'Measurement Row';
}

function getRowObservationTitle(section, index, direction, groupIndex = null) {
  if (section === 'towerFootingResistance' && Number.isInteger(groupIndex)) {
    const group = state.draft.towerFootingResistance?.[groupIndex];
    const foot = group?.readings?.[index]?.measurementPointLocation || TOWER_FOOT_POINTS[index] || `Foot-${index + 1}`;
    const tower = safeText(group?.mainLocationTower, `Tower ${groupIndex + 1}`);
    return `Tower Footing Resistance Measurement & Analysis | ${tower} | ${foot}`;
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
  const measured = toNumber(row?.measuredZs);
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
      totalCurrentItotal:
        currentValues.length === TOWER_FOOT_POINTS.length ? round(currentValues.reduce((sum, value) => sum + value, 0) / 1000, 4) : null,
      hasAnyInput
    });
  });

  return summaries;
}

function deriveTowerFootingAssessment(group, groupSummary) {
  const standard = STANDARD_GUIDANCE.towerFootingResistance;
  const zsat = toNumber(group?.standardTolerableImpedanceZsat) ?? 10;
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
  const direction1Average = average((source.soilResistivity.direction1 || []).map((row) => toNumber(row.resistivity)));
  const direction2Average = average((source.soilResistivity.direction2 || []).map((row) => toNumber(row.resistivity)));
  const overallAverage = average([direction1Average, direction2Average].filter((value) => Number.isFinite(value)));
  return {
    direction1Average: round(direction1Average, 2),
    direction2Average: round(direction2Average, 2),
    overallAverage: round(overallAverage, 2),
    category: getSoilCategory(overallAverage)
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

function buildAssessmentSummary(source) {
  const soil = calculateSoilSummary(source);
  const actions = countActionRows(source);
  const selected = selectedTests(source);
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

function builderReadiness() {
  const project = state.draft.project;
  const projectComplete = ['projectNo', 'clientName', 'siteLocation', 'workOrder', 'reportDate', 'engineerName'].filter((field) => safeText(project[field], '')).length;
  const selectionComplete = selectedTestCount(state.draft) ? 1 : 0;
  const score = Math.round(((projectComplete + selectionComplete) / 7) * 100);
  return {
    score,
    projectComplete,
    selectionComplete
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
    towerFootingResistance: 'Rows sharing the same tower location are assessed together using grouped total impedance and current.'
  };
  return map[stepId] || 'Continue entering the field measurements for the selected scope.';
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
    `
  );
}

function actionButtonsCell(section, index, row, direction) {
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

  const { section, index, direction, remark, photos, rowId, groupIndex } = state.observationEditor;
  const title = getRowObservationTitle(section, index, direction, groupIndex);

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

function soilRowHtml(direction, row, index) {
  return `
    <tr>
      <td><input class="table-input" value="${escapeHtml(row.spacing)}" data-section="soilResistivity" data-direction="${direction}" data-index="${index}" data-field="spacing" /></td>
      <td><input class="table-input" value="${escapeHtml(row.resistivity)}" data-section="soilResistivity" data-direction="${direction}" data-index="${index}" data-field="resistivity" /></td>
      ${actionButtonsCell('soilResistivity', index, row, direction)}
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

function renderSoilStep() {
  const summary = calculateSoilSummary(state.draft);
  return sectionCard(
    'Soil Resistivity Test',
    'Measurement Sheet',
    `
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
      <div class="split-grid">
        ${tableSurface(
          'Direction 1',
          'Keep the measured columns exactly as per the sheet.',
          '<tr><th>Spacing of Probes (m)</th><th>Resistivity rho (ohm-m)</th><th>Action</th></tr>',
          state.draft.soilResistivity.direction1.map((row, index) => soilRowHtml('direction1', row, index)).join(''),
          '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="soilResistivity" data-direction="direction1">Add Measurement Row</button></div>'
        )}
        ${tableSurface(
          'Direction 2',
          'Second direction for the mean calculation.',
          '<tr><th>Spacing of Probes (m)</th><th>Resistivity rho (ohm-m)</th><th>Action</th></tr>',
          state.draft.soilResistivity.direction2.map((row, index) => soilRowHtml('direction2', row, index)).join(''),
          '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="soilResistivity" data-direction="direction2">Add Measurement Row</button></div>'
        )}
      </div>
      <label class="field">
        <span>Site Notes</span>
        <textarea class="text-area" rows="3" data-bind="soilResistivity.notes" placeholder="Ground condition, nearby structures, or observations">${escapeHtml(state.draft.soilResistivity.notes)}</textarea>
      </label>
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
              <td class="tower-group-cell" rowspan="${readings.length}">${autoTextCell(assessment.totalImpedanceZt === null ? '-' : String(assessment.totalImpedanceZt))}</td>
              <td class="tower-group-cell" rowspan="${readings.length}">${autoTextCell(assessment.totalCurrentItotal === null ? '-' : String(assessment.totalCurrentItotal))}</td>
              <td class="tower-group-cell" rowspan="${readings.length}">
                <input
                  class="table-input"
                  value="${escapeHtml(safeText(group.standardTolerableImpedanceZsat, '10'))}"
                  data-section="towerFootingResistance"
                  data-index="${groupIndex}"
                  data-field="standardTolerableImpedanceZsat"
                />
              </td>
              <td class="tower-group-cell" rowspan="${readings.length}">
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
    tableSurface(
      'Tower Footing Resistance Measurement & Analysis',
      'Each tower location contains 4 fixed footing rows: Foot-1, Foot-2, Foot-3, and Foot-4.',
      '<tr><th>Sr. No.</th><th>Main Location – Tower</th><th>Measurement Point Location</th><th>Foot to Earthing Connection Status</th><th>Measured Current I (mA)</th><th>Measured Impedance (ohm)</th><th>Total Impedance Zt (ohm)</th><th>Total Current Itotal (A)</th><th>Standard Tolerable Impedance value Zsat</th><th>Remarks</th><th>Action</th></tr>',
      state.draft.towerFootingResistance.map((group, groupIndex) => towerGroupRowsHtml(group, groupIndex, summaries)).join(''),
      '<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="towerFootingResistance">Add New Tower Location</button></div>'
    )
  );
}

function genericStep(title, subtitle, section, columns, statusRenderer) {
  return sectionCard(
    title,
    subtitle,
    tableSurface(
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
    )
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
    return genericStep('Continuity Test', 'Measurement Sheet', 'continuityTest', [
      { key: 'srNo', label: 'Sr. No.' },
      { key: 'mainLocation', label: 'Main Location' },
      { key: 'measurementPoint', label: 'Measurement Point' },
      { key: 'resistance', label: 'Resistance (ohm)' },
      { key: 'impedance', label: 'Impedance (ohm)' },
      {
        key: 'status',
        label: 'Status',
        render: (row) => {
          const assessment = deriveContinuityAssessment(row);
          return pill(assessment.status.tone, assessment.status.label);
        }
      },
      {
        key: 'comment',
        label: 'Comment',
        render: (row) => autoTextCell(deriveContinuityAssessment(row).comment)
      }
    ]);
  }
  if (step === 'loopImpedanceTest') {
    return genericStep('Loop Impedance Test', 'Measurement Sheet', 'loopImpedanceTest', [
      { key: 'srNo', label: 'Sr. No.' },
      { key: 'mainLocation', label: 'Main Location' },
      { key: 'panelEquipment', label: 'Panel / Equipment' },
      { key: 'measuredZs', label: 'Measured Zs (ohm)' },
      {
        key: 'status',
        label: 'Status',
        render: (row) => {
          const assessment = deriveLoopAssessment(row);
          return pill(assessment.status.tone, assessment.status.label);
        }
      },
      {
        key: 'remarks',
        label: 'Remarks',
        render: (row) => autoTextCell(deriveLoopAssessment(row).comment)
      }
    ]);
  }
  if (step === 'prospectiveFaultCurrent') {
    return genericStep('Prospective Fault Current', 'Measurement Sheet', 'prospectiveFaultCurrent', [
      { key: 'srNo', label: 'Sr. No.' },
      { key: 'location', label: 'Location' },
      { key: 'feederTag', label: 'Feeder & Tag' },
      { key: 'deviceType', label: 'Device Type', type: 'select', options: ['ACB', 'MCCB', 'MCB', 'SFU'] },
      { key: 'deviceRating', label: 'Rating (A)' },
      { key: 'breakingCapacity', label: 'Breaking (kA)' },
      { key: 'measuredPoints', label: 'Measured Points' },
      { key: 'loopImpedance', label: 'Loop Z (ohm)' },
      { key: 'prospectiveFaultCurrent', label: 'PFC' },
      { key: 'voltage', label: 'Voltage' },
      {
        key: 'status',
        label: 'Status',
        render: (row) => {
          const assessment = deriveFaultAssessment(row);
          return pill(assessment.status.tone, assessment.status.label);
        }
      },
      {
        key: 'comment',
        label: 'Comment',
        render: (row) => autoTextCell(deriveFaultAssessment(row).comment)
      }
    ]);
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
    ]);
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
    ]);
  }
  if (step === 'towerFootingResistance') {
    return renderTowerFootingStep();
  }
  return '';
}

function builderView() {
  const steps = getSteps();
  const assessment = buildAssessmentSummary(state.draft);
  const readiness = builderReadiness();
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
            <p>${stepGuidance(currentStep)}</p>
            ${renderProgressBar('Project completeness', Math.round((readiness.projectComplete / 6) * 100), readiness.projectComplete === 6 ? 'healthy' : 'warning')}
            ${renderProgressBar('Overall setup', readiness.score, readiness.score > 84 ? 'healthy' : readiness.score > 50 ? 'warning' : 'critical')}
          </section>

          <section class="mini-surface aside-panel">
            <p class="section-kicker">Selected Modules</p>
            <h4>${selectedTestCount(state.draft)} section(s)</h4>
            ${renderSelectedModuleTags(state.draft)}
          </section>

          <section class="mini-surface aside-panel">
            <p class="section-kicker">Reference Standards</p>
            <h4>Typical standards</h4>
            ${renderStandardsSummary(state.draft)}
          </section>

          <section class="mini-surface aside-panel">
            <p class="section-kicker">Live Assessment</p>
            <div class="aside-stats">
              <div><span>Mean Soil</span><strong>${assessment.soil.overallAverage === null ? '-' : `${assessment.soil.overallAverage} ohm-m`}</strong></div>
              <div><span>Soil Category</span><strong>${assessment.soil.category.label}</strong></div>
              <div><span>Critical Electrodes</span><strong>${assessment.actions.electrodeCritical}</strong></div>
              <div><span>Action Points</span><strong>${assessment.actions.total}</strong></div>
            </div>
          </section>

          <section class="mini-surface aside-panel">
            <p class="section-kicker">Current Step</p>
            <h4>${escapeHtml(steps[state.stepIndex].label)}</h4>
            <p>Use the step pills above to move between completed sections and keep the report aligned with actual site work.</p>
          </section>
        </aside>
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
            <th>Total Current Itotal (A)</th>
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
                             <td class="tower-group-cell" rowspan="${readings.length}">${escapeHtml(String(assessment.zsat))}</td>
                             <td class="tower-group-cell" rowspan="${readings.length}">${autoTextCell(assessment.comment)}</td>`
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

      ${report.tests.soilResistivity
        ? reportSectionBlock(
            'Soil Resistivity Test',
            `
              <div class="metric-strip">
                <article class="metric-box"><span>Direction 1 Average</span><strong>${soil.direction1Average === null ? '-' : `${soil.direction1Average} ohm-m`}</strong></article>
                <article class="metric-box"><span>Direction 2 Average</span><strong>${soil.direction2Average === null ? '-' : `${soil.direction2Average} ohm-m`}</strong></article>
                <article class="metric-box"><span>Classification</span><strong>${escapeHtml(soil.category.label)}</strong>${pill(soil.category.tone, soil.category.label)}</article>
              </div>
              ${simpleDetailTable(
                [
                  { label: 'Spacing of Probes (m)', key: 'spacing' },
                  { label: 'Direction 1 Resistivity (ohm-m)', key: 'direction1' },
                  { label: 'Direction 2 Resistivity (ohm-m)', key: 'direction2' }
                ],
                report.soilResistivity.direction1.map((row, index) => ({
                  spacing: row.spacing,
                  direction1: row.resistivity,
                  direction2: report.soilResistivity.direction2[index] ? report.soilResistivity.direction2[index].resistivity : '-'
                })),
                {
                  observationRenderer: (_row, index) => {
                    return [
                      renderObservationCard(report.soilResistivity.direction1[index], 'Direction 1 Observation'),
                      renderObservationCard(report.soilResistivity.direction2[index], 'Direction 2 Observation')
                    ]
                      .filter(Boolean)
                      .join('');
                  }
                }
              )}
              ${
                report.soilResistivity.notes
                  ? `<p class="detail-note"><strong>Site Notes:</strong> ${escapeHtml(report.soilResistivity.notes)}</p>`
                  : ''
              }
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
                  label: 'Status',
                  render: (row) => {
                    const assessment = deriveElectrodeAssessment(row);
                    return pill(assessment.status.tone, assessment.status.label);
                  }
                },
                {
                  label: 'Comment / Observation',
                  render: (row) => autoTextCell(deriveElectrodeAssessment(row).comment)
                }
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
                { label: 'Main Location', key: 'mainLocation' },
                { label: 'Panel / Equipment', key: 'panelEquipment' },
                { label: 'Measured Zs (ohm)', key: 'measuredZs' },
                {
                  label: 'Status',
                  render: (row) => {
                    const assessment = deriveLoopAssessment(row);
                    return pill(assessment.status.tone, assessment.status.label);
                  }
                },
                {
                  label: 'Remarks',
                  render: (row) => autoTextCell(safeText(row.remarks, deriveLoopAssessment(row).comment))
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
                { label: 'Location', key: 'location' },
                { label: 'Feeder & Tag', key: 'feederTag' },
                { label: 'Device', key: 'deviceType' },
                { label: 'Rating (A)', key: 'deviceRating' },
                { label: 'Breaking (kA)', key: 'breakingCapacity' },
                { label: 'Measured Points', key: 'measuredPoints' },
                { label: 'Loop Z (ohm)', key: 'loopImpedance' },
                { label: 'PFC', key: 'prospectiveFaultCurrent' },
                { label: 'Voltage', key: 'voltage' },
                {
                  label: 'Status',
                  render: (row) => {
                    const assessment = deriveFaultAssessment(row);
                    return pill(assessment.status.tone, assessment.status.label);
                  }
                },
                {
                  label: 'Comment',
                  render: (row) => autoTextCell(safeText(row.comment, deriveFaultAssessment(row).comment))
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
                  label: 'Status',
                  render: (row) => {
                    const assessment = deriveRiserAssessment(row);
                    return pill(assessment.status.tone, assessment.status.label);
                  }
                },
                {
                  label: 'Comment',
                  render: (row) => autoTextCell(safeText(row.comment, deriveRiserAssessment(row).comment))
                }
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

function openObservationEditor(section, index, direction, groupIndex = null) {
  const row = getSectionRow(state.draft, section, index, direction, groupIndex);
  if (!row) {
    return;
  }
  state.observationEditor = {
    section,
    index,
    direction,
    groupIndex,
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
    state.observationEditor.groupIndex
  );
  if (!row) {
    closeObservationEditor();
    return;
  }
  row.rowId = state.observationEditor.rowId;
  row.rowObservation = String(state.observationEditor.remark || '').trim();
  row.rowPhotos = cloneRowPhotos(state.observationEditor.photos).filter((photo) => photo.dataUrl);
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
  });
}

function addRow(section, direction) {
  if (section === 'soilResistivity') {
    const nextSpacing = SOIL_SPACING_PRESETS[state.draft.soilResistivity[direction].length] || '';
    state.draft.soilResistivity[direction].push(defaultSoilRow(nextSpacing));
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
    state.draft.loopImpedanceTest.push(defaultLoopRow(String(state.draft.loopImpedanceTest.length + 1)));
    return;
  }
  if (section === 'prospectiveFaultCurrent') {
    state.draft.prospectiveFaultCurrent.push(defaultFaultRow(String(state.draft.prospectiveFaultCurrent.length + 1)));
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

function removeRow(section, index, direction) {
  if (
    state.observationEditor &&
    state.observationEditor.section === section &&
    safeText(state.observationEditor.direction, '') === safeText(direction, '')
  ) {
    closeObservationEditor();
  }
  if (section === 'soilResistivity') {
    if (state.draft.soilResistivity[direction].length === 1) {
      showToast('Keep at least one row in each soil direction.', 'warning');
      return;
    }
    state.draft.soilResistivity[direction].splice(index, 1);
    return;
  }
  if (state.draft[section].length === 1) {
    showToast('Keep at least one row in the selected section.', 'warning');
    return;
  }
  state.draft[section].splice(index, 1);
  renumberSection(section);
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

async function exportPdf(id) {
  state.exporting = true;
  render();
  try {
    const result = await api(`/api/reports/${encodeURIComponent(id)}/pdf`, { method: 'POST' });
    window.open(result.pdfUrl, '_blank', 'noopener');
    showToast('PDF generated successfully.', 'healthy');
  } catch (error) {
    showToast(error.message, 'critical');
  } finally {
    state.exporting = false;
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
    render();
    return;
  }

  if (action === 'go-dashboard') {
    closeObservationEditor();
    state.view = 'dashboard';
    state.activeReport = null;
    render();
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

  if (action === 'add-row') {
    addRow(button.dataset.section, button.dataset.direction);
    render();
    return;
  }

  if (action === 'remove-row') {
    removeRow(button.dataset.section, Number(button.dataset.index), button.dataset.direction);
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
      button.dataset.groupIndex === undefined ? null : Number(button.dataset.groupIndex)
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
    const field = target.dataset.field;
    const direction = target.dataset.direction;

    if (section === 'soilResistivity') {
      state.draft.soilResistivity[direction][index][field] = target.value;
    } else if (section === 'towerFootingResistance') {
      if (Number.isInteger(readingIndex)) {
        state.draft.towerFootingResistance[index].readings[readingIndex][field] = target.value;
      } else {
        state.draft.towerFootingResistance[index][field] = target.value;
      }
    } else {
      state.draft[section][index][field] = target.value;
    }
    rerenderPreservingFocus();
  }
});

document.addEventListener('change', async (event) => {
  const target = event.target;
  if (target.id === 'zohoProjectPicker') {
    applyZohoProject(target.value);
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
    const field = target.dataset.field;
    const direction = target.dataset.direction;
    if (section === 'soilResistivity') {
      state.draft.soilResistivity[direction][index][field] = target.value;
    } else if (section === 'towerFootingResistance') {
      if (Number.isInteger(readingIndex)) {
        state.draft.towerFootingResistance[index].readings[readingIndex][field] = target.value;
      } else {
        state.draft.towerFootingResistance[index][field] = target.value;
      }
    } else {
      state.draft[section][index][field] = target.value;
    }
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
  render();
  await loadCatalog();
  await loadZohoProjects();
  await loadReports();
}

init();
