const REPORT_MODEL_FALLBACK = ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0', '4.5', '5.0'];
const PHASE_MEASURED_POINTS = ['R-E', 'Y-E', 'B-E'];
const TOWER_FOOT_POINTS = ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'];
const reportModelExports = (() => {
  try {
    return require('./reportModel');
  } catch (_error) {
    return {};
  }
})();
const SOIL_SPACING_PRESETS =
  Array.isArray(reportModelExports.SOIL_SPACING_PRESETS) && reportModelExports.SOIL_SPACING_PRESETS.length
    ? reportModelExports.SOIL_SPACING_PRESETS
    : REPORT_MODEL_FALLBACK;

function asText(value, fallback = '') {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || fallback;
}

function normalizeSpacing(value) {
  const raw = asText(value);
  if (!raw) {
    return '';
  }

  const numeric = Number.parseFloat(raw);
  if (!Number.isFinite(numeric)) {
    return raw;
  }

  const fixed = numeric.toFixed(1);
  return SOIL_SPACING_PRESETS.includes(fixed) ? fixed : String(numeric);
}

function normalizeSoilRow(row = {}) {
  return {
    spacing: normalizeSpacing(row.spacing),
    resistivity: asText(row.resistivity)
  };
}

function normalizeUncertainField(item = {}) {
  return {
    path: asText(item.path),
    reason: asText(item.reason)
  };
}

function normalizeWarningList(warnings) {
  return (Array.isArray(warnings) ? warnings : []).map((item) => asText(item)).filter(Boolean);
}

function normalizePhaseRows(rows, rowMapper) {
  return (Array.isArray(rows) ? rows : [])
    .map((row, index) => rowMapper(row, index))
    .filter((row) => Object.values(row).some(Boolean));
}

function validateSoilOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'soilResistivity') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Soil Resistivity.');
  }

  const rawLocations = Array.isArray(payload.locations) ? payload.locations : [];
  if (!rawLocations.length) {
    throw new Error('No soil resistivity locations were extracted from the uploaded document.');
  }

  const locations = rawLocations
    .map((location, index) => {
      const direction1 = (Array.isArray(location.direction1) ? location.direction1 : [])
        .map(normalizeSoilRow)
        .filter((row) => row.spacing || row.resistivity);
      const direction2 = (Array.isArray(location.direction2) ? location.direction2 : [])
        .map(normalizeSoilRow)
        .filter((row) => row.spacing || row.resistivity);

      if (!direction1.length && !direction2.length) {
        return null;
      }

      return {
        name: asText(location.name, `Location ${index + 1}`),
        drivenElectrodeDiameter: asText(location.drivenElectrodeDiameter),
        drivenElectrodeLength: asText(location.drivenElectrodeLength),
        notes: asText(location.notes),
        direction1,
        direction2
      };
    })
    .filter(Boolean);

  if (!locations.length) {
    throw new Error('The uploaded sheet did not contain any usable soil resistivity rows.');
  }

  return {
    sheetType: 'soilResistivity',
    locations,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateElectrodeOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'electrodeResistance') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Earth Electrode Resistance.');
  }

  const rows = (Array.isArray(payload.rows) ? payload.rows : [])
    .map((row) => ({
      tag: asText(row.tag),
      location: asText(row.location),
      electrodeType: asText(row.electrodeType, 'Rod'),
      materialType: asText(row.materialType, 'Copper'),
      length: asText(row.length),
      diameter: asText(row.diameter),
      resistanceWithoutGrid: asText(row.resistanceWithoutGrid),
      resistanceWithGrid: asText(row.resistanceWithGrid)
    }))
    .filter((row) =>
      row.tag ||
      row.location ||
      row.length ||
      row.diameter ||
      row.resistanceWithoutGrid ||
      row.resistanceWithGrid
    );

  if (!rows.length) {
    throw new Error('The uploaded sheet did not contain any usable earth electrode rows.');
  }

  return {
    sheetType: 'electrodeResistance',
    rows,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateContinuityOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'continuityTest') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Continuity Test.');
  }

  const rows = (Array.isArray(payload.rows) ? payload.rows : [])
    .map((row, index) => ({
      srNo: asText(row.srNo, String(index + 1)),
      mainLocation: asText(row.mainLocation),
      measurementPoint: asText(row.measurementPoint),
      resistance: asText(row.resistance),
      impedance: asText(row.impedance)
    }))
    .filter((row) => row.mainLocation || row.measurementPoint || row.resistance || row.impedance);

  if (!rows.length) {
    throw new Error('The uploaded sheet did not contain any usable continuity rows.');
  }

  return {
    sheetType: 'continuityTest',
    rows,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateLoopImpedanceOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'loopImpedanceTest') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Loop Impedance Test.');
  }

  const groups = (Array.isArray(payload.groups) ? payload.groups : [])
    .map((group, groupIndex) => {
      const rows = normalizePhaseRows(group.rows, (row, rowIndex) => ({
        measuredPoints: asText(row.measuredPoints, PHASE_MEASURED_POINTS[rowIndex] || PHASE_MEASURED_POINTS[0]),
        loopImpedance: asText(row.loopImpedance),
        voltage: asText(row.voltage, '230')
      }));

      if (!rows.length) {
        return null;
      }

      return {
        location: asText(group.location),
        feederTag: asText(group.feederTag),
        deviceType: asText(group.deviceType, 'MCB'),
        deviceRating: asText(group.deviceRating),
        breakingCapacity: asText(group.breakingCapacity),
        rows: rows.map((row, rowIndex) => ({
          ...row,
          measuredPoints: asText(row.measuredPoints, PHASE_MEASURED_POINTS[rowIndex] || PHASE_MEASURED_POINTS[0])
        }))
      };
    })
    .filter(Boolean);

  if (!groups.length) {
    throw new Error('The uploaded sheet did not contain any usable loop impedance feeder groups.');
  }

  return {
    sheetType: 'loopImpedanceTest',
    groups,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateProspectiveFaultOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'prospectiveFaultCurrent') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Prospective Fault Current.');
  }

  const groups = (Array.isArray(payload.groups) ? payload.groups : [])
    .map((group) => {
      const rows = normalizePhaseRows(group.rows, (row, rowIndex) => ({
        measuredPoints: asText(row.measuredPoints, PHASE_MEASURED_POINTS[rowIndex] || PHASE_MEASURED_POINTS[0]),
        loopImpedance: asText(row.loopImpedance),
        prospectiveFaultCurrent: asText(row.prospectiveFaultCurrent),
        voltage: asText(row.voltage, '230')
      }));

      if (!rows.length) {
        return null;
      }

      return {
        location: asText(group.location),
        feederTag: asText(group.feederTag),
        deviceType: asText(group.deviceType, 'MCB'),
        deviceRating: asText(group.deviceRating),
        breakingCapacity: asText(group.breakingCapacity),
        rows: rows.map((row, rowIndex) => ({
          ...row,
          measuredPoints: asText(row.measuredPoints, PHASE_MEASURED_POINTS[rowIndex] || PHASE_MEASURED_POINTS[0])
        }))
      };
    })
    .filter(Boolean);

  if (!groups.length) {
    throw new Error('The uploaded sheet did not contain any usable prospective fault current feeder groups.');
  }

  return {
    sheetType: 'prospectiveFaultCurrent',
    groups,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateRiserIntegrityOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'riserIntegrityTest') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Riser / Grid Integrity Test.');
  }

  const rows = (Array.isArray(payload.rows) ? payload.rows : [])
    .map((row, index) => ({
      srNo: asText(row.srNo, String(index + 1)),
      mainLocation: asText(row.mainLocation),
      measurementPoint: asText(row.measurementPoint),
      resistanceTowardsEquipment: asText(row.resistanceTowardsEquipment),
      resistanceTowardsGrid: asText(row.resistanceTowardsGrid)
    }))
    .filter((row) => row.mainLocation || row.measurementPoint || row.resistanceTowardsEquipment || row.resistanceTowardsGrid);

  if (!rows.length) {
    throw new Error('The uploaded sheet did not contain any usable riser integrity rows.');
  }

  return {
    sheetType: 'riserIntegrityTest',
    rows,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateEarthContinuityOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'earthContinuityTest') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Earth Continuity Test.');
  }

  const rows = (Array.isArray(payload.rows) ? payload.rows : [])
    .map((row, index) => ({
      srNo: asText(row.srNo, String(index + 1)),
      tag: asText(row.tag),
      locationBuildingName: asText(row.locationBuildingName),
      distance: asText(row.distance),
      measuredValue: asText(row.measuredValue)
    }))
    .filter((row) => row.tag || row.locationBuildingName || row.distance || row.measuredValue);

  if (!rows.length) {
    throw new Error('The uploaded sheet did not contain any usable earth continuity rows.');
  }

  return {
    sheetType: 'earthContinuityTest',
    rows,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

function validateTowerFootingOcrPayload(payload = {}) {
  const sheetType = asText(payload.sheetType);
  if (sheetType !== 'towerFootingResistance') {
    throw new Error('Gemini OCR returned an unexpected sheet type for Tower Footing Resistance.');
  }

  const groups = (Array.isArray(payload.groups) ? payload.groups : [])
    .map((group, groupIndex) => {
      const readings = (Array.isArray(group.readings) ? group.readings : [])
        .map((reading, readingIndex) => ({
          measurementPointLocation: asText(reading.measurementPointLocation, TOWER_FOOT_POINTS[readingIndex] || TOWER_FOOT_POINTS[0]),
          footToEarthingConnectionStatus: asText(reading.footToEarthingConnectionStatus, 'Given'),
          measuredCurrentMa: asText(reading.measuredCurrentMa),
          measuredImpedance: asText(reading.measuredImpedance)
        }))
        .filter((reading) => reading.measurementPointLocation || reading.measuredCurrentMa || reading.measuredImpedance);

      if (!readings.length) {
        return null;
      }

      return {
        srNo: asText(group.srNo, String(groupIndex + 1)),
        mainLocationTower: asText(group.mainLocationTower, `Tower Location ${groupIndex + 1}`),
        readings: TOWER_FOOT_POINTS.map((foot, readingIndex) => ({
          measurementPointLocation: foot,
          footToEarthingConnectionStatus: asText(readings[readingIndex]?.footToEarthingConnectionStatus, 'Given'),
          measuredCurrentMa: asText(readings[readingIndex]?.measuredCurrentMa),
          measuredImpedance: asText(readings[readingIndex]?.measuredImpedance)
        }))
      };
    })
    .filter(Boolean);

  if (!groups.length) {
    throw new Error('The uploaded sheet did not contain any usable tower footing groups.');
  }

  return {
    sheetType: 'towerFootingResistance',
    groups,
    warnings: normalizeWarningList(payload.warnings),
    uncertainFields: (Array.isArray(payload.uncertainFields) ? payload.uncertainFields : [])
      .map(normalizeUncertainField)
      .filter((item) => item.path || item.reason)
  };
}

module.exports = {
  validateSoilOcrPayload,
  validateElectrodeOcrPayload,
  validateContinuityOcrPayload,
  validateLoopImpedanceOcrPayload,
  validateProspectiveFaultOcrPayload,
  validateRiserIntegrityOcrPayload,
  validateEarthContinuityOcrPayload,
  validateTowerFootingOcrPayload,
  normalizeSpacing
};
