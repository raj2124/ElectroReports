// Field-level validation of a normalized report against the sheet definitions.
//
//  - errors   : invalid data (not a number, out of range, unknown option, a code the
//               sheet does not offer, an instrument not allowed for the test). They
//               block the save.
//  - warnings : incomplete or unusual entries (missing dimensions, probe depth outside
//               10-15 cm, missing serial number, expired calibration). They never block
//               a save on site; they are shown, and `review` ones go on the reviewer's
//               checklist. Expired calibration also blocks approval of a FINAL report
//               until a reviewer acknowledges it (see reportStore.transitionReport).
//
// Pure: it reads the report and the definitions only. Nothing here changes data.

const {
  SHEETS,
  JOB_HEADER,
  SOIL_SPACING_PRESETS,
  GROUND_CONDITIONS,
  getSheet,
  optionValue,
  optionValues,
  instrumentsForSheet,
  flattenFields
} = require('./sheetDefinitions');

const NUMBER_PATTERN = /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/;
const DASHES = new Set(['', '-', '–', '—']);
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;
const SOIL_DIRECTIONS = ['direction1', 'direction2', 'direction3', 'direction4', 'direction5', 'direction6', 'direction7', 'direction8'];

const asText = (value) => String(value === undefined || value === null ? '' : value).trim();
const isBlank = (value) => DASHES.has(asText(value));
const isNumeric = (value) => NUMBER_PATTERN.test(asText(value));
const asNumber = (value) => (isNumeric(value) ? Number(asText(value)) : null);

function validateReport(report, context = {}) {
  const errors = [];
  const warnings = [];
  const add = (list, item) => list.push(item);
  const error = (item) => add(errors, { severity: 'error', ...item });
  const warn = (item) => add(warnings, { severity: 'warning', ...item });

  const tests = (report && report.tests) || {};
  const project = (report && report.project) || {};

  // ---- job level -----------------------------------------------------------------
  if (!isBlank(project.groundCondition) && !GROUND_CONDITIONS.includes(project.groundCondition)) {
    error({ code: 'BAD_OPTION', path: 'project.groundCondition', message: `Ground condition must be one of ${GROUND_CONDITIONS.join(', ')}.` });
  }
  if (!isBlank(project.soilSpacingPreset) && !SOIL_SPACING_PRESETS[project.soilSpacingPreset]) {
    error({ code: 'BAD_OPTION', path: 'project.soilSpacingPreset', message: `Soil spacing set must be one of ${Object.keys(SOIL_SPACING_PRESETS).join(', ')}.` });
  }
  if (!isBlank(project.assumedSoilResistivity)) {
    const assumed = asNumber(project.assumedSoilResistivity);
    if (assumed === null || assumed <= 0) {
      error({ code: 'NOT_NUMBER', path: 'project.assumedSoilResistivity', message: 'Assumed soil resistivity must be a number above 0 (ohm-m).' });
    } else if (context.rho !== null && context.rho !== undefined && !context.rhoAssumed) {
      error({ code: 'ASSUMED_RHO_WITH_MEASURED', path: 'project.assumedSoilResistivity', message: 'An assumed soil resistivity can only be given when soil resistivity was not measured; remove it so the report has one rho.' });
    }
  }

  // ---- soil (header per location, readings per direction) -------------------------
  if (tests.soilResistivity) {
    const soil = getSheet('soilResistivity');
    ((report.soilResistivity && report.soilResistivity.locations) || []).forEach((location, locationIndex) => {
      const where = `${soil.code} location ${locationIndex + 1}`;
      const path = `soilResistivity.locations.${locationIndex}`;
      const hasSoilReading = SOIL_DIRECTIONS.some((key) => ((location && location[key]) || []).some((row) => !isBlank(row.resistivity)));
      if (hasSoilReading && isBlank(location.inputUnit)) {
        warn({ code: 'SELECTION_MISSING', sheet: soil.id, path: `${path}.inputUnit`, message: `${where}: readings are entered but the reading unit (R or rho) is not ticked.` });
      }
      if (!isBlank(location.inputUnit) && !['ohm', 'ohm-m'].includes(location.inputUnit)) {
        error({ code: 'BAD_OPTION', sheet: soil.id, path: `${path}.inputUnit`, message: `${where}: reading unit must be R (ohm) or rho (ohm-m).` });
      }
      checkNumber(location.probeDepthCm, { min: 0, exclusiveMin: true }, `${where}: probe depth (cm)`, `${path}.probeDepthCm`, soil.id, error);
      if (!isBlank(location.probeDepthCm) && isNumeric(location.probeDepthCm)) {
        const depth = asNumber(location.probeDepthCm);
        if (depth > 0 && (depth < 10 || depth > 15)) {
          warn({ code: 'PROBE_DEPTH', sheet: soil.id, path: `${path}.probeDepthCm`, message: `${where}: probe depth ${depth} cm is outside 10–15 cm (IS 3043:2018 Cl. 40.4.1).` });
        }
      }
      SOIL_DIRECTIONS.forEach((key) => {
        checkNumber(location.bearings && location.bearings[key], { min: 0, max: 360 }, `${where}: ${key.replace('direction', 'direction ')} bearing (°)`, `${path}.bearings.${key}`, soil.id, error);
        ((location && location[key]) || []).forEach((row, rowIndex) => {
          checkNumber(row.spacing, { min: 0, exclusiveMin: true }, `${where} ${key.replace('direction', 'direction ')} row ${rowIndex + 1}: probe spacing (m)`, `${path}.${key}.${rowIndex}.spacing`, soil.id, error);
          checkNumber(row.resistivity, { min: 0, exclusiveMin: true }, `${where} ${key.replace('direction', 'direction ')} row ${rowIndex + 1}: reading`, `${path}.${key}.${rowIndex}.resistivity`, soil.id, error);
          if (location.inputUnit !== 'ohm' && isNumeric(row.resistivity)) {
            const rho = asNumber(row.resistivity);
            if (rho > 0 && (rho < 1 || rho > 10000)) {
              warn({ code: 'RHO_RANGE', sheet: soil.id, path: `${path}.${key}.${rowIndex}.resistivity`, message: `${where}: resistivity ${rho} ohm-m is outside 1–10 000 ohm-m (IS 3043:2018 Cl. 40.1.1).` });
            }
          }
        });
      });
    });
  }

  // ---- every other sheet -----------------------------------------------------------
  SHEETS.filter((sheet) => sheet.id !== 'soilResistivity' && tests[sheet.id]).forEach((sheet) => {
    const header = (report.sheetHeaders && report.sheetHeaders[sheet.id]) || {};
    const rows = Array.isArray(report[sheet.id]) ? report[sheet.id] : [];
    const label = sheet.code === 'legacy' ? sheet.title : sheet.code;

    validateHeader(sheet, header, label, project, { error, warn });

    rows.forEach((row, rowIndex) => {
      const where = `${label} row ${rowIndex + 1}`;
      warnMissingSelections(sheet, row, rowIndex, label, warn);
      validateFields(sheet, flattenFields(sheet.rowFields), row, `${sheet.id}.${rowIndex}`, where, error);
      if (sheet.layout === 'towerGroups') {
        (row.readings || []).forEach((reading, readingIndex) => {
          validateFields(sheet, sheet.rowFields.filter((field) => !field.merge && field.kind !== 'fixed'), reading, `${sheet.id}.${rowIndex}.readings.${readingIndex}`, `${label} tower ${rowIndex + 1} foot ${readingIndex + 1}`, error);
        });
      }
    });

    if (sheet.id === 'electrodeResistance') {
      rows.forEach((row, rowIndex) => warnElectrodeDimensions(row, rowIndex, label, warn));
    }
    if (sheet.id === 'stepTouchVoltage') {
      warnStepTouch(rows, header, label, warn);
    }
  });

  return { errors, warnings };
}

// ---- helpers ----------------------------------------------------------------------

function checkNumber(value, rule, label, path, sheetId, error) {
  if (isBlank(value)) {
    return;
  }
  if (!isNumeric(value)) {
    error({ code: 'NOT_NUMBER', sheet: sheetId, path, message: `${label} must be a number.` });
    return;
  }
  const number = Number(asText(value));
  if (rule.min !== undefined && (rule.exclusiveMin ? number <= rule.min : number < rule.min)) {
    error({ code: 'OUT_OF_RANGE', sheet: sheetId, path, message: `${label} must be ${rule.exclusiveMin ? 'above' : 'at least'} ${rule.min}.` });
  }
  if (rule.max !== undefined && number > rule.max) {
    error({ code: 'OUT_OF_RANGE', sheet: sheetId, path, message: `${label} must be at most ${rule.max}.` });
  }
}

function validateFields(sheet, fields, values, pathPrefix, where, error) {
  fields.forEach((field) => {
    if (field.legacy || ['serial', 'fixed', 'text', 'date', 'tick'].includes(field.kind)) {
      return;
    }
    const value = values ? values[field.key] : undefined;
    const label = `${where}: ${field.label}`;
    const path = `${pathPrefix}.${field.key}`;
    if (field.kind === 'number') {
      checkNumber(value, { min: field.min, exclusiveMin: field.exclusiveMin, max: field.max }, label, path, sheet.id, error);
    } else if (field.kind === 'ticks' && field.multi) {
      (Array.isArray(value) ? value : []).forEach((code) => {
        if (!optionValues(field).includes(code)) {
          error({ code: 'CODE_NOT_ALLOWED', sheet: sheet.id, path, message: `${label}: ${code} is not offered on this sheet (${optionValues(field).join(', ')}).` });
        }
      });
    } else if (field.kind === 'ticks' && !isBlank(value) && !optionValues(field).includes(value)) {
      error({ code: 'BAD_OPTION', sheet: sheet.id, path, message: `${label}: "${value}" is not one of ${optionValues(field).join(', ')}.` });
    }
  });
}

function validateHeader(sheet, header, label, project, { error, warn }) {
  const base = `sheetHeaders.${sheet.id}`;

  // Sheet-specific header cells (numbers, options).
  validateFields(sheet, flattenFields(sheet.headerFields).filter((field) => !field.readOnly), header, base, `${label} header`, error);

  // Instrument, serial and calibration.
  if (!isBlank(header.instrument)) {
    const allowed = instrumentsForSheet(sheet.id).map((instrument) => instrument.id);
    if (!allowed.includes(header.instrument)) {
      error({ code: 'INSTRUMENT_NOT_ALLOWED', sheet: sheet.id, path: `${base}.instrument`, message: `${label}: this instrument is not on the list for this test (${allowed.join(', ')}).` });
    } else {
      if (isBlank(header.serialNo)) {
        warn({ code: 'SERIAL_MISSING', sheet: sheet.id, path: `${base}.serialNo`, message: `${label}: instrument serial number is not recorded.` });
      }
      if (isBlank(header.calibrationDue)) {
        warn({ code: 'CAL_MISSING', sheet: sheet.id, path: `${base}.calibrationDue`, review: true, message: `${label}: instrument calibration due date is not recorded.` });
      }
    }
  }
  if (!isBlank(header.calibrationDue)) {
    if (!ISO_DATE.test(header.calibrationDue)) {
      error({ code: 'BAD_DATE', sheet: sheet.id, path: `${base}.calibrationDue`, message: `${label}: calibration due date must be a date (yyyy-mm-dd).` });
    } else if (ISO_DATE.test(asText(project.reportDate)) && header.calibrationDue < project.reportDate) {
      warn({
        code: 'CAL_EXPIRED',
        sheet: sheet.id,
        path: `${base}.calibrationDue`,
        review: true,
        blocksApproval: true,
        instrument: header.instrument,
        serialNo: header.serialNo,
        calibrationDue: header.calibrationDue,
        workDate: project.reportDate,
        message: `${label}: instrument calibration expired on ${header.calibrationDue}, before the work date ${project.reportDate}.`
      });
    }
  }
}

// A row (or tower foot) with a reading must also have the selections the assessment needs
// (electrode type and material, device type, foot connection). They are never pre-filled.
function warnMissingSelections(sheet, row, rowIndex, label, warn) {
  const needed = sheet.rowFields.filter((field) => field.needsSelection);
  if (!needed.length || !sheet.readingKeys) {
    return;
  }
  const check = (target, hasReading, where, pathPrefix) => {
    if (!hasReading) {
      return;
    }
    const missing = needed.filter((field) => isBlank(target[field.key]) && (target[field.key] === undefined || !Array.isArray(target[field.key])));
    if (missing.length) {
      warn({ code: 'SELECTION_MISSING', sheet: sheet.id, path: pathPrefix, message: `${where}: has a reading but ${missing.map((field) => field.label).join(' and ')} ${missing.length > 1 ? 'are' : 'is'} not ticked.` });
    }
  };
  if (sheet.layout === 'towerGroups') {
    (row.readings || []).forEach((reading, readingIndex) => {
      check(reading, sheet.readingKeys.some((key) => !isBlank(reading[key])), `${label} tower ${rowIndex + 1} foot ${readingIndex + 1}`, `${sheet.id}.${rowIndex}.readings.${readingIndex}`);
    });
    return;
  }
  check(row, sheet.readingKeys.some((key) => !isBlank(row[key])), `${label} row ${rowIndex + 1}${row.tag ? ` (${row.tag})` : ''}`, `${sheet.id}.${rowIndex}`);
}

function warnElectrodeDimensions(row, rowIndex, label, warn) {
  const hasReading = ['resistanceWithoutGrid', 'resistanceWithGrid', 'checkR50', 'checkR70'].some((key) => !isBlank(row[key]));
  if (!hasReading) {
    return;
  }
  const type = asText(row.electrodeType) || 'Rod';
  const where = `${label} row ${rowIndex + 1}${row.tag ? ` (${row.tag})` : ''}`;
  if (type === 'Plate') {
    const hasPlate = (!isBlank(row.plateLengthM) && !isBlank(row.plateWidthM)) || !isBlank(row.plateArea);
    if (!hasPlate) {
      warn({ code: 'MISSING_DIMS', sheet: 'electrodeResistance', path: `electrodeResistance.${rowIndex}`, message: `${where}: plate size (m × m) is not recorded, so the limit cannot be calculated.` });
    }
  } else if (isBlank(row.length) || isBlank(row.diameter)) {
    warn({ code: 'MISSING_DIMS', sheet: 'electrodeResistance', path: `electrodeResistance.${rowIndex}`, message: `${where}: length and ${type === 'Strip' ? 'strip width' : 'diameter'} are not recorded, so the limit cannot be calculated.` });
  }
}

function warnStepTouch(rows, header, label, warn) {
  const hasData = rows.some((row) => ['v1', 'v2', 'v3', 'vmax'].some((key) => !isBlank(row[key])));
  if (hasData) {
    const missing = [];
    if (isBlank(header.faultCurrentKA)) missing.push('earth fault current If');
    if (isBlank(header.faultClearingTimeS)) missing.push('fault clearing time ts');
    if (missing.length) {
      warn({ code: 'HEADER_INCOMPLETE', sheet: 'stepTouchVoltage', path: 'sheetHeaders.stepTouchVoltage', message: `${label}: ${missing.join(' and ')} not recorded in the header, so voltages cannot be assessed.` });
    }
  }
  rows.forEach((row, rowIndex) => {
    const volts = ['v1', 'v2', 'v3'].map((key) => asNumber(row[key])).filter((value) => value !== null);
    const vmax = asNumber(row.vmax);
    if (volts.length && vmax !== null && vmax < Math.max(...volts)) {
      warn({ code: 'VMAX_BELOW_READINGS', sheet: 'stepTouchVoltage', path: `stepTouchVoltage.${rowIndex}.vmax`, message: `${label} row ${rowIndex + 1}: Vmax ${vmax} V is below the largest of V1–V3 (${Math.max(...volts)} V).` });
    }
  });
}

module.exports = { validateReport, isNumeric, asNumber, isBlank, JOB_HEADER, optionValue };
