// Data-driven entry screens for the approved measurement sheets MS-01..MS-09.
//
// Everything here is generated from the sheet definitions the server serves at
// GET /api/sheets (src/sheetDefinitions.js): the header cells and row columns, in the
// order and with the labels printed on the sheets, the options of each tick group, and
// the default values of a new row. There are no limits, statuses or rules in this file;
// those come from the server's assessment (see app.js, rowAssessment / sourceAssessment).
//
// Loaded before app.js. It uses app.js helpers (state, escapeHtml, safeText, pill,
// buildRowId, rowAssessment, ...) only when functions run, never at load time.

let SHEET_DEFS = null;
let ACCEPTANCE_PROFILES = [];

const SOIL_DIRECTION_KEYS = ['direction1', 'direction2', 'direction3', 'direction4', 'direction5', 'direction6', 'direction7', 'direction8'];
const SOIL_PAGE_SIZE = 4;
const LEGACY_MATERIALS = { Copper: 'Cu', 'Copper Bonded': 'Cu-bonded' };
// Equipment ids kept on the project only so the (unchanged) PDF instrument list still has a value.
const LEGACY_EQUIPMENT_IDS = ['mi3152', 'mi3290', 'kyoritsu4118a'];

function setSheetDefinitions(definitions) {
  SHEET_DEFS = definitions;
}

function setAcceptanceProfiles(profiles) {
  ACCEPTANCE_PROFILES = Array.isArray(profiles) ? profiles : [];
}

function allSheets() {
  return SHEET_DEFS ? SHEET_DEFS.sheets : [];
}

function sheetById(id) {
  return allSheets().find((sheet) => sheet.id === id) || null;
}

function optionList(field) {
  return (field.options || []).map((option) => (option && typeof option === 'object' ? option : { value: option, label: option }));
}

function flattenSheetFields(fields) {
  return (fields || []).flatMap((field) => (field.kind === 'group' ? field.fields : [field]));
}

function defaultFieldValue(field) {
  if (field.kind === 'ticks' && field.multi) {
    return [];
  }
  if (field.kind === 'tick') {
    return false;
  }
  return Object.prototype.hasOwnProperty.call(field, 'default') ? field.default : '';
}

// ---- New rows, groups, locations, headers --------------------------------------------

function newSheetRow(sheetId, overrides = {}) {
  const sheet = sheetById(sheetId);
  const row = { rowId: buildRowId(sheetId), rowObservation: '', rowPhotos: [] };
  flattenSheetFields(sheet ? sheet.rowFields : []).forEach((field) => {
    if (field.kind === 'serial') {
      row[field.key] = '';
    } else if (field.kind === 'fixed') {
      row[field.key] = field.values ? field.values[0] : '';
    } else {
      row[field.key] = defaultFieldValue(field);
    }
  });
  return Object.assign(row, overrides);
}

// Loop and PFC rows come in groups of three (R, Y, B phase) that share the merged cells.
function newPhaseGroup(sheetId, startIndex = 1) {
  const sheet = sheetById(sheetId);
  const phaseField = sheet.rowFields.find((field) => field.kind === 'fixed');
  return phaseField.values.map((phase, index) => newSheetRow(sheetId, { srNo: String(startIndex + index), [phaseField.key]: phase }));
}

function newTowerGroup(srNo = '') {
  const sheet = sheetById('towerFootingResistance');
  const feet = (sheet.rowFields.find((field) => field.kind === 'fixed') || { values: [] }).values;
  return {
    groupId: buildRowId('tower-group'),
    srNo,
    mainLocationTower: '',
    readings: feet.map((foot) => {
      const reading = newSheetRow('towerFootingResistance', { measurementPointLocation: foot, footLabel: foot });
      delete reading.srNo;
      delete reading.mainLocationTower;
      return reading;
    }),
    totalImpedanceZt: '',
    totalCurrentItotal: '',
    standardTolerableImpedanceZsat: String(sheet.zsat === undefined ? '' : sheet.zsat),
    remarks: ''
  };
}

function soilPresetFor(presetId) {
  const presets = SHEET_DEFS ? SHEET_DEFS.soilSpacingPresets : {};
  return presets[presetId] || presets[SHEET_DEFS ? SHEET_DEFS.defaultSoilSpacingPreset : 'etpl-default'] || { spacings: [], directions: 4 };
}

function newSoilRow(spacing = '') {
  return { rowId: buildRowId('soil'), spacing, resistivity: '', remark: '', rowObservation: '', rowPhotos: [] };
}

function newSoilLocation(name = 'Location 1', presetId = 'etpl-default') {
  const preset = soilPresetFor(presetId);
  const location = {
    locationId: buildRowId('soil-location'),
    name,
    inputUnit: '',
    probeDepthCm: '',
    weather: '',
    bearings: Object.fromEntries(SOIL_DIRECTION_KEYS.map((key) => [key, ''])),
    drivenElectrodeDiameter: '',
    drivenElectrodeLength: '',
    notes: ''
  };
  SOIL_DIRECTION_KEYS.forEach((key, index) => {
    location[key] = index < preset.directions ? preset.spacings.map((spacing) => newSoilRow(spacing)) : [];
  });
  return location;
}

function newSheetHeader(sheetId) {
  const sheet = sheetById(sheetId);
  const header = { instrument: '', serialNo: '', calibrationDue: '', timeStart: '', timeFinish: '' };
  flattenSheetFields(sheet ? sheet.headerFields : []).forEach((field) => {
    if (!field.readOnly) {
      header[field.key] = field.kind === 'tick' ? false : '';
    }
  });
  return header;
}

function createDraft() {
  const today = new Date().toISOString().slice(0, 10);
  const draft = {
    project: {
      projectNo: '',
      clientName: '',
      siteLocation: '',
      workOrder: '',
      reportDate: today,
      engineerName: '',
      clientRepresentative: '',
      groundCondition: '',
      soilSpacingPreset: SHEET_DEFS ? SHEET_DEFS.defaultSoilSpacingPreset : 'etpl-default',
      equipmentSelections: [...LEGACY_EQUIPMENT_IDS],
      acceptanceProfile: 'ETPL-DEFAULT',
      assumedSoilResistivity: '',
      assumedSoilResistivitySource: '',
      zohoProjectId: '',
      zohoProjectName: '',
      zohoProjectOwner: '',
      zohoProjectStage: ''
    },
    tests: {},
    sheetHeaders: {},
    soilResistivity: { locations: [] },
    ocrImports: {}
  };
  allSheets().forEach((sheet) => {
    draft.tests[sheet.id] = sheet.id === 'soilResistivity' || sheet.id === 'electrodeResistance';
    if (sheet.id === 'soilResistivity') {
      draft.soilResistivity.locations = [newSoilLocation('Location 1', draft.project.soilSpacingPreset)];
    } else if (sheet.layout === 'phaseGroups') {
      draft[sheet.id] = newPhaseGroup(sheet.id, 1);
    } else if (sheet.layout === 'towerGroups') {
      draft[sheet.id] = [newTowerGroup('1')];
    } else {
      draft[sheet.id] = [newSheetRow(sheet.id, { srNo: '1' })];
    }
  });
  return draft;
}

// ---- Restoring a saved draft (older drafts lack the newer fields) ----------------------

function upgradeRow(sheetId, row) {
  const upgraded = { ...newSheetRow(sheetId), ...row };
  if (LEGACY_MATERIALS[upgraded.materialType]) {
    upgraded.materialType = LEGACY_MATERIALS[upgraded.materialType];
  }
  if (!Array.isArray(upgraded.conditionCodes) && Object.prototype.hasOwnProperty.call(newSheetRow(sheetId), 'conditionCodes')) {
    upgraded.conditionCodes = [];
  }
  return upgraded;
}

function upgradeSoilLocation(location) {
  const base = newSoilLocation(safeText(location && location.name, 'Location 1'), 'etpl-default');
  const upgraded = { ...base, ...location };
  upgraded.bearings = { ...base.bearings, ...(location && location.bearings) };
  SOIL_DIRECTION_KEYS.forEach((key) => {
    upgraded[key] = Array.isArray(location && location[key]) ? location[key].map((row) => ({ ...newSoilRow(), ...row })) : [];
  });
  return upgraded;
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
  if (source.sheetHeaders && typeof source.sheetHeaders === 'object') {
    draft.sheetHeaders = cloneSerializable(source.sheetHeaders);
  }
  if (source.soilResistivity && Array.isArray(source.soilResistivity.locations) && source.soilResistivity.locations.length) {
    draft.soilResistivity = { locations: cloneSerializable(source.soilResistivity.locations).map(upgradeSoilLocation) };
  }
  allSheets()
    .filter((sheet) => sheet.id !== 'soilResistivity')
    .forEach((sheet) => {
      if (!Array.isArray(source[sheet.id]) || !source[sheet.id].length) {
        return;
      }
      const rows = cloneSerializable(source[sheet.id]);
      draft[sheet.id] =
        sheet.layout === 'towerGroups'
          ? rows.map((group) => ({ ...newTowerGroup(), ...group, readings: (group.readings || []).map((reading) => upgradeRow(sheet.id, reading)) }))
          : rows.map((row) => upgradeRow(sheet.id, row));
    });
  draft.ocrImports = normalizeRestoredOcrImports(source.ocrImports);
  return draft;
}

// ---- Is there anything entered? (for "recover your draft" and OCR overwrite prompts) ----

function fieldDiffersFromDefault(field, value) {
  if (field.kind === 'serial' || field.kind === 'fixed') {
    return false;
  }
  const fallback = defaultFieldValue(field);
  if (Array.isArray(fallback)) {
    return Array.isArray(value) && value.length > 0;
  }
  return String(value === undefined || value === null ? '' : value).trim() !== String(fallback).trim();
}

function rowHasEntry(sheetId, row) {
  const sheet = sheetById(sheetId);
  if (!sheet || !row) {
    return false;
  }
  return Boolean(
    flattenSheetFields(sheet.rowFields).some((field) => fieldDiffersFromDefault(field, row[field.key])) ||
      rowHasObservationData(row)
  );
}

function sheetHasData(sheetId, source = state.draft) {
  const sheet = sheetById(sheetId);
  if (!sheet) {
    return false;
  }
  const rows = Array.isArray(source && source[sheetId]) ? source[sheetId] : [];
  if (sheet.layout === 'towerGroups') {
    return rows.some(
      (group) =>
        safeText(group.mainLocationTower, '') || (group.readings || []).some((reading) => rowHasEntry(sheetId, reading))
    );
  }
  return rows.some((row) => rowHasEntry(sheetId, row));
}

function sheetHeaderHasData(sheetId, source = state.draft) {
  const header = source && source.sheetHeaders && source.sheetHeaders[sheetId];
  return Boolean(header && Object.values(header).some((value) => (typeof value === 'boolean' ? value : safeText(value, ''))));
}

const draftHasMeaningfulElectrodeData = (source) => sheetHasData('electrodeResistance', source);
const draftHasMeaningfulContinuityData = (source) => sheetHasData('continuityTest', source);
const draftHasMeaningfulLoopData = (source) => sheetHasData('loopImpedanceTest', source);
const draftHasMeaningfulFaultData = (source) => sheetHasData('prospectiveFaultCurrent', source);
const draftHasMeaningfulRiserData = (source) => sheetHasData('riserIntegrityTest', source);
const draftHasMeaningfulEarthContinuityData = (source) => sheetHasData('earthContinuityTest', source);
const draftHasMeaningfulTowerData = (source) => sheetHasData('towerFootingResistance', source);
const draftHasMeaningfulGroundData = (source) => sheetHasData('groundResistance', source);
const draftHasMeaningfulStepTouchData = (source) => sheetHasData('stepTouchVoltage', source);

function draftHasMeaningfulSheetHeaders(source) {
  return allSheets().some((sheet) => sheetHeaderHasData(sheet.id, source));
}

// ---- Editing: one function applies every kind of input -------------------------------

function sheetHeaderFor(sheetId) {
  if (!state.draft.sheetHeaders) {
    state.draft.sheetHeaders = {};
  }
  if (!state.draft.sheetHeaders[sheetId]) {
    state.draft.sheetHeaders[sheetId] = newSheetHeader(sheetId);
  }
  return state.draft.sheetHeaders[sheetId];
}

function sameSpacing(a, b) {
  const x = Number.parseFloat(a);
  const y = Number.parseFloat(b);
  return Number.isFinite(x) && Number.isFinite(y) ? x === y : String(a) === String(b);
}

function codeOrder(code) {
  const legend = SHEET_DEFS ? SHEET_DEFS.conditionCodes.map((entry) => entry.code) : [];
  return legend.indexOf(code);
}

// The value an input holds, by kind. Checkboxes: a single tick is a boolean; condition
// codes toggle membership of the array.
function readInputValue(target, current) {
  if (target.dataset.kind === 'codes') {
    const set = new Set(Array.isArray(current) ? current : []);
    if (target.checked) {
      set.add(target.dataset.code);
    } else {
      set.delete(target.dataset.code);
    }
    return [...set].sort((a, b) => codeOrder(a) - codeOrder(b));
  }
  if (target.type === 'checkbox') {
    return target.checked;
  }
  return target.value;
}

// Returns true when the input belonged to a sheet and the draft was updated.
function applyFieldInput(target) {
  const ds = target.dataset;
  const draft = state.draft;

  if (ds.soilCell !== undefined) {
    const location = draft.soilResistivity.locations[Number(ds.locationIndex)];
    if (!location) {
      return false;
    }
    if (!Array.isArray(location[ds.direction])) {
      location[ds.direction] = [];
    }
    let row = location[ds.direction].find((entry) => sameSpacing(entry.spacing, ds.spacing));
    if (!row) {
      row = newSoilRow(ds.spacing);
      location[ds.direction].push(row);
    }
    row[ds.field] = target.value;
    return true;
  }

  if (ds.soilLocation !== undefined) {
    const location = draft.soilResistivity.locations[Number(ds.soilLocation)];
    if (!location) {
      return false;
    }
    if (ds.bearing) {
      location.bearings = location.bearings || {};
      location.bearings[ds.bearing] = target.value;
    } else {
      location[ds.field] = target.value;
    }
    return true;
  }

  if (ds.headerSheet !== undefined) {
    const header = sheetHeaderFor(ds.headerSheet);
    header[ds.field] = readInputValue(target, header[ds.field]);
    return true;
  }

  if (ds.section !== undefined) {
    const sheet = sheetById(ds.section);
    const rows = draft[ds.section];
    const index = Number(ds.index);
    if (!sheet || !Array.isArray(rows) || !rows[index]) {
      return false;
    }
    if (sheet.layout === 'towerGroups' && ds.readingIndex !== undefined) {
      const reading = rows[index].readings[Number(ds.readingIndex)];
      reading[ds.field] = readInputValue(target, reading[ds.field]);
      return true;
    }
    if (ds.groupSync === 'true' && sheet.layout === 'phaseGroups') {
      const size = sheet.groupSize || 3;
      const start = Math.floor(index / size) * size;
      const value = readInputValue(target, rows[index][ds.field]);
      for (let offset = 0; offset < size; offset += 1) {
        if (rows[start + offset]) {
          rows[start + offset][ds.field] = Array.isArray(value) ? [...value] : value;
        }
      }
      return true;
    }
    rows[index][ds.field] = readInputValue(target, rows[index][ds.field]);
    return true;
  }
  return false;
}

// ---- Row operations -------------------------------------------------------------------

function renumberSheet(sheetId) {
  const sheet = sheetById(sheetId);
  const rows = state.draft[sheetId];
  if (!sheet || !Array.isArray(rows)) {
    return;
  }
  rows.forEach((row, index) => {
    if (Object.prototype.hasOwnProperty.call(row, 'srNo')) {
      row.srNo = String(index + 1);
    }
  });
  if (sheet.layout === 'phaseGroups') {
    const phaseField = sheet.rowFields.find((field) => field.kind === 'fixed');
    rows.forEach((row, index) => {
      row[phaseField.key] = phaseField.values[index % phaseField.values.length];
    });
  }
}

function addSheetRow(sheetId) {
  const sheet = sheetById(sheetId);
  const rows = state.draft[sheetId];
  if (!sheet || !Array.isArray(rows)) {
    return;
  }
  if (sheet.layout === 'phaseGroups') {
    rows.push(...newPhaseGroup(sheetId, rows.length + 1));
  } else if (sheet.layout === 'towerGroups') {
    rows.push(newTowerGroup(String(rows.length + 1)));
  } else {
    rows.push(newSheetRow(sheetId, { srNo: String(rows.length + 1) }));
  }
}

function removeSheetRow(sheetId, index) {
  const sheet = sheetById(sheetId);
  const rows = state.draft[sheetId];
  if (!sheet || !Array.isArray(rows)) {
    return;
  }
  if (state.observationEditor && state.observationEditor.section === sheetId) {
    closeObservationEditor();
  }
  if (sheet.layout === 'phaseGroups') {
    const size = sheet.groupSize || 3;
    if (rows.length <= size) {
      showToast('Keep at least one 3-phase group on this sheet.', 'warning');
      return;
    }
    rows.splice(Math.floor(index / size) * size, size);
  } else if (rows.length === 1) {
    showToast('Keep at least one row on this sheet.', 'warning');
    return;
  } else {
    rows.splice(index, 1);
  }
  renumberSheet(sheetId);
}

function addSoilLocation() {
  const locations = state.draft.soilResistivity.locations;
  locations.push(newSoilLocation(`Location ${locations.length + 1}`, state.draft.project.soilSpacingPreset));
}

function removeSoilLocation(locationIndex) {
  const locations = state.draft.soilResistivity.locations;
  if (locations.length === 1) {
    showToast('Keep at least one soil location in the report.', 'warning');
    return;
  }
  locations.splice(locationIndex, 1);
}

function addSoilSpacing(locationIndex, value) {
  const location = state.draft.soilResistivity.locations[locationIndex];
  const spacing = String(value || '').trim();
  if (!location) {
    return;
  }
  if (!/^\d+(\.\d+)?$/.test(spacing) || Number(spacing) <= 0) {
    showToast('Enter the probe spacing as a number above 0 (m).', 'warning');
    return;
  }
  if (!Array.isArray(location.direction1)) {
    location.direction1 = [];
  }
  if (!location.direction1.some((row) => sameSpacing(row.spacing, spacing))) {
    location.direction1.push(newSoilRow(spacing));
  }
}

function removeSoilSpacing(locationIndex, spacing) {
  const location = state.draft.soilResistivity.locations[locationIndex];
  if (!location) {
    return;
  }
  SOIL_DIRECTION_KEYS.forEach((key) => {
    location[key] = (location[key] || []).filter((row) => !sameSpacing(row.spacing, spacing));
  });
}

// Changing the project's spacing set re-seeds locations that hold no readings yet;
// locations with readings keep them (the matrix shows the union of spacings).
function applySoilPreset() {
  const preset = state.draft.project.soilSpacingPreset;
  state.draft.soilResistivity.locations = state.draft.soilResistivity.locations.map((location, index) => {
    return soilLocationHasData(location) && SOIL_DIRECTION_KEYS.some((key) => (location[key] || []).some((row) => safeText(row.resistivity, '')))
      ? location
      : { ...newSoilLocation(safeText(location.name, `Location ${index + 1}`), preset), locationId: location.locationId, inputUnit: location.inputUnit, probeDepthCm: location.probeDepthCm, weather: location.weather, notes: location.notes };
  });
}

// Handles the sheet-related data-action buttons. Returns true when it handled one.
// ---- Tick picker (tablet landscape) ------------------------------------------------------
// The whole page is re-rendered on every input, so the open picker is remembered here and
// reopened by render() through restoreTickPicker().

let openTickPicker = null;

function closeTickPicker() {
  if (!openTickPicker) {
    return;
  }
  openTickPicker = null;
  document.querySelectorAll('.tick-cell.is-open').forEach((cell) => {
    const group = cell.querySelector('.tick-group');
    cell.classList.remove('is-open');
    if (group && group.hasAttribute('popover')) {
      try {
        group.hidePopover();
      } catch (error) {
        // already closed
      }
      group.removeAttribute('popover');
    }
  });
}

function restoreTickPicker() {
  if (!openTickPicker) {
    return;
  }
  const cell = document.querySelector(`[data-tick-cell="${openTickPicker}"]`);
  const group = cell && cell.querySelector('.tick-group');
  const button = cell && cell.querySelector('.tick-summary');
  if (!group || !button || typeof group.showPopover !== 'function' || window.getComputedStyle(button).display === 'none') {
    openTickPicker = null;
    return;
  }
  group.setAttribute('popover', 'manual');
  group.showPopover();
  const rect = button.getBoundingClientRect();
  group.style.top = `${Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - group.offsetHeight - 8))}px`;
  group.style.left = `${Math.max(8, Math.min(rect.left, window.innerWidth - group.offsetWidth - 8))}px`;
}

document.addEventListener('click', (event) => {
  if (openTickPicker && event.target && event.target.closest && !event.target.closest('.tick-cell')) {
    closeTickPicker();
  }
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    closeTickPicker();
  }
});
document.addEventListener('change', (event) => {
  // Choosing one of several closes the picker; condition codes stay open for more ticks.
  if (openTickPicker && event.target && event.target.type === 'radio') {
    openTickPicker = null;
  }
});
window.addEventListener('resize', closeTickPicker);

// ---- Faster entry ---------------------------------------------------------------------------
// Fill down, copy previous row, Enter / Tab flow, a new row when the last one is used, and
// pasting a block copied from Excel. All of them change the draft the same way typing does.

const isEmptyValue = (value) => (Array.isArray(value) ? value.length === 0 : value === undefined || value === null || String(value).trim() === '');
const cloneValue = (value) => (Array.isArray(value) ? [...value] : value);

// The objects holding one column's values, in reading order (a phase group is one target for
// its shared columns; a tower foot column has one target per foot).
function columnTargets(sheet, field) {
  const rows = state.draft[sheet.id] || [];
  if (sheet.layout === 'towerGroups') {
    return field.merge ? rows.map((group) => [group]) : rows.flatMap((group) => (group.readings || []).map((reading) => [reading]));
  }
  if (sheet.layout === 'phaseGroups' && field.merge) {
    const size = sheet.groupSize || 3;
    const groups = [];
    for (let start = 0; start < rows.length; start += size) {
      groups.push(rows.slice(start, start + size));
    }
    return groups;
  }
  return rows.map((row) => [row]);
}

// Empty cells of a column take the first value entered in it.
function fillDown(sheetId, key) {
  const sheet = sheetById(sheetId);
  const field = sheet && sheet.rowFields.find((entry) => entry.key === key);
  if (!field) {
    return false;
  }
  const targets = columnTargets(sheet, field);
  const source = targets.find((group) => !isEmptyValue(group[0][key]));
  if (!source) {
    return false;
  }
  targets.forEach((group) => {
    if (isEmptyValue(group[0][key])) {
      group.forEach((target) => {
        target[key] = cloneValue(source[0][key]);
      });
    }
  });
  return true;
}

// Type, material, dimensions and device fields of the row above.
function copyPreviousRow(sheetId, index) {
  const sheet = sheetById(sheetId);
  const rows = state.draft[sheetId];
  if (!sheet || !Array.isArray(rows) || !rows[index]) {
    return false;
  }
  const size = sheet.layout === 'phaseGroups' ? sheet.groupSize || 3 : 1;
  const start = index - (index % size);
  if (start - size < 0) {
    return false;
  }
  const fields = sheet.rowFields.filter((field) => field.copyable && !field.legacy);
  for (let offset = 0; offset < size; offset += 1) {
    if (rows[start + offset] && rows[start - size + offset]) {
      fields.forEach((field) => {
        rows[start + offset][field.key] = cloneValue(rows[start - size + offset][field.key]);
      });
    }
  }
  return true;
}

// Called by app.js after a sheet cell changed: a new row appears when the last row is used.
function autoAddRow(target) {
  const ds = target && target.dataset;
  const sheet = ds && ds.section !== undefined ? sheetById(ds.section) : null;
  if (!sheet || !['rows', 'phaseGroups'].includes(sheet.layout)) {
    return false;
  }
  const rows = state.draft[sheet.id] || [];
  const size = sheet.layout === 'phaseGroups' ? sheet.groupSize || 3 : 1;
  if (Number(ds.index) < rows.length - size || !rows.slice(rows.length - size).some((row) => rowHasEntry(sheet.id, row))) {
    return false;
  }
  addSheetRow(sheet.id);
  return true;
}

function cellSelector(section, index, field, readingIndex) {
  return `[data-section="${section}"][data-index="${index}"][data-field="${field}"]${readingIndex !== undefined && readingIndex !== null ? `[data-reading-index="${readingIndex}"]` : ''}`;
}

// Enter: the next row of the same column (a new row after the last one).
function moveDown(target) {
  const ds = target.dataset;
  const sheet = sheetById(ds.section);
  const rows = state.draft[ds.section];
  const field = sheet && sheet.rowFields.find((entry) => entry.key === ds.field);
  if (!sheet || !field || !Array.isArray(rows)) {
    return;
  }
  const index = Number(ds.index);
  let next;
  if (sheet.layout === 'towerGroups') {
    const reading = ds.readingIndex === undefined ? null : Number(ds.readingIndex);
    if (reading === null || field.merge) {
      next = { index: index + 1 };
    } else if (reading + 1 < (rows[index].readings || []).length) {
      next = { index, reading: reading + 1 };
    } else {
      next = { index: index + 1, reading: 0 };
    }
  } else {
    const size = sheet.layout === 'phaseGroups' && field.merge ? sheet.groupSize || 3 : 1;
    next = { index: index + size };
  }
  let selector = () => cellSelector(ds.section, next.index, ds.field, next.reading);
  if (!document.querySelector(selector())) {
    if (next.index >= rows.length && sheet.layout !== 'soilMatrix') {
      addSheetRow(ds.section);
      scheduleDraftAutosave();
      render();
    }
  }
  const element = document.querySelector(selector());
  if (element) {
    element.focus();
    if (typeof element.select === 'function') {
      element.select();
    }
  }
}

function matchOption(field, raw) {
  const wanted = String(raw).trim().toLowerCase();
  return optionList(field).find((option) => String(option.value).toLowerCase() === wanted || String(option.label).toLowerCase() === wanted) || null;
}

function pastedValue(field, raw) {
  const text = String(raw).trim();
  if (field.kind === 'ticks' && field.multi) {
    const tokens = text.toUpperCase().split(/[,;\s]+/).filter(Boolean);
    return optionList(field).map((option) => option.value).filter((code) => tokens.includes(String(code).toUpperCase()));
  }
  if (field.kind === 'ticks') {
    const option = matchOption(field, text);
    return option ? option.value : undefined;
  }
  if (field.kind === 'tick') {
    return /^(1|true|yes|y|x|✓)$/i.test(text);
  }
  return text;
}

// Paste a tab-separated block (copied from Excel) from the focused cell, across and down.
function pasteGrid(target, grid) {
  const ds = target.dataset;
  if (ds.soilCell !== undefined) {
    const table = target.closest('table');
    const trs = [...table.querySelectorAll('tbody tr.sheet-row')];
    const row0 = trs.indexOf(target.closest('tr'));
    const col0 = [...target.closest('tr').querySelectorAll('[data-soil-cell]')].indexOf(target);
    grid.forEach((line, r) => {
      const inputs = trs[row0 + r] ? [...trs[row0 + r].querySelectorAll('[data-soil-cell]')] : [];
      line.forEach((raw, c) => {
        if (inputs[col0 + c]) {
          applyFieldInput({ dataset: inputs[col0 + c].dataset, type: 'text', value: String(raw).trim() });
        }
      });
    });
    return true;
  }
  const sheet = sheetById(ds.section);
  const rows = state.draft[ds.section];
  if (!sheet || !Array.isArray(rows)) {
    return false;
  }
  const columns = visibleRowFields(sheet, rows).filter((field) => !['serial', 'fixed'].includes(field.kind) && !field.legacy);
  const startColumn = columns.findIndex((field) => field.key === ds.field);
  if (startColumn < 0) {
    return false;
  }
  const size = sheet.layout === 'phaseGroups' ? sheet.groupSize || 3 : 1;
  let rowIndex = Number(ds.index);
  let reading = ds.readingIndex === undefined ? 0 : Number(ds.readingIndex);
  grid.forEach((line) => {
    while (rows.length <= rowIndex) {
      addSheetRow(sheet.id);
    }
    line.forEach((raw, c) => {
      const field = columns[startColumn + c];
      if (!field) {
        return;
      }
      const value = pastedValue(field, raw);
      if (value === undefined || (field.kind !== 'tick' && field.kind !== 'ticks' && String(raw).trim() === '')) {
        return;
      }
      if (sheet.layout === 'towerGroups') {
        if (field.merge) {
          if (reading === 0) rows[rowIndex][field.key] = value;
        } else if ((rows[rowIndex].readings || [])[reading]) {
          rows[rowIndex].readings[reading][field.key] = value;
        }
        return;
      }
      if (field.merge && sheet.layout === 'phaseGroups') {
        if (rowIndex % size === 0) {
          for (let offset = 0; offset < size; offset += 1) {
            if (rows[rowIndex + offset]) rows[rowIndex + offset][field.key] = cloneValue(value);
          }
        }
        return;
      }
      rows[rowIndex][field.key] = cloneValue(value);
    });
    if (sheet.layout === 'towerGroups') {
      reading += 1;
      if (reading >= (rows[rowIndex].readings || []).length) {
        reading = 0;
        rowIndex += 1;
      }
    } else {
      rowIndex += 1;
    }
  });
  return true;
}

document.addEventListener('keydown', (event) => {
  const target = event.target;
  if (event.key === 'Enter' && !event.shiftKey && !event.isComposing && target && target.matches && target.matches('input.table-input[data-section]') && state.view === 'builder') {
    event.preventDefault();
    moveDown(target);
  }
});

document.addEventListener('paste', (event) => {
  const target = event.target;
  if (!target || !target.matches || !target.matches('input[data-section], input[data-soil-cell]') || state.view !== 'builder') {
    return;
  }
  const text = (event.clipboardData && event.clipboardData.getData('text/plain')) || '';
  const trimmed = text.replace(/\r/g, '').replace(/\n+$/, '');
  if (!/[\t\n]/.test(trimmed)) {
    return; // one value: the ordinary paste
  }
  event.preventDefault();
  if (pasteGrid(target, trimmed.split('\n').map((line) => line.split('\t')))) {
    scheduleDraftAutosave();
    rerenderPreservingFocus();
  }
});

// ---- Step status dots -------------------------------------------------------------------------
// not-started | in-progress | issues | complete. Built from the server's assessments and
// validation, not from rules of its own. A row marked Not Measurable counts as complete.

const STEP_STATUS_LABEL = { 'not-started': 'Not started', 'in-progress': 'In progress', issues: 'Has issues', complete: 'Complete' };

function soilStepStatus(problems) {
  const locations = (state.draft.soilResistivity && state.draft.soilResistivity.locations) || [];
  const keys = ['direction1', 'direction2', 'direction3', 'direction4', 'direction5', 'direction6', 'direction7', 'direction8'];
  const filled = (row) => !isEmptyValue(row && row.resistivity);
  const any = locations.some((location) => keys.some((key) => (location[key] || []).some(filled)));
  if (!any) {
    return problems ? 'issues' : 'not-started';
  }
  if (problems) {
    return 'issues';
  }
  const complete = locations.every((location) => !isEmptyValue(location.inputUnit) && keys.every((key) => (location[key] || []).length === 0 || (location[key] || []).every(filled)));
  return complete ? 'complete' : 'in-progress';
}

function stepStatus(stepId) {
  const project = state.draft.project;
  if (stepId === 'project') {
    const required = ['projectNo', 'clientName', 'siteLocation', 'workOrder', 'engineerName'];
    const count = required.filter((key) => !isEmptyValue(project[key])).length;
    if (count === 0) return 'not-started';
    return count === required.length && !isEmptyValue(project.reportDate) ? 'complete' : 'in-progress';
  }
  if (stepId === 'selection') {
    return Object.keys(state.draft.tests || {}).some((id) => state.draft.tests[id]) ? 'complete' : 'not-started';
  }
  const sheet = sheetById(stepId);
  if (!sheet) {
    return 'not-started';
  }
  const validation = assessmentSummary(state.draft).validation || { errors: [], warnings: [] };
  const problems = validation.errors.some((item) => item.sheet === stepId) || validation.warnings.some((item) => item.sheet === stepId);
  if (sheet.id === 'soilResistivity') {
    return soilStepStatus(problems);
  }
  const rows = (Array.isArray(state.draft[stepId]) ? state.draft[stepId] : []).filter((row) => (sheet.layout === 'towerGroups' ? safeText(row.mainLocationTower, '') || (row.readings || []).some((reading) => rowHasEntry(stepId, reading)) : rowHasEntry(stepId, row)));
  if (!rows.length) {
    return problems ? 'issues' : 'not-started';
  }
  const assessments = rows.map((row) => rowAssessment(stepId, row));
  if (problems || assessments.some((assessment) => assessment.status.tone === 'critical')) {
    return 'issues';
  }
  const unfinished = assessments.some((assessment) => assessment.status.tone === 'neutral' && assessment.status.label !== 'Not Measurable');
  return unfinished ? 'in-progress' : 'complete';
}

// ---- Explanation popover (hover, focus or tap on an "i" button) ------------------------------
// One element outside #app, so the page re-render does not destroy it.

let explainPinned = null;

function explainPopover() {
  let pop = document.getElementById('explainPop');
  if (!pop) {
    pop = document.createElement('div');
    pop.id = 'explainPop';
    pop.className = 'explain-pop';
    pop.setAttribute('popover', 'manual');
    pop.setAttribute('role', 'tooltip');
    document.body.appendChild(pop);
  }
  return pop;
}

function showExplain(button) {
  const pop = explainPopover();
  pop.textContent = button.dataset.explain;
  if (typeof pop.showPopover === 'function' && !pop.matches(':popover-open')) {
    pop.showPopover();
  }
  const rect = button.getBoundingClientRect();
  const width = pop.offsetWidth;
  pop.style.top = `${Math.max(8, Math.min(rect.bottom + 6, window.innerHeight - pop.offsetHeight - 8))}px`;
  pop.style.left = `${Math.max(8, Math.min(rect.right - width, window.innerWidth - width - 8))}px`;
}

function hideExplain() {
  explainPinned = null;
  const pop = document.getElementById('explainPop');
  if (pop && typeof pop.hidePopover === 'function' && pop.matches(':popover-open')) {
    pop.hidePopover();
  }
}

document.addEventListener('mouseover', (event) => {
  const button = event.target.closest && event.target.closest('.explain-btn');
  if (button && !explainPinned) {
    showExplain(button);
  }
});
document.addEventListener('mouseout', (event) => {
  if (event.target.closest && event.target.closest('.explain-btn') && !explainPinned) {
    hideExplain();
  }
});
document.addEventListener('focusin', (event) => {
  const button = event.target.closest && event.target.closest('.explain-btn');
  if (button) {
    showExplain(button);
  }
});
document.addEventListener('focusout', (event) => {
  if (event.target.closest && event.target.closest('.explain-btn') && !explainPinned) {
    hideExplain();
  }
});
document.addEventListener('click', (event) => {
  const button = event.target.closest && event.target.closest('.explain-btn');
  if (button) {
    if (explainPinned === button.dataset.explain) {
      hideExplain();
    } else {
      explainPinned = button.dataset.explain;
      showExplain(button);
    }
    return;
  }
  if (explainPinned) {
    hideExplain();
  }
});

function handleSheetAction(action, button) {
  if (action === 'add-row' && sheetById(button.dataset.section)) {
    addSheetRow(button.dataset.section);
    return true;
  }
  if (action === 'remove-row' && sheetById(button.dataset.section)) {
    removeSheetRow(button.dataset.section, Number(button.dataset.index));
    return true;
  }
  if (action === 'remove-tower-group') {
    removeSheetRow('towerFootingResistance', Number(button.dataset.groupIndex));
    return true;
  }
  if (action === 'add-soil-location') {
    addSoilLocation();
    return true;
  }
  if (action === 'remove-soil-location') {
    removeSoilLocation(Number(button.dataset.locationIndex));
    return true;
  }
  if (action === 'add-soil-spacing') {
    const input = document.getElementById(`soilSpacingInput-${button.dataset.locationIndex}`);
    addSoilSpacing(Number(button.dataset.locationIndex), input ? input.value : '');
    return true;
  }
  if (action === 'remove-soil-spacing') {
    removeSoilSpacing(Number(button.dataset.locationIndex), button.dataset.spacing);
    return true;
  }
  if (action === 'fill-down') {
    fillDown(button.dataset.section, button.dataset.field);
    scheduleDraftAutosave();
    return true;
  }
  if (action === 'copy-prev-row') {
    copyPreviousRow(button.dataset.section, Number(button.dataset.index));
    scheduleDraftAutosave();
    return true;
  }
  if (action === 'open-tick-picker') {
    openTickPicker = openTickPicker === button.dataset.picker ? null : button.dataset.picker;
    return true;
  }
  if (action === 'soil-page') {
    state.soilPage = Number(button.dataset.page) || 0;
    return true;
  }
  return false;
}

// ---- Rendering: controls --------------------------------------------------------------

const esc = (value) => escapeHtml(value === undefined || value === null ? '' : value);

function dataAttrs(attrs) {
  return Object.entries(attrs)
    .filter(([, value]) => value !== undefined && value !== null && value !== false)
    .map(([key, value]) => `data-${key.replace(/[A-Z]/g, (c) => `-${c.toLowerCase()}`)}="${esc(value === true ? 'true' : value)}"`)
    .join(' ');
}

// A tick group: radio buttons (choose one) or checkboxes (condition codes, choose any),
// each with a 44 px tap target.
function tickSummaryText(field, value) {
  const multi = Boolean(field.multi);
  const chosen = optionList(field).filter((option) => (multi ? Array.isArray(value) && value.includes(option.value) : String(value === undefined || value === null ? '' : value) === String(option.value)));
  return chosen.length ? chosen.map((option) => option.label).join(', ') : 'Select';
}

// In a table cell the tick group is wrapped with a summary button. Wide screens show the
// ticks; on tablets in landscape (1024-1279 px) the ticks collapse to the summary and open
// as a picker on tap (see styles.css).
function renderTicks(field, value, attrs, groupName, extraClass = '', inTable = false) {
  const group = renderTickGroup(field, value, attrs, groupName, extraClass);
  if (!inTable) {
    return group;
  }
  const open = openTickPicker === groupName;
  return `<div class="tick-cell ${open ? 'is-open' : ''}" data-tick-cell="${esc(groupName)}"><button class="tick-summary" type="button" data-action="open-tick-picker" data-picker="${esc(groupName)}" aria-haspopup="true" aria-expanded="${open}" aria-label="${esc(field.label)}: choose"><span class="tick-summary-text">${esc(tickSummaryText(field, value))}</span><span class="tick-summary-caret" aria-hidden="true">&#9662;</span></button>${group}</div>`;
}

function renderTickGroup(field, value, attrs, groupName, extraClass = '') {
  const multi = Boolean(field.multi);
  const legend = SHEET_DEFS ? Object.fromEntries(SHEET_DEFS.conditionCodes.map((entry) => [entry.code, entry.text])) : {};
  return `
    <div class="tick-group ${multi ? 'tick-group-multi' : ''} ${extraClass}" role="${multi ? 'group' : 'radiogroup'}" aria-label="${esc(field.label)}">
      ${optionList(field)
        .map((option) => {
          const checked = multi ? Array.isArray(value) && value.includes(option.value) : String(value === undefined || value === null ? '' : value) === String(option.value);
          const extra = multi ? { kind: 'codes', code: option.value } : {};
          return `
            <label class="tick" ${field.codeLegend && legend[option.value] ? `title="${esc(legend[option.value])}"` : ''}>
              <input type="${multi ? 'checkbox' : 'radio'}" name="${esc(groupName)}" value="${esc(option.value)}" ${checked ? 'checked' : ''} ${dataAttrs({ ...attrs, ...extra })} />
              <span class="tick-box" aria-hidden="true"></span>
              <span class="tick-label">${esc(option.label)}</span>
            </label>
          `;
        })
        .join('')}
    </div>
  `;
}

// An empty date shows "dd/mm/yyyy" in the placeholder colour; the browser's own empty date
// text (which can look like a real, greyed-out date) is hidden until the field is focused.
function dateControl(inputHtml, value) {
  return `<span class="date-field ${value ? '' : 'is-empty'}">${inputHtml}<span class="date-placeholder" aria-hidden="true">dd/mm/yyyy</span></span>`;
}

function renderTextControl(field, value, attrs, cssClass = 'table-input') {
  const isNumber = field.kind === 'number';
  const placeholder = field.placeholder || field.unit;
  const input = `<input class="${cssClass}" type="${field.kind === 'date' ? 'date' : 'text'}" ${isNumber ? 'inputmode="decimal"' : ''} autocomplete="off" aria-label="${esc(field.label)}" ${placeholder && field.kind !== 'date' ? `placeholder="${esc(placeholder)}"` : ''} value="${esc(value)}" ${dataAttrs(attrs)} />`;
  return field.kind === 'date' ? dateControl(input, value) : input;
}

function renderControl(field, value, attrs, groupName, cssClass = 'table-input', inTable = false) {
  if (field.kind === 'ticks') {
    return renderTicks(field, value, attrs, groupName, '', inTable);
  }
  if (field.kind === 'tick') {
    return `<label class="tick"><input type="checkbox" ${value ? 'checked' : ''} ${dataAttrs(attrs)} /><span class="tick-box" aria-hidden="true"></span><span class="tick-label">${esc(field.tickLabel || field.label)}</span></label>`;
  }
  if (field.kind === 'fixed') {
    return `<span class="table-input-static">${esc(field.fromProfile ? (field.value === undefined ? '' : field.value) : value)}</span>`;
  }
  return renderTextControl(field, value, attrs, cssClass);
}

// ---- Rendering: problems, header, notes -----------------------------------------------

function renderProblems(sheetId) {
  const validation = (assessmentSummary(state.draft).validation) || { errors: [], warnings: [] };
  const mine = (item) => item.sheet === sheetId;
  const errors = validation.errors.filter(mine);
  const warnings = validation.warnings.filter(mine);
  if (!errors.length && !warnings.length) {
    return '';
  }
  const list = (items, tone, title) =>
    items.length
      ? `<div class="sheet-problems sheet-problems-${tone}" role="${tone === 'critical' ? 'alert' : 'status'}"><strong>${esc(title)}</strong><ul>${items.map((item) => `<li>${esc(item.message)}</li>`).join('')}</ul></div>`
      : '';
  return `${list(errors, 'critical', 'Fix before saving')}${list(warnings, 'warning', 'Check these (saved with the report)')}`;
}

function renderSheetIntro(sheet) {
  return `<div class="sheet-ref-row"><p class="sheet-ref"><strong>${esc(sheet.code === 'legacy' ? 'Retired sheet' : `${sheet.code} · ${sheet.docNo}`)}</strong> · ${esc(sheet.reference)}</p>${sheet.retired ? '' : `<button class="button button-secondary print-sheets-button" type="button" data-action="open-print-sheets" data-sheet="${esc(sheet.id)}">Print sheets for site</button>`}</div>`;
}

function renderNotes(sheet) {
  const diagram = typeof sheetDiagram === 'function' ? sheetDiagram(sheet.id) : '';
  if ((!sheet.notes || !sheet.notes.length) && !diagram) {
    return '';
  }
  return `<details class="sheet-notes"><summary>How to record this sheet</summary><div class="howto">${diagram ? `<figure class="howto-figure"><figcaption>SET-UP</figcaption>${diagram}</figure>` : ''}<div class="howto-steps"><strong>HOW TO RECORD</strong><ul>${(sheet.notes || []).map((note) => `<li>${esc(note)}</li>`).join('')}</ul></div></div></details>`;
}

function renderCodeLegend() {
  if (!SHEET_DEFS) {
    return '';
  }
  return `<p class="sheet-legend"><strong>Condition codes:</strong> ${SHEET_DEFS.conditionCodes.map((entry) => `<span><b>${esc(entry.code)}</b> ${esc(entry.text)}</span>`).join(' · ')}</p>`;
}

function headerCell(label, content, wide = false) {
  return `<div class="header-cell ${wide ? 'header-cell-wide' : ''}"><span class="header-cell-label">${esc(label)}</span>${content}</div>`;
}

function renderHeaderCells(sheet, header) {
  return flattenHeaderEntries(sheet.headerFields)
    .map((entry) => {
      if (entry.kind === 'group') {
        return headerCell(
          entry.label,
          `<div class="header-group">${entry.fields
            .filter((field) => showIfMatches(field, header))
            .map((field) => `<label class="header-subfield"><span>${esc(field.label)}${field.unit ? ` (${esc(field.unit)})` : ''}</span>${renderControl(field, header[field.key], { headerSheet: sheet.id, field: field.key }, `${sheet.id}-${field.key}`)}</label>`)
            .join('')}</div>`,
          true
        );
      }
      if (!showIfMatches(entry, header)) {
        return '';
      }
      if (entry.readOnly) {
        return headerCell(entry.label, `<span class="table-input-static">${esc(sheet.zsat === undefined ? '-' : sheet.zsat)} Ω</span><small class="header-hint">from the project's acceptance profile</small>`);
      }
      return headerCell(entry.label, renderControl(entry, header[entry.key], { headerSheet: sheet.id, field: entry.key }, `${sheet.id}-${entry.key}`), entry.kind === 'ticks');
    })
    .join('');
}

function flattenHeaderEntries(fields) {
  return fields || [];
}

function showIfMatches(field, values) {
  return !field.showIf || String((values || {})[field.showIf.key] || '') === String(field.showIf.equals);
}

// The printed instrument row plus the sheet's own header cells, with the job-level cells
// (client, project, work date, ground condition) shown for reference.
function renderSheetHeader(sheet) {
  return renderSheetHeaderBase(sheet, { includeSheetCells: true });
}

// ---- Rendering: status cells ------------------------------------------------------------

function flagChips(flags) {
  return (flags || []).map((flag) => `<span class="flag-chip" title="${esc(flag.message)}">${esc(flag.code === 'FOP_DISAGREE' ? 'Re-test: 61.8% unconfirmed' : flag.code === 'FOOT_NOT_GIVEN' ? 'Foot without earth connection' : flag.code)}</span>`).join('');
}

// ---- Rendering: cells and tables --------------------------------------------------------

function visibleRowFields(sheet, rows) {
  return sheet.rowFields.filter((field) => {
    if (field.legacy) {
      return false;
    }
    if (field.showIf) {
      return rows.some((row) => showIfMatches(field, row));
    }
    return true;
  });
}

function renderRowCell(sheet, field, row, index, rowNumber, extra = {}, columnIndex = -1) {
  const attrs = { section: sheet.id, index, field: field.key, ...extra };
  let content;
  if (field.kind === 'serial') {
    content = `<span class="table-input-static">${esc(rowNumber)}</span>`;
  } else if (field.kind === 'fixed') {
    const position = field.values.indexOf(row[field.key]);
    content = `<span class="table-input-static">${esc(position >= 0 ? field.display[position] : row[field.key])}</span>`;
  } else if (field.showIf && !showIfMatches(field, row)) {
    content = '<span class="table-input-static">–</span>';
  } else {
    content = renderControl(field, row[field.key], attrs, `${sheet.id}-${index}-${field.key}-${extra.readingIndex === undefined ? '' : extra.readingIndex}`, 'table-input', true);
  }
  const sticky = columnIndex >= 0 && columnIndex < 2 ? ` sticky-left-${columnIndex + 1}` : '';
  return `<td data-label="${esc(field.label)}" class="sheet-cell sheet-cell-${field.kind}${sticky}">${content}</td>`;
}

// Column widths. Sheets that fit the page (MS-01, 03, 06, 07) take the printed proportions
// (field.width = the fraction in make_sheets_v3.py) as percentages. (Plain percentages: a
// max()/calc() width on <col> is ignored by browsers.)
// Dense sheets (MS-02, 04, 05, 08, 09) give every column a minimum width in pixels so its
// content reads clearly (tick options on one line, numbers whole, Remark >= 160 px, Photo
// ref >= 80 px); the table is as wide as those minimums need and scrolls inside its frame.
const TRAILING_SHARE = { status: 13, action: 4.5, tower: 5, note: 4.5 };
const TRAILING_PX = { status: 150, action: 56, tower: 56, note: 56 };
const REMARK_KEYS = new Set(['remark', 'remarks', 'comment', 'observation']);
const LEFT_STICKY_PX = [44, 150];

function tickOptionPx(option) {
  return 27 + 4 + Math.round(String(option.label).length * 6.8) + 8;
}

function columnMinWidth(field) {
  if (field.kind === 'serial') return LEFT_STICKY_PX[0];
  if (field.kind === 'fixed') return 60;
  if (field.kind === 'number') return 92;
  if (field.kind === 'tick') return 130;
  if (field.kind === 'ticks') {
    const options = optionList(field);
    return Math.max(80, options.reduce((sum, option) => sum + tickOptionPx(option), 0) + 4 * (options.length - 1) + 8);
  }
  if (field.key === 'photoRef') return 90;
  if (REMARK_KEYS.has(field.key)) return 170;
  return field.key === 'tag' || field.key === 'location' || field.key === 'mainLocationTower' ? LEFT_STICKY_PX[1] : 140;
}

function renderColgroup(columns, trailing = ['status', 'action'], dense = false) {
  if (dense) {
    const printed = columns.map((field, index) => `<col class="col-${esc(field.kind)}" style="width:${index === 1 ? Math.max(LEFT_STICKY_PX[1], columnMinWidth(field)) : columnMinWidth(field)}px">`);
    const extra = trailing.map((name) => `<col class="col-${name}" style="width:${TRAILING_PX[name]}px">`);
    return `<colgroup>${printed.join('')}${extra.join('')}</colgroup>`;
  }
  const trailingTotal = trailing.reduce((sum, name) => sum + TRAILING_SHARE[name], 0);
  const printedTotal = columns.reduce((sum, field) => sum + (field.width || 8), 0);
  const unit = (100 - trailingTotal) / printedTotal;
  const printed = columns.map((field) => `<col class="col-${esc(field.kind)}" style="width:${((field.width || 8) * unit).toFixed(3)}%">`);
  const extra = trailing.map((name) => `<col class="col-${name}" style="width:${TRAILING_SHARE[name]}%">`);
  return `<colgroup>${printed.join('')}${extra.join('')}</colgroup>`;
}

// Everything tableSurface needs to draw a sheet table.
function sheetTableOptions(sheet, columns, trailing, extraClass = '') {
  const dense = Boolean(sheet.dense);
  const minPx = dense ? columns.reduce((sum, field, index) => sum + (index === 1 ? Math.max(LEFT_STICKY_PX[1], columnMinWidth(field)) : columnMinWidth(field)), 0) + trailing.reduce((sum, name) => sum + TRAILING_PX[name], 0) : 0;
  return {
    className: `sheet-table ${dense ? 'sheet-dense' : ''} ${extraClass}`.trim(),
    colgroup: renderColgroup(columns, trailing, dense),
    tableStyle: dense ? `min-width:${minPx}px` : ''
  };
}

const FILLABLE_KINDS = new Set(['text', 'number', 'ticks', 'tick']);

function renderSheetHead(columns, withStatus, sheetId = '') {
  const label = (field) => {
    const fill = sheetId && FILLABLE_KINDS.has(field.kind) && !field.legacy ? `<button class="fill-down" type="button" data-action="fill-down" data-section="${esc(sheetId)}" data-field="${esc(field.key)}" aria-label="Fill down ${esc(field.label)}" title="Fill the empty cells in this column with the first value entered">&#8615;</button>` : '';
    return `<th scope="col"><span class="th-label">${esc(field.label)}</span>${fill}</th>`;
  };
  return `<tr>${columns.map((field, index) => label(field).replace('<th ', `<th class="${index < 2 ? `sticky-left-${index + 1}` : ''}" `)).join('')}${withStatus ? '<th class="sticky-right-2" scope="col">Assessment</th>' : ''}<th class="sticky-right-1" scope="col">Action</th></tr>`;
}

// Status cell: a compact badge, priority and the calculated limit stay beside the readings;
// the full explanation opens on hover, focus or tap. `explain` overrides the text (the
// loop and fault sheets show one explanation per feeder, on its first phase row).
function renderStatusCell(sheetId, row, explain) {
  const assessment = rowAssessment(sheetId, row);
  const limit = assessment.limit === null || assessment.limit === undefined ? '' : `<span class="status-limit">Limit ${Number(assessment.limit).toFixed(2)}</span>`;
  const text = explain === undefined ? assessment.comment : explain;
  const button = text ? `<button class="explain-btn" type="button" data-explain="${esc(text)}" aria-label="Why this status">i</button>` : '';
  return `
    <td class="sheet-status-cell sticky-right-2" data-label="Status">
      <div class="status-line">${pill(assessment.status.tone, assessment.status.label)}${assessment.priority ? `<span class="soft-chip">${esc(assessment.priority)}</span>` : ''}${button}</div>
      ${limit}
      ${flagChips(assessment.flags)}
      <div class="auto-eval-text status-explain-inline">${esc(assessment.comment)}</div>
    </td>
  `;
}

// Row action cell with "copy previous row" ahead of remove / observation.
function sheetActionCell(sheet, index, row, withCopy) {
  const cell = actionButtonsCell(sheet.id, index, row).replace('class="table-action-cell"', 'class="table-action-cell sticky-right-1"');
  const copy = withCopy ? `<button class="icon-button icon-button-copy" type="button" data-action="copy-prev-row" data-section="${esc(sheet.id)}" data-index="${index}" aria-label="Copy type, material, dimensions and device from the previous row" title="Copy from previous row">&#10697;</button>` : '';
  return cell.replace('<div class="table-action-buttons">', `<div class="table-action-buttons">${copy}`);
}

function hasCopyable(sheet) {
  return sheet.rowFields.some((field) => field.copyable && !field.legacy);
}

function renderRowsLayout(sheet) {
  const rows = state.draft[sheet.id] || [];
  const columns = visibleRowFields(sheet, rows);
  const copy = hasCopyable(sheet);
  return tableSurface(
    `${sheet.code === 'legacy' ? sheet.title : `${sheet.code} ${sheet.title}`}`,
    'Add rows as needed for the executed field scope.',
    renderSheetHead(columns, true, sheet.id),
    rows
      .map((row, index) => `<tr class="sheet-row">${columns.map((field, columnIndex) => renderRowCell(sheet, field, row, index, index + 1, {}, columnIndex)).join('')}${renderStatusCell(sheet.id, row)}${sheetActionCell(sheet, index, row, copy && index > 0)}</tr>`)
      .join(''),
    `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${esc(sheet.id)}">Add row</button></div>`,
    sheetTableOptions(sheet, columns, ['status', 'action'])
  );
}

// One explanation per feeder: the distinct phase comments, joined.
function groupExplanation(sheetId, rows, start, size) {
  const labelled = rows.slice(start, start + size).map((row) => ({ phase: row.measuredPoints, text: rowAssessment(sheetId, row).comment })).filter((entry) => entry.text);
  const distinct = [...new Set(labelled.map((entry) => entry.text))];
  if (distinct.length <= 1) {
    return distinct[0] || '';
  }
  return labelled.map((entry) => `${entry.phase ? `${entry.phase}: ` : ''}${entry.text}`).join('\n');
}

function renderPhaseGroupsLayout(sheet) {
  const rows = state.draft[sheet.id] || [];
  const size = sheet.groupSize || 3;
  const columns = visibleRowFields(sheet, rows);
  const copy = hasCopyable(sheet);
  const body = rows
    .map((row, index) => {
      const offset = index % size;
      const groupNumber = Math.floor(index / size) + 1;
      const first = rows[index - offset] || row;
      const cells = columns
        .map((field, columnIndex) => {
          const sticky = columnIndex < 2 ? ` sticky-left-${columnIndex + 1}` : '';
          if (field.merge) {
            if (offset !== 0) {
              return '';
            }
            const merged = field.kind === 'serial' ? `<span class="table-input-static">${groupNumber}</span>` : renderControl(field, first[field.key], { section: sheet.id, index, field: field.key, groupSync: true }, `${sheet.id}-${index}-${field.key}`, 'table-input', true);
            return `<td rowspan="${size}" data-label="${esc(field.label)}" class="phase-group-cell sheet-cell sheet-cell-${field.kind}${sticky}">${merged}</td>`;
          }
          return renderRowCell(sheet, field, row, index, index + 1, {}, columnIndex);
        })
        .join('');
      const explain = offset === 0 ? groupExplanation(sheet.id, rows, index, size) : '';
      return `<tr class="sheet-row ${offset === 0 ? 'tower-group-row-start' : ''}">${cells}${renderStatusCell(sheet.id, row, explain)}${sheetActionCell(sheet, index, row, copy && offset === 0 && index > 0)}</tr>`;
    })
    .join('');
  return tableSurface(`${sheet.code} ${sheet.title}`, 'Each feeder is three readings, one per phase; the shared cells are entered once.', renderSheetHead(columns, true, sheet.id), body, `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${esc(sheet.id)}">Add feeder</button></div>`, sheetTableOptions(sheet, columns, ['status', 'action']));
}

function towerObservationButton(groupIndex, readingIndex, reading) {
  const has = rowHasObservationData(reading);
  return `<button class="icon-button icon-button-observation ${has ? 'icon-button-observation-active' : ''}" type="button" data-action="open-row-observation" data-section="towerFootingResistance" data-index="${readingIndex}" data-group-index="${groupIndex}" aria-label="${has ? 'Edit row observation' : 'Add row observation'}" title="${has ? 'Edit row observation' : 'Add row observation'}"><span class="icon-button-symbol">+</span></button>`;
}

function renderTowerLayout(sheet) {
  const groups = state.draft[sheet.id] || [];
  const feetField = sheet.rowFields.find((field) => field.kind === 'fixed');
  const columns = sheet.rowFields.filter((field) => !field.legacy);
  const body = groups
    .map((group, groupIndex) => {
      const assessment = rowAssessment(sheet.id, group);
      const feet = group.readings || [];
      return feet
        .map((reading, readingIndex) => {
          const first = readingIndex === 0;
          const cells = columns
            .map((field, columnIndex) => {
              const sticky = columnIndex < 2 ? ` sticky-left-${columnIndex + 1}` : '';
              if (field.merge) {
                if (!first) {
                  return '';
                }
                const content =
                  field.kind === 'serial'
                    ? `<span class="table-input-static">${groupIndex + 1}</span>`
                    : renderControl(field, group[field.key], { section: sheet.id, index: groupIndex, field: field.key }, `${sheet.id}-${groupIndex}-${field.key}`, 'table-input', true);
                return `<td rowspan="${feet.length}" data-label="${esc(field.label)}" class="tower-group-cell sheet-cell${sticky}">${content}</td>`;
              }
              if (field.kind === 'fixed') {
                const position = feetField.values.indexOf(reading.measurementPointLocation);
                return `<td data-label="${esc(field.label)}" class="tower-foot-label"><span class="table-input-static">${esc(position >= 0 ? feetField.display[position] : reading.measurementPointLocation)}</span></td>`;
              }
              return renderRowCell(sheet, field, reading, groupIndex, groupIndex + 1, { readingIndex }, columnIndex);
            })
            .join('');
          const num = (value) => (value === null || value === undefined ? '-' : esc(value));
          const status = first
            ? `<td rowspan="${feet.length}" class="sheet-status-cell tower-group-cell sticky-right-3" data-label="Assessment">
                 <div class="tower-totals tower-totals-${esc(assessment.status.tone)}">Zt ${num(assessment.totalImpedanceZt)} Ω · Zsat ${assessment.zsat === null || assessment.zsat === undefined ? esc(sheet.zsat) : esc(assessment.zsat)} Ω · <strong>${esc(assessment.status.label)}</strong></div>
                 <div class="status-line">${flagChips(assessment.flags)}${assessment.remark ? `<button class="explain-btn" type="button" data-explain="${esc(assessment.remark)}" aria-label="Why this status">i</button>` : ''}</div>
                 <div class="auto-eval-text status-explain-inline">ΣI ${num(assessment.totalCurrentItotal)} mA. ${esc(assessment.remark)}</div>
               </td>`
            : '';
          const action = first
            ? `<td rowspan="${feet.length}" class="table-action-cell tower-group-action-cell sticky-right-2"><button class="icon-button icon-button-danger tower-group-remove" type="button" data-action="remove-tower-group" data-group-index="${groupIndex}" aria-label="Remove tower ${groupIndex + 1}" title="Remove tower"><span class="icon-button-symbol">&times;</span></button></td>`
            : '';
          return `<tr class="${first ? 'tower-group-row-start' : ''} ${readingIndex === feet.length - 1 ? 'tower-group-row-end' : ''}">${cells}${status}${action}<td class="table-action-cell sticky-right-1">${towerObservationButton(groupIndex, readingIndex, reading)}</td></tr>`;
        })
        .join('');
    })
    .join('');
  const head = `<tr>${columns.map((field, index) => `<th class="${index < 2 ? `sticky-left-${index + 1}` : ''}" scope="col"><span class="th-label">${esc(field.label)}</span>${FILLABLE_KINDS.has(field.kind) ? `<button class="fill-down" type="button" data-action="fill-down" data-section="${esc(sheet.id)}" data-field="${esc(field.key)}" aria-label="Fill down ${esc(field.label)}" title="Fill the empty cells in this column with the first value entered">&#8615;</button>` : ''}</th>`).join('')}<th class="sticky-right-3" scope="col">Assessment</th><th class="sticky-right-2" scope="col">Tower</th><th class="sticky-right-1" scope="col">Note</th></tr>`;
  return tableSurface(
    `${sheet.code} ${sheet.title}`,
    'Each tower has four fixed feet (F-1 to F-4); the tower totals are calculated.',
    head,
    body,
    `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${esc(sheet.id)}">Add tower</button></div>`,
    sheetTableOptions(sheet, columns, ['status', 'tower', 'note'], 'tower-sheet-table')
  );
}

// ---- Rendering: soil (MS-01) --------------------------------------------------------------

function soilSpacingsFor(location, preset) {
  const spacings = [...preset.spacings];
  SOIL_DIRECTION_KEYS.forEach((key) => {
    (location[key] || []).forEach((row) => {
      if (safeText(row.spacing, '') && !spacings.some((spacing) => sameSpacing(spacing, row.spacing))) {
        spacings.push(String(row.spacing));
      }
    });
  });
  return spacings.sort((a, b) => Number(a) - Number(b));
}

function soilDirectionCount(location, preset) {
  let count = preset.directions;
  SOIL_DIRECTION_KEYS.forEach((key, index) => {
    if ((location[key] || []).some((row) => safeText(row.resistivity, ''))) {
      count = Math.max(count, index + 1);
    }
  });
  return count;
}

function renderSoilLocationHeader(sheet, location, locationIndex, locationAssessment) {
  const attrs = (field) => ({ soilLocation: locationIndex, field: field.key });
  const named = Object.fromEntries(sheet.headerFields.map((field) => [field.key, field]));
  const unitField = { ...named.inputUnit };
  return `
    <div class="header-grid">
      ${headerCell('Location name', renderTextControl(named.name, location.name, attrs(named.name)))}
      ${headerCell('Reading unit', renderTicks(unitField, location.inputUnit, attrs(unitField), `soil-${locationIndex}-unit`), true)}
      ${headerCell('Probe depth (cm)', renderTextControl(named.probeDepthCm, location.probeDepthCm, attrs(named.probeDepthCm)))}
      ${['bearingsA', 'bearingsB']
        .map((key) => {
          const group = named[key];
          return headerCell(
            group.label,
            `<div class="header-group">${group.fields.map((field) => `<label class="header-subfield"><span>${esc(field.label)} (°)</span>${renderTextControl(field, (location.bearings || {})[field.key], { soilLocation: locationIndex, bearing: field.key })}</label>`).join('')}</div>`,
            true
          );
        })
        .join('')}
      ${headerCell('Weather', renderTicks(named.weather, location.weather, attrs(named.weather), `soil-${locationIndex}-weather`), true)}
      <div class="header-cell"><span class="header-cell-label">Corrosivity (IS 3043 Table 4)</span>${locationAssessment ? pill(locationAssessment.category.tone, locationAssessment.category.label) : ''}</div>
    </div>
  `;
}

function renderSoilMatrix(sheet, location, locationIndex, preset, locationAssessment, page) {
  const directions = SOIL_DIRECTION_KEYS.slice(page * SOIL_PAGE_SIZE, page * SOIL_PAGE_SIZE + SOIL_PAGE_SIZE);
  const spacings = soilSpacingsFor(location, preset);
  const unit = location.inputUnit === 'ohm' ? 'R (Ω)' : 'ρ (Ω·m)';
  const cell = (direction, spacing, field) => {
    const row = (location[direction] || []).find((entry) => sameSpacing(entry.spacing, spacing));
    return `<input class="table-input" type="text" inputmode="decimal" autocomplete="off" aria-label="${esc(`${direction.replace('direction', 'Direction ')} reading at ${spacing} m`)}" placeholder="${esc(unit)}" value="${esc(row ? row[field] : '')}" ${dataAttrs({ soilCell: true, locationIndex, direction, spacing, field })} />`;
  };
  const remarkRow = (spacing) => {
    const row = (location.direction1 || []).find((entry) => sameSpacing(entry.spacing, spacing));
    return `<input class="table-input" type="text" autocomplete="off" aria-label="Remark at ${esc(spacing)} m" value="${esc(row ? row.remark : '')}" ${dataAttrs({ soilCell: true, locationIndex, direction: 'direction1', spacing, field: 'remark' })} />`;
  };
  const header = `<tr><th>Sr.</th><th>Probe spacing a (m)</th>${directions.map((direction) => `<th>${esc(direction.replace('direction', 'Direction '))} reading</th>`).join('')}${page === 0 ? '<th>Remark</th>' : ''}</tr>`;
  const body = spacings
    .map((spacing, index) => {
      const custom = !preset.spacings.some((entry) => sameSpacing(entry, spacing));
      return `<tr class="sheet-row"><td data-label="Sr.">${index + 1}</td><td data-label="Probe spacing a (m)"><span class="table-input-static">${esc(spacing)}</span>${custom ? `<button class="icon-button icon-button-danger" type="button" data-action="remove-soil-spacing" data-location-index="${locationIndex}" data-spacing="${esc(spacing)}" aria-label="Remove spacing ${esc(spacing)} m">&times;</button>` : ''}</td>${directions.map((direction) => `<td data-label="${esc(direction.replace('direction', 'Direction '))} reading">${cell(direction, spacing, 'resistivity')}</td>`).join('')}${page === 0 ? `<td data-label="Remark">${remarkRow(spacing)}</td>` : ''}</tr>`;
    })
    .join('');
  const means = locationAssessment && locationAssessment.directionAverages ? locationAssessment.directionAverages : {};
  const meanRow = `<tr class="sheet-mean-row"><td></td><td><strong>Mean ρ (Ω·m)</strong></td>${directions.map((direction) => `<td><strong>${means[direction] === null || means[direction] === undefined ? '-' : esc(means[direction])}</strong></td>`).join('')}${page === 0 ? `<td><strong>Location ${locationAssessment && locationAssessment.overallAverage !== null && locationAssessment.overallAverage !== undefined ? esc(locationAssessment.overallAverage) : '-'}</strong></td>` : ''}</tr>`;
  const fixedShare = 4 + 10 + (page === 0 ? 18 : 0);
  const directionShare = ((100 - fixedShare) / Math.max(directions.length, 1)).toFixed(3);
  const colgroup = `<colgroup><col class="col-serial" style="width:4%"><col class="col-fixed" style="width:10%">${directions.map(() => `<col class="col-number" style="width:${directionShare}%">`).join('')}${page === 0 ? '<col class="col-text" style="width:18%">' : ''}</colgroup>`;
  return `<div class="table-wrap sheet-table-wrap"><table class="data-table sheet-table soil-matrix">${colgroup}<thead>${header}</thead><tbody>${body}${meanRow}</tbody></table></div>`;
}

function renderSoilStep() {
  const sheet = sheetById('soilResistivity');
  const preset = soilPresetFor(state.draft.project.soilSpacingPreset);
  const assessment = assessmentSummary(state.draft);
  const soil = assessment.soil;
  const locations = getSoilLocations(state.draft);
  return sectionCard(
    `${sheet.code} ${sheet.title}`,
    'Measurement Sheet',
    `
      ${renderSheetIntro(sheet)}
      ${renderSoilOcrAssist()}
      ${renderProblems(sheet.id)}
      ${renderSheetHeaderBase(sheet, { includeSheetCells: false })}
      <div class="metric-strip">
        <article class="metric-box"><span>Mean soil resistivity</span><strong>${soil.overallAverage === null ? '-' : `${soil.overallAverage} ohm-m`}</strong>${pill(soil.category.tone, soil.category.label)}</article>
        ${soil.corrosionAllowancePct === null || soil.corrosionAllowancePct === undefined ? '' : `<article class="metric-box"><span>Steel corrosion allowance (Table 5)</span><strong>${soil.corrosionAllowancePct}%</strong></article>`}
        <article class="metric-box"><span>Spacing set</span><strong>${esc(preset.label || '')}</strong></article>
      </div>
      <div class="stack-grid">
        ${locations
          .map((location, locationIndex) => {
            const locationAssessment = (soil.locations || [])[locationIndex];
            const directionCount = soilDirectionCount(location, preset);
            const pages = Math.ceil(directionCount / SOIL_PAGE_SIZE);
            const page = Math.min(state.soilPage || 0, pages - 1);
            return `
              <section class="mini-surface mini-surface-location">
                <div class="mini-surface-head mini-surface-head-location">
                  <div><h4>${esc(safeText(location.name, `Location ${locationIndex + 1}`))}</h4><p>${directionCount} direction${directionCount === 1 ? '' : 's'} · ${esc(preset.label || '')} spacings</p></div>
                  <div class="mini-surface-head-actions">${locations.length > 1 ? `<button class="button button-ghost" type="button" data-action="remove-soil-location" data-location-index="${locationIndex}">Remove location</button>` : ''}</div>
                </div>
                ${renderSoilLocationHeader(sheet, location, locationIndex, locationAssessment)}
                ${pages > 1 ? `<div class="soil-pages" role="tablist">${Array.from({ length: pages }, (_, p) => `<button class="step-pill ${p === page ? 'step-pill-active' : ''}" type="button" data-action="soil-page" data-page="${p}">Directions ${p * SOIL_PAGE_SIZE + 1}–${Math.min(directionCount, (p + 1) * SOIL_PAGE_SIZE)}</button>`).join('')}</div>` : ''}
                ${renderSoilMatrix(sheet, location, locationIndex, preset, locationAssessment, page)}
                <div class="mini-surface-foot soil-add-spacing">
                  <label class="field"><span>Add a probe spacing (m)</span><input class="table-input" id="soilSpacingInput-${locationIndex}" type="text" inputmode="decimal" autocomplete="off" placeholder="e.g. 7.5" /></label>
                  <button class="button button-secondary" type="button" data-action="add-soil-spacing" data-location-index="${locationIndex}">Add spacing</button>
                </div>
                <label class="field"><span>Site notes</span><textarea class="text-area" rows="2" data-bind="soilResistivity.locations.${locationIndex}.notes" placeholder="Ground condition, nearby structures, or observations">${esc(location.notes)}</textarea></label>
              </section>
            `;
          })
          .join('')}
      </div>
      <div class="mini-surface-foot"><button class="button button-secondary" type="button" data-action="add-soil-location">Add another location</button></div>
      ${renderNotes(sheet)}
    `
  );
}

function renderSheetHeaderBase(sheet, { includeSheetCells }) {
  const project = state.draft.project;
  const header = (state.draft.sheetHeaders && state.draft.sheetHeaders[sheet.id]) || newSheetHeader(sheet.id);
  const instrumentField = { key: 'instrument', label: 'Instrument used', kind: 'ticks', options: sheet.instrumentOptions };
  return `
    <div class="sheet-header">
      <div class="sheet-job-line">
        <span><b>Client</b> ${esc(project.clientName || '-')}</span>
        <span><b>Project</b> ${esc(project.projectNo || '-')}</span>
        <span><b>Work date</b> ${esc(project.reportDate || '-')}</span>
        <span><b>Ground</b> ${esc(project.groundCondition || '-')}</span>
        <button class="button button-ghost button-small" type="button" data-action="go-step" data-step-index="0">Edit job details</button>
      </div>
      <div class="header-grid">
        ${headerCell('Instrument used', renderTicks(instrumentField, header.instrument, { headerSheet: sheet.id, field: 'instrument' }, `${sheet.id}-instrument`), true)}
        ${headerCell('Instrument Sr. No.', renderTextControl({ label: 'Instrument serial number', kind: 'text' }, header.serialNo, { headerSheet: sheet.id, field: 'serialNo' }))}
        ${headerCell('Calibration due', renderTextControl({ label: 'Calibration due', kind: 'date' }, header.calibrationDue, { headerSheet: sheet.id, field: 'calibrationDue' }))}
        ${headerCell('Time start', renderTextControl({ label: 'Time start', kind: 'text' }, header.timeStart, { headerSheet: sheet.id, field: 'timeStart' }))}
        ${headerCell('Time finish', renderTextControl({ label: 'Time finish', kind: 'text' }, header.timeFinish, { headerSheet: sheet.id, field: 'timeFinish' }))}
        ${includeSheetCells ? renderHeaderCells(sheet, header) : ''}
      </div>
    </div>
  `;
}

// ---- Rendering: a builder step ----------------------------------------------------------------

const SHEET_OCR_ASSIST = {
  electrodeResistance: () => renderElectrodeOcrAssist(),
  continuityTest: () => renderContinuityOcrAssist(),
  loopImpedanceTest: () => renderLoopOcrAssist(),
  prospectiveFaultCurrent: () => renderFaultOcrAssist(),
  riserIntegrityTest: () => renderRiserOcrAssist(),
  towerFootingResistance: () => renderTowerOcrAssist(),
  earthContinuityTest: () => renderEarthContinuityOcrAssist()
};

function renderSheetStep(sheetId) {
  const sheet = sheetById(sheetId);
  if (!sheet) {
    return sectionCard('Loading', 'Measurement Sheet', '<p class="muted">The sheet definitions are still loading.</p>');
  }
  if (sheet.id === 'soilResistivity') {
    return renderSoilStep();
  }
  const hasCodes = sheet.rowFields.some((field) => field.codeLegend);
  const table = sheet.layout === 'phaseGroups' ? renderPhaseGroupsLayout(sheet) : sheet.layout === 'towerGroups' ? renderTowerLayout(sheet) : renderRowsLayout(sheet);
  const banner = sheet.id === 'electrodeResistance' ? `<div class="limit-banner">${esc(assessmentSummary(state.draft).banners.electrode)}</div>` : '';
  return sectionCard(
    `${sheet.code === 'legacy' ? sheet.title : `${sheet.code} ${sheet.title}`}`,
    'Measurement Sheet',
    `
      ${renderSheetIntro(sheet)}
      ${SHEET_OCR_ASSIST[sheet.id] ? SHEET_OCR_ASSIST[sheet.id]() : ''}
      ${banner}
      ${renderProblems(sheet.id)}
      ${renderSheetHeader(sheet)}
      ${table}
      ${hasCodes ? renderCodeLegend() : ''}
      ${renderNotes(sheet)}
    `
  );
}

// ---- Rendering: job-level header cells in the project step -------------------------------------

function renderJobHeaderFields() {
  const project = state.draft.project;
  const assessment = assessmentSummary(state.draft);
  const soilSelected = Boolean(state.draft.tests.soilResistivity);
  const showAssumed = !soilSelected || assessment.basis.rho === null || assessment.basis.rhoAssumed || safeText(project.assumedSoilResistivity, '');
  const profiles = ACCEPTANCE_PROFILES.length ? ACCEPTANCE_PROFILES : [{ id: project.acceptanceProfile || 'ETPL-DEFAULT', description: '' }];
  const groundField = { label: 'Ground condition', multi: false, options: (SHEET_DEFS ? SHEET_DEFS.groundConditions : ['Dry', 'Damp', 'Wet']) };
  const presets = SHEET_DEFS ? SHEET_DEFS.soilSpacingPresets : {};
  const projectProblems = ((assessment.validation || { errors: [], warnings: [] }).errors || []).filter((item) => !item.sheet);
  return `
    <div class="field-grid">
      ${renderField('Client representative', project.clientRepresentative, 'project.clientRepresentative', 'text', 'Name of the client representative')}
      <label class="field">
        <span>Acceptance profile</span>
        <select class="text-input" data-bind="project.acceptanceProfile">
          ${profiles.map((profile) => `<option value="${esc(profile.id)}" ${profile.id === project.acceptanceProfile ? 'selected' : ''}>${esc(profile.id)}${profile.description ? ` – ${esc(profile.description)}` : ''}</option>`).join('')}
        </select>
        <small class="field-hint">Limits and their sources come from this profile in the Standards Acceptance List.</small>
      </label>
      <label class="field">
        <span>Soil probe-spacing set</span>
        <select class="text-input" data-bind="project.soilSpacingPreset">
          ${Object.entries(presets).map(([id, preset]) => `<option value="${esc(id)}" ${id === project.soilSpacingPreset ? 'selected' : ''}>${esc(preset.label)} (${preset.spacings.join(', ')} m; ${preset.directions} directions)</option>`).join('')}
        </select>
      </label>
    </div>
    <div class="header-cell header-cell-wide job-ground">
      <span class="header-cell-label">Ground condition</span>
      ${renderTicks(groundField, project.groundCondition, { bind: 'project.groundCondition' }, 'project-ground')}
    </div>
    ${
      showAssumed
        ? `<div class="assumed-rho-panel">
             <p class="section-kicker">Soil resistivity not measured?</p>
             <p>If soil was not measured, enter an assumed rho and its source; it is printed as an assumption next to every limit it is used for.${assessment.basis.rho !== null && !assessment.basis.rhoAssumed ? ' <strong>Soil has been measured, so remove the assumed value.</strong>' : ''}</p>
             <div class="field-grid">
               ${renderField('Assumed soil resistivity (Ω·m)', project.assumedSoilResistivity, 'project.assumedSoilResistivity', 'text', 'e.g. 30')}
               ${renderField('Source of the assumption', project.assumedSoilResistivitySource, 'project.assumedSoilResistivitySource', 'text', 'e.g. client geotechnical report GT-12')}
             </div>
           </div>`
        : ''
    }
    ${projectProblems.length ? `<div class="sheet-problems sheet-problems-critical" role="alert"><strong>Fix before saving</strong><ul>${projectProblems.map((item) => `<li>${esc(item.message)}</li>`).join('')}</ul></div>` : ''}
  `;
}

// ---- Rendering: report detail (saved reports) --------------------------------------------------

function displayValue(field, value, sheet) {
  if (field.kind === 'ticks' && field.multi) {
    return Array.isArray(value) && value.length ? value.join(', ') : '–';
  }
  if (field.kind === 'ticks') {
    const match = optionList(field).find((option) => String(option.value) === String(value));
    return match ? match.label : safeText(value, '–');
  }
  if (field.kind === 'tick') {
    return value ? 'Yes' : '–';
  }
  return safeText(value, '–');
}

function renderDetailHeader(report, sheet) {
  const header = (report.sheetHeaders && report.sheetHeaders[sheet.id]) || {};
  const instrument = (sheet.instrumentOptions || []).find((option) => option.value === header.instrument);
  const cells = [
    ['Instrument', instrument ? instrument.label : safeText(header.instrument, '–')],
    ['Serial no.', safeText(header.serialNo, '–')],
    ['Calibration due', safeText(header.calibrationDue, '–')],
    ['Time', [header.timeStart, header.timeFinish].filter(Boolean).join(' – ') || '–']
  ];
  flattenSheetFields(sheet.headerFields).forEach((field) => {
    if (field.readOnly) {
      cells.push([field.label, sheet.zsat === undefined ? '–' : `${sheet.zsat} Ω (acceptance profile)`]);
    } else if (header[field.key] !== undefined && header[field.key] !== '' && header[field.key] !== false) {
      cells.push([field.label.replace(/ \(.*\)$/, ''), displayValue(field, header[field.key], sheet)]);
    }
  });
  return `<div class="detail-grid">${cells.map(([label, value]) => `<div><span>${esc(label)}</span><strong>${esc(value)}</strong></div>`).join('')}</div>`;
}

function renderSheetDetail(report, sheet) {
  if (sheet.id === 'soilResistivity') {
    return '';
  }
  const rows = Array.isArray(report[sheet.id]) ? report[sheet.id] : [];
  const columns = visibleRowFields(sheet, rows).filter((field) => field.kind !== 'serial');
  let table;
  if (sheet.layout === 'towerGroups') {
    table = renderTowerDetailTable(report, sheet, rows, columns, sheet.rowFields.find((field) => field.kind === 'fixed'));
  } else {
    table = simpleDetailTable(
      [
        ...columns.map((field) => ({
          label: field.label,
          render: (row) => {
            if (field.kind === 'fixed') {
              const position = field.values.indexOf(row[field.key]);
              return esc(position >= 0 ? field.display[position] : row[field.key]);
            }
            return esc(displayValue(field, row[field.key], sheet));
          }
        })),
        {
          label: 'Assessment',
          render: (row) => {
            const assessment = rowAssessment(sheet.id, row);
            const limit = assessment.limit === null || assessment.limit === undefined ? '' : `<span class="status-limit">Limit ${Number(assessment.limit).toFixed(2)}</span>`;
            return `${pill(assessment.status.tone, assessment.status.label)}${assessment.priority ? ` <span class="soft-chip">${esc(assessment.priority)}</span>` : ''} ${limit}${flagChips(assessment.flags)}<div class="auto-eval-text">${esc(assessment.remark)}</div>`;
          }
        }
      ],
      rows
    );
  }
  return `${renderDetailHeader(report, sheet)}${table}`;
}

function renderTowerDetailTable(report, sheet, groups, columns, feetField) {
  const perFoot = columns.filter((field) => !field.merge && field.kind !== 'fixed');
  return `
    <div class="table-wrap">
      <table class="data-table tower-data-table">
        <thead><tr><th>Sr.</th><th>Tower no. / location</th><th>Foot</th>${perFoot.map((field) => `<th>${esc(field.label)}</th>`).join('')}<th>Assessment</th></tr></thead>
        <tbody>
          ${groups
            .map((group, groupIndex) => {
              const assessment = rowAssessment(sheet.id, group);
              const feet = group.readings || [];
              return feet
                .map((reading, readingIndex) => {
                  const position = feetField.values.indexOf(reading.measurementPointLocation);
                  const first = readingIndex === 0;
                  const observation = renderObservationCard(reading, `${safeText(group.mainLocationTower, `Tower ${groupIndex + 1}`)} | ${safeText(reading.measurementPointLocation, 'Foot')}`);
                  return `
                    <tr class="${first ? 'tower-group-row-start' : ''} ${readingIndex === feet.length - 1 ? 'tower-group-row-end' : ''}">
                      ${first ? `<td class="tower-group-cell" rowspan="${feet.length}">${groupIndex + 1}</td><td class="tower-group-cell" rowspan="${feet.length}">${esc(group.mainLocationTower)}</td>` : ''}
                      <td class="tower-foot-label">${esc(position >= 0 ? feetField.display[position] : reading.measurementPointLocation)}</td>
                      ${perFoot.map((field) => `<td>${esc(displayValue(field, reading[field.key], sheet))}</td>`).join('')}
                      ${first ? `<td class="tower-group-cell" rowspan="${feet.length}">${pill(assessment.status.tone, assessment.status.label)}${flagChips(assessment.flags)}<div class="status-limit">Zt ${assessment.totalImpedanceZt === null || assessment.totalImpedanceZt === undefined ? '-' : esc(assessment.totalImpedanceZt)} Ω · Zsat ${assessment.zsat === null || assessment.zsat === undefined ? esc(sheet.zsat) : esc(assessment.zsat)} Ω</div><div class="auto-eval-text">${esc(assessment.remark)}</div></td>` : ''}
                    </tr>
                    ${observation ? `<tr class="detail-table-observation"><td colspan="${perFoot.length + 4}"><div class="row-observation-stack">${observation}</div></td></tr>` : ''}
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

function renderReviewChecklist(report) {
  const items = (sourceAssessment(report).reviewItems) || [];
  if (!items.length) {
    return '';
  }
  return `
    <div class="review-checklist">
      <p class="section-kicker">Reviewer checklist</p>
      <ul>${items.map((item) => `<li class="${item.blocksApproval ? 'review-item-blocking' : ''}">${item.blocksApproval ? '<strong>Blocks final approval until acknowledged:</strong> ' : ''}${esc(item.text)}</li>`).join('')}</ul>
    </div>
  `;
}

// PDF layout for the new sheets and fields arrives in Session 6C; say so where the
// export button is, rather than letting sections silently be missing.
function renderExportNote(report) {
  const missing = ['groundResistance', 'stepTouchVoltage'].filter((id) => report.tests && report.tests[id]).map((id) => (sheetById(id) ? sheetById(id).code : id));
  return missing.length
    ? `<p class="export-note" role="note">The PDF does not include ${esc(missing.join(' and '))} yet; its layout arrives in Session 6C. The data is saved and shown on this page.</p>`
    : '';
}

// Soil (MS-01) on a saved report: header facts, the matrix of every direction, the means
// and the corrosivity class, all from the stored data and the server's assessment.
function renderSoilDetail(report) {
  const assessment = sourceAssessment(report);
  const soil = assessment.soil;
  const sheet = sheetById('soilResistivity');
  const preset = soilPresetFor(report.project && report.project.soilSpacingPreset);
  const header = (report.sheetHeaders && report.sheetHeaders.soilResistivity) || {};
  const instrument = sheet ? (sheet.instrumentOptions || []).find((option) => option.value === header.instrument) : null;
  return `
    <div class="metric-strip">
      <article class="metric-box"><span>Mean soil resistivity</span><strong>${soil.overallAverage === null ? '-' : `${soil.overallAverage} ohm-m`}</strong>${pill(soil.category.tone, soil.category.label)}</article>
      ${soil.corrosionAllowancePct === null || soil.corrosionAllowancePct === undefined ? '' : `<article class="metric-box"><span>Steel corrosion allowance (Table 5)</span><strong>${soil.corrosionAllowancePct}%</strong></article>`}
    </div>
    ${instrument || header.serialNo || header.calibrationDue ? `<div class="detail-grid"><div><span>Instrument</span><strong>${esc(instrument ? instrument.label : safeText(header.instrument, '–'))}</strong></div><div><span>Serial no.</span><strong>${esc(safeText(header.serialNo, '–'))}</strong></div><div><span>Calibration due</span><strong>${esc(safeText(header.calibrationDue, '–'))}</strong></div></div>` : ''}
    ${(soil.locations || [])
      .map((locationSummary, locationIndex) => {
        const location = getSoilLocations(report)[locationIndex] || {};
        const directionCount = soilDirectionCount(location, preset);
        const spacings = soilSpacingsFor(location, preset).filter((spacing) => SOIL_DIRECTION_KEYS.some((key) => (location[key] || []).some((row) => sameSpacing(row.spacing, spacing) && safeText(row.resistivity, ''))));
        const directions = SOIL_DIRECTION_KEYS.slice(0, directionCount).filter((key) => (location[key] || []).some((row) => safeText(row.resistivity, '')));
        const unit = location.inputUnit === 'ohm' ? 'R (Ω)' : 'ρ (Ω·m)';
        const bearings = SOIL_DIRECTION_KEYS.filter((key) => safeText((location.bearings || {})[key], '')).map((key) => `${key.replace('direction', 'D')} ${(location.bearings || {})[key]}°`).join('  ');
        return `
          <div class="detail-subsection">
            <div class="detail-subsection-head"><h4>${esc(locationSummary.name)}</h4>${pill(locationSummary.category.tone, locationSummary.category.label)}</div>
            <div class="detail-grid">
              <div><span>Reading unit</span><strong>${esc(unit)}</strong></div>
              <div><span>Probe depth</span><strong>${esc(safeText(location.probeDepthCm, '–'))}${safeText(location.probeDepthCm, '') ? ' cm' : ''}</strong></div>
              <div><span>Weather</span><strong>${esc(safeText(location.weather, '–'))}</strong></div>
              <div><span>Bearings</span><strong>${esc(bearings || '–')}</strong></div>
              <div><span>Location mean</span><strong>${locationSummary.overallAverage === null ? '-' : `${locationSummary.overallAverage} ohm-m`}</strong></div>
            </div>
            ${simpleDetailTable(
              [
                { label: 'Probe spacing a (m)', key: 'spacing' },
                ...directions.map((key) => ({ label: `${key.replace('direction', 'Direction ')} reading (${unit})`, key })),
                { label: 'Remark', key: 'remark' }
              ],
              spacings.map((spacing) => {
                const row = { spacing, remark: ((location.direction1 || []).find((entry) => sameSpacing(entry.spacing, spacing)) || {}).remark || '' };
                directions.forEach((key) => {
                  const match = (location[key] || []).find((entry) => sameSpacing(entry.spacing, spacing));
                  row[key] = match && safeText(match.resistivity, '') ? match.resistivity : '–';
                });
                return row;
              }),
              { observationRenderer: () => '' }
            )}
            <p class="detail-note"><strong>Direction means (ohm-m):</strong> ${directions.map((key) => `${key.replace('direction', 'D')} ${locationSummary.directionAverages && locationSummary.directionAverages[key] !== null && locationSummary.directionAverages[key] !== undefined ? esc(locationSummary.directionAverages[key]) : '-'}`).join(' · ') || '–'}</p>
            ${safeText(location.notes, '') ? `<p class="detail-note"><strong>Site notes:</strong> ${esc(location.notes)}</p>` : ''}
          </div>
        `;
      })
      .join('')}
  `;
}
