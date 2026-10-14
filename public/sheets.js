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
    inputUnit: 'ohm-m',
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

function renderTextControl(field, value, attrs, cssClass = 'table-input') {
  const isNumber = field.kind === 'number';
  return `<input class="${cssClass}" type="${field.kind === 'date' ? 'date' : 'text'}" ${isNumber ? 'inputmode="decimal"' : ''} autocomplete="off" aria-label="${esc(field.label)}" ${field.unit ? `placeholder="${esc(field.unit)}"` : ''} value="${esc(value)}" ${dataAttrs(attrs)} />`;
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
  return `<p class="sheet-ref"><strong>${esc(sheet.code === 'legacy' ? 'Retired sheet' : `${sheet.code} · ${sheet.docNo}`)}</strong> · ${esc(sheet.reference)}</p>`;
}

function renderNotes(sheet) {
  if (!sheet.notes || !sheet.notes.length) {
    return '';
  }
  return `<details class="sheet-notes"><summary>How to record this sheet</summary><ul>${sheet.notes.map((note) => `<li>${esc(note)}</li>`).join('')}</ul></details>`;
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

function renderStatusCell(sheetId, row) {
  const assessment = rowAssessment(sheetId, row);
  const limit = assessment.limit === null || assessment.limit === undefined ? '' : `<span class="status-limit">Limit ${Number(assessment.limit).toFixed(2)}</span>`;
  return `
    <td class="sheet-status-cell" data-label="Status">
      ${pill(assessment.status.tone, assessment.status.label)}
      ${assessment.priority ? `<span class="soft-chip">${esc(assessment.priority)}</span>` : ''}
      ${limit}
      ${flagChips(assessment.flags)}
      <div class="auto-eval-text">${esc(assessment.comment)}</div>
    </td>
  `;
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

function renderRowCell(sheet, field, row, index, rowNumber, extra = {}) {
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
  return `<td data-label="${esc(field.label)}" class="sheet-cell sheet-cell-${field.kind}">${content}</td>`;
}

// Column widths come from the printed sheet (field.width = the fraction in
// make_sheets_v3.py). The printed columns share what is left after the status and action
// columns. (Plain percentages: a max()/calc() width on <col> is ignored by browsers.)
const TRAILING_SHARE = { status: 13, action: 4.5, tower: 7, note: 4.5 };

function renderColgroup(columns, trailing = ['status', 'action']) {
  const trailingTotal = trailing.reduce((sum, name) => sum + TRAILING_SHARE[name], 0);
  const printedTotal = columns.reduce((sum, field) => sum + (field.width || 8), 0);
  const unit = (100 - trailingTotal) / printedTotal;
  const printed = columns.map((field) => `<col class="col-${esc(field.kind)}" style="width:${((field.width || 8) * unit).toFixed(3)}%">`);
  const extra = trailing.map((name) => `<col class="col-${name}" style="width:${TRAILING_SHARE[name]}%">`);
  return `<colgroup>${printed.join('')}${extra.join('')}</colgroup>`;
}

function renderSheetHead(columns, withStatus) {
  return `<tr>${columns.map((field) => `<th>${esc(field.label)}</th>`).join('')}${withStatus ? '<th>Assessment</th>' : ''}<th>Action</th></tr>`;
}

function renderRowsLayout(sheet) {
  const rows = state.draft[sheet.id] || [];
  const columns = visibleRowFields(sheet, rows);
  return tableSurface(
    `${sheet.code === 'legacy' ? sheet.title : `${sheet.code} ${sheet.title}`}`,
    'Add rows as needed for the executed field scope.',
    renderSheetHead(columns, true),
    rows
      .map((row, index) => `<tr class="sheet-row">${columns.map((field) => renderRowCell(sheet, field, row, index, index + 1)).join('')}${renderStatusCell(sheet.id, row)}${actionButtonsCell(sheet.id, index, row)}</tr>`)
      .join(''),
    `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${esc(sheet.id)}">Add row</button></div>`,
    { className: 'sheet-table', colgroup: renderColgroup(columns) }
  );
}

function renderPhaseGroupsLayout(sheet) {
  const rows = state.draft[sheet.id] || [];
  const size = sheet.groupSize || 3;
  const columns = visibleRowFields(sheet, rows);
  const body = rows
    .map((row, index) => {
      const offset = index % size;
      const groupNumber = Math.floor(index / size) + 1;
      const first = rows[index - offset] || row;
      const cells = columns
        .map((field) => {
          if (field.merge) {
            if (offset !== 0) {
              return '';
            }
            const merged = field.kind === 'serial' ? `<span class="table-input-static">${groupNumber}</span>` : renderControl(field, first[field.key], { section: sheet.id, index, field: field.key, groupSync: true }, `${sheet.id}-${index}-${field.key}`, 'table-input', true);
            return `<td rowspan="${size}" data-label="${esc(field.label)}" class="phase-group-cell sheet-cell sheet-cell-${field.kind}">${merged}</td>`;
          }
          return renderRowCell(sheet, field, row, index, index + 1);
        })
        .join('');
      return `<tr class="sheet-row ${offset === 0 ? 'tower-group-row-start' : ''}">${cells}${renderStatusCell(sheet.id, row)}${actionButtonsCell(sheet.id, index, row)}</tr>`;
    })
    .join('');
  return tableSurface(`${sheet.code} ${sheet.title}`, 'Each feeder is three readings, one per phase; the shared cells are entered once.', renderSheetHead(columns, true), body, `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${esc(sheet.id)}">Add feeder</button></div>`, { className: 'sheet-table', colgroup: renderColgroup(columns) });
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
            .map((field) => {
              if (field.merge) {
                if (!first) {
                  return '';
                }
                const content =
                  field.kind === 'serial'
                    ? `<span class="table-input-static">${groupIndex + 1}</span>`
                    : renderControl(field, group[field.key], { section: sheet.id, index: groupIndex, field: field.key }, `${sheet.id}-${groupIndex}-${field.key}`, 'table-input', true);
                return `<td rowspan="${feet.length}" data-label="${esc(field.label)}" class="tower-group-cell sheet-cell">${content}</td>`;
              }
              if (field.kind === 'fixed') {
                const position = feetField.values.indexOf(reading.measurementPointLocation);
                return `<td data-label="${esc(field.label)}" class="tower-foot-label"><span class="table-input-static">${esc(position >= 0 ? feetField.display[position] : reading.measurementPointLocation)}</span></td>`;
              }
              return renderRowCell(sheet, field, reading, groupIndex, groupIndex + 1, { readingIndex });
            })
            .join('');
          const status = first
            ? `<td rowspan="${feet.length}" class="sheet-status-cell tower-group-cell" data-label="Assessment">
                 ${pill(assessment.status.tone, assessment.status.label)}
                 ${flagChips(assessment.flags)}
                 <div class="status-limit">Zt ${assessment.totalImpedanceZt === null || assessment.totalImpedanceZt === undefined ? '-' : esc(assessment.totalImpedanceZt)} Ω · Σ I ${assessment.totalCurrentItotal === null || assessment.totalCurrentItotal === undefined ? '-' : esc(assessment.totalCurrentItotal)} mA · Zsat ${assessment.zsat === null || assessment.zsat === undefined ? esc(sheet.zsat) : esc(assessment.zsat)} Ω</div>
                 <div class="auto-eval-text">${esc(assessment.remark)}</div>
               </td>`
            : '';
          const action = first
            ? `<td rowspan="${feet.length}" class="table-action-cell tower-group-action-cell"><button class="button button-ghost tower-group-remove" type="button" data-action="remove-tower-group" data-group-index="${groupIndex}">Remove tower</button></td>`
            : '';
          return `<tr class="${first ? 'tower-group-row-start' : ''} ${readingIndex === feet.length - 1 ? 'tower-group-row-end' : ''}">${cells}${status}${action}<td class="table-action-cell">${towerObservationButton(groupIndex, readingIndex, reading)}</td></tr>`;
        })
        .join('');
    })
    .join('');
  return tableSurface(
    `${sheet.code} ${sheet.title}`,
    'Each tower has four fixed feet (F-1 to F-4); the tower totals are calculated.',
    `<tr>${columns.map((field) => `<th>${esc(field.label)}</th>`).join('')}<th>Assessment</th><th>Tower</th><th>Photo / note</th></tr>`,
    body,
    `<div class="mini-surface-foot"><button class="button button-secondary" data-action="add-row" data-section="${esc(sheet.id)}">Add tower</button></div>`,
    { className: 'sheet-table tower-sheet-table', colgroup: renderColgroup(columns, ['status', 'tower', 'note']) }
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
