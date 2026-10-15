// What a printed sheet looks like, derived from the sheet definitions (src/sheetDefinitions.js,
// the source of truth for fields and order) plus the few print-only numbers that live in
// reference/generators/make_sheets_v3.py (orientation, row height, row count, legend, fill).
// A test compares the two.
const {
  SHEETS,
  INSTRUMENTS,
  CONDITION_CODE_LEGEND,
  SOIL_SPACING_PRESETS,
  DEFAULT_SOIL_PRESET,
  GROUND_CONDITIONS,
  optionValue
} = require('../sheetDefinitions');

// land: landscape, n: base row count, rh: nominal row height, min: min_rows, legend: condition
// code legend printed, fill: grow the table to fill the page.
const PRINT = Object.freeze({
  soilResistivity: { land: false, n: 10, rh: 30, min: 0, legend: false, fill: false },
  electrodeResistance: { land: true, n: 10, rh: 27, min: 10, legend: true, fill: true },
  continuityTest: { land: false, n: 19, rh: 26, min: 19, legend: true, fill: true },
  loopImpedanceTest: { land: true, n: 12, rh: 20, min: 12, legend: true, fill: true },
  prospectiveFaultCurrent: { land: true, n: 12, rh: 20, min: 12, legend: true, fill: true },
  riserIntegrityTest: { land: false, n: 19, rh: 26, min: 19, legend: true, fill: true },
  groundResistance: { land: false, n: 18, rh: 27, min: 18, legend: true, fill: true },
  stepTouchVoltage: { land: true, n: 12, rh: 24, min: 12, legend: false, fill: true },
  towerFootingResistance: { land: false, n: 20, rh: 21, min: 20, legend: true, fill: true }
});

// Text columns pre-printed from earlier reports of the same project (tags and locations only;
// readings, instrument, serial and ticks always stay blank). `identity` is how two rows are
// recognised as the same physical point across days (the same rule as the multi-day stitching in
// reportCollectionModel.js): 'first' = the first of these that is filled in (tag, else location),
// 'all' = all of them together.
const PREPRINT = Object.freeze({
  electrodeResistance: { keys: ['tag', 'location'], identity: 'first' },
  continuityTest: { keys: ['tag', 'mainLocation', 'referencePoint', 'measurementPoint'], identity: 'all', identityKeys: ['mainLocation', 'measurementPoint'] },
  loopImpedanceTest: { keys: ['location', 'feederTag'], identity: 'all' },
  prospectiveFaultCurrent: { keys: ['location', 'feederTag'], identity: 'all' },
  riserIntegrityTest: { keys: ['mainLocation', 'measurementPoint'], identity: 'all' },
  groundResistance: { keys: ['tag', 'location'], identity: 'first' },
  stepTouchVoltage: { keys: ['location'], identity: 'all' },
  towerFootingResistance: { keys: ['mainLocationTower'], identity: 'all' }
});

const HEADER_FRACTIONS = [0.42, 0.33, 0.25];

const CODE_LEGEND = CONDITION_CODE_LEGEND.map(({ code, text }) => `${code} ${text}`).join(' · ');

const printable = (field) => !field.extra && !field.legacy;
const optionLabel = (option) => (option && typeof option === 'object' ? option.label : option);

// One printed column.
function columnFor(field) {
  const base = { key: field.key, label: field.label, w: field.width, merge: Boolean(field.merge) };
  switch (field.kind) {
    case 'serial': return { ...base, k: 'sr' };
    case 'fixed': return { ...base, k: 'fixed', values: field.display || field.values || [] };
    case 'ticks': return { ...base, k: 'ticks', options: (field.options || []).map(optionLabel), values: (field.options || []).map(optionValue) };
    case 'text': return { ...base, k: 'text' };
    default: return { ...base, k: 'digits' };
  }
}

// What a header cell prints in its value area.
function headerCell(field) {
  if (field.kind === 'ticks') {
    return { key: field.key, ticks: (field.options || []).map(optionLabel).map((label) => (label === 'Other' ? 'Other ____' : label)), values: (field.options || []).map(optionValue) };
  }
  if (field.kind === 'tick') {
    return { key: field.key, ticks: [field.tickLabel || field.label], values: [true] };
  }
  if (field.kind === 'group' && field.fields.every((sub) => sub.unit === '°')) {
    return { key: field.key, text: field.fields.map((sub) => `${sub.label} _____°`).join('   '), subKeys: field.fields.map((sub) => ({ key: sub.key, label: sub.label, unit: sub.unit })) };
  }
  return { key: field.key, text: '' };
}

function buildSpec(sheet) {
  const print = PRINT[sheet.id];
  const cells = sheet.headerFields.filter(printable).map((field) => ({ label: field.label, ...headerCell(field) }));
  const headerRows = [];
  for (let index = 0; index < cells.length; index += 3) {
    headerRows.push(cells.slice(index, index + 3).map((cell, position) => ({ ...cell, frac: HEADER_FRACTIONS[position] })));
  }
  return {
    id: sheet.id,
    code: sheet.code,
    docNo: sheet.docNo,
    title: `${sheet.code}  ${sheet.title} – Measurement Sheet`,
    reference: sheet.reference,
    landscape: print.land,
    baseRows: print.n,
    rowHeight: print.rh,
    minRows: print.min,
    group: sheet.groupSize || 1,
    legend: print.legend,
    fill: print.fill,
    instruments: sheet.instruments.map((id) => INSTRUMENTS.find((instrument) => instrument.id === id)),
    columns: sheet.rowFields.filter(printable).map(columnFor),
    headerRows,
    notes: sheet.notes,
    preprint: PREPRINT[sheet.id] || { keys: [], identity: 'all' }
  };
}

const SPECS = Object.fromEntries(SHEETS.filter((sheet) => !sheet.retired).map((sheet) => [sheet.id, buildSpec(sheet)]));
const PRINTABLE_SHEET_IDS = Object.keys(SPECS);

// The soil sheet prints the default probe-spacing set (the approved layout: 10 spacings x 4
// directions). Other spacing sets use the entry screen, or extra sheets.
function soilSpacings() {
  return SOIL_SPACING_PRESETS[DEFAULT_SOIL_PRESET].spacings.map((value) => Number(value).toFixed(2));
}

module.exports = { SPECS, PRINTABLE_SHEET_IDS, PRINT, PREPRINT, CODE_LEGEND, GROUND_CONDITIONS, soilSpacings };
