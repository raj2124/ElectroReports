// What is on a printed sheet page, as a list of cells to read (Session 6I). Derived from the print spec
// (src/sheetPrint/spec.js: the columns and header cells exactly as printed) and the sheet definitions
// (src/sheetDefinitions.js: kinds, limits, options), so a scan can only ever contain cells that exist on the
// page that was issued. Serial numbers and fixed columns (spacing, phase, foot) come from the row's position and
// are never read. Pre-printed text (tags, locations from earlier days) comes from the issued record.
//
// A cell has a stable path: "h.<key>" in the header, "r<printed row>.<key>" in the table (a merged column of a
// phase or tower group lives on the first row of its group).
//   kind  text | number | ticks (options; multi = several may be ticked) | tick (a single box)

const { SPECS, soilSpacings, GROUND_CONDITIONS } = require('../sheetPrint/spec');
const { getSheet } = require('../sheetDefinitions');

const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const flatten = (fields) => (fields || []).flatMap((field) => (field.kind === 'group' ? field.fields : [field]));

function optionsOf(labels, values) {
  return labels.map((label, index) => ({ value: values ? values[index] : label, label }));
}

// issued: { sheetKey, rowCount, rowOffset, page, preRows } (an issued_sheets record, or one made from a manual choice).
function buildForm(issued) {
  const spec = SPECS[issued.sheetKey];
  const sheet = getSheet(issued.sheetKey);
  if (!spec || !sheet) throw new Error(`There is no printable sheet "${issued.sheetKey}".`);
  const rowCount = issued.rowCount || spec.baseRows;
  const group = spec.group;
  const rowDefs = Object.fromEntries(flatten(sheet.rowFields).map((field) => [field.key, field]));
  const headerDefs = Object.fromEntries(flatten(sheet.headerFields).map((field) => [field.key, field]));
  const cells = [];
  const header = (key, label, kind, extra = {}) => cells.push({ path: `h.${key}`, section: 'header', row: null, key, label, kind, ...extra });

  // the instrument row and the footer: on every sheet
  // (the soil sheet keeps no instrument row in a report: its header lives on each location)
  if (issued.sheetKey !== 'soilResistivity') {
    header('instrument', 'Instrument used', 'ticks', { options: optionsOf(spec.instruments.map((entry) => entry.label), spec.instruments.map((entry) => entry.id)) });
    header('serialNo', 'Instrument Sr. No.', 'text');
    header('calibrationDue', 'Calibration due (dd/mm/yyyy)', 'text', { date: true });
    header('timeStart', 'Time start', 'text');
    header('timeFinish', 'Time finish', 'text');
  }
  header('clientRepresentative', 'Client representative', 'text', { job: true });
  header('groundCondition', 'Ground condition', 'ticks', { job: true, options: optionsOf(GROUND_CONDITIONS) });
  // the sheet's own header cells
  spec.headerRows.flat().forEach((cell) => {
    if (cell.subKeys) {
      cell.subKeys.forEach((sub) => header(sub.key, `${cell.label} ${sub.label}`, 'number', { def: headerDefs[sub.key], unit: sub.unit }));
      return;
    }
    const top = (sheet.headerFields || []).find((field) => field.key === cell.key);
    if (top && top.kind === 'group') {
      top.fields.forEach((sub) => header(sub.key, `${cell.label} · ${sub.label}`, sub.kind === 'number' ? 'number' : 'text', { def: sub }));
      return;
    }
    const def = headerDefs[cell.key];
    if (!def || def.kind === 'fixed' || def.readOnly) return;
    if (cell.ticks && def.kind === 'tick') header(cell.key, cell.label, 'tick', { def });
    else if (cell.ticks) header(cell.key, cell.label, 'ticks', { def, options: optionsOf(cell.ticks.map((label) => String(label).replace(/ ?_+$/, '')), cell.values), multi: Boolean(def.multi) });
    else header(cell.key, cell.label, def.kind === 'number' ? 'number' : 'text', { def });
  });

  const spacings = issued.sheetKey === 'soilResistivity' ? soilSpacings() : null;
  const preRows = Array.isArray(issued.preRows) ? issued.preRows : [];
  for (let r = 0; r < rowCount; r += 1) {
    const first = r % group === 0;
    const item = Math.floor(r / group);
    spec.columns.forEach((col) => {
      if (col.k === 'sr' || col.k === 'fixed') return;
      if (col.merge && !first) return;
      const def = rowDefs[col.key];
      const base = { path: `r${r}.${col.key}`, section: 'row', row: r, item, key: col.key, label: col.label, def };
      if (col.k === 'ticks') {
        cells.push({ ...base, kind: 'ticks', options: optionsOf(col.options, col.values), multi: Boolean(def && def.multi) });
        return;
      }
      if (col.k === 'text') {
        const pre = (preRows[item] || {})[col.key];
        if (text(pre) && (group === 1 || col.merge)) {
          cells.push({ ...base, kind: 'text', preprinted: true, value: text(pre) });
          return;
        }
        cells.push({ ...base, kind: 'text' });
        return;
      }
      cells.push({ ...base, kind: 'number' });
    });
  }
  return { sheetKey: issued.sheetKey, sheetId: issued.sheetId || '', code: spec.code, title: spec.title, layout: sheet.layout, group, rowCount, rowOffset: issued.rowOffset || 0, page: issued.page || 1, landscape: spec.landscape, spacings, columns: spec.columns.map((col) => ({ key: col.key, label: col.label, k: col.k, merge: Boolean(col.merge), values: col.values || null })), cells };
}

module.exports = { buildForm };
