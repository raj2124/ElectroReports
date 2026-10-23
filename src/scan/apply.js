// Puts a read sheet into a daily report (Session 6I). Pure: it takes the report (or the starting point of a new day) and
// returns the report-shaped input to hand to normalizeReportInput, plus what it overwrote.
//
//   - A blank cell never overwrites anything. A cell with content writes its value.
//   - Every row / group / location carries `sheetRow`, its place on the printed sheet. A scan overwrites the row that has
//     the same sheetRow, and otherwise inserts a new row in sheetRow order. A row typed by hand has no sheetRow and counts
//     by its position in the table (a tower group counts 4 per position: it is 4 printed rows).
//       flat sheets and phase groups : the printed row number (rowOffset + row)
//       tower groups                 : the printed row number of the group's first foot
//       soil                         : the page index (one location per page)
//   - Pre-printed text (tags, locations) is written together with the row it belongs to, but a row that holds only
//     pre-printed text is not created.

const { SPECS } = require('../sheetPrint/spec');
const { SOIL_SPACING_PRESETS, DEFAULT_SOIL_PRESET } = require('../sheetDefinitions');
const { hasContent, parseDate, normaliseNumber } = require('./cells');

const PHASES = ['R-E', 'Y-E', 'B-E'];
const FEET = ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'];
const clone = (value) => JSON.parse(JSON.stringify(value));
const text = (value) => String(value === undefined || value === null ? '' : value).trim();

// the value a cell writes
function valueOf(def, cell) {
  if (def.kind === 'ticks') return def.multi ? [...cell.selected] : cell.selected[0] || '';
  if (def.kind === 'tick') return text(cell.value) === 'true';
  if (def.date) return parseDate(cell.value) || text(cell.value);
  return def.kind === 'number' ? normaliseNumber(cell.value) : text(cell.value);
}

const isEmptyValue = (value) => value === undefined || value === null || value === '' || value === false || (Array.isArray(value) && !value.length);
const sameValue = (a, b) => JSON.stringify(Array.isArray(a) ? [...a].sort() : text(a)) === JSON.stringify(Array.isArray(b) ? [...b].sort() : text(b));

// Fields of one printed row (or the merged cells of its group): { key: value } for the cells with content.
function patchOf(form, cells, row, { merged = null } = {}) {
  const patch = {};
  let real = false;
  form.cells.filter((def) => def.section === 'row' && def.row === row && (merged === null || Boolean(form.columns.find((col) => col.key === def.key && col.merge)) === merged)).forEach((def) => {
    const cell = cells[def.path];
    if (!hasContent(def, cell)) return;
    patch[def.key] = valueOf(def, cell);
    if (!def.preprinted) real = true;
  });
  return { patch, real };
}

const effective = (entry, index, unit) => (Number.isInteger(entry.sheetRow) ? entry.sheetRow : index * unit);

// Merges `incoming` ([{ sheetRow, merge(existing) -> next, make() -> new }]) into `list`.
function mergeList(list, incoming, unit, onOverwrite) {
  const out = [...list];
  incoming.forEach((item) => {
    const at = out.findIndex((entry, index) => effective(entry, index, unit) === item.sheetRow);
    if (at >= 0) {
      onOverwrite(out[at], item.sheetRow);
      out[at] = item.merge(out[at]);
      return;
    }
    const before = out.findIndex((entry, index) => effective(entry, index, unit) > item.sheetRow);
    out.splice(before < 0 ? out.length : before, 0, item.make());
  });
  return out;
}

// returns { candidate, overwrites: [{ where, key, label, before, after }], previous: [row snapshots], written: n }
function applyScan({ report, form, cells }) {
  const candidate = clone(report);
  const overwrites = [];
  const previous = [];
  const snapshots = new Set();
  const labelOf = (key) => (form.cells.find((def) => def.key === key) || {}).label || key;
  let written = 0;

  const noteOverwrites = (existing, patch, where, snapshotKey) => {
    const diff = Object.keys(patch).filter((key) => !isEmptyValue(existing[key]) && !isEmptyValue(patch[key]) && !sameValue(existing[key], patch[key]));
    diff.forEach((key) => overwrites.push({ where, key, label: labelOf(key), before: existing[key], after: patch[key] }));
    if (diff.length && !snapshots.has(snapshotKey)) {
      snapshots.add(snapshotKey);
      previous.push({ where, row: clone(existing) });
    }
  };

  // ---- header ----
  const header = {};
  const job = {};
  const soilHeader = {};
  form.cells.filter((def) => def.section === 'header').forEach((def) => {
    const cell = cells[def.path];
    if (!hasContent(def, cell)) return;
    if (form.layout === 'soilMatrix') {
      if (!def.job) soilHeader[def.key] = valueOf(def, cell);
    }
    if (def.job) job[def.key] = valueOf(def, cell);
    else if (form.layout !== 'soilMatrix') header[def.key] = valueOf(def, cell);
  });
  if (Object.keys(header).length) {
    const existing = (candidate.sheetHeaders && candidate.sheetHeaders[form.sheetKey]) || {};
    noteOverwrites(existing, header, `${form.code} header`, 'header');
    candidate.sheetHeaders = { ...(candidate.sheetHeaders || {}), [form.sheetKey]: { ...existing, ...header } };
    written += Object.keys(header).length;
  }
  if (Object.keys(job).length) {
    noteOverwrites(candidate.project || {}, job, 'Job details', 'job');
    candidate.project = { ...(candidate.project || {}), ...job };
    written += Object.keys(job).length;
  }

  // ---- table ----
  const base = form.rowOffset;
  if (form.layout === 'soilMatrix') {
    const preset = SOIL_SPACING_PRESETS[DEFAULT_SOIL_PRESET];
    const directionKeys = ['direction1', 'direction2', 'direction3', 'direction4'];
    const location = { ...Object.fromEntries(['name', 'inputUnit', 'probeDepthCm', 'weather'].filter((key) => soilHeader[key] !== undefined).map((key) => [key, soilHeader[key]])) };
    const bearings = Object.fromEntries(directionKeys.filter((key) => form.cells.some((def) => def.section === 'header' && def.key === key) && soilHeader[key] !== undefined).map((key) => [key, soilHeader[key]]));
    const rows = {};
    directionKeys.forEach((key) => { rows[key] = []; });
    let any = Object.keys(location).length > 0 || Object.keys(bearings).length > 0;
    for (let r = 0; r < form.rowCount; r += 1) {
      const { patch } = patchOf(form, cells, r);
      directionKeys.forEach((key) => {
        if (patch[key] !== undefined || (key === 'direction1' && patch.remark)) {
          rows[key].push({ spacing: preset.spacings[r], resistivity: patch[key] !== undefined ? patch[key] : '', remark: key === 'direction1' ? patch.remark || '' : '' });
          any = true;
        }
      });
    }
    if (any) {
      const locations = (candidate.soilResistivity && candidate.soilResistivity.locations) || [];
      const incoming = [{
        sheetRow: base,
        merge: (existing) => {
          noteOverwrites({ ...existing, ...(existing.bearings || {}) }, { ...location, ...bearings }, `${form.code} location "${existing.name || base + 1}"`, `soil${base}`);
          const merged = { ...existing, ...location, bearings: { ...(existing.bearings || {}), ...bearings }, sheetRow: base };
          directionKeys.forEach((key) => {
            const current = Array.isArray(existing[key]) ? [...existing[key]] : [];
            rows[key].forEach((row) => {
              const at = current.findIndex((entry) => text(entry.spacing) === text(row.spacing));
              if (at >= 0) {
                if (!isEmptyValue(current[at].resistivity) && !isEmptyValue(row.resistivity) && !sameValue(current[at].resistivity, row.resistivity)) {
                  overwrites.push({ where: `${form.code} ${text(existing.name)} spacing ${row.spacing} m`, key, label: `${labelOf(key)} @ ${row.spacing} m`, before: current[at].resistivity, after: row.resistivity });
                  if (!snapshots.has(`soil${base}`)) { snapshots.add(`soil${base}`); previous.push({ where: `${form.code} location`, row: clone(existing) }); }
                }
                current[at] = { ...current[at], ...Object.fromEntries(Object.entries(row).filter(([, value]) => !isEmptyValue(value))) };
              } else current.push(row);
            });
            merged[key] = current;
          });
          return merged;
        },
        make: () => ({ sheetRow: base, name: location.name || `Location ${locations.length + 1}`, ...location, bearings, ...rows })
      }];
      candidate.soilResistivity = { ...(candidate.soilResistivity || {}), locations: mergeList(locations, incoming, 1, () => {}) };
      written += 1;
    }
  } else if (form.layout === 'towerGroups') {
    const groups = Array.isArray(candidate[form.sheetKey]) ? candidate[form.sheetKey] : [];
    const incoming = [];
    for (let item = 0; item < form.rowCount / form.group; item += 1) {
      const first = item * form.group;
      const mergedPart = patchOf(form, cells, first, { merged: true });
      const feet = FEET.map((foot, k) => ({ foot, ...patchOf(form, cells, first + k, { merged: false }) }));
      if (!mergedPart.real && !feet.some((entry) => entry.real)) continue;
      const sheetRow = base + first;
      const readingPatch = (entry) => ({ measurementPointLocation: entry.foot, ...entry.patch });
      incoming.push({
        sheetRow,
        merge: (existing) => {
          noteOverwrites(existing, mergedPart.patch, `${form.code} tower ${existing.mainLocationTower || item + 1}`, `g${sheetRow}`);
          const readings = FEET.map((foot, k) => {
            const old = (existing.readings || []).find((reading) => reading.measurementPointLocation === foot) || (existing.readings || [])[k] || { measurementPointLocation: foot };
            noteOverwrites(old, feet[k].patch, `${form.code} tower ${existing.mainLocationTower || item + 1} ${foot}`, `g${sheetRow}`);
            return { ...old, ...readingPatch(feet[k]) };
          });
          return { ...existing, ...mergedPart.patch, readings, sheetRow };
        },
        make: () => ({ sheetRow, ...mergedPart.patch, readings: feet.map(readingPatch) })
      });
      written += 1;
    }
    candidate[form.sheetKey] = mergeList(groups, incoming, form.group, () => {});
  } else {
    const rows = Array.isArray(candidate[form.sheetKey]) ? candidate[form.sheetKey] : [];
    const incoming = [];
    const phased = form.layout === 'phaseGroups';
    for (let r = 0; r < form.rowCount; r += 1) {
      let patch;
      let real;
      if (phased) {
        const firstRow = r - (r % form.group);
        const mergedPart = patchOf(form, cells, firstRow, { merged: true });
        const own = patchOf(form, cells, r, { merged: false });
        let groupReal = mergedPart.real;
        for (let k = 0; k < form.group; k += 1) groupReal = groupReal || patchOf(form, cells, firstRow + k, { merged: false }).real;
        patch = { ...mergedPart.patch, ...own.patch, measuredPoints: PHASES[r % form.group] };
        real = groupReal;
      } else {
        ({ patch, real } = patchOf(form, cells, r));
      }
      if (!real) continue;
      const sheetRow = base + r;
      incoming.push({
        sheetRow,
        merge: (existing) => {
          noteOverwrites(existing, patch, `${form.code} row ${sheetRow + 1}`, `r${sheetRow}`);
          return { ...existing, ...patch, sheetRow };
        },
        make: () => ({ sheetRow, ...patch })
      });
      written += 1;
    }
    candidate[form.sheetKey] = mergeList(rows, incoming, 1, () => {});
  }
  candidate.tests = { ...(candidate.tests || {}), [form.sheetKey]: true };
  return { candidate, overwrites, previous, written };
}

module.exports = { applyScan, PHASES, FEET, valueOf };
