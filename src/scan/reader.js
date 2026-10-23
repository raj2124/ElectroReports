// Reads one page: the primary model, then (only if something is doubtful) the escalation model, and the deterministic
// checks (Session 6I). The result is { cells: { [path]: cell }, models, escalated }.
//
// Escalation: when any cell is low-confidence, illegible, not a number, multi-ticked or an unclear unit, the whole page is
// read again by the stronger model. For a cell that was doubtful the stronger model's reading is taken if it is sure of it;
// wherever the two reads disagree the cell is flagged READS_DIFFER (and keeps the other reading as `alt`), whatever
// either model said about its own confidence.

const { cleanCell, flagsFor, flag, sameContent } = require('./cells');

const DOUBTFUL = new Set(['LOW_CONFIDENCE', 'ILLEGIBLE', 'NON_NUMERIC', 'MULTI_TICK', 'UNIT_UNCLEAR', 'INVALID_OPTION']);

function toCells(form, raw, threshold, source) {
  const byPath = new Map((raw && Array.isArray(raw.cells) ? raw.cells : []).filter((entry) => entry && typeof entry.path === 'string').map((entry) => [entry.path, entry]));
  const cells = {};
  form.cells.forEach((def) => {
    if (def.preprinted) {
      cells[def.path] = { value: def.value, selected: [], confidence: 1, illegible: false, flags: [], source: 'preprinted' };
      return;
    }
    const cell = cleanCell(def, byPath.get(def.path) || { confidence: 0.99 }, { source });
    cell.flags = [...cell.flags, ...flagsFor(def, cell, threshold)];
    cells[def.path] = cell;
  });
  return cells;
}

const doubtful = (cell) => cell.flags.some((entry) => DOUBTFUL.has(entry.code));

// provider: { read({ png, form, model }) }; models: { primary, escalation }; threshold: confidence below which a cell is flagged.
async function readPage({ form, png, provider, models, threshold = 0.8 }) {
  const first = toCells(form, await provider.read({ png, form, model: models.primary }), threshold, 'primary');
  const needsSecond = models.escalation && models.escalation !== models.primary && form.cells.some((def) => !def.preprinted && doubtful(first[def.path]));
  if (!needsSecond) return { cells: first, models: { primary: models.primary, escalation: null }, escalated: false };

  const second = toCells(form, await provider.read({ png, form, model: models.escalation }), threshold, 'escalated');
  const cells = {};
  form.cells.forEach((def) => {
    const a = first[def.path];
    const b = second[def.path];
    if (def.preprinted) {
      cells[def.path] = a;
      return;
    }
    const differ = !sameContent(def, a, b);
    const useSecond = doubtful(a) && !doubtful(b) && !b.illegible;
    const chosen = useSecond ? b : a;
    const other = useSecond ? a : b;
    const flags = [...chosen.flags];
    if (differ) flags.push(flag('READS_DIFFER'));
    cells[def.path] = { ...chosen, flags, ...(differ ? { alt: def.kind === 'ticks' ? { selected: other.selected } : { value: other.value } } : {}) };
  });
  return { cells, models: { primary: models.primary, escalation: models.escalation }, escalated: true };
}

// After the engineer edits cells: the same checks again; an edited cell no longer counts as "not sure".
function applyEdits(form, cells, edits, threshold = 0.8) {
  const next = { ...cells };
  form.cells.forEach((def) => {
    const edit = edits && edits[def.path];
    if (!edit || def.preprinted) return;
    const cell = cleanCell(def, { value: edit.value, selected: edit.selected, confidence: 1, illegible: false }, { source: 'edited' });
    cell.flags = [...cell.flags, ...flagsFor(def, cell, threshold)];
    next[def.path] = { ...cell, edited: true, was: (cells[def.path] || {}).edited ? cells[def.path].was : { value: (cells[def.path] || {}).value, selected: (cells[def.path] || {}).selected } };
  });
  return next;
}

module.exports = { readPage, applyEdits, toCells };
