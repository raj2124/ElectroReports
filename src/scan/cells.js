// A read cell and the reasons it needs the engineer's eye (Session 6I).
//   cell = { value, selected, confidence, illegible, flags: [{ code, text }], source, alt? }
// `value` is a string for text / number / tick cells ('' = blank); `selected` is an array of option values for
// ticks cells. The model transcribes; every check here is deterministic and only ever ADDS a flag.

const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const NUMBER = /^[-+]?(\d+([.,]\d*)?|[.,]\d+)$/;

const FLAG_TEXT = {
  LOW_CONFIDENCE: 'The reader was not sure of this cell.',
  ILLEGIBLE: 'The reader could not make this out.',
  NON_NUMERIC: 'Not a number.',
  OUT_OF_RANGE: 'Outside the allowed range for this field.',
  MULTI_TICK: 'More than one box is ticked where only one is allowed.',
  UNIT_UNCLEAR: 'The reading unit (R or ρ) is not clearly ticked.',
  INVALID_OPTION: 'A tick was read that is not one of this cell\'s options.',
  BAD_DATE: 'Not a date (dd/mm/yyyy).',
  READS_DIFFER: 'Two reads of the page disagree on this cell.'
};
const flag = (code, extra = '') => ({ code, text: `${FLAG_TEXT[code]}${extra ? ` ${extra}` : ''}` });

// "0,52" -> "0.52" when it is a number; anything else is left exactly as read.
function normaliseNumber(value) {
  const raw = text(value).replace(/\s+/g, '');
  return NUMBER.test(raw) ? raw.replace(',', '.') : text(value);
}

// "5/11/2026", "05-11-26", "2026-11-05" -> "2026-11-05"; '' when it is not a date.
function parseDate(value) {
  const raw = text(value);
  let match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(raw);
  if (!match) {
    match = /^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})$/.exec(raw);
    if (match) match = [null, match[3].length === 2 ? `20${match[3]}` : match[3], match[2].padStart(2, '0'), match[1].padStart(2, '0')];
  }
  if (!match) return '';
  const iso = `${match[1]}-${match[2]}-${match[3]}`;
  const date = new Date(`${iso}T00:00:00Z`);
  return !Number.isNaN(date.getTime()) && date.toISOString().slice(0, 10) === iso ? iso : '';
}

// One cell as the provider gave it, cleaned to the form's cell kind. Never throws.
function cleanCell(cellDef, raw, { source = 'primary' } = {}) {
  const input = raw && typeof raw === 'object' ? raw : {};
  const confidence = Number.isFinite(Number(input.confidence)) ? Math.min(1, Math.max(0, Number(input.confidence))) : 0;
  const illegible = input.illegible === true;
  const cell = { value: '', selected: [], confidence, illegible, flags: [], source };
  if (cellDef.kind === 'ticks') {
    const allowed = new Set(cellDef.options.map((option) => String(option.value)));
    const given = (Array.isArray(input.selected) ? input.selected : []).map((entry) => String(entry));
    cell.selected = [...new Set(given.filter((entry) => allowed.has(entry)))];
    if (given.length !== cell.selected.length) cell.flags.push(flag('INVALID_OPTION'));
    cell.value = '';
  } else if (cellDef.kind === 'tick') {
    cell.value = input.selected && input.selected.length ? 'true' : input.value === true || text(input.value).toLowerCase() === 'true' ? 'true' : '';
  } else {
    cell.value = illegible ? text(input.value) : cellDef.kind === 'number' ? normaliseNumber(input.value) : text(input.value);
  }
  return cell;
}

// The deterministic checks for a cell (also run again after the engineer edits it).
function flagsFor(cellDef, cell, threshold) {
  const flags = [];
  const blank = cellDef.kind === 'ticks' ? !cell.selected.length : !text(cell.value);
  if (cell.illegible) flags.push(flag('ILLEGIBLE'));
  else if (cell.source !== 'edited' && cell.confidence < threshold) flags.push(flag('LOW_CONFIDENCE', `(${Math.round(cell.confidence * 100)}%)`));
  if (cellDef.kind === 'number' && !blank) {
    if (!NUMBER.test(text(cell.value).replace(/\s+/g, ''))) flags.push(flag('NON_NUMERIC'));
    else {
      const number = Number(normaliseNumber(cell.value));
      const def = cellDef.def || {};
      const below = Number.isFinite(def.min) && (def.exclusiveMin ? number <= def.min : number < def.min);
      const above = Number.isFinite(def.max) && number > def.max;
      if (below || above) flags.push(flag('OUT_OF_RANGE', `(${[Number.isFinite(def.min) ? `${def.exclusiveMin ? '>' : '≥'} ${def.min}` : '', Number.isFinite(def.max) ? `≤ ${def.max}` : ''].filter(Boolean).join(', ')})`));
    }
  }
  if (cellDef.kind === 'ticks' && !cellDef.multi && cell.selected.length > 1) flags.push(flag('MULTI_TICK'));
  if (cellDef.key === 'inputUnit' && cell.selected.length !== 1) flags.push(flag('UNIT_UNCLEAR'));
  if (cellDef.date && !blank && !parseDate(cell.value)) flags.push(flag('BAD_DATE'));
  return flags;
}

// A cell holds something (a blank cell is not data, and never overwrites).
function hasContent(cellDef, cell) {
  if (!cell) return false;
  return cellDef.kind === 'ticks' ? cell.selected.length > 0 : Boolean(text(cell.value));
}

const sameContent = (cellDef, a, b) => (cellDef.kind === 'ticks' ? [...a.selected].sort().join('|') === [...b.selected].sort().join('|') : text(a.value) === text(b.value));

module.exports = { FLAG_TEXT, flag, cleanCell, flagsFor, hasContent, sameContent, normaliseNumber, parseDate };
