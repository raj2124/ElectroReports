// Who gets credit for a unit of site work (Session 6G). A "unit" is what one Site Work Log entry in
// Zoho People counts: a row of a measurement table, a soil location, or a tower footing group.
//
//   enteredBy      the FIRST person who entered the unit. Set once, never overwritten.
//   lastEditedBy   whoever last changed its content. For audit only; it never moves the credit.
//
// The stamps are put on by the server (never read from the browser: the report normaliser rebuilds
// every row and drops them), by comparing the saved report with the one it replaces. A unit that was
// already in the stored report without a stamp (entered before 6G) stays unstamped: the log credits
// the report's creator for it instead of guessing.

const { TEST_LIBRARY } = require('./reportModel');

const STAMPS = ['enteredBy', 'lastEditedBy'];

// Where the units of each test live and which id identifies one. Anything else is a plain row array.
const UNIT_SOURCES = {
  soilResistivity: { list: (report) => (report.soilResistivity && report.soilResistivity.locations) || [], idKey: 'locationId', label: 'location' },
  towerFootingResistance: { list: (report) => report.towerFootingResistance || [], idKey: 'groupId', label: 'tower' }
};

function sourceFor(testId) {
  return UNIT_SOURCES[testId] || { list: (report) => report[testId] || [], idKey: 'rowId', label: 'row' };
}

const withoutStamps = (unit) => {
  const copy = { ...unit };
  STAMPS.forEach((key) => delete copy[key]);
  return copy;
};

// Key order must not matter: the same content is "unchanged" however it was serialised.
function stable(value) {
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`;
  if (value && typeof value === 'object') return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${stable(value[key])}`).join(',')}}`;
  return JSON.stringify(value === undefined ? null : value);
}

// [{ testId, testLabel, unitId, unit }] for every unit of every test the report carries.
function listUnits(report) {
  const units = [];
  TEST_LIBRARY.forEach((test) => {
    const { list, idKey } = sourceFor(test.id);
    (Array.isArray(list(report)) ? list(report) : []).forEach((unit) => {
      if (unit && unit[idKey]) units.push({ testId: test.id, testLabel: test.label, unitId: String(unit[idKey]), unit });
    });
  });
  return units;
}

const stampOf = (actor, at) => ({ userId: actor.userId || actor.id || '', name: actor.name || '', email: String(actor.email || '').toLowerCase(), at });

// `next` is the normalised report about to be saved, `previous` the stored one (null for a new
// report). Returns a copy of `next` with the stamps put on.
function stampReport(next, previous, actor, at = new Date().toISOString()) {
  const before = new Map(listUnits(previous || {}).map((entry) => [`${entry.testId}/${entry.unitId}`, entry.unit]));
  const stamp = (testId, idKey) => (unit) => {
    const old = before.get(`${testId}/${unit[idKey]}`);
    const out = withoutStamps(unit);
    if (!old) {
      out.enteredBy = stampOf(actor, at);
      return out;
    }
    if (old.enteredBy) out.enteredBy = old.enteredBy;
    if (stable(withoutStamps(old)) !== stable(withoutStamps(out))) {
      out.lastEditedBy = stampOf(actor, at);
    } else if (old.lastEditedBy) {
      out.lastEditedBy = old.lastEditedBy;
    }
    return out;
  };
  const result = { ...next };
  TEST_LIBRARY.forEach((test) => {
    const { idKey } = sourceFor(test.id);
    const apply = stamp(test.id, idKey);
    if (test.id === 'soilResistivity') {
      if (result.soilResistivity && Array.isArray(result.soilResistivity.locations)) {
        result.soilResistivity = { ...result.soilResistivity, locations: result.soilResistivity.locations.map(apply) };
      }
    } else if (Array.isArray(result[test.id])) {
      result[test.id] = result[test.id].map(apply);
    }
  });
  return result;
}

module.exports = { STAMPS, listUnits, stampReport, stable };
