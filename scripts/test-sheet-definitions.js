// The app's entry screens must use the same fields and order as the approved printed
// sheets. This compares src/sheetDefinitions.js with the sheet generator
// (reference/generators/make_sheets_v3.py) and with the register's instrument list.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { SHEETS, getSheet, instrumentsForSheet, flattenFields, publicDefinitions, SOIL_SPACING_PRESETS } = require('../src/sheetDefinitions');
const { loadRegister } = require('../src/acceptance');

let passed = 0;
function check(name, fn) {
  fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const source = fs.readFileSync(path.join(__dirname, '..', 'reference', 'generators', 'make_sheets_v3.py'), 'utf8');
const INSTR_IDS = { 3290: 'mi3290', 3152: 'mi3152', 4105: 'kyoritsu4105a', 4118: 'kyoritsu4118a', 3295: 'mi3295' };
const unescape = (text) => text.replace(/\\n/g, ' ').replace(/\s+/g, ' ').trim();

// One printed sheet per make(...) call.
const printed = {};
source.split(/\nmake\(c, /).slice(1).forEach((chunk) => {
  const docNo = /"ET-D03-FRM-(\d{3})"/.exec(chunk)[1];
  const title = /"(MS-0\d  [^"]+)"/.exec(chunk)[1];
  const instruments = [...(/\[("\d+"(?:,\s*"\d+")*)\]/.exec(chunk)[1].matchAll(/"(\d+)"/g))].map((m) => INSTR_IDS[m[1]]);
  const colsBlock = chunk.slice(chunk.indexOf('['), chunk.indexOf('extra=') > 0 ? chunk.indexOf('notes=') : chunk.length);
  const columns = [...colsBlock.matchAll(/col\("((?:[^"\\]|\\.)*)"/g)].map((m) => unescape(m[1]));
  const extraStart = chunk.indexOf('extra=');
  const extraBlock = extraStart > 0 ? chunk.slice(extraStart, chunk.indexOf('fig=', extraStart) > 0 ? chunk.indexOf('fig=', extraStart) : chunk.length) : '';
  const headers = [...extraBlock.matchAll(/\("([^"]+)",\s*"[^"]*",\s*\.\d+\)/g)].map((m) => unescape(m[1]));
  printed[`ET-D03-FRM-${docNo}`] = { title, instruments, columns, headers };
});

const printedRowLabels = (sheet) => sheet.rowFields.filter((field) => !field.extra && !field.legacy).map((field) => field.label);
const printedHeaderLabels = (sheet) => sheet.headerFields.filter((field) => !field.extra).map((field) => field.label);

check('the generator defines the nine approved sheets', () => {
  assert.deepStrictEqual(Object.keys(printed).sort(), Array.from({ length: 9 }, (_, i) => `ET-D03-FRM-00${i + 1}`));
});

check('every sheet has the printed document number, code and title', () => {
  SHEETS.filter((sheet) => !sheet.retired).forEach((sheet) => {
    const sheetPrint = printed[sheet.docNo];
    assert.ok(sheetPrint, `${sheet.id}: ${sheet.docNo} is not a printed sheet`);
    assert.ok(sheetPrint.title.startsWith(sheet.code), `${sheet.id}: code ${sheet.code} vs "${sheetPrint.title}"`);
    assert.ok(sheetPrint.title.includes(sheet.title), `${sheet.id}: title "${sheet.title}" vs "${sheetPrint.title}"`);
  });
});

check('row columns match the printed columns, same labels, same order', () => {
  SHEETS.filter((sheet) => !sheet.retired).forEach((sheet) => {
    assert.deepStrictEqual(printedRowLabels(sheet), printed[sheet.docNo].columns, `${sheet.code} columns`);
  });
});

check('header cells match the printed header cells, same labels, same order', () => {
  SHEETS.filter((sheet) => !sheet.retired).forEach((sheet) => {
    assert.deepStrictEqual(printedHeaderLabels(sheet), printed[sheet.docNo].headers, `${sheet.code} header`);
  });
});

check('instruments per sheet match the printed ticks and the register list', () => {
  const register = loadRegister().instruments;
  SHEETS.filter((sheet) => !sheet.retired).forEach((sheet) => {
    assert.deepStrictEqual([...sheet.instruments].sort(), [...printed[sheet.docNo].instruments].sort(), `${sheet.code} printed ticks`);
    assert.deepStrictEqual(instrumentsForSheet(sheet.id).map((instrument) => instrument.id).sort(), [...sheet.instruments].sort(), `${sheet.code} register`);
    assert.ok(register[sheet.registerKey], `${sheet.code}: no register instruments for ${sheet.registerKey}`);
  });
});

check('condition-code ticks per sheet are the printed ones', () => {
  const codesFor = (id) => getSheet(id).rowFields.find((field) => field.key === 'conditionCodes').options;
  assert.deepStrictEqual(codesFor('electrodeResistance'), ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8']);
  ['continuityTest', 'riserIntegrityTest', 'groundResistance'].forEach((id) => assert.deepStrictEqual(codesFor(id), ['C1', 'C2', 'C6', 'C7'], id));
  assert.deepStrictEqual(codesFor('towerFootingResistance'), ['C1', 'C2', 'C5', 'C6']);
  ['loopImpedanceTest', 'prospectiveFaultCurrent', 'stepTouchVoltage'].forEach((id) => {
    assert.ok(!getSheet(id).rowFields.some((field) => field.key === 'conditionCodes'), `${id} has no code column`);
  });
});

check('printed option lists match (types, materials, devices, trip curves, methods, surfaces)', () => {
  const options = (sheetId, key) => getSheet(sheetId).rowFields.find((field) => field.key === key).options.map((o) => (typeof o === 'object' ? o.label : o));
  assert.deepStrictEqual(options('electrodeResistance', 'electrodeType'), ['Plate', 'Pipe', 'Rod', 'Strip']);
  assert.deepStrictEqual(options('electrodeResistance', 'materialType'), ['Cu', 'Cu-bonded', 'GI', 'MS', 'CI', 'SS']);
  assert.deepStrictEqual(options('electrodeResistance', 'method'), ['61.8%', 'Clamp', 'Slope']);
  assert.deepStrictEqual(options('loopImpedanceTest', 'deviceType'), ['ACB', 'MCCB', 'MCB', 'SFU', 'Fuse']);
  assert.deepStrictEqual(options('prospectiveFaultCurrent', 'deviceType'), ['ACB', 'MCCB', 'MCB', 'SFU', 'Fuse']);
  assert.deepStrictEqual(options('loopImpedanceTest', 'tripCurve'), ['B', 'C', 'D', 'Adj.']);
  assert.deepStrictEqual(options('groundResistance', 'method'), ['61.8%', 'Slope', 'Clamp']);
  assert.deepStrictEqual(options('stepTouchVoltage', 'measurementType'), ['Step', 'Touch-structure', 'Touch-mesh', 'Touch-fence', 'Touch-gate']);
  assert.deepStrictEqual(options('stepTouchVoltage', 'surface'), ['Soil', 'Gravel', 'Concrete']);
  assert.deepStrictEqual(options('towerFootingResistance', 'footToEarthingConnectionStatus'), ['Given', 'Not given']);
  for (const [needle, id] of [['"Dry', 'groundCondition']]) assert.ok(source.includes('☐ Dry') && source.includes('☐ Wet'), id);
});

check('soil spacing presets are the approved ones', () => {
  assert.deepStrictEqual(SOIL_SPACING_PRESETS['etpl-default'].spacings, ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0', '4.5', '5.0']);
  assert.strictEqual(SOIL_SPACING_PRESETS['etpl-default'].directions, 4);
  assert.deepStrictEqual(SOIL_SPACING_PRESETS['is3043-substation'].spacings, ['2', '5', '10', '15', '25', '50']);
  assert.strictEqual(SOIL_SPACING_PRESETS['is3043-substation'].directions, 8);
});

check('field keys are unique per sheet and every field has a printed label', () => {
  SHEETS.forEach((sheet) => {
    const keys = flattenFields(sheet.rowFields).map((field) => field.key);
    assert.strictEqual(new Set(keys).size, keys.length, `${sheet.id} duplicate row keys`);
    flattenFields([...sheet.headerFields, ...sheet.rowFields]).forEach((field) => assert.ok(field.label, `${sheet.id}.${field.key} has no label`));
  });
});

check('definitions are JSON-safe and carry Zsat from the resolved profile for MS-09', () => {
  const published = JSON.parse(JSON.stringify(publicDefinitions()));
  assert.strictEqual(published.sheets.length, SHEETS.length);
  assert.strictEqual(published.sheets.find((sheet) => sheet.id === 'towerFootingResistance').zsat, 10);
  assert.ok(published.sheets.find((sheet) => sheet.id === 'electrodeResistance').commonHeaderFields[0].options.length >= 2);
});

console.log(`\n${passed} sheet-definition checks passed.`);
