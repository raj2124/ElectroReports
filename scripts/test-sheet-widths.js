// The entry tables use the printed column proportions: every printed row column's `width`
// in src/sheetDefinitions.js must equal the `w` of the same col(...) in the sheet generator.
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const { SHEETS } = require('../src/sheetDefinitions');

const source = fs.readFileSync(path.join(__dirname, '..', 'reference', 'generators', 'make_sheets_v3.py'), 'utf8');
const printed = {};
source.split(/\nmake\(c, /).slice(1).forEach((chunk) => {
  const docNo = /"(ET-D03-FRM-\d{3})"/.exec(chunk)[1];
  const colsBlock = chunk.slice(chunk.indexOf('['), chunk.indexOf('extra=') > 0 ? chunk.indexOf('notes=') : chunk.length);
  printed[docNo] = [...colsBlock.matchAll(/col\("(?:[^"\\]|\\.)*",\s*([\d.]+)/g)].map((m) => Number(m[1]));
});

let passed = 0;
SHEETS.filter((sheet) => !sheet.retired).forEach((sheet) => {
  const widths = sheet.rowFields.filter((field) => !field.extra && !field.legacy).map((field) => field.width);
  assert.deepStrictEqual(widths, printed[sheet.docNo], `${sheet.code} column widths`);
  passed += 1;
  console.log(`  ok - ${sheet.code} widths match the printed sheet`);
});
console.log(`\n${passed} width checks passed.`);
