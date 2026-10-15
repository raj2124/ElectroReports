// Printable measurement sheets (Session 6B.2a): the generated PDFs must match the approved
// v3 layout (reference/generators/make_sheets_v3.py and its sample PDF) and the sheet
// definitions (src/sheetDefinitions.js): same sheets, orientation, columns, order, row counts.
// Also: pre-printed rows, continuation pages, engineer, fonts, sheet IDs and QR codes, and the
// issued_sheets table that lets a scan be matched back to exactly what was printed.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const fontkit = require('fontkit');
const jsQR = require('jsqr');

const { SHEETS, INSTRUMENTS, instrumentsForSheet } = require('../src/sheetDefinitions');
const { loadRegister } = require('../src/acceptance');
const { createReportStore } = require('../src/reportStore');
const { SPECS, PRINTABLE_SHEET_IDS, PRINT } = require('../src/sheetPrint/spec');
const { renderSheetsPdf, pageCapacity, A4 } = require('../src/sheetPrint/render');
const { newSheetId, parseSheetId } = require('../src/sheetPrint/ids');
const { printSheets, planPages, SheetPrintError } = require('../src/sheetPrint');
const { Canvas } = require('../src/sheetPrint/canvas');
const { resolveFonts, BUNDLED_DIR } = require('../src/sheetPrint/fonts');
const { openPdf, pageCount, pageSize, pageTextLines, pageChars, renderPageRgba } = require('../src/pdfRaster');

const ROOT = path.join(__dirname, '..');
const GENERATOR = fs.readFileSync(path.join(ROOT, 'reference', 'generators', 'make_sheets_v3.py'), 'utf8');
const REFERENCE_PDF = fs.readFileSync(path.join(ROOT, 'reference', 'samples', 'ESHA_Measurement_Sheets_v3_MS-01_to_MS-09.pdf'));
// Row counts printed on the approved sample (MS-01..MS-09). Grouped sheets count rows, not groups.
const REFERENCE_ROWS = { soilResistivity: 10, electrodeResistance: 10, continuityTest: 19, loopImpedanceTest: 12, prospectiveFaultCurrent: 12, riserIntegrityTest: 19, groundResistance: 18, stepTouchVoltage: 12, towerFootingResistance: 20 };
const M = 26;

let passed = 0;
const tests = [];
function check(name, fn) {
  tests.push({ name, fn });
}

const squash = (text) => String(text).replace(/\s+/g, ' ').trim();
const compact = (text) => String(text).replace(/\s+/g, '');
const project = { projectNo: 'PRJ-2026-001', clientName: 'Sample Client', siteLocation: 'Main Plant, Surat', workOrder: 'WO-628' };
const blankPage = (sheetKey, extra = {}) => ({ sheetKey, sheetId: newSheetId(), project, workDate: '2026-10-14', engineer: 'Test Engineer', page: 1, pages: 1, preItems: [], itemOffset: 0, ...extra });

// ---- generator constants ----------------------------------------------------------------------
const v2 = JSON.parse(/V2 = (\{[^}]+\})/.exec(GENERATOR)[1]);
const generator = {};
GENERATOR.split(/\nmake\(c, /).slice(1).forEach((chunk) => {
  const docNo = /"(ET-D03-FRM-\d{3})"/.exec(chunk)[1];
  const geometry = /\)\],\s*(\d+),\s*(\d+)/.exec(chunk);
  generator[docNo] = {
    land: chunk.startsWith('True'),
    n: Number(geometry[1]),
    rh: Number(geometry[2]),
    group: Number((/group=(\d+)/.exec(chunk) || [0, 1])[1]),
    min: v2[(/min_rows=V2\["(\d+)"\]/.exec(chunk) || [])[1]] || 0,
    legend: !/legend=False/.test(chunk),
    fill: !/fill=False/.test(chunk)
  };
});

check('the nine printable sheets are exactly the approved MS-01..MS-09', () => {
  assert.deepStrictEqual(PRINTABLE_SHEET_IDS, SHEETS.filter((sheet) => !sheet.retired).map((sheet) => sheet.id));
  assert.strictEqual(PRINTABLE_SHEET_IDS.length, 9);
  assert.deepStrictEqual(Object.keys(generator).sort(), PRINTABLE_SHEET_IDS.map((id) => SPECS[id].docNo).sort());
});

check('print geometry (orientation, row height, base rows, min rows, group, legend, fill) equals the generator', () => {
  SHEETS.filter((sheet) => !sheet.retired).forEach((sheet) => {
    const printed = generator[sheet.docNo];
    const mine = PRINT[sheet.id];
    assert.strictEqual(mine.land, printed.land, `${sheet.code} orientation`);
    assert.strictEqual(mine.n, printed.n, `${sheet.code} base rows`);
    assert.strictEqual(mine.rh, printed.rh, `${sheet.code} row height`);
    assert.strictEqual(mine.min, printed.min, `${sheet.code} min rows`);
    assert.strictEqual(mine.legend, printed.legend, `${sheet.code} legend`);
    assert.strictEqual(mine.fill, printed.fill, `${sheet.code} fill`);
    assert.strictEqual(SPECS[sheet.id].group, printed.group || 1, `${sheet.code} group`);
  });
});

check('MI 3295 is a normal instrument: no "optional / confirm" wording in the definitions, register or print', () => {
  const instrument = INSTRUMENTS.find((item) => item.id === 'mi3295');
  assert.strictEqual(instrument.label, 'Metrel MI 3295 Step & Contact Voltage Meter');
  assert.ok(loadRegister().instruments.step_touch.includes('MI 3295'));
  assert.ok(!JSON.stringify(loadRegister().instruments).includes('to confirm'));
  assert.ok(instrumentsForSheet('stepTouchVoltage').some((item) => item.id === 'mi3295'));
  assert.ok(!GENERATOR.includes('confirm availability'));
  assert.ok(!SPECS.stepTouchVoltage.notes.some((note) => /confirm availability/i.test(note)));
});

// ---- generated PDF vs reference PDF vs sheet definitions -----------------------------------------
function columnBounds(spec, width) {
  const total = spec.columns.reduce((sum, column) => sum + column.w, 0);
  const cw = width - 2 * M;
  let x = M;
  return spec.columns.map((column) => {
    const left = x;
    x += (column.w / total) * cw;
    return [left, x];
  });
}

// Header labels per column and the Sr. numbers of the rows, read back from a PDF page.
function readTable(pdf, index, spec) {
  const { width } = pageSize(pdf, index);
  const lines = pageTextLines(pdf, index);
  const bounds = columnBounds(spec, width);
  const srLine = lines.find((line) => line.text.trim() === 'Sr.' && line.x < bounds[0][1]);
  assert.ok(srLine, `${spec.code}: no Sr. header on the page`);
  const centreY = srLine.y + srLine.h / 2;
  const columnOf = (line) => bounds.findIndex(([left, right]) => line.x + line.w / 2 >= left && line.x + line.w / 2 < right);
  // Header labels are read character by character: neighbouring labels can touch each other, so
  // a text line may run across two columns.
  const headerChars = pageChars(pdf, index).filter((ch) => Math.abs((ch.y0 + ch.y1) / 2 - centreY) <= 14);
  const labels = spec.columns.map((_, column) => {
    const mine = headerChars
      .filter((ch) => {
        const middle = (ch.x0 + ch.x1) / 2;
        return middle >= bounds[column][0] && middle < bounds[column][1];
      })
      .sort((a, b) => Math.round((a.y0 + a.y1) / 2 / 3) - Math.round((b.y0 + b.y1) / 2 / 3) || a.x0 - b.x0);
    let text = '';
    let previous = null;
    mine.forEach((ch) => {
      if (previous && Math.abs((ch.y0 + ch.y1) / 2 - (previous.y0 + previous.y1) / 2) > 3) {
        text += ' ';
      }
      text += ch.c;
      previous = ch;
    });
    return squash(text);
  });
  const serials = lines
    .filter((line) => line.y > centreY + 14 && columnOf(line) === 0 && /^\d+$/.test(line.text.trim()))
    .map((line) => Number(line.text.trim()));
  const headerText = headerChars.map((ch) => ch.c).join('').replace(/\s+/g, '');
  return { labels, headerText, serials, lines, width };
}

const generatedCache = {};
async function generatedBlank() {
  if (!generatedCache.pdf) {
    const { buffer } = await renderSheetsPdf(PRINTABLE_SHEET_IDS.map((id) => blankPage(id)));
    generatedCache.buffer = buffer;
    generatedCache.pdf = await openPdf(buffer);
  }
  return generatedCache.pdf;
}

check('generated PDF: nine pages in the approved orientation and A4 size, same as the reference', async () => {
  const generated = await generatedBlank();
  const reference = await openPdf(REFERENCE_PDF);
  assert.strictEqual(pageCount(generated), 9);
  assert.strictEqual(pageCount(reference), 9);
  PRINTABLE_SHEET_IDS.forEach((id, index) => {
    const mine = pageSize(generated, index);
    const theirs = pageSize(reference, index);
    assert.ok(Math.abs(mine.width - theirs.width) < 0.5 && Math.abs(mine.height - theirs.height) < 0.5, `${SPECS[id].code} page size`);
    assert.strictEqual(mine.width > mine.height, PRINT[id].land, `${SPECS[id].code} orientation`);
    assert.ok(Math.abs(Math.max(mine.width, mine.height) - A4.h) < 0.5);
  });
});

check('generated PDF: columns match sheetDefinitions and the reference PDF, in the same order', async () => {
  const generated = await generatedBlank();
  const reference = await openPdf(REFERENCE_PDF);
  PRINTABLE_SHEET_IDS.forEach((id, index) => {
    const spec = SPECS[id];
    const sheet = SHEETS.find((item) => item.id === id);
    const expected = sheet.rowFields.filter((field) => !field.extra && !field.legacy).map((field) => squash(field.label));
    const mine = readTable(generated, index, spec);
    const theirs = readTable(reference, index, spec);
    assert.deepStrictEqual(mine.labels, expected, `${spec.code} generated columns vs definitions (order and text)`);
    // The approved sample's own headers can run a little wider than their column (MS-07), so the
    // reference is compared by the characters in its header band, not column by column.
    const sorted = (text) => [...text].sort().join('');
    assert.strictEqual(sorted(theirs.headerText), sorted(compact(expected.join(''))), `${spec.code} reference header text vs definitions`);
    assert.strictEqual(sorted(mine.headerText), sorted(theirs.headerText), `${spec.code} generated vs reference header text`);
  });
});

check('generated PDF: row counts equal the reference (none fewer) and the computed capacity', async () => {
  const generated = await generatedBlank();
  const reference = await openPdf(REFERENCE_PDF);
  PRINTABLE_SHEET_IDS.forEach((id, index) => {
    const spec = SPECS[id];
    const perRow = (serials) => serials.length * spec.group;
    const mine = perRow(readTable(generated, index, spec).serials);
    const theirs = perRow(readTable(reference, index, spec).serials);
    assert.strictEqual(theirs, REFERENCE_ROWS[id], `${spec.code} reference sample rows`);
    assert.strictEqual(mine, theirs, `${spec.code} rows: generated ${mine} vs reference ${theirs}`);
    assert.strictEqual(pageCapacity(spec).rows, REFERENCE_ROWS[id], `${spec.code} computed capacity`);
    assert.deepStrictEqual(readTable(generated, index, spec).serials, Array.from({ length: mine / spec.group }, (_, n) => n + 1));
  });
});

check('generated PDF: title, document number, header cells, instruments, notes and footer', async () => {
  const generated = await generatedBlank();
  const reference = await openPdf(REFERENCE_PDF);
  PRINTABLE_SHEET_IDS.forEach((id, index) => {
    const spec = SPECS[id];
    const sheet = SHEETS.find((item) => item.id === id);
    const text = compact(pageTextLines(generated, index).map((line) => line.text).join(' '));
    const referenceText = compact(pageTextLines(reference, index).map((line) => line.text).join(' '));
    assert.ok(text.includes(compact(`${spec.docNo} | Rev 00`)), `${spec.code} doc no`);
    assert.ok(referenceText.includes(compact(spec.title)) && referenceText.includes(compact(spec.docNo)), `${spec.code} reference title`);
    assert.ok(text.includes(compact(spec.title)), `${spec.code} title`);
    assert.ok(text.includes(compact(`${spec.title.split(' – ')[0]} | ${spec.docNo}`)), `${spec.code} footer`);
    assert.ok(text.includes('INTERNAL'));
    ['Client', 'Project No.', 'Work Order No.', 'Site / Unit', 'Work date', 'Sheet', 'ETPL engineer(s)', 'Client representative', 'Ground condition', 'Instrument used:', 'Time start / finish']
      .forEach((label) => assert.ok(text.includes(compact(label)), `${spec.code} header cell "${label}"`));
    sheet.headerFields.filter((field) => !field.extra).forEach((field) => assert.ok(text.includes(compact(field.label)), `${spec.code} header "${field.label}"`));
    instrumentsForSheet(id).forEach((instrument) => assert.ok(text.includes(compact(instrument.label)), `${spec.code} instrument ${instrument.label}`));
    sheet.notes.forEach((note) => assert.ok(text.includes(compact(note)), `${spec.code} note "${note.slice(0, 30)}..."`));
    assert.ok(text.includes('SET-UP') && text.includes('HOWTORECORD'), `${spec.code} band`);
    assert.strictEqual(text.includes('Conditioncodes:'), PRINT[id].legend, `${spec.code} condition-code legend`);
    if (PRINT[id].legend) {
      assert.ok(text.includes(compact('C1 Corroded')) && text.includes(compact('OTH Other – write remark')));
    }
    ['ETPL engineer – name & signature', 'Client representative – name & signature'].forEach((label) => assert.ok(text.includes(compact(label)), `${spec.code} signature "${label}"`));
  });
  const stepText = compact(pageTextLines(generated, PRINTABLE_SHEET_IDS.indexOf('stepTouchVoltage')).map((line) => line.text).join(' '));
  assert.ok(stepText.includes('Metrel MI 3295 Step & Contact Voltage Meter'.replace(/\s+/g, '')));
  assert.ok(!/optional|confirmavailability/i.test(stepText));
});

check('soil sheet prints the approved spacing set; fixed columns repeat per row', async () => {
  const generated = await generatedBlank();
  const lines = pageTextLines(generated, 0).map((line) => line.text.trim());
  ['0.50', '1.00', '1.50', '2.00', '2.50', '3.00', '3.50', '4.00', '4.50', '5.00'].forEach((spacing) => assert.ok(lines.includes(spacing), `spacing ${spacing}`));
  const loop = pageTextLines(generated, PRINTABLE_SHEET_IDS.indexOf('loopImpedanceTest')).map((line) => line.text.trim());
  assert.strictEqual(loop.filter((text) => text === 'R-E').length, 4);
  assert.strictEqual(loop.filter((text) => text === 'B-E').length, 4);
  const tower = pageTextLines(generated, PRINTABLE_SHEET_IDS.indexOf('towerFootingResistance')).map((line) => line.text.trim());
  assert.strictEqual(tower.filter((text) => text === 'F-1').length, 5 + 1, 'five tower blocks, plus the foot label in the SET-UP diagram');
});

check('pre-filled header cells: client, project, work order, site, work date, engineer; readings stay blank', async () => {
  const generated = await generatedBlank();
  const text = pageTextLines(generated, 1).map((line) => line.text.trim());
  ['Sample Client', 'PRJ-2026-001', 'WO-628', 'Main Plant, Surat', '14 / 10 / 2026', '1 of 1', 'Test Engineer'].forEach((value) => assert.ok(text.includes(value), `header value ${value}`));
  const numeric = text.filter((line) => /^[\d.]+$/.test(line));
  assert.deepStrictEqual(numeric, ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'], 'only the Sr. numbers are printed in the table body');
});

// ---- fonts ----------------------------------------------------------------------------------
check('fonts: Arial when installed, bundled Liberation Sans (with its OFL licence) otherwise; every glyph exists', async () => {
  const bundled = { family: 'Liberation Sans', regular: path.join(BUNDLED_DIR, 'LiberationSans-Regular.ttf'), bold: path.join(BUNDLED_DIR, 'LiberationSans-Bold.ttf') };
  assert.ok(fs.existsSync(bundled.regular) && fs.existsSync(bundled.bold));
  assert.ok(/SIL OPEN FONT LICENSE/i.test(fs.readFileSync(path.join(BUNDLED_DIR, 'LICENSE-Liberation-OFL.txt'), 'utf8')));
  const chosen = resolveFonts();
  assert.ok(chosen.family === 'Arial' || chosen.family === 'Liberation Sans');
  if (fs.existsSync('/System/Library/Fonts/Supplemental/Arial.ttf')) {
    assert.strictEqual(chosen.family, 'Arial');
  }
  const drawn = new Set();
  const original = Canvas.prototype.text;
  Canvas.prototype.text = function recordText(x, y, value, options) {
    [...String(value)].forEach((ch) => drawn.add(ch));
    return original.call(this, x, y, value, options);
  };
  try {
    const pages = PRINTABLE_SHEET_IDS.map((id) => blankPage(id, { preItems: [{ tag: 'Ω ρ °' }] }));
    const withBundled = await renderSheetsPdf(pages, { fonts: bundled });
    assert.ok(withBundled.buffer.length > 50000);
    const pdf = await openPdf(withBundled.buffer);
    assert.strictEqual(pageCount(pdf), 9);
    [chosen, bundled].forEach((fonts) => {
      [fonts.regular, fonts.bold].forEach((file) => {
        const font = fontkit.openSync(file);
        const missing = [...drawn].filter((ch) => ch.trim() && !font.hasGlyphForCodePoint(ch.codePointAt(0)));
        assert.deepStrictEqual(missing, [], `${path.basename(file)} lacks glyphs: ${missing.join(' ')}`);
      });
    });
  } finally {
    Canvas.prototype.text = original;
  }
});

// ---- sheet IDs and QR codes ----------------------------------------------------------------------
check('sheet IDs: ER-XXXX-XXXXX, unique, check character catches a mistyped or look-alike character', () => {
  const ids = new Set();
  for (let i = 0; i < 5000; i += 1) {
    const id = newSheetId();
    assert.ok(/^ER-[A-HJ-NP-Z2-9]{4}-[A-HJ-NP-Z2-9]{5}$/.test(id), id);
    assert.strictEqual(parseSheetId(id), id);
    assert.strictEqual(parseSheetId(`ER1|${id.toLowerCase()}`), id);
    ids.add(id);
  }
  assert.strictEqual(ids.size, 5000);
  const id = newSheetId();
  const body = id.slice(3).replace('-', '');
  const swapped = `ER-${body[1]}${body[0]}${body.slice(2, 4)}-${body.slice(4)}`;
  assert.ok(body[0] === body[1] || parseSheetId(swapped) === null, 'a transposition is rejected');
  assert.strictEqual(parseSheetId(id.slice(0, -1) + (id.endsWith('A') ? 'B' : 'A')), null);
  assert.strictEqual(parseSheetId('hello'), null);
  assert.strictEqual(parseSheetId(''), null);
});

check('every printed page carries a QR code that scans back to its own sheet ID', async () => {
  const pages = PRINTABLE_SHEET_IDS.map((id) => blankPage(id));
  const { buffer } = await renderSheetsPdf(pages);
  const pdf = await openPdf(buffer);
  for (let index = 0; index < pages.length; index += 1) {
    const raster = renderPageRgba(pdf, index, 150);
    const code = jsQR(raster.data, raster.width, raster.height);
    assert.ok(code, `${SPECS[pages[index].sheetKey].code}: QR not found`);
    assert.strictEqual(code.data, `ER1|${pages[index].sheetId}`);
    const text = pageTextLines(pdf, index).map((line) => line.text).join(' ');
    assert.ok(text.includes(`Sheet ID ${pages[index].sheetId}`), 'the sheet ID is also printed in words');
  }
});

// ---- pre-print, continuation pages, engineer, issued sheets --------------------------------------
function makeStore() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'er-print-'));
  return createReportStore({ filePath: path.join(dir, 'reports.db') });
}

const pit = (tag, location) => ({ tag, location });

check('pre-print: tags and locations from earlier reports of the same project, de-duplicated, in order', async () => {
  const store = makeStore();
  store.addReport({ project: { ...project, reportDate: '2026-10-10' }, electrodeResistance: [pit('ep-01', 'TR-1'), pit('EP-02', 'TR-1')] });
  store.addReport({ project: { ...project, reportDate: '2026-10-12' }, electrodeResistance: [pit('EP-03', 'MCC'), pit('EP-02', 'TR-1 bay'), pit('EP-01', 'TR-1')] });
  store.addReport({ project: { ...project, reportDate: '2026-10-14' }, electrodeResistance: [pit('EP-SAME-DAY', 'x')] });
  store.addReport({ project: { ...project, reportDate: '2026-10-20' }, electrodeResistance: [pit('EP-LATER', 'x')] });
  store.addReport({ project: { ...project, projectNo: 'OTHER-1', reportDate: '2026-10-09' }, electrodeResistance: [pit('EP-OTHER', 'x')] });
  const result = await printSheets({ project, workDate: '2026-10-14', sheets: ['electrodeResistance'], engineerName: 'A. Engineer' }, { store });
  assert.strictEqual(result.pages.length, 1);
  assert.deepStrictEqual(result.pages[0].preRows.map((row) => row.tag), ['EP-01', 'EP-02', 'EP-03']);
  assert.strictEqual(result.pages[0].preRows[1].location, 'TR-1 bay', 'the latest wording of a point is printed');
  const pdf = await openPdf(result.buffer);
  const text = pageTextLines(pdf, 0).map((line) => line.text.trim());
  ['EP-01', 'EP-02', 'EP-03'].forEach((tag) => assert.ok(text.includes(tag)));
  ['EP-SAME-DAY', 'EP-LATER', 'EP-OTHER'].forEach((tag) => assert.ok(!text.includes(tag), `${tag} must not be pre-printed`));
  assert.strictEqual(pageSize(pdf, 0).width > pageSize(pdf, 0).height, true);
  assert.strictEqual(readTable(pdf, 0, SPECS.electrodeResistance).serials.length, 10, 'blank rows fill the page: no fewer than the reference');
});

check('pre-print: the report being worked on is not a source for its own sheets; no earlier report means blank rows', async () => {
  const store = makeStore();
  const own = store.addReport({ project: { ...project, reportDate: '2026-10-01' }, electrodeResistance: [pit('EP-OWN', 'x')] });
  const withOwn = await printSheets({ project, workDate: '2026-10-14', sheets: ['electrodeResistance'], reportId: own.id }, { store });
  assert.deepStrictEqual(withOwn.pages[0].preRows, []);
  const fresh = await printSheets({ project: { ...project, projectNo: 'NEW-1' }, workDate: '2026-10-14', sheets: ['continuityTest', 'electrodeResistance'] }, { store: makeStore() });
  assert.deepStrictEqual(fresh.pages.map((page) => page.sheetCode), ['MS-02', 'MS-03'], 'printed in MS order');
  fresh.pages.forEach((page) => assert.deepStrictEqual(page.preRows, []));
});

check('continuation pages: more points than rows prints more pages, each with its own ID, numbering continues', async () => {
  const store = makeStore();
  const pits = Array.from({ length: 12 }, (_, i) => pit(`EP-${String(i + 1).padStart(2, '0')}`, `Loc ${i + 1}`));
  store.addReport({ project: { ...project, reportDate: '2026-10-10' }, electrodeResistance: pits });
  const result = await printSheets({ project, workDate: '2026-10-14', sheets: ['electrodeResistance'], engineerName: 'A' }, { store });
  assert.strictEqual(result.pages.length, 2);
  assert.notStrictEqual(result.pages[0].sheetId, result.pages[1].sheetId);
  assert.deepStrictEqual(result.pages.map((page) => [page.page, page.pages, page.rowOffset, page.preRows.length]), [[1, 2, 0, 10], [2, 2, 10, 2]]);
  const pdf = await openPdf(result.buffer);
  assert.strictEqual(pageCount(pdf), 2);
  const second = readTable(pdf, 1, SPECS.electrodeResistance);
  assert.deepStrictEqual(second.serials, Array.from({ length: 10 }, (_, i) => 11 + i), 'Sr. continues from 11');
  const text = pageTextLines(pdf, 1).map((line) => line.text.trim());
  assert.ok(text.includes('EP-11') && text.includes('EP-12') && !text.includes('EP-10'));
  assert.ok(text.includes('2 of 2'));
});

check('grouped sheets pre-print one block per feeder / tower and paginate by blocks', async () => {
  const store = makeStore();
  const feeders = Array.from({ length: 5 }, (_, i) => ['R-E', 'Y-E', 'B-E'].map((phase) => ({ location: `Panel ${i + 1}`, feederTag: `F-${i + 1}`, measuredPoints: phase })));
  const towers = Array.from({ length: 6 }, (_, i) => ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'].map((foot) => ({ mainLocationTower: `Tower ${i + 1}`, measurementPointLocation: foot })));
  store.addReport({ project: { ...project, reportDate: '2026-10-10' }, loopImpedanceTest: feeders.flat(), towerFootingResistance: towers.flat() });
  const result = await printSheets({ project, workDate: '2026-10-14', sheets: ['loopImpedanceTest', 'towerFootingResistance'] }, { store });
  const byCode = (code) => result.pages.filter((page) => page.sheetCode === code);
  assert.deepStrictEqual(byCode('MS-04').map((page) => page.preRows.length), [4, 1]);
  assert.deepStrictEqual(byCode('MS-04').map((page) => page.rowOffset), [0, 12]);
  assert.deepStrictEqual(byCode('MS-09').map((page) => page.preRows.length), [5, 1]);
  const pdf = await openPdf(result.buffer);
  const loop = pageTextLines(pdf, 0).map((line) => line.text.trim());
  assert.ok(loop.includes('Panel 1') && loop.includes('F-4') && !loop.includes('Panel 5'));
  const loop2 = readTable(pdf, 1, SPECS.loopImpedanceTest);
  assert.deepStrictEqual(loop2.serials, [5, 6, 7, 8], 'feeder numbers continue on the continuation page');
  assert.strictEqual(loop2.lines.filter((line) => line.text.trim() === 'Panel 5').length, 1);
});

check('engineer: the Zoho task owner when a task is chosen, else the typed name, with a warning on a failed lookup', async () => {
  const store = makeStore();
  const zohoProject = { ...project, zohoProjectId: '1001' };
  const getTasks = async (id) => {
    assert.strictEqual(id, '1001');
    return [{ id: 't1', name: 'Soil testing', ownerName: 'Rakesh Patel' }, { id: 't2', name: 'No owner', ownerName: '' }];
  };
  const byTask = await printSheets({ project: zohoProject, workDate: '2026-10-14', sheets: ['soilResistivity'], taskId: 't1', engineerName: 'Typed Name' }, { store, getTasks });
  assert.strictEqual(byTask.engineer, 'Rakesh Patel');
  assert.deepStrictEqual(byTask.warnings, []);
  assert.ok(pageTextLines(await openPdf(byTask.buffer), 0).some((line) => line.text.trim() === 'Rakesh Patel'));
  const typed = await printSheets({ project: zohoProject, workDate: '2026-10-14', sheets: ['soilResistivity'], engineerName: 'Typed Name' }, { store, getTasks });
  assert.strictEqual(typed.engineer, 'Typed Name');
  const noOwner = await printSheets({ project: zohoProject, workDate: '2026-10-14', sheets: ['soilResistivity'], taskId: 't2', engineerName: 'Typed Name' }, { store, getTasks });
  assert.strictEqual(noOwner.engineer, 'Typed Name');
  assert.strictEqual(noOwner.warnings.length, 1);
  const failing = await printSheets({ project: zohoProject, workDate: '2026-10-14', sheets: ['soilResistivity'], taskId: 't1', engineerName: 'Typed Name' }, { store, getTasks: async () => { throw new Error('Zoho down'); } });
  assert.strictEqual(failing.engineer, 'Typed Name');
  assert.strictEqual(failing.warnings.length, 1);
  const blank = await printSheets({ project, workDate: '2026-10-14', sheets: ['soilResistivity'] }, { store });
  assert.strictEqual(blank.engineer, '');
});

check('every printed page is stored with its ID, project, date, test, page and pre-printed rows', async () => {
  const store = makeStore();
  store.addReport({ project: { ...project, reportDate: '2026-10-10' }, electrodeResistance: [pit('EP-01', 'TR-1')], continuityTest: [{ tag: 'C-1', mainLocation: 'MCC', referencePoint: 'Bar A', measurementPoint: 'Motor 1' }] });
  const result = await printSheets({ project: { ...project, zohoProjectId: 'Z1' }, workDate: '2026-10-14', sheets: ['soilResistivity', 'electrodeResistance', 'continuityTest'], engineerName: 'A. Engineer' }, { store });
  assert.strictEqual(result.pages.length, 3);
  result.pages.forEach((page) => {
    const stored = store.getIssuedSheet(page.sheetId);
    assert.ok(stored, `${page.sheetCode} stored`);
    assert.strictEqual(stored.batchId, result.batchId);
    assert.strictEqual(stored.project.projectNo, 'PRJ-2026-001');
    assert.strictEqual(stored.project.zohoProjectId, 'Z1');
    assert.strictEqual(stored.workDate, '2026-10-14');
    assert.strictEqual(stored.sheetKey, page.sheetKey);
    assert.strictEqual(stored.docNo, SPECS[page.sheetKey].docNo);
    assert.strictEqual([stored.page, stored.pages, stored.rowOffset, stored.rowCount].join(','), [1, 1, 0, REFERENCE_ROWS[page.sheetKey]].join(','));
    assert.strictEqual(stored.engineer, 'A. Engineer');
    assert.deepStrictEqual(stored.preRows, page.preRows);
  });
  assert.deepStrictEqual(store.getIssuedSheet(result.pages[2].sheetId).preRows, [{ tag: 'C-1', mainLocation: 'MCC', referencePoint: 'Bar A', measurementPoint: 'Motor 1' }]);
  assert.strictEqual(store.listIssuedSheets({ batchId: result.batchId }).length, 3);
  assert.strictEqual(store.listIssuedSheets({ projectNo: 'PRJ-2026-001', workDate: '2026-10-14' }).length, 3);
  assert.strictEqual(store.listIssuedSheets({ workDate: '2030-01-01' }).length, 0);
  assert.strictEqual(store.getIssuedSheet('ER-NOPE-NOPE2'), null);
  // A later print is a new batch with new IDs; earlier IDs stay matchable.
  const again = await printSheets({ project, workDate: '2026-10-14', sheets: ['soilResistivity'] }, { store });
  assert.notStrictEqual(again.batchId, result.batchId);
  assert.ok(store.getIssuedSheet(result.pages[0].sheetId));
  assert.strictEqual(store.listIssuedSheets({ projectNo: 'PRJ-2026-001' }).length, 4);
});

check('requests are validated: work date, project number, tests, retired and unknown sheets', async () => {
  const store = makeStore();
  const bad = async (body, pattern) => {
    await assert.rejects(() => printSheets(body, { store }), (error) => error instanceof SheetPrintError && error.status === 400 && pattern.test(error.message));
  };
  await bad({ project, sheets: ['soilResistivity'] }, /work date/i);
  await bad({ project, workDate: '14/10/2026', sheets: ['soilResistivity'] }, /work date/i);
  await bad({ project: { ...project, projectNo: '' }, workDate: '2026-10-14', sheets: ['soilResistivity'] }, /project number/i);
  await bad({ project, workDate: '2026-10-14', sheets: [] }, /at least one test/i);
  await bad({ project, workDate: '2026-10-14' }, /at least one test/i);
  await bad({ project, workDate: '2026-10-14', sheets: ['earthContinuityTest'] }, /no printed sheet/i);
  await bad({ project, workDate: '2026-10-14', sheets: ['nonsense'] }, /no printed sheet/i);
  assert.strictEqual(store.listIssuedSheets().length, 0, 'a refused request issues no IDs');
});

check('a filled test sheet prints values into the cells (used by the OCR tests), and its QR still scans', async () => {
  const filled = {
    instrument: 'mi3290',
    serialNo: 'SN-4471',
    calibrationDue: '12/03/2027',
    timeStart: '09:10',
    timeFinish: '11:40',
    job: { groundCondition: 'Damp', clientRepresentative: 'S. Mehta' },
    headerValues: { currentElectrodeDistanceM: '31.5', auxRodsOk: true, check90180: 'Yes' },
    rows: [{ tag: 'EP-01', location: 'TR-1', electrodeType: 'Rod', materialType: 'GI', length: '3.0', diameter: '16', resistanceWithoutGrid: '2.45', checkR50: '2.41', checkR70: '2.49', resistanceWithGrid: '1.12', method: '61.8%', conditionCodes: ['C1', 'C6'] }, { tag: 'EP-02', resistanceWithoutGrid: '4.2' }]
  };
  const page = blankPage('electrodeResistance', { filled });
  const { buffer } = await renderSheetsPdf([page]);
  const pdf = await openPdf(buffer);
  const text = pageTextLines(pdf, 0).map((line) => line.text.trim());
  ['SN-4471', '12/03/2027', '09:10 / 11:40', 'S. Mehta', '31.5', 'EP-01', 'TR-1', '3.0', '16', '2.45', '2.41', '2.49', '1.12', 'EP-02', '4.2'].forEach((value) => assert.ok(text.some((line) => line.includes(value)), `filled value ${value}`));
  const raster = renderPageRgba(pdf, 0, 150);
  assert.strictEqual(jsQR(raster.data, raster.width, raster.height).data, `ER1|${page.sheetId}`);
  // The blank sheet is unchanged by the fill option: same rows, none of the values.
  const blank = await openPdf((await renderSheetsPdf([blankPage('electrodeResistance')])).buffer);
  assert.ok(!pageTextLines(blank, 0).some((line) => line.text.includes('SN-4471') || line.text.includes('2.45')));
});

check('planPages gives each page of a batch one timestamp and batch id, and rows match the capacity', () => {
  const pages = planPages({ project, workDate: '2026-10-14', sheets: PRINTABLE_SHEET_IDS, reportId: '' }, [], 'E');
  assert.strictEqual(pages.length, 9);
  assert.strictEqual(new Set(pages.map((page) => page.batchId)).size, 1);
  assert.strictEqual(new Set(pages.map((page) => page.sheetId)).size, 9);
  pages.forEach((page) => assert.strictEqual(page.rowCount, REFERENCE_ROWS[page.sheetKey]));
});

(async () => {
  for (const { name, fn } of tests) {
    try {
      await fn();
    } catch (error) {
      console.error(`  FAIL - ${name}`);
      throw error;
    }
    passed += 1;
    console.log(`  ok - ${name}`);
  }
  console.log(`\n${passed} printable-sheet checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
