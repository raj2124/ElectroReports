// Session 6I, scanned sheets, on mocks (no provider is called): the form of each printed sheet, the cell checks, PDF / image
// splitting and QR reading (also on a degraded photo), the page reader with its escalation, the merge into a daily report by
// sheetRow for every layout, and the whole upload -> read -> edit -> accept flow including its rules.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createReportStore } = require('../src/reportStore');
const { SPECS } = require('../src/sheetPrint/spec');
const { renderSheetsPdf } = require('../src/sheetPrint/render');
const { newSheetId } = require('../src/sheetPrint/ids');
const { buildForm } = require('../src/scan/form');
const { cleanCell, flagsFor, normaliseNumber, parseDate } = require('../src/scan/cells');
const { splitUpload, UploadError } = require('../src/scan/split');
const { findSheetId } = require('../src/scan/qr');
const { readPage, applyEdits } = require('../src/scan/reader');
const { applyScan } = require('../src/scan/apply');
const { createMockProvider } = require('../src/scan/mockProvider');
const { createAiGuard, ProviderError } = require('../src/aiGuard');
const { createScanService, ScanError } = require('../src/scan/service');
const { normalizeReportInput } = require('../src/reportModel');
const { issueFilled, withoutQr } = require('./lib/scan-fixtures');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'scan-'));
const newStore = () => createReportStore({ filePath: path.join(dir, `${Math.random().toString(36).slice(2)}.db`) });
const PROJECT = { projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej' };

function setup() {
  const store = newStore();
  const anup = store.addUser({ email: 'anup@elegrow.com', name: 'Anup Parekh', roles: ['engineer'] });
  const nidhi = store.addUser({ email: 'nidhi@elegrow.com', name: 'Nidhi Shah', roles: ['engineer'] });
  const reviewer = store.addUser({ email: 'rev@elegrow.com', name: 'A. P.', roles: ['reviewer'] });
  const admin = store.addUser({ email: 'admin@elegrow.com', name: 'H. S. Patil', roles: ['admin'] });
  const workspace = store.addProject({ ...PROJECT, workOrder: 'W', engineerName: 'E' });
  return { store, anup, nidhi, reviewer, admin, workspace };
}
const cellOf = (form, path_) => form.cells.find((cell) => cell.path === path_);
const flagCodes = (cell) => cell.flags.map((entry) => entry.code).sort();
const ELECTRODE_ROWS = [
  { tag: 'P-1', location: 'Yard', electrodeType: 'Rod', materialType: 'GI', length: '3', diameter: '20', resistanceWithoutGrid: '0.62', resistanceWithGrid: '0.52', method: '61.8%', conditionCodes: ['C1', 'C6'], observation: 'corroded clamp' },
  { tag: 'P-2', location: 'Gate', electrodeType: 'Plate', materialType: 'Cu', length: '0.6', diameter: '3', resistanceWithGrid: '1.20', method: '61.8%' }
];

(async () => {
  // ---- the form of a printed page -----------------------------------------------------------------
  await check('the form lists the cells of each printed sheet: no serial or fixed columns, merged columns once per group, pre-printed text from the issued record', () => {
    const flat = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10, preRows: [{ tag: 'P-9', location: 'Old yard' }] });
    assert.ok(!flat.cells.some((cell) => ['srNo', 'spacing', 'measuredPoints'].includes(cell.key)));
    assert.deepStrictEqual([cellOf(flat, 'r0.tag').preprinted, cellOf(flat, 'r0.tag').value, Boolean(cellOf(flat, 'r1.tag').preprinted)], [true, 'P-9', false]);
    assert.deepStrictEqual(cellOf(flat, 'r0.conditionCodes').options.map((option) => option.value), ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8']);
    assert.strictEqual(cellOf(flat, 'r0.conditionCodes').multi, true);
    assert.strictEqual(cellOf(flat, 'r0.electrodeType').multi, false);
    assert.strictEqual(cellOf(flat, 'h.instrument').options.length, 2);
    const loop = buildForm({ sheetKey: 'loopImpedanceTest', rowCount: 12 });
    assert.ok(cellOf(loop, 'r0.location') && !cellOf(loop, 'r1.location') && cellOf(loop, 'r3.location'), 'merged columns sit on the first row of each 3-row group');
    assert.ok(cellOf(loop, 'r1.loopImpedance') && cellOf(loop, 'r2.voltage'), 'readings are per phase row');
    const tower = buildForm({ sheetKey: 'towerFootingResistance', rowCount: 20 });
    assert.ok(cellOf(tower, 'r0.mainLocationTower') && !cellOf(tower, 'r1.mainLocationTower') && cellOf(tower, 'r4.mainLocationTower'));
    assert.deepStrictEqual(['h.lineName', 'h.lineVoltageTowerType'].map((path_) => Boolean(cellOf(tower, path_))), [true, true]);
    const soil = buildForm({ sheetKey: 'soilResistivity', rowCount: 10 });
    assert.ok(cellOf(soil, 'h.inputUnit') && cellOf(soil, 'h.direction3') && cellOf(soil, 'r9.direction4') && !cellOf(soil, 'h.instrument'), 'soil: location header, 10 spacing rows, no instrument row');
    const step = buildForm({ sheetKey: 'stepTouchVoltage', rowCount: 12 });
    assert.ok(cellOf(step, 'h.faultCurrentKA') && cellOf(step, 'h.concreteThicknessM'), 'grouped header cells are read one by one');
    Object.keys(SPECS).forEach((sheetKey) => assert.ok(buildForm({ sheetKey, rowCount: SPECS[sheetKey].baseRows }).cells.length > 20, sheetKey));
    assert.throws(() => buildForm({ sheetKey: 'nope' }), /no printable sheet/);
  });

  // ---- cells and flags -----------------------------------------------------------------------------
  await check('flags: low confidence, illegible, not a number, out of range, two ticks on a single choice, unclear unit, bad date, invalid option; comma decimals are read as numbers', () => {
    const form = buildForm({ sheetKey: 'soilResistivity', rowCount: 10 });
    const number = cellOf(form, 'r0.direction1');
    const check_ = (def, raw, threshold = 0.8) => { const cell = cleanCell(def, raw); cell.flags.push(...flagsFor(def, cell, threshold)); return cell; };
    assert.deepStrictEqual(flagCodes(check_(number, { value: '12.5', confidence: 0.95 })), []);
    assert.deepStrictEqual(flagCodes(check_(number, { value: '12.5', confidence: 0.79 })), ['LOW_CONFIDENCE']);
    assert.deepStrictEqual(flagCodes(check_(number, { value: '12.5', confidence: 0.5 }, 0.4)), [], 'the threshold is configurable');
    assert.deepStrictEqual(flagCodes(check_(number, { value: '1?5', confidence: 0.99, illegible: true })), ['ILLEGIBLE', 'NON_NUMERIC']);
    assert.deepStrictEqual(flagCodes(check_(number, { value: 'abc', confidence: 0.99 })), ['NON_NUMERIC']);
    assert.deepStrictEqual(flagCodes(check_(number, { value: '0', confidence: 0.99 })), ['OUT_OF_RANGE'], 'a soil reading must be above 0');
    assert.strictEqual(check_(number, { value: '0,52', confidence: 0.99 }).value, '0.52');
    assert.strictEqual(normaliseNumber(' 1 2,5 '), '12.5');
    assert.deepStrictEqual(flagCodes(check_(cellOf(form, 'h.direction1'), { value: '400', confidence: 0.99 })), ['OUT_OF_RANGE'], 'a bearing is at most 360');
    const unit = cellOf(form, 'h.inputUnit');
    assert.deepStrictEqual(flagCodes(check_(unit, { selected: ['ohm', 'ohm-m'], confidence: 0.99 })), ['MULTI_TICK', 'UNIT_UNCLEAR']);
    assert.deepStrictEqual(flagCodes(check_(unit, { selected: [], confidence: 0.99 })), ['UNIT_UNCLEAR']);
    assert.deepStrictEqual(flagCodes(check_(unit, { selected: ['ohm'], confidence: 0.99 })), []);
    assert.deepStrictEqual(flagCodes(check_(unit, { selected: ['ohm', 'bogus'], confidence: 0.99 })), ['INVALID_OPTION'], 'a tick that is no option is dropped and flagged');
    const flat = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10 });
    assert.deepStrictEqual(flagCodes(check_(cellOf(flat, 'r0.conditionCodes'), { selected: ['C1', 'C6'], confidence: 0.99 })), [], 'several condition codes are fine');
    assert.deepStrictEqual(flagCodes(check_(cellOf(flat, 'r0.electrodeType'), { selected: ['Rod', 'Pipe'], confidence: 0.99 })), ['MULTI_TICK']);
    assert.deepStrictEqual(flagCodes(check_(cellOf(flat, 'h.calibrationDue'), { value: '31/02/2027', confidence: 0.99 })), ['BAD_DATE']);
    assert.strictEqual(parseDate('5/11/26'), '2026-11-05');
    assert.strictEqual(parseDate('2026-11-05'), '2026-11-05');
    assert.strictEqual(parseDate('05.11.2026'), '2026-11-05');
    assert.strictEqual(parseDate('nope'), '');
  });

  // ---- splitting and the QR code -------------------------------------------------------------------
  await check('a multi-page PDF is split into pages and every page\'s QR is found; a PNG is one page; bad uploads are refused', async () => {
    const { store } = setup();
    const a = await issueFilled(store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    const b = await issueFilled(store, { sheetKey: 'loopImpedanceTest', project: PROJECT, workDate: '2026-10-12' });
    const c = await issueFilled(store, { sheetKey: 'towerFootingResistance', project: PROJECT, workDate: '2026-10-12' });
    // one PDF of three pages
    const { buffer } = await renderSheetsPdf([a, b, c].map((sheet, index) => ({ ...sheet.pageSet[0], sheetId: sheet.issued.sheetId, filled: { rows: index === 0 ? ELECTRODE_ROWS : [] } })));
    const pages = await splitUpload({ buffer, mimetype: 'application/pdf' });
    assert.strictEqual(pages.length, 3);
    assert.ok(pages.every((page) => Math.max(page.width, page.height) <= 2300 && page.png.slice(0, 4).toString('hex') === '89504e47'));
    assert.deepStrictEqual(pages.map((page) => findSheetId(page.rgba).sheetId), [a, b, c].map((sheet) => sheet.issued.sheetId));
    const image = await splitUpload({ buffer: pages[0].png, mimetype: 'image/png' });
    assert.strictEqual(image.length, 1);
    assert.strictEqual(findSheetId(image[0].rgba).sheetId, a.issued.sheetId);
    await assert.rejects(() => splitUpload({ buffer: Buffer.from('hello world'), mimetype: 'text/plain' }), UploadError);
    await assert.rejects(() => splitUpload({ buffer: Buffer.from('%PDF-1.4 garbage'), mimetype: 'application/pdf' }), (error) => error instanceof UploadError && /could not be opened/.test(error.message));
    await assert.rejects(() => splitUpload({ buffer, mimetype: 'application/pdf', maxPages: 2 }), (error) => error.code === 'TOO_MANY_PAGES');
  });

  await check('the QR search retries scales and contrast: a washed-out, noisy, shrunken page still scans; a page without a code does not', async () => {
    const { store } = setup();
    const a = await issueFilled(store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    const [page] = await splitUpload({ buffer: a.pdf, mimetype: 'application/pdf' });
    const { data, width, height } = page.rgba;
    let seed = 7;
    const noise = () => { seed = (seed * 16807) % 2147483647; return (seed / 2147483647 - 0.5) * 70; };
    const washed = new Uint8ClampedArray(data.length);
    for (let i = 0; i < data.length; i += 4) {
      for (let k = 0; k < 3; k += 1) washed[i + k] = 110 + data[i + k] * 0.45 + noise(); // low contrast, grey paper, noise
      washed[i + 3] = 255;
    }
    const hit = findSheetId({ data: washed, width, height });
    assert.ok(hit, 'found in the degraded page');
    assert.strictEqual(hit.sheetId, a.issued.sheetId);
    const blank = { data: new Uint8ClampedArray(width * height * 4).fill(255), width, height };
    assert.strictEqual(findSheetId(blank), null);
  });

  // ---- the reader ----------------------------------------------------------------------------------
  const MODELS = { primary: 'primary-model', escalation: 'escalation-model' };
  await check('the reader escalates only when something is doubtful; the stronger read replaces a doubtful cell and every disagreement is flagged', async () => {
    const { store } = setup();
    const sheet = await issueFilled(store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    const form = buildForm({ ...sheet.issued, sheetId: sheet.issued.sheetId });
    const calls = [];
    const provider = (truth) => createMockProvider({ truth: { [sheet.issued.sheetId]: truth }, primaryModel: MODELS.primary });
    const spy = (inner) => ({ read: async (args) => { calls.push(args.model); return inner.read(args); } });
    // clean read: one call, no flags, pre-printed cells (none here) excluded
    const clean = await readPage({ form, png: Buffer.alloc(0), provider: spy(provider({ primary: sheet.truth })), models: MODELS });
    assert.deepStrictEqual([calls, clean.escalated], [[MODELS.primary], false]);
    assert.ok(Object.values(clean.cells).every((cell) => cell.flags.length === 0));
    assert.strictEqual(clean.cells['r0.resistanceWithGrid'].value, '0.52');
    assert.deepStrictEqual(clean.cells['r0.conditionCodes'].selected, ['C1', 'C6']);
    // a doubtful cell: second read
    calls.length = 0;
    const doubtful = { ...sheet.truth, 'r0.resistanceWithGrid': { value: '0.5Z', confidence: 0.55 }, 'r1.tag': { value: 'P-2', confidence: 0.99 } };
    const better = { ...sheet.truth, 'r1.tag': { value: 'P-3', confidence: 0.99 } };
    const second = await readPage({ form, png: Buffer.alloc(0), provider: spy(provider({ primary: doubtful, escalation: better })), models: MODELS });
    assert.deepStrictEqual([calls, second.escalated], [[MODELS.primary, MODELS.escalation], true]);
    assert.strictEqual(second.cells['r0.resistanceWithGrid'].value, '0.52', 'the sure stronger read replaces the doubtful one');
    assert.deepStrictEqual(flagCodes(second.cells['r0.resistanceWithGrid']), ['READS_DIFFER']);
    assert.deepStrictEqual(second.cells['r0.resistanceWithGrid'].alt, { value: '0.5Z' });
    assert.deepStrictEqual(flagCodes(second.cells['r1.tag']), ['READS_DIFFER'], 'a confident disagreement is flagged too');
    assert.strictEqual(second.cells['r1.tag'].value, 'P-2', 'and the primary value stays when it was not doubtful');
    // still doubtful after the second read
    const still = await readPage({ form, png: Buffer.alloc(0), provider: provider({ primary: doubtful, escalation: { ...doubtful } }), models: MODELS });
    assert.ok(flagCodes(still.cells['r0.resistanceWithGrid']).includes('LOW_CONFIDENCE'));
    // a provider failure surfaces as a ProviderError, nothing is partial
    await assert.rejects(() => readPage({ form, png: Buffer.alloc(0), provider: provider({ fail: 'QUOTA' }), models: MODELS }), (error) => error instanceof ProviderError && error.code === 'QUOTA');
    // cells the model did not return are blank and confident; unknown paths are ignored
    const partial = await readPage({ form, png: Buffer.alloc(0), provider: { read: async () => ({ cells: [{ path: 'r0.tag', value: 'X', confidence: 0.9 }, { path: 'r99.nope', value: '1', confidence: 1 }] }) }, models: { primary: 'p' } });
    assert.strictEqual(partial.cells['r0.tag'].value, 'X');
    assert.ok(!('r99.nope' in partial.cells) && partial.cells['r5.length'].value === '' && partial.cells['r5.length'].flags.length === 0);
  });

  await check('edits clear "not sure" but not real problems, and keep what was read', async () => {
    const form = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10 });
    const cells = { 'r0.resistanceWithGrid': { value: '0.5Z', selected: [], confidence: 0.4, illegible: false, flags: [{ code: 'LOW_CONFIDENCE', text: '' }, { code: 'NON_NUMERIC', text: '' }], source: 'primary' }, 'r0.length': { value: '3', selected: [], confidence: 0.99, illegible: false, flags: [], source: 'primary' } };
    const edited = applyEdits(form, cells, { 'r0.resistanceWithGrid': { value: '0,52' }, 'r0.length': { value: '-4' } });
    assert.deepStrictEqual([edited['r0.resistanceWithGrid'].value, edited['r0.resistanceWithGrid'].flags, edited['r0.resistanceWithGrid'].edited, edited['r0.resistanceWithGrid'].was.value], ['0.52', [], true, '0.5Z']);
    assert.deepStrictEqual(flagCodes(edited['r0.length']), ['OUT_OF_RANGE']);
    const again = applyEdits(form, edited, { 'r0.resistanceWithGrid': { value: '0.53' } });
    assert.strictEqual(again['r0.resistanceWithGrid'].was.value, '0.5Z', 'the original reading is kept across several edits');
  });

  // ---- merging into a daily report ------------------------------------------------------------------
  const readAll = async (sheet) => (await readPage({ form: sheet.form, png: Buffer.alloc(0), provider: createMockProvider({ truth: { [sheet.issued.sheetId || '_']: { primary: sheet.truth }, [sheet.form.sheetKey]: { primary: sheet.truth } }, primaryModel: MODELS.primary }), models: { primary: MODELS.primary } })).cells;
  const BASE = () => ({ project: { ...PROJECT, reportDate: '2026-10-12', workOrder: 'W', engineerName: 'E' }, tests: {}, sheetHeaders: {}, soilResistivity: { locations: [] } });

  await check('flat sheet: rows are written at their printed position with sheetRow; a blank cell never overwrites; a second scan overwrites the same row and reports what it replaced', async () => {
    const { store } = setup();
    const sheet = await issueFilled(store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', instrument: 'mi3290', serialNo: 'SN1', calibrationDue: '14/03/2027', timeStart: '09:10', timeFinish: '11:40', job: { clientRepresentative: 'R. Mehta', groundCondition: 'Dry' }, rows: ELECTRODE_ROWS });
    const first = applyScan({ report: BASE(), form: sheet.form, cells: await readAll(sheet) });
    assert.deepStrictEqual(first.overwrites, []);
    const rows = first.candidate.electrodeResistance;
    assert.deepStrictEqual(rows.map((row) => [row.sheetRow, row.tag, row.resistanceWithGrid]), [[0, 'P-1', '0.52'], [1, 'P-2', '1.20']]);
    assert.deepStrictEqual(rows[0].conditionCodes, ['C1', 'C6']);
    assert.deepStrictEqual(first.candidate.sheetHeaders.electrodeResistance, { instrument: 'mi3290', serialNo: 'SN1', calibrationDue: '2027-03-14', timeStart: '09:10', timeFinish: '11:40' });
    assert.deepStrictEqual([first.candidate.project.clientRepresentative, first.candidate.project.groundCondition, first.candidate.tests.electrodeResistance], ['R. Mehta', 'Dry', true]);
    // it goes through the real normaliser, validation included
    const normalized = normalizeReportInput(first.candidate);
    assert.deepStrictEqual(normalized.electrodeResistance.map((row) => [row.sheetRow, row.tag]), [[0, 'P-1'], [1, 'P-2']]);
    assert.strictEqual(normalized.sheetHeaders.electrodeResistance.instrument, 'mi3290');

    // a page with rowOffset 10 inserts after; rows 3 and 11 sit between and after in order; existing row 0 is overwritten, blank cells keep their value
    const later = { ...BASE(), ...first.candidate, electrodeResistance: [...rows, { sheetRow: 11, tag: 'P-12', resistanceWithGrid: '0.9' }] };
    const form2 = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10, rowOffset: 0, preRows: [] });
    const cells2 = await readAll({ ...sheet, form: form2, issued: sheet.issued, truth: { 'r0.resistanceWithGrid': { value: '0.55' }, 'r3.tag': { value: 'P-4' }, 'r3.resistanceWithGrid': { value: '2.0' } } });
    const merged = applyScan({ report: later, form: form2, cells: cells2 });
    assert.deepStrictEqual(merged.candidate.electrodeResistance.map((row) => [row.sheetRow, row.tag]), [[0, 'P-1'], [1, 'P-2'], [3, 'P-4'], [11, 'P-12']]);
    assert.strictEqual(merged.candidate.electrodeResistance[0].resistanceWithGrid, '0.55');
    assert.strictEqual(merged.candidate.electrodeResistance[0].location, 'Yard', 'a blank cell does not erase what is there');
    assert.deepStrictEqual(merged.overwrites.map((entry) => [entry.key, entry.before, entry.after]), [['resistanceWithGrid', '0.52', '0.55']]);
    assert.deepStrictEqual(merged.previous.map((entry) => entry.row.resistanceWithGrid), ['0.52'], 'the previous row is kept');
    // a second page of the same sheet starts at its rowOffset
    const page2 = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10, rowOffset: 10 });
    const cells3 = await readAll({ form: page2, issued: { sheetId: '' }, truth: { 'r0.tag': { value: 'P-11' }, 'r0.resistanceWithGrid': { value: '3.1' } } });
    const third = applyScan({ report: merged.candidate, form: page2, cells: cells3 });
    assert.deepStrictEqual(third.candidate.electrodeResistance.map((row) => row.sheetRow), [0, 1, 3, 10, 11]);
  });

  await check('hand-typed rows count by their position; a scan row lands on top of the one at that position', async () => {
    const hand = { ...BASE(), tests: { electrodeResistance: true }, electrodeResistance: [{ rowId: 'a', tag: 'H-1', resistanceWithGrid: '1.0' }, { rowId: 'b', tag: 'H-2', resistanceWithGrid: '2.0' }] };
    const form = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10 });
    const cells = await readAll({ form, issued: { sheetId: '' }, truth: { 'r1.resistanceWithGrid': { value: '2.5' }, 'r4.tag': { value: 'S-5' }, 'r4.resistanceWithGrid': { value: '0.7' } } });
    const result = applyScan({ report: hand, form, cells });
    assert.deepStrictEqual(result.candidate.electrodeResistance.map((row) => [row.rowId || null, row.tag, row.resistanceWithGrid]), [['a', 'H-1', '1.0'], ['b', 'H-2', '2.5'], [null, 'S-5', '0.7']]);
    assert.deepStrictEqual(result.overwrites.map((entry) => entry.where), ['MS-02 row 2']);
  });

  await check('pre-printed tags are written with their row, but a row holding only pre-printed text is not created', async () => {
    const form = buildForm({ sheetKey: 'electrodeResistance', rowCount: 10, preRows: [{ tag: 'P-1', location: 'Yard' }, { tag: 'P-2', location: 'Gate' }] });
    const cells = await readAll({ form, issued: { sheetId: '' }, truth: { 'r1.resistanceWithGrid': { value: '1.1' } } });
    const result = applyScan({ report: BASE(), form, cells });
    assert.deepStrictEqual(result.candidate.electrodeResistance.map((row) => [row.sheetRow, row.tag, row.location, row.resistanceWithGrid]), [[1, 'P-2', 'Gate', '1.1']]);
  });

  await check('phase groups: the merged cells are copied onto all three phase rows, each phase keeps its own reading', async () => {
    const { store } = setup();
    const sheet = await issueFilled(store, { sheetKey: 'loopImpedanceTest', project: PROJECT, workDate: '2026-10-12', instrument: 'kyoritsu4118a', rows: [{ location: 'MCC-1', feederTag: 'F-12', deviceType: 'MCCB', deviceRating: '63', tripCurve: 'C', loopImpedance: '0.31', voltage: '231' }, { loopImpedance: '0.33', voltage: '230' }, { loopImpedance: '0.35' }] });
    const result = applyScan({ report: BASE(), form: sheet.form, cells: await readAll(sheet) });
    const rows = result.candidate.loopImpedanceTest;
    assert.deepStrictEqual(rows.map((row) => [row.sheetRow, row.measuredPoints, row.location, row.feederTag, row.deviceType, row.deviceRating, row.loopImpedance]), [[0, 'R-E', 'MCC-1', 'F-12', 'MCCB', '63', '0.31'], [1, 'Y-E', 'MCC-1', 'F-12', 'MCCB', '63', '0.33'], [2, 'B-E', 'MCC-1', 'F-12', 'MCCB', '63', '0.35']]);
    const normalized = normalizeReportInput(result.candidate);
    assert.strictEqual(normalized.loopImpedanceTest.length, 3);
    assert.deepStrictEqual(normalized.loopImpedanceTest.map((row) => row.srNo), ['1', '2', '3']);
  });

  await check('tower groups: a group of four feet is one item at the printed row of its first foot', async () => {
    const { store } = setup();
    const feet = (statusAndReadings) => statusAndReadings.map(([status, current, impedance]) => ({ footToEarthingConnectionStatus: status, measuredCurrentMa: current, measuredImpedance: impedance }));
    const rows = [{ mainLocationTower: 'T-12', ...feet([['Given', '12', '4.1']])[0] }, ...feet([['Given', '11', '4.3'], ['Not given', '', ''], ['Given', '13', '3.9']])];
    const sheet = await issueFilled(store, { sheetKey: 'towerFootingResistance', project: PROJECT, workDate: '2026-10-12', header: { lineName: '220 kV Dahej' }, rows });
    const cells = await readAll(sheet);
    const result = applyScan({ report: BASE(), form: sheet.form, cells });
    const [group] = result.candidate.towerFootingResistance;
    assert.deepStrictEqual([group.sheetRow, group.mainLocationTower, group.readings.map((reading) => reading.measurementPointLocation), group.readings.map((reading) => reading.measuredImpedance || '')], [0, 'T-12', ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'], ['4.1', '4.3', '', '3.9']]);
    assert.strictEqual(group.readings[2].footToEarthingConnectionStatus, 'Not given');
    assert.strictEqual(result.candidate.sheetHeaders.towerFootingResistance.lineName, '220 kV Dahej');
    const normalized = normalizeReportInput(result.candidate);
    assert.strictEqual(normalized.towerFootingResistance[0].sheetRow, 0);
    assert.strictEqual(normalized.towerFootingResistance[0].readings[1].measuredCurrentMa, '11');
    // the second tower of the page sits 4 printed rows later, and a hand-typed first group counts as position 0
    const twice = applyScan({ report: { ...BASE(), towerFootingResistance: [{ mainLocationTower: 'T-1', readings: [] }] }, form: sheet.form, cells });
    assert.strictEqual(twice.overwrites.length, 1);
    assert.strictEqual(twice.overwrites[0].key, 'mainLocationTower');
  });

  await check('soil: one location per page from its header, readings at the printed spacings, remark on direction 1; a re-scan of the page overwrites the location', async () => {
    const { store } = setup();
    const readings = Array.from({ length: 10 }, (_, k) => ({ direction1: String(40 + k), direction2: String(41 + k), direction3: k < 8 ? String(39 + k) : '', direction4: k < 8 ? String(42 + k) : '', remark: k === 0 ? 'wet' : '' }));
    const sheet = await issueFilled(store, { sheetKey: 'soilResistivity', project: PROJECT, workDate: '2026-10-12', header: { name: 'Switchyard A', inputUnit: 'ohm-m', probeDepthCm: '12', direction1: '10', direction2: '190', direction3: '100', direction4: '280', weather: 'Sunny' }, rows: readings });
    const cells = await readAll(sheet);
    const result = applyScan({ report: BASE(), form: sheet.form, cells });
    const [location] = result.candidate.soilResistivity.locations;
    assert.deepStrictEqual([location.sheetRow, location.name, location.inputUnit, location.probeDepthCm, location.weather, location.bearings.direction3], [0, 'Switchyard A', 'ohm-m', '12', 'Sunny', '100']);
    assert.deepStrictEqual([location.direction1.length, location.direction3.length, location.direction1[0].spacing, location.direction1[0].resistivity, location.direction1[0].remark, location.direction1[9].spacing], [10, 8, '0.5', '40', 'wet', '5.0']);
    const normalized = normalizeReportInput({ ...result.candidate, tests: { soilResistivity: true } });
    assert.strictEqual(normalized.soilResistivity.locations[0].sheetRow, 0);
    const again = applyScan({ report: { ...BASE(), soilResistivity: { locations: normalized.soilResistivity.locations } }, form: sheet.form, cells: await readAll({ form: sheet.form, issued: { sheetId: '' }, truth: { 'r0.direction1': { value: '45' } } }) });
    assert.deepStrictEqual(again.overwrites.map((entry) => [entry.key, entry.before, entry.after]), [['direction1', '40', '45']]);
    assert.strictEqual(again.candidate.soilResistivity.locations.length, 1, 'same page: same location');
    const page2 = buildForm({ sheetKey: 'soilResistivity', rowCount: 10, rowOffset: 1 });
    const other = applyScan({ report: { ...BASE(), soilResistivity: { locations: normalized.soilResistivity.locations } }, form: page2, cells: await readAll({ form: page2, issued: { sheetId: '' }, truth: { 'h.name': { value: 'Switchyard B' }, 'r0.direction1': { value: '50' }, 'r0.direction2': { value: '51' } } }) });
    assert.deepStrictEqual(other.candidate.soilResistivity.locations.map((entry) => entry.name), ['Switchyard A', 'Switchyard B']);
  });

  // ---- the whole flow, with the real store ---------------------------------------------------------------
  const flow = (extra = {}) => {
    const env = setup();
    const truth = {};
    const provider = createMockProvider({ truth, primaryModel: MODELS.primary });
    const guard = createAiGuard({ store: env.store });
    const scansDir = path.join(dir, `scans-${Math.random().toString(36).slice(2)}`);
    const service = createScanService({ store: env.store, config: { dir: scansDir, minConfidence: 0.8, maxPages: 5, ...extra }, provider, guard, models: MODELS });
    return { ...env, truth, provider, guard, service, dir: scansDir };
  };

  await check('upload -> read -> review -> accept: the scan is identified by its QR, flagged cells are shown, nothing is in a report until Accept, and Accept creates the day, credited to the person who accepts', async () => {
    const f = flow();
    const sheet = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', instrument: 'mi3290', rows: ELECTRODE_ROWS });
    f.truth[sheet.issued.sheetId] = { primary: { ...sheet.truth, 'r0.resistanceWithGrid': { value: '0.5?', confidence: 0.6 } }, escalation: sheet.truth };
    const [queued] = await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', fileName: 'visit.pdf', user: f.anup });
    assert.strictEqual(queued.status, 'queued');
    await f.service.idle();
    const detail = f.service.detail(queued.id, f.anup);
    assert.deepStrictEqual([detail.status, detail.sheetId, detail.sheetKey, detail.workDate, detail.projectNo, detail.escalated], ['read', sheet.issued.sheetId, 'electrodeResistance', '2026-10-12', PROJECT.projectNo, true]);
    assert.match(detail.identified, /^QR/);
    assert.deepStrictEqual(detail.models, { primary: MODELS.primary, escalation: MODELS.escalation });
    assert.strictEqual(detail.cells['r0.resistanceWithGrid'].value, '0.52');
    assert.deepStrictEqual(flagCodes(detail.cells['r0.resistanceWithGrid']), ['READS_DIFFER']);
    assert.strictEqual(detail.form.cells.length, buildForm({ ...sheet.issued }).cells.length);
    assert.deepStrictEqual(f.store.listReports(''), [], 'nothing reaches a report before Accept');
    assert.strictEqual(f.service.image(queued.id, f.anup).slice(0, 4).toString('hex'), '89504e47');
    // the engineer corrects a flagged cell and accepts
    const edited = f.service.edit(queued.id, f.anup, { 'r0.resistanceWithGrid': { value: '0.53' } });
    assert.deepStrictEqual(edited.cells['r0.resistanceWithGrid'].was, { value: '0.52', selected: [] });
    const result = f.service.accept(queued.id, f.anup, {});
    assert.deepStrictEqual([result.createdDay, result.written > 0, result.overwrites], [true, true, 0]);
    const reports = f.store.listReports('');
    assert.strictEqual(reports.length, 1);
    const day = reports[0];
    assert.deepStrictEqual([day.kind, day.collectionDate, day.projectId, day.reviewStatus, day.createdBy], ['daily', '2026-10-12', f.workspace.id, 'draft', f.anup.id]);
    assert.deepStrictEqual(day.electrodeResistance.map((row) => [row.sheetRow, row.tag, row.resistanceWithGrid, row.enteredBy.email]), [[0, 'P-1', '0.53', 'anup@elegrow.com'], [1, 'P-2', '1.20', 'anup@elegrow.com']]);
    assert.strictEqual(day.sheetHeaders.electrodeResistance.instrument, 'mi3290');
    const after = f.store.getScan(queued.id);
    assert.deepStrictEqual([after.status, after.acceptedReportId, after.createdDay], ['accepted', day.id, true]);
    assert.throws(() => f.service.accept(queued.id, f.anup, {}), (error) => error instanceof ScanError && error.status === 409);
    assert.throws(() => f.service.edit(queued.id, f.anup, {}), (error) => error.code === 'NOT_READ');
  });

  await check('a second scan of the same page overwrites the same rows only after confirmation, the previous values are kept in the scan record, and the first enterer keeps the credit', async () => {
    const f = flow();
    const first = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    f.truth[first.issued.sheetId] = { primary: first.truth };
    const [one] = await f.service.ingest({ buffer: first.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    f.service.accept(one.id, f.anup, {});
    // Nidhi cannot add to Anup's day; an admin can
    f.truth[first.issued.sheetId] = { primary: { ...first.truth, 'r0.resistanceWithGrid': { value: '0.47' }, 'r1.tag': { value: 'P-2B' } } };
    const [two] = await f.service.ingest({ buffer: first.pdf, mimetype: 'application/pdf', user: f.nidhi });
    await f.service.idle();
    assert.throws(() => f.service.accept(two.id, f.nidhi, { confirm: true }), (error) => error.status === 403 && error.code === 'NOT_OWNER');
    const [three] = await f.service.ingest({ buffer: first.pdf, mimetype: 'application/pdf', user: f.admin });
    await f.service.idle();
    let asked;
    try { f.service.accept(three.id, f.admin, {}); } catch (error) { asked = error; }
    assert.deepStrictEqual([asked.code, asked.status, asked.extra.overwrites.map((entry) => [entry.key, entry.before, entry.after])], ['NEEDS_CONFIRM', 409, [['resistanceWithGrid', '0.52', '0.47'], ['tag', 'P-2', 'P-2B']]]);
    assert.strictEqual(f.store.listReports('')[0].electrodeResistance[0].resistanceWithGrid, '0.52', 'nothing was written without the confirmation');
    const done = f.service.accept(three.id, f.admin, { confirm: true });
    assert.strictEqual(done.overwrites, 2);
    const day = f.store.listReports('')[0];
    assert.deepStrictEqual(day.electrodeResistance.map((row) => [row.tag, row.resistanceWithGrid, row.enteredBy.email]), [['P-1', '0.47', 'anup@elegrow.com'], ['P-2B', '1.20', 'anup@elegrow.com']]);
    assert.strictEqual(day.electrodeResistance[0].lastEditedBy.email, 'admin@elegrow.com');
    const record = f.store.getScan(three.id);
    assert.deepStrictEqual(record.previous.map((entry) => entry.row.resistanceWithGrid), ['0.52', '1.20']);
    assert.strictEqual(record.overwrites.length, 2);
    assert.strictEqual(f.store.listReports('').length, 1, 'one day per date');
  });

  await check('accept rules: a day that is no longer a draft, a project with no workspace, and a sheet with nothing on it are refused', async () => {
    const f = flow();
    const sheet = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    f.truth[sheet.issued.sheetId] = { primary: sheet.truth };
    const [one] = await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    f.service.accept(one.id, f.anup, {});
    f.store.transition(f.store.listReports('')[0].id, 'in_review', { actor: 'P' });
    f.truth[sheet.issued.sheetId] = { primary: { ...sheet.truth, 'r0.tag': { value: 'ZZ' } } };
    const [two] = await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    assert.throws(() => f.service.accept(two.id, f.anup, { confirm: true }), (error) => error.code === 'NOT_DRAFT');
    // no workspace
    const stranger = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: { projectNo: '999-NOPE', clientName: 'X' }, workDate: '2026-10-13', rows: ELECTRODE_ROWS });
    f.truth[stranger.issued.sheetId] = { primary: stranger.truth };
    const [three] = await f.service.ingest({ buffer: stranger.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    assert.throws(() => f.service.accept(three.id, f.anup, {}), (error) => error.code === 'NO_WORKSPACE');
    // blank sheet
    const blank = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-14' });
    const [four] = await f.service.ingest({ buffer: blank.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    assert.throws(() => f.service.accept(four.id, f.anup, {}), (error) => error.code === 'NOTHING_TO_ACCEPT');
    // a reading the validation refuses is reported, nothing is written
    const bad = await issueFilled(f.store, { sheetKey: 'continuityTest', project: PROJECT, workDate: '2026-10-15', rows: [{ resistance: '0.1' }] });
    f.truth[bad.issued.sheetId] = { primary: { ...bad.truth, 'r0.resistance': { value: '-3', confidence: 0.99 } } };
    const [five] = await f.service.ingest({ buffer: bad.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    let refused;
    try { f.service.accept(five.id, f.anup, {}); } catch (error) { refused = error; }
    assert.ok(refused && refused.code === 'INVALID_REPORT' && refused.status === 422, refused && refused.message);
    assert.strictEqual(f.store.listReports('').length, 1);
  });

  await check('a page whose QR will not scan is identified by the printed sheet id the provider reads; an id this app never issued, or no id at all, waits for the engineer, whose choice then drives the read', async () => {
    const f = flow();
    const sheet = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    const [page] = await splitUpload({ buffer: sheet.pdf, mimetype: 'application/pdf' });
    const smudged = withoutQr(page);
    assert.strictEqual(findSheetId((await splitUpload({ buffer: smudged, mimetype: 'image/png' }))[0].rgba), null, 'the test page really has no readable code');
    // the provider reads the printed id
    f.truth.__fallbackIds = [sheet.issued.sheetId];
    f.truth[sheet.issued.sheetId] = { primary: sheet.truth };
    const [viaId] = await f.service.ingest({ buffer: smudged, mimetype: 'image/png', user: f.anup });
    await f.service.idle();
    const identified = f.service.detail(viaId.id, f.anup);
    assert.deepStrictEqual([identified.status, identified.identified, identified.sheetId], ['read', 'printed sheet id', sheet.issued.sheetId]);
    // an id nobody issued
    f.truth.__fallbackIds = [newSheetId()];
    const [unknown] = await f.service.ingest({ buffer: smudged, mimetype: 'image/png', user: f.anup });
    await f.service.idle();
    const waiting = f.service.detail(unknown.id, f.anup);
    assert.strictEqual(waiting.status, 'needs_sheet');
    assert.match(waiting.note, /was not issued by this app/);
    // an id that is not even well formed
    f.truth.__fallbackIds = ['garbage'];
    const [garbled] = await f.service.ingest({ buffer: smudged, mimetype: 'image/png', user: f.anup });
    await f.service.idle();
    assert.match(f.service.detail(garbled.id, f.anup).note, /QR code could not be read/);
    // the engineer says which sheet it is
    assert.throws(() => f.service.assign(unknown.id, f.anup, { project: f.workspace.id, workDate: 'x', sheetKey: 'electrodeResistance' }), (error) => error.code === 'BAD_DATE');
    assert.throws(() => f.service.assign(unknown.id, f.anup, { project: 'nope', workDate: '2026-10-12', sheetKey: 'electrodeResistance' }), (error) => error.code === 'BAD_PROJECT');
    assert.throws(() => f.service.assign(unknown.id, f.anup, { project: f.workspace.id, workDate: '2026-10-12', sheetKey: 'nope' }), (error) => error.code === 'BAD_SHEET');
    f.truth.electrodeResistance = { primary: sheet.truth };
    f.service.assign(unknown.id, f.anup, { project: f.workspace.id, workDate: '2026-10-12', sheetKey: 'electrodeResistance', rowOffset: 10 });
    await f.service.idle();
    const read = f.service.detail(unknown.id, f.anup);
    assert.deepStrictEqual([read.status, read.identified, read.form.rowOffset, read.projectNo], ['read', 'chosen by the engineer', 10, PROJECT.projectNo]);
    f.service.accept(unknown.id, f.anup, {});
    assert.deepStrictEqual(f.store.listReports('')[0].electrodeResistance.map((row) => row.sheetRow), [10, 11], 'the chosen row offset places the rows');
    assert.throws(() => f.service.assign(viaId.id, f.anup, { project: f.workspace.id, workDate: '2026-10-12', sheetKey: 'electrodeResistance' }), (error) => error.code === 'NOT_ASSIGNABLE');
  });

  await check('a provider failure leaves the scan failed with the reason and Retry reads it; after repeated quota failures the provider is paused instead of hammered', async () => {
    const f = flow();
    const sheet = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    f.truth[sheet.issued.sheetId] = { fail: 'QUOTA' };
    const scans = [];
    for (let i = 0; i < 4; i += 1) scans.push((await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', user: f.anup }))[0]);
    await f.service.idle();
    const states = scans.map((scan) => f.service.detail(scan.id, f.anup));
    assert.ok(states.every((detail) => detail.status === 'failed'));
    assert.deepStrictEqual([states[0].error.code, states[0].error.retryable], ['QUOTA', true]);
    assert.strictEqual(states[3].error.code, 'COOLING', 'the fourth call was not even made');
    assert.strictEqual(f.guard.status('gemini').state, 'cooling');
    const usage = f.store.aiUsageSummary();
    assert.deepStrictEqual(usage.map((row) => [row.provider, row.kind, row.calls, row.ok]), [['gemini', 'ocr', 3, 0]], 'only the calls that were made are counted');
    // the provider is back
    f.guard.reset('gemini');
    f.truth[sheet.issued.sheetId] = { primary: sheet.truth };
    f.service.retry(scans[0].id, f.anup);
    await f.service.idle();
    assert.strictEqual(f.service.detail(scans[0].id, f.anup).status, 'read');
    assert.throws(() => f.service.retry(scans[0].id, f.anup), (error) => error.code === 'NOT_FAILED');
    assert.ok(!JSON.stringify(f.store.getScan(scans[0].id)).includes('API'), 'no key text anywhere in the record');
  });

  await check('visibility and discard: your own scans, or all of them for reviewer and above; only the uploader (or an admin) changes one; an accepted scan stays', async () => {
    const f = flow();
    const sheet = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    f.truth[sheet.issued.sheetId] = { primary: sheet.truth };
    const [scan] = await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    assert.deepStrictEqual([f.service.list(f.anup).length, f.service.list(f.nidhi).length, f.service.list(f.nidhi, { all: true }).length, f.service.list(f.reviewer, { all: true }).length], [1, 0, 0, 1]);
    assert.throws(() => f.service.detail(scan.id, f.nidhi), (error) => error.status === 404);
    assert.ok(f.service.detail(scan.id, f.reviewer).canChange === false);
    assert.throws(() => f.service.edit(scan.id, f.reviewer, {}), (error) => error.code === 'NOT_OWNER');
    assert.throws(() => f.service.discard(scan.id, f.nidhi), (error) => error.status === 404);
    f.service.accept(scan.id, f.anup, {});
    assert.throws(() => f.service.discard(scan.id, f.anup), (error) => error.code === 'ACCEPTED');
    const [other] = await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    assert.strictEqual(f.service.discard(other.id, f.anup).status, 'discarded');
  });

  await check('admin purge deletes the scans of a date range, images included, and nothing else', async () => {
    const f = flow();
    const sheet = await issueFilled(f.store, { sheetKey: 'electrodeResistance', project: PROJECT, workDate: '2026-10-12', rows: ELECTRODE_ROWS });
    f.truth[sheet.issued.sheetId] = { primary: sheet.truth };
    const [scan] = await f.service.ingest({ buffer: sheet.pdf, mimetype: 'application/pdf', user: f.anup });
    await f.service.idle();
    const file = path.join(f.dir, f.store.getScan(scan.id).imagePath);
    assert.ok(fs.existsSync(file));
    const today = new Date().toISOString().slice(0, 10);
    assert.throws(() => f.service.purge({ from: '2026-12-31', to: '2026-01-01' }), (error) => error.code === 'BAD_RANGE');
    assert.deepStrictEqual(f.service.purge({ from: '2001-01-01', to: '2001-12-31' }), { removed: 0, from: '2001-01-01', to: '2001-12-31' });
    assert.ok(f.store.getScan(scan.id));
    const result = f.service.purge({ from: today, to: today });
    assert.strictEqual(result.removed, 1);
    assert.strictEqual(f.store.getScan(scan.id), null);
    assert.ok(!fs.existsSync(file), 'the image is gone');
    assert.strictEqual(f.store.listReports('').length, 0);
  });

  console.log(`${passed} checks passed`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
