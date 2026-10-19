// Review and approval (Session 6C): draft -> in_review -> checked -> approved for daily and final
// reports, the reviewer checklist, acknowledgements, cover initials from the transitions,
// re-issue as R1/R2, and the per-client logo store.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { buildReviewState, revisionRows, initialsOf } = require('../src/review');
const { generateFinalReportPdf } = require('../src/finalReport/pdf');
const { parseLogoDataUrl, resolveClientLogo, clientKeys, LogoError } = require('../src/finalReport/logos');
const { openPdf, pageTextLines } = require('../src/pdfRaster');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'review-'));
const store = createReportStore({ filePath: path.join(dir, 'review.db') });
const project = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'Gujarat Fluorochemicals Ltd', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' });

function sampleDaily(date, extra = {}) {
  const payload = buildPayload();
  const normalized = normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: date, clientId: 'Z-77' }, ...extra });
  return store.addReport({ ...normalized, kind: 'daily', projectId: project.id, collectionDate: date });
}

function tickAll(id) {
  const ticks = Object.fromEntries(buildReviewState(store, store.getReport(id)).checklist.filter((item) => item.manual).map((item) => [item.id, true]));
  store.setReviewTicks(id, ticks);
}

const daily = sampleDaily('2026-09-14');
const state = (id) => buildReviewState(store, store.getReport(id));

(async () => {
  await check('initials come from names', () => {
    assert.strictEqual(initialsOf('H. S. Patil'), 'HSP');
    assert.strictEqual(initialsOf('A. P.'), 'AP');
    assert.strictEqual(initialsOf('JCS'), 'JCS');
    assert.strictEqual(initialsOf(''), '');
  });

  await check('daily review: submit, tick, mark as checked, approve; each step is its own recorded transition', () => {
    assert.deepStrictEqual(state(daily.id).available, { submit: true, check: false, approve: false, return: false, reissue: false });
    store.transition(daily.id, 'in_review', { actor: 'H. S. Patil' });
    assert.deepStrictEqual(state(daily.id).available, { submit: false, check: true, approve: false, return: true, reissue: false });
    assert.throws(() => store.transition(daily.id, 'approved', { actor: 'J. C. Shah' }), (error) => error.code === 'NOT_CHECKED');
    assert.throws(() => store.markChecked(daily.id, { actor: 'A. P.' }), (error) => error.code === 'CHECKLIST_INCOMPLETE');
    tickAll(daily.id);
    store.markChecked(daily.id, { actor: 'A. P.', note: 'Counts agree' });
    const checked = state(daily.id);
    assert.deepStrictEqual([checked.checked.by, checked.available.approve, checked.available.check], ['A. P.', true, false]);
    const { report, warnings } = store.transition(daily.id, 'approved', { actor: 'J. C. Shah' });
    assert.strictEqual(report.reviewStatus, 'approved');
    assert.deepStrictEqual(warnings, []);
    assert.deepStrictEqual(store.listTransitions(daily.id).map((row) => row.action), ['submit', 'check', 'approve']);
    assert.deepStrictEqual(state(daily.id).initials, { prepared: 'HSP', checked: 'AP', approved: 'JCS' });
    assert.strictEqual(state(daily.id).checklist.length, 1, 'a daily report has only the counts item');
  });

  await check('approved reports cannot be ticked or checked again; returning clears the checks', () => {
    assert.throws(() => store.setReviewTicks(daily.id, { counts: false }), (error) => error.code === 'INVALID_STATE');
    const d = sampleDaily('2026-09-15');
    store.transition(d.id, 'in_review', { actor: 'P' });
    tickAll(d.id);
    store.markChecked(d.id, { actor: 'C' });
    store.transition(d.id, 'draft', { actor: 'C', note: 'redo' });
    assert.strictEqual(state(d.id).checklist[0].ticked, false);
    store.transition(d.id, 'in_review', { actor: 'P' });
    assert.strictEqual(state(d.id).checked, null, 'a new submit needs a new check');
    assert.throws(() => store.transition(d.id, 'approved', { actor: 'A' }), (error) => error.code === 'NOT_CHECKED');
  });

  await check('same person checking and approving is allowed but warned about', () => {
    const d = sampleDaily('2026-09-16');
    store.transition(d.id, 'in_review', { actor: 'P' });
    tickAll(d.id);
    store.markChecked(d.id, { actor: 'A. P.' });
    const { report, warnings } = store.transition(d.id, 'approved', { actor: 'a. p.' });
    assert.strictEqual(report.reviewStatus, 'approved');
    assert.strictEqual(warnings.length, 1);
    assert.ok(/both checked and approved/.test(warnings[0]));
  });

  const { report: final } = store.assembleFinal(project.id, { actor: 'lead' });

  await check('final checklist: standard 6.4 items, sanity check not run, unit code and planned-quantity warnings', () => {
    const review = state(final.id);
    assert.deepStrictEqual(review.checklist.map((item) => item.id), ['counts', 'p1', 'assumptions', 'lastyear', 'otherunit']);
    assert.strictEqual(review.checklist[0].auto, 'pass');
    assert.strictEqual(review.checklist[1].auto, 'warn', 'P1 pits have no photo reference in the sample');
    assert.strictEqual(review.sanityCheck, 'not run');
    assert.deepStrictEqual(review.acknowledgements.map((item) => item.key), ['p1Photos', 'sanityCheck']);
    assert.ok(review.warnings.some((warning) => /Planned quantity not set/.test(warning)));
    assert.strictEqual(review.consistency.ok, true);
    assert.strictEqual(review.documentNo, '024_FI_ESHA');
  });

  await check('missing unit code raises a checklist warning and the SITE fallback', () => {
    const payload = buildPayload();
    const d = store.addReport({ ...normalizeReportInput(payload), kind: 'daily', projectId: project.id, collectionDate: '2026-09-17' });
    const review = state(d.id);
    assert.ok(review.warnings.some((warning) => /Unit code not set/.test(warning)));
    assert.strictEqual(review.documentNo, '024_SITE_ESHA');
  });

  await check('final approval needs the check, then the sanity-check acknowledgement; cover initials and revision row follow', async () => {
    store.transition(final.id, 'in_review', { actor: 'H. S. Patil' });
    tickAll(final.id);
    store.markChecked(final.id, { actor: 'A. P.' });
    assert.throws(() => store.transition(final.id, 'approved', { actor: 'J. C. Shah' }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED' && /Sanity check not run/.test(error.message) && /No photo reference for P1 pits EP-03, EP-07, EP-08/.test(error.message));
    const approved = store.transition(final.id, 'approved', { actor: 'J. C. Shah', designation: 'Head - Earthing', acknowledgements: { sanityCheck: 'Provider unavailable; narrative read against the tables.', p1Photos: 'Photos taken on site; to follow with Annexure B.' } }).report;
    assert.strictEqual(approved.reviewStatus, 'approved');
    assert.ok(approved.approvedAt);
    const note = store.listTransitions(final.id).pop().note;
    assert.ok(note.includes('Sanity check not run - acknowledged: Provider unavailable') && note.includes('P1 photo missing - acknowledged: Photos taken on site'));
    assert.strictEqual(store.listTransitions(final.id).pop().designation, 'Head - Earthing');
    const rows = revisionRows(store, store.getReport(final.id));
    assert.deepStrictEqual(rows.map((row) => [row.revision, row.issue, row.prepared, row.checked, row.approved]), [['R0', 'First Issue', 'HSP', 'AP', 'JCS']]);
    assert.strictEqual(rows[0].date, approved.approvedAt.slice(0, 10));
  });

  let r1;
  await check('re-issue after approval creates R1 as a new draft; the approved R0 stays; the cover lists both', async () => {
    assert.strictEqual(store.reissueReport('missing', {}), null);
    const d = sampleDaily('2026-09-18');
    assert.throws(() => store.reissueReport(d.id, { actor: 'x' }), (error) => error.code === 'INVALID_STATE');
    r1 = store.reissueReport(final.id, { actor: 'H. S. Patil', reason: 'Corrected pit EP-07 location' });
    assert.strictEqual(r1.reviewStatus, 'draft');
    assert.strictEqual(r1.revision, 1);
    assert.strictEqual(r1.revisionOf, final.id);
    assert.strictEqual(r1.approvedAt, undefined);
    assert.strictEqual(store.getReport(final.id).reviewStatus, 'approved');
    assert.throws(() => store.reissueReport(final.id, { actor: 'x' }), (error) => error.code === 'ALREADY_REISSUED');
    const rows = revisionRows(store, r1);
    assert.deepStrictEqual(rows.map((row) => [row.revision, row.issue]), [['R0', 'First Issue'], ['R1', 'Corrected pit EP-07 location']]);
    assert.deepStrictEqual([rows[1].prepared, rows[1].checked, rows[1].approved], ['', '', '']);
    // R0 keeps the initials of its own review on the R1 cover
    assert.deepStrictEqual([rows[0].prepared, rows[0].checked, rows[0].approved], ['HSP', 'AP', 'JCS']);
    const r1Cover = pageTextLines(await openPdf((await generateFinalReportPdf(r1, { revisions: rows, status: 'draft' })).buffer), 0).map((fragment) => fragment.text).join(' ');
    assert.ok(/R0 \S+ First Issue HSP AP JCS/.test(r1Cover.replace(/\s+/g, ' ')), r1Cover);

    // R1 goes through its own review; R2 follows from R1.
    store.transition(r1.id, 'in_review', { actor: 'H. S. Patil' });
    tickAll(r1.id);
    store.markChecked(r1.id, { actor: 'A. P.' });
    store.transition(r1.id, 'approved', { actor: 'J. C. Shah', acknowledgements: { sanityCheck: 'Read by hand.', p1Photos: 'Photos follow.' } });
    const r2 = store.reissueReport(r1.id, { actor: 'H. S. Patil', reason: 'Client comments' });
    assert.strictEqual(r2.revision, 2);
    const pdf = await generateFinalReportPdf(r2, { revisions: revisionRows(store, r2), status: r2.reviewStatus });
    const doc = await openPdf(pdf.buffer);
    const cover = pageTextLines(doc, 0).map((fragment) => fragment.text).join(' ');
    ['R0', 'R1', 'R2', 'First Issue', 'Corrected pit EP-07 location', 'Client comments', 'HSP', 'JCS'].forEach((text) => assert.ok(cover.includes(text), `cover lacks ${text}`));
    const header = pageTextLines(doc, 2).map((fragment) => fragment.text).join(' ');
    assert.ok(header.includes('024_FI_ESHA | R2'));
  });

  await check('a FINAL report with no instrument recorded cannot be approved (a hard block, not an acknowledgement)', () => {
    const bare = store.addProject({ projectNo: '025-D03', clientName: 'No Instruments Ltd' });
    const payload = buildPayload();
    delete payload.sheetHeaders;
    const day = store.addReport({ ...normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI' } }), kind: 'daily', projectId: bare.id, collectionDate: '2026-09-14' });
    store.transition(day.id, 'in_review', { actor: 'H. S. Patil' });
    tickAll(day.id);
    store.markChecked(day.id, { actor: 'A. P.' });
    assert.strictEqual(store.transition(day.id, 'approved', { actor: 'J. C. Shah' }).report.reviewStatus, 'approved', 'a daily is not blocked');
    const { report: noInstruments } = store.assembleFinal(bare.id, {});
    store.transition(noInstruments.id, 'in_review', { actor: 'H. S. Patil' });
    tickAll(noInstruments.id);
    store.markChecked(noInstruments.id, { actor: 'A. P.' });
    const review = state(noInstruments.id);
    assert.deepStrictEqual(review.blockers.map((item) => item.code), ['NO_INSTRUMENTS']);
    const everything = { sanityCheck: 'x', p1Photos: 'x', calibration: 'x' };
    assert.throws(() => store.transition(noInstruments.id, 'approved', { actor: 'J. C. Shah', acknowledgements: everything }), (error) => error.code === 'NO_INSTRUMENTS' && /No instrument is recorded/.test(error.message));
    // recording the instrument clears the block
    store.updateReport(noInstruments.id, (current) => ({ ...current, calibrationRecords: [{ instrument: 'mi3290', serialNo: 'MI3290-1', calibrationDue: '2027-03-31', workDate: '2026-09-14' }] }));
    assert.deepStrictEqual(state(noInstruments.id).blockers, []);
    assert.strictEqual(store.transition(noInstruments.id, 'approved', { actor: 'J. C. Shah', acknowledgements: everything }).report.reviewStatus, 'approved');
  });

  await check('a P1 pit without a photo needs the photo or a typed acknowledgement; with photos nothing is asked', () => {
    const photo = store.addProject({ projectNo: '026-D03', clientName: 'Photos Ltd' });
    const payload = buildPayload();
    payload.electrodeResistance = payload.electrodeResistance.map((row) => ({ ...row, photoRef: ['EP-03', 'EP-07', 'EP-08'].includes(row.tag) ? `IMG_${row.tag}` : '' }));
    store.addReport({ ...normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI' } }), kind: 'daily', projectId: photo.id, collectionDate: '2026-09-14' });
    const { report: withPhotos } = store.assembleFinal(photo.id, {});
    assert.deepStrictEqual(state(withPhotos.id).acknowledgements.map((item) => item.key), ['sanityCheck']);
    assert.strictEqual(state(withPhotos.id).checklist.find((item) => item.id === 'p1').auto, 'pass');
    store.updateReport(withPhotos.id, (current) => ({ ...current, electrodeResistance: current.electrodeResistance.map((row) => (row.tag === 'EP-07' ? { ...row, photoRef: '' } : row)) }));
    const asks = state(withPhotos.id).acknowledgements;
    assert.deepStrictEqual(asks.map((item) => item.key), ['p1Photos', 'sanityCheck']);
    assert.ok(/EP-07/.test(asks[0].text) && !/EP-03/.test(asks[0].text));
    store.transition(withPhotos.id, 'in_review', { actor: 'H. S. Patil' });
    tickAll(withPhotos.id);
    store.markChecked(withPhotos.id, { actor: 'A. P.' });
    assert.throws(() => store.transition(withPhotos.id, 'approved', { actor: 'J. C. Shah', acknowledgements: { sanityCheck: 'x' } }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED' && error.items.map((item) => item.key).join() === 'p1Photos');
    assert.throws(() => store.transition(withPhotos.id, 'approved', { actor: 'J. C. Shah', acknowledgements: { sanityCheck: 'x', p1Photos: '   ' } }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED');
    store.transition(withPhotos.id, 'approved', { actor: 'J. C. Shah', acknowledgements: { sanityCheck: 'x', p1Photos: 'Photo of EP-07 to follow.' } });
    assert.ok(store.listTransitions(withPhotos.id).pop().note.includes('P1 photo missing - acknowledged: Photo of EP-07 to follow.'));
  });

  await check('approval block: full names, designations and dates come from the recorded steps', () => {
    const { approvalBlock } = require('../src/review');
    const block = approvalBlock(store, store.getReport(final.id));
    assert.deepStrictEqual([block.prepared.name, block.checked.name, block.approved.name, block.approved.designation], ['H. S. Patil', 'A. P.', 'J. C. Shah', 'Head - Earthing']);
    assert.ok(/^\d{4}-\d{2}-\d{2}$/.test(block.approved.date));
  });

  await check('transition rows written before the action column existed are classified from their statuses', () => {
    const { transitionAction } = require('../src/review');
    assert.strictEqual(transitionAction({ fromStatus: 'draft', toStatus: 'in_review' }), 'submit');
    assert.strictEqual(transitionAction({ fromStatus: 'in_review', toStatus: 'approved' }), 'approve');
    assert.strictEqual(transitionAction({ fromStatus: 'in_review', toStatus: 'draft' }), 'return');
  });

  await check('client logos: PNG/JPG only, 1 MB max, checked by content; per client with name fallback; report override wins', () => {
    const png = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.alloc(64, 1)]);
    const jpg = Buffer.concat([Buffer.from([0xff, 0xd8, 0xff, 0xe0]), Buffer.alloc(64, 2)]);
    const url = (mime, buffer) => `data:${mime};base64,${buffer.toString('base64')}`;
    assert.strictEqual(parseLogoDataUrl(url('image/png', png)).mime, 'image/png');
    assert.strictEqual(parseLogoDataUrl(url('image/jpeg', jpg)).mime, 'image/jpeg');
    assert.throws(() => parseLogoDataUrl(url('image/gif', png)), LogoError);
    assert.throws(() => parseLogoDataUrl(url('image/png', Buffer.from('not an image at all'))), /not a valid PNG or JPEG/);
    assert.throws(() => parseLogoDataUrl(url('image/png', Buffer.concat([png, Buffer.alloc(1024 * 1024)]))), /limit is 1 MB/);
    assert.throws(() => parseLogoDataUrl('http://example.com/x.png'), LogoError);

    const target = { project: { clientId: 'Z-77', clientName: 'Gujarat Fluorochemicals Ltd' } };
    assert.deepStrictEqual(clientKeys(target.project), ['zoho:Z-77', 'name:gujarat fluorochemicals ltd']);
    assert.strictEqual(resolveClientLogo(store, target), null);
    store.setClientLogo({ key: 'name:gujarat fluorochemicals ltd', clientName: 'Gujarat Fluorochemicals Ltd', mime: 'image/png', data: png });
    assert.ok(resolveClientLogo(store, target).equals(png), 'falls back to the client name');
    store.setClientLogo({ key: 'zoho:Z-77', clientName: 'Gujarat Fluorochemicals Ltd', mime: 'image/jpeg', data: jpg });
    assert.ok(resolveClientLogo(store, target).equals(jpg), 'the Zoho client id is preferred');
    assert.ok(resolveClientLogo(store, { ...target, clientLogoOverride: url('image/png', png) }).equals(png), 'the report override wins');
    assert.ok(resolveClientLogo(store, { project: { clientName: 'GUJARAT   fluorochemicals ltd' } }).equals(png), 'names compare ignoring case and spacing');
    assert.strictEqual(store.deleteClientLogo('zoho:Z-77'), true);
    assert.ok(resolveClientLogo(store, target).equals(png));
  });

  await check('logos live in the database only: nothing under the repository is written', () => {
    const tracked = require('child_process').execSync('git ls-files', { cwd: path.join(__dirname, '..') }).toString();
    assert.ok(!/client[-_]?logo/i.test(tracked.replace(/src\/finalReport\/logos\.js/g, '')));
  });

  console.log(`\n${passed} review checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
