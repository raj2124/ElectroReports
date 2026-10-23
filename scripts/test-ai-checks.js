// Session 6I, the secondary AI checks, on mocks (nothing is sent to OpenAI): the sanity check of a final report's narrative and the
// line it feeds on the review checklist (passed / issues / not run, and the typed acknowledgement before approval), the rule that
// a verifier may not introduce a figure, the remark wording (digit and verdict filter, template fallback), provider failures that
// never block, the guard that pauses a provider, and the HTTP routes.
const assert = require('assert');
const fs = require('fs');
const net = require('net');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');
const { createReportStore } = require('../src/reportStore');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { buildReviewState, sanityState } = require('../src/review');
const { createAiChecks, createMockChecksClient, templateWording, acceptableWording, sanityHashOf } = require('../src/aiChecks');
const { createAiGuard, ProviderError } = require('../src/aiGuard');
const { devEnv, createCaller } = require('./lib/dev-auth');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}
const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ai-checks-'));
const CONFIG = { openai: { apiKey: '', sanityModel: 'sanity-model', observationModel: 'observation-model' }, scan: { aiChecksUseMock: true } };

// a final report with a narrative the way the PDF export stores it
function finalWith(narrativeOverride) {
  const store = createReportStore({ filePath: path.join(dir, `${Math.random().toString(36).slice(2)}.db`) });
  const project = store.addProject({ projectNo: '024-D03_TS-26-GFL', clientName: 'GFL' });
  const payload = buildPayload();
  store.addReport({ ...normalizeReportInput({ ...payload, project: { ...payload.project, unit: 'FI Plant', unitCode: 'FI', reportDate: '2026-10-01' } }), kind: 'daily', projectId: project.id, collectionDate: '2026-10-01' });
  const { report } = store.assembleFinal(project.id, { actor: 'lead' });
  // plain wording with no figure of its own (the template's own rounded soil average would not match the raw results)
  const base = { keyFindings: ['Each test is summarised in the results section.', 'Corrective actions are listed in the plan.', 'Further detail follows in the report.'], conditionStatement: 'The condition of the earthing system is described in the results section.', note: 'test' };
  if (narrativeOverride !== null) {
    store.updateReport(report.id, (current) => ({ ...current, reportNarrative: { ...base, aiWritten: true, ...(narrativeOverride || {}) } }));
  }
  return { store, id: report.id, get: () => store.getReport(report.id) };
}
const tick = (store, id) => store.setReviewTicks(id, Object.fromEntries(buildReviewState(store, store.getReport(id)).checklist.filter((item) => item.manual).map((item) => [item.id, true])));
const checks = (store, client) => createAiChecks({ store, config: CONFIG, guard: createAiGuard({ store }), client });

(async () => {
  await check('wording: a draft is plain wording only: no digit, no verdict word, no length overrun; the template is built from the condition codes alone', () => {
    assert.strictEqual(templateWording(['C1', 'C6']), 'Observed: corrosion is visible and a connection is loose. Inspect and rectify as per the corrective action plan.');
    assert.match(templateWording([]), /flagged by the assessment/);
    assert.match(templateWording(['C1', 'C2', 'C6']), /corrosion is visible, the strip or test link is disconnected and a connection is loose/);
    assert.deepStrictEqual(['Corrosion is visible at the clamp; clean and re-make the joint.', 'The pit reads 12 ohm', 'This pit failed', 'Critical corrosion', 'It is compliant with the standard', 'Resistance is within limits', 'Not acceptable as found', 'x'.repeat(300), ''].map(acceptableWording), [true, false, false, false, false, false, false, false, false]);
    ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'OTH'].forEach((code) => assert.ok(acceptableWording(templateWording([code])), `template for ${code} is itself plain wording`));
  });

  await check('suggested wording: the model\'s plain draft is used; a draft with a number or a verdict is replaced by the standard wording, with a note; unknown codes are ignored', async () => {
    const store = createReportStore({ filePath: path.join(dir, 'w.db') });
    const answer = (text) => ({ mock: true, wording: async () => ({ text, model: 'm' }), sanity: async () => ({}) });
    const good = await checks(store, answer('Corrosion is visible at the connection; clean and re-make the joint.')).suggestWording({ sheetLabel: 'Earth Electrode (Pit) Resistance', conditionCodes: ['C1'] });
    assert.deepStrictEqual([good.source, good.text, good.status.ok, good.model], ['ai', 'Corrosion is visible at the connection; clean and re-make the joint.', true, 'm']);
    for (const bad of ['Pit resistance is 12 ohm and corroded.', 'This pit has failed.', 'The connection is critical.']) {
      const result = await checks(store, answer(bad)).suggestWording({ conditionCodes: ['C1'] });
      assert.deepStrictEqual([result.source, result.text, result.status.ok, /number or a verdict/.test(result.status.note)], ['template', templateWording(['C1']), true, true], bad);
    }
    const codes = await checks(store, createMockChecksClient()).suggestWording({ conditionCodes: ['C6', 'ZZZ', 'C1', 'C6'] });
    assert.strictEqual(codes.text, templateWording(['C6', 'C1']));
  });

  await check('a provider failure never blocks the wording: the standard wording comes back with the reason; repeated failures pause the provider', async () => {
    const store = createReportStore({ filePath: path.join(dir, 'g.db') });
    const guard = createAiGuard({ store });
    const client = createMockChecksClient({ failWith: 'QUOTA' });
    const service = createAiChecks({ store, config: CONFIG, guard, client });
    const first = await service.suggestWording({ conditionCodes: ['C5'] });
    assert.deepStrictEqual([first.source, first.text, first.status.ok, first.status.code], ['template', templateWording(['C5']), false, 'QUOTA']);
    await service.suggestWording({ conditionCodes: ['C5'] });
    await service.suggestWording({ conditionCodes: ['C5'] });
    const paused = await service.suggestWording({ conditionCodes: ['C5'] });
    assert.deepStrictEqual([paused.status.code, /paused after repeated failures/.test(paused.status.reason), guard.status('openai').state], ['COOLING', true, 'cooling']);
    assert.deepStrictEqual(store.aiUsageSummary().map((row) => [row.provider, row.kind, row.calls, row.ok]), [['openai', 'wording', 3, 0]], 'the paused call was not made and not counted');
    assert.ok(!JSON.stringify(paused).match(/sk-|api key/i));
  });

  await check('sanity check: a clean narrative passes and the review line says so; nothing extra to acknowledge', async () => {
    const t = finalWith({});
    const before = buildReviewState(t.store, t.get());
    assert.deepStrictEqual([before.sanityCheck, before.acknowledgements.some((item) => item.key === 'sanityCheck')], ['not run', true]);
    const result = await checks(t.store, createMockChecksClient()).sanityCheck(t.get());
    assert.deepStrictEqual([result.status, result.flags, result.dropped, result.model, Boolean(result.narrativeHash)], ['passed', [], 0, 'mock-sanity', true]);
    const after = buildReviewState(t.store, t.get());
    assert.deepStrictEqual([after.sanityCheck, after.sanity.state, after.acknowledgements.map((item) => item.key).filter((key) => key.startsWith('sanity'))], ['passed', 'passed', []]);
  });

  await check('sanity check: a contradiction (a figure the results do not hold) is shown on the checklist and needs a typed acknowledgement before approval, never a hard block', async () => {
    const t = finalWith({ keyFindings: ['The average soil resistivity is 987.65 ohm-m.', 'All pits are within their limits.', 'Corrective action is listed in the plan.'] });
    const result = await checks(t.store, createMockChecksClient()).sanityCheck(t.get());
    assert.strictEqual(result.status, 'issues');
    assert.match(result.flags[0].issue, /987\.65/);
    const state = buildReviewState(t.store, t.get());
    assert.deepStrictEqual([state.sanityCheck, state.sanity.flags.length], ['issues', 1]);
    const ask = state.acknowledgements.find((item) => item.key === 'sanityIssues');
    assert.ok(ask && ask.required && /987\.65/.test(ask.text) && /read them/.test(ask.text));
    assert.ok(!state.acknowledgements.some((item) => item.key === 'sanityCheck'), 'run, so the "not run" note is not asked for');
    t.store.transition(t.id, 'in_review', { actor: 'P' });
    tick(t.store, t.id);
    t.store.markChecked(t.id, { actor: 'C' });
    const rest = Object.fromEntries(state.acknowledgements.filter((item) => item.key !== 'sanityIssues').map((item) => [item.key, 'noted']));
    assert.throws(() => t.store.transition(t.id, 'approved', { actor: 'A', acknowledgements: rest }), (error) => error.code === 'ACKNOWLEDGEMENT_REQUIRED' && /possible contradiction/.test(error.message));
    const approved = t.store.transition(t.id, 'approved', { actor: 'A', acknowledgements: { ...rest, sanityIssues: 'Read; the 987.65 is a typing slip in the narrative, corrected in the next issue.' } });
    assert.strictEqual(approved.report.reviewStatus, 'approved');
    assert.match(t.store.listTransitions(t.id).pop().note, /Sanity check issues - acknowledged: Read; the 987\.65/);
  });

  await check('a verifier may not introduce a figure: a flag that carries a number found in neither the results nor the narrative is dropped', async () => {
    const t = finalWith({});
    const invented = { mock: true, wording: async () => ({}), sanity: async () => ({ consistent: false, flags: [{ area: 'Soil', severity: 'critical', issue: 'The soil resistivity should be 4321.9 ohm-m.' }], notes: 'x', model: 'm' }) };
    const result = await checks(t.store, invented).sanityCheck(t.get());
    assert.deepStrictEqual([result.status, result.flags.length, result.dropped], ['passed', 0, 1]);
    const quoting = { mock: true, wording: async () => ({}), sanity: async ({ narrative }) => ({ consistent: false, flags: [{ area: 'Findings', severity: 'bogus-level', issue: `The narrative says "${narrative.conditionStatement.slice(0, 40)}" which contradicts the status.` }, { area: '', severity: 'info', issue: '' }], notes: 'n', model: 'm' }) };
    const second = await checks(t.store, quoting).sanityCheck(t.get());
    assert.deepStrictEqual([second.status, second.flags.length, second.flags[0].severity], ['issues', 1, 'warning'], 'empty flags are dropped and an unknown severity becomes a warning');
  });

  await check('sanity check failures never block: no narrative, a provider error, a paused provider: the answer says why, the report is left as it was, "not run" still applies', async () => {
    const none = finalWith(null);
    const noNarrative = await checks(none.store, createMockChecksClient()).sanityCheck(none.get());
    assert.deepStrictEqual([noNarrative.status, noNarrative.code], ['error', 'NO_NARRATIVE']);
    const t = finalWith({});
    for (const code of ['QUOTA', 'AUTH', 'TIMEOUT', 'UPSTREAM']) {
      const result = await checks(t.store, createMockChecksClient({ failWith: code })).sanityCheck(t.get());
      assert.deepStrictEqual([result.status, result.code], ['error', code]);
    }
    assert.strictEqual(t.get().aiSanityCheck, undefined, 'a failed run leaves nothing behind');
    assert.strictEqual(buildReviewState(t.store, t.get()).sanityCheck, 'not run');
    // an earlier good result survives a later failed run
    await checks(t.store, createMockChecksClient()).sanityCheck(t.get());
    await checks(t.store, createMockChecksClient({ failWith: 'QUOTA' })).sanityCheck(t.get());
    assert.strictEqual(buildReviewState(t.store, t.get()).sanityCheck, 'passed');
  });

  await check('a check only counts for the narrative it looked at: a changed narrative, or results that changed, makes it "not run" again; an older stored check is read as it was', async () => {
    const t = finalWith({});
    await checks(t.store, createMockChecksClient()).sanityCheck(t.get());
    assert.strictEqual(sanityState(t.get()).state, 'passed');
    t.store.updateReport(t.id, (current) => ({ ...current, reportNarrative: { ...current.reportNarrative, conditionStatement: `${current.reportNarrative.conditionStatement} Reworded.` } }));
    const stale = sanityState(t.get());
    assert.deepStrictEqual([stale.state, stale.reason], ['not run', 'The narrative changed after the check ran.']);
    assert.ok(buildReviewState(t.store, t.get()).acknowledgements.find((item) => item.key === 'sanityCheck').text.includes('The narrative changed'));
    assert.notStrictEqual(sanityHashOf(t.get()), '');
    // a check kept by the earlier version of the app (no status, no hash)
    assert.strictEqual(sanityState({ aiSanityCheck: { consistent: true, flags: [] }, kind: 'final' }).state, 'passed');
    const legacy = sanityState({ aiSanityCheck: { consistent: false, flags: [{ area: 'a', issue: 'i' }] }, kind: 'final' });
    assert.deepStrictEqual([legacy.state, legacy.flags.length], ['issues', 1]);
    assert.strictEqual(sanityState({}).state, 'not run');
  });

  // ---- HTTP -----------------------------------------------------------------------------------------------
  const port = await freePort();
  const server = spawn(process.execPath, ['src/server.js'], {
    cwd: path.join(__dirname, '..'),
    env: { ...process.env, ...devEnv(), ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'http', 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ZOHO_USE_MOCK: 'true', GRAPH_USE_MOCK: 'true', AI_CHECKS_USE_MOCK: 'true', OCR_USE_MOCK: 'true', OPENAI_API_KEY: '', GEMINI_API_KEY: '', ANTHROPIC_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  try {
    await new Promise((resolve, reject) => {
      server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
      setTimeout(() => reject(new Error('server did not start')), 15000);
    });
    const call = createCaller(`http://127.0.0.1:${port}`);
    const created = await call('POST', '/api/projects', { projectNo: '024-D03_TS-26-GFL', clientName: 'GFL', siteLocation: 'Dahej', workOrder: 'W', engineerName: 'E' }, 'engineer@elegrow.com');
    const payload = buildPayload();
    const day = await call('POST', `/api/projects/${created.body.projectId}/days`, { workDate: '2026-10-12', report: { ...payload, project: { ...payload.project, reportDate: '2026-10-12' } } }, 'engineer@elegrow.com');
    const finalId = (await call('POST', `/api/projects/${created.body.projectId}/final/assemble`, {}, 'engineer@elegrow.com')).body.report.id;

    await check('routes: the sanity check is for a reviewer on a final report that is not approved; the answer carries the new checklist line', async () => {
      assert.strictEqual((await call('POST', `/api/reports/${finalId}/ai/sanity`, {}, 'engineer@elegrow.com')).status, 403);
      assert.strictEqual((await call('POST', `/api/reports/${day.body.id}/ai/sanity`, {}, 'reviewer@elegrow.com')).status, 400, 'a daily report has no narrative to check');
      const run = await call('POST', `/api/reports/${finalId}/ai/sanity`, {}, 'reviewer@elegrow.com');
      assert.strictEqual(run.status, 200, JSON.stringify(run.body).slice(0, 300));
      assert.deepStrictEqual([run.body.result.status, run.body.review.sanityCheck, run.body.review.sanity.state], ['passed', 'passed', 'passed']);
      assert.ok(!run.body.review.acknowledgements.some((item) => item.key.startsWith('sanity')));
      assert.ok((await call('GET', '/api/admin/audit?action=ai.')).body.entries.some((entry) => entry.action === 'ai.sanity'));
      assert.strictEqual((await call('GET', `/api/reports/${finalId}/review`)).body.sanityCheck, 'passed');
    });

    await check('routes: wording suggestions are for engineers; a bad sheet is refused; the answer is plain text and says where it came from', async () => {
      const ok = await call('POST', '/api/observations/suggest', { sheetKey: 'electrodeResistance', pointKind: 'pit', conditionCodes: ['C1', 'C6'] }, 'engineer@elegrow.com');
      assert.strictEqual(ok.status, 200);
      assert.deepStrictEqual([ok.body.source, ok.body.text, ok.body.status.ok], ['ai', templateWording(['C1', 'C6']), true]);
      assert.strictEqual((await call('POST', '/api/observations/suggest', { sheetKey: 'nope' }, 'engineer@elegrow.com')).status, 400);
      assert.strictEqual((await call('POST', '/api/observations/suggest', { sheetKey: 'electrodeResistance' }, 'approver@elegrow.com')).status, 403);
      assert.match((await call('POST', '/api/observations/suggest', { sheetKey: 'electrodeResistance', conditionCodes: ['<script>', 'C5'] }, 'engineer@elegrow.com')).body.text, /the electrode is damaged/);
    });

    await check('routes: the status screen shows the mode, models and pause state of both providers and never a key; admin sees the call counts', async () => {
      const status = (await call('GET', '/api/ai/status')).body;
      assert.deepStrictEqual([status.checks.mock, status.checks.guard.state, status.sheetReader.mock, status.sheetReader.guard.state, status.sheetReader.minConfidence], [true, 'ok', true, 'ok', 0.8]);
      assert.ok(status.checks.sanityModel && status.sheetReader.model);
      assert.ok(!/api.?key|secret|sk-/i.test(JSON.stringify(status).replace(/"apiKeyConfigured"[^,}]*/g, '')), JSON.stringify(status).slice(0, 200));
      const usage = await call('GET', '/api/admin/ai-usage');
      assert.deepStrictEqual(usage.body.usage.map((row) => [row.provider, row.kind]).sort(), [['openai', 'sanity'], ['openai', 'wording']].sort());
      assert.strictEqual((await call('GET', '/api/admin/ai-usage', undefined, 'engineer@elegrow.com')).status, 403);
    });
  } finally {
    server.kill();
  }
  console.log(`${passed} checks passed`);
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
