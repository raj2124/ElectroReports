// Final report and daily progress report PDFs (Session 6C): render the make_report.py sample data
// and assert sections, counts and statuses (same golden numbers as test:esha), the consistency
// block, and the Claude-unavailable narrative fallback.
const assert = require('assert');
const { normalizeReportInput } = require('../src/reportModel');
const { buildPayload } = require('./fixtures/golden-esha');
const { buildAllTestsPayload } = require('./fixtures/golden-all-tests');
const { buildFinalReportData, ConsistencyError } = require('../src/finalReport/data');
const { generateFinalReportPdf } = require('../src/finalReport/pdf');
const { generateDailyProgressPdf } = require('../src/finalReport/dailyPdf');
const { generateNarrative, narrativeInput, unsupportedFigures } = require('../src/finalReport/narrative');
const { openPdf, pageCount, pageTextLines } = require('../src/pdfRaster');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

function sampleReport(extra = {}) {
  const payload = buildPayload();
  payload.project = {
    ...payload.project,
    unit: 'FI Plant',
    unitCode: 'FI',
    siteLocation: '12/A GIDC Industrial Estate, Dahej, Bharuch, Gujarat 392130',
    plannedQuantities: { soilResistivity: 2, electrodeResistance: 10, loopImpedanceTest: 4, prospectiveFaultCurrent: 4, riserIntegrityTest: 6 }
  };
  return normalizeReportInput({ ...payload, ...extra });
}

async function textOf(buffer) {
  const pdf = await openPdf(buffer);
  const pages = [];
  for (let i = 0; i < pageCount(pdf); i += 1) {
    pages.push(pageTextLines(pdf, i).map((fragment) => fragment.text).join('\n'));
  }
  return pages;
}

const goodClient = (output) => ({ messages: { create: async () => ({ stop_reason: 'end_turn', model: 'claude-opus-5-5', content: [{ type: 'text', text: JSON.stringify(output) }] }) } });
const failingClient = { messages: { create: async () => { throw new Error('529 overloaded'); } } };

(async () => {
  const report = sampleReport();
  const result = await generateFinalReportPdf(report, {
    revisions: [{ revision: 'R0', date: '2026-09-20', issue: 'First Issue', prepared: 'HSP', checked: 'AP', approved: 'JCS' }],
    status: 'approved'
  });
  const pages = await textOf(result.buffer);
  const all = pages.join('\n').replace(/\s+/g, ' ');

  await check('A4 PDF with cover, control/index and the sections of standard 5.1 in order', () => {
    assert.strictEqual(result.buffer.subarray(0, 5).toString(), '%PDF-');
    const order = ['Executive Summary', 'Introduction & Scope of Work', 'Basis of Assessment', 'Instruments Used', 'Testing Methodology', 'Results', 'Corrective Action Plan', 'Recommendations for Maintenance & Re-testing', 'Annexures'];
    const positions = order.map((title) => pages.findIndex((page, i) => i >= 2 && page.includes(`${order.indexOf(title) + 1}. ${title}`)));
    assert.ok(positions.every((position) => position >= 2), `sections missing: ${positions}`);
    assert.deepStrictEqual(positions, [...positions].sort((a, b) => a - b));
    assert.ok(pages[1].includes('Document Control & Distribution') && pages[1].includes('Index'));
    assert.ok(all.includes('Earthing System Health Assessment') && all.includes('CONFIDENTIAL'));
    assert.ok(all.includes('IS 3043:2018 Cl. 38.4.2'));
  });

  await check('cover carries document no., Zoho project, work order, internal no. and the revision table with initials', () => {
    const cover = pages[0];
    ['024_FI_ESHA', '024-D03_TS-26-GFL', '1002-1261201769', 'Gujarat Fluorochemicals Ltd', 'R0', 'First Issue', 'HSP', 'AP', 'JCS', '20-09-2026', 'Prepared By', 'Checked By', 'Approved By'].forEach((text) => assert.ok(cover.includes(text), `cover lacks ${text}`));
    assert.ok(!all.includes('DRAFT – NOT APPROVED'), 'approved report has no draft marker');
  });

  await check('index lists exactly the sections present and their page numbers are right', () => {
    const index = pages[1];
    result.data.tests.forEach((test) => assert.ok(all.includes(test.label)));
    ['6. Results', '7. Corrective Action Plan'].forEach((title) => assert.ok(index.includes(title)));
    const lines = index.split('\n');
    const resultsPage = pages.findIndex((page, i) => i >= 2 && page.includes('6. Results')) + 1;
    assert.ok(lines.includes(String(resultsPage)), `index should show page ${resultsPage} for Results`);
  });

  await check('executive summary: golden counts, statuses and electrode priorities', () => {
    const summary = pages.find((page) => page.includes('A. Overall summary of tests'));
    ['88.9%', '75.0%', '100.0%', '83.3%', '>Attention Required', 'Healthy'].forEach((text) => assert.ok(summary.includes(text), `summary lacks ${text}`));
    assert.ok(summary.includes('EP-03, EP-07, EP-08') && summary.includes('EP-05, EP-09'));
    assert.ok(summary.includes('EP-01, EP-02, EP-04, EP-06, EP-10'));
    assert.deepStrictEqual(result.data.summaryTable.map((row) => [row.total, row.healthy, row.unhealthy, row.notAssessed]), [[9, 8, 1, 1], [4, 3, 1, 0], [3, 3, 0, 1], [6, 5, 1, 0]]);
  });

  await check('results show every pit, panel and riser with its status; EP-08 is not measured; limits are the golden ones', () => {
    ['EP-07', 'EP-08', '>Permissible Limit', 'Above limit', 'Within limit', 'Not measurable', 'Warning', '6.83', '4.21', '7.12', 'not measured', 'PMCC-1', 'APDB-2', 'Not Assessable', 'R-202 (R1)', '1.42', '25.68'].forEach((text) => assert.ok(all.includes(text), `results lack ${text}`));
    assert.ok(all.includes('Moderately corrosive'));
  });

  await check('Basis of Assessment prints the resolved profile with sources and clauses, and the declared soil resistivity', () => {
    const basis = pages.find((page) => page.includes('3. Basis of Assessment'));
    ['ETPL acceptance criterion', 'Standard', 'IS 3043:2018', 'Declared assumptions'].forEach((text) => assert.ok(all.includes(text)));
    assert.ok(basis.includes('Basis of Assessment'));
    assert.ok(all.includes('No assumed values were used'));
  });

  await check('CAP: defects one line per point, data gaps grouped, same lines as the data layer', () => {
    const cap = pages.filter((page) => page.includes('Corrective Action Plan')).join('\n');
    ['EP-03', 'EP-07', 'EP-08', 'C1: Corroded', 'MCC-2 7F2 Feeder', 'Data gap', 'PSC (L-N) not recorded', 'APDB-2 Incomer'].forEach((text) => assert.ok(all.includes(text), `CAP lacks ${text}`));
    assert.ok(cap.length > 0);
    assert.strictEqual(result.data.cap.filter((line) => line.kind === 'gap').length, 2);
  });

  await check('electrode chart is present and there are no pie charts', () => {
    assert.ok(all.includes('Measured R (navy) vs calculated limit (grey)'));
  });

  // ---- Session 6C.2b report fixes, rendered from the all-tests fixture ---------------------------
  const allReport = normalizeReportInput(buildAllTestsPayload());
  const full = await generateFinalReportPdf(allReport, {
    status: 'approved',
    approvals: { prepared: { name: 'H. S. Patil', designation: 'Site Engineer', date: '2026-09-18' }, checked: { name: 'A. S. Mahto', designation: 'Senior Engineer', date: '2026-09-19' }, approved: { name: 'J. C. Shah', designation: 'Head - Earthing Services', date: '2026-09-20' } }
  });
  const fullPages = await textOf(full.buffer);
  const fullText = fullPages.join('\n').replace(/\s+/g, ' ');

  await check('all nine sheets render in one report with every section', () => {
    ['Soil Resistivity', 'MS-02', 'MS-03', 'MS-04', 'MS-05', 'MS-06', 'MS-07', 'MS-08', 'MS-09'].forEach((text) => assert.ok(fullText.includes(text), `lacks ${text}`));
  });

  await check('cover has no "[client logo]" placeholder, with or without a logo', async () => {
    assert.ok(!fullText.includes('[client logo'));
    assert.ok(!all.includes('[client logo'));
  });

  await check('executive summary: "Assessed" column and "(n not assessed)" in the status', () => {
    const summary = fullPages.find((page) => page.includes('A. Overall summary of tests')).replace(/\s+/g, ' ');
    assert.ok(summary.includes('Assessed') && !/\bTotal\b/.test(summary.split('B. Earth electrode')[0]));
    assert.ok(summary.includes('>Attention Required (1 not assessed)') && summary.includes('Healthy (1 not assessed)'));
  });

  await check('instruments table has Range and Accuracy; annexures say "provided separately" unless attached', async () => {
    const instruments = fullPages.find((page, i) => i >= 2 && page.includes('4. Instruments Used')).replace(/\s+/g, ' ');
    ['Range', 'Accuracy', 'Metrel MI 3290 GX Earth Analyser', '0.01 Ω – 19.9 kΩ', 'Kyoritsu 4118A'].forEach((text) => assert.ok(instruments.includes(text), `instruments lack ${text}`));
    const annexures = fullPages.find((page, i) => i >= 2 && page.includes('9. Annexures')).replace(/\s+/g, ' ');
    assert.strictEqual((annexures.match(/Provided separately/g) || []).length, 4);
    const attached = await textOf((await generateFinalReportPdf(allReport, { status: 'approved', annexures: ['C', 'D'] })).buffer);
    const attachedPage = attached.find((page, i) => i >= 2 && page.includes('9. Annexures')).replace(/\s+/g, ' ');
    assert.strictEqual((attachedPage.match(/Attached/g) || []).length, 2);
    assert.strictEqual((attachedPage.match(/Provided separately/g) || []).length, 2);
  });

  await check('introduction: purpose, client representative, ETPL team and applicable standards', () => {
    ['Purpose', 'was engaged by Gujarat Fluorochemicals Ltd', 'ETPL team', 'H. S. Patil', 'R. K. Shah', 'Applicable standards', 'IS 3043:2018', 'IEEE 80-2013', 'IEEE 81-2012', 'IS/IEC 60364', 'CEA (Measures relating to Safety and Electric Supply) Regulations, 2023'].forEach((text) => assert.ok(fullText.includes(text), `intro lacks ${text}`));
  });

  await check('methodology: current-probe distance, lead nulling and safety (LOTO, PPE, permit)', () => {
    ['Current-probe distance', 'Test-lead nulling', 'lock-out/tag-out (LOTO)', 'PPE', 'work permit'].forEach((text) => assert.ok(fullText.includes(text), `methodology lacks ${text}`));
  });

  await check('basis: the loop row source is readable (no "Cl. IS 3043")', () => {
    assert.ok(!fullText.includes('Cl. IS 3043'));
    assert.ok(fullText.includes('IS 3043:2018 Cl. 42 (method)'));
  });

  await check('soil results: R beside rho, test conditions and a rho-vs-spacing chart per location', () => {
    ['R Direction 1', 'ρ Direction 1', 'weather: Sunny', 'ground condition: Dry', 'Apparent resistivity ρ against probe spacing – Utility Yard', 'Apparent resistivity ρ against probe spacing – Reactor Bay'].forEach((text) => assert.ok(fullText.includes(text), `soil lacks ${text}`));
  });

  await check('electrode results: test condition, 50/70% check, Status vs Priority', () => {
    ['Test condition', 'Isolated', '50/70% check', 'Agree', 'Disagree', 'Status (R vs limit)', 'Within limit', 'Above limit', 'Priority'].forEach((text) => assert.ok(fullText.includes(text), `electrodes lack ${text}`));
  });

  await check('document control: approval block with names, designations and dates; index has a sub-entry per test', () => {
    const control = fullPages[1].replace(/\s+/g, ' ');
    ['Approval', 'H. S. Patil', 'Site Engineer', '18-09-2026', 'A. S. Mahto', 'Senior Engineer', 'J. C. Shah', 'Head - Earthing Services', '20-09-2026'].forEach((text) => assert.ok(control.includes(text), `control lacks ${text}`));
    ['Soil Resistivity', 'Earth Electrode Resistance Test (MS-02)', 'Loop Impedance Test (MS-04)', 'Tower Footing Resistance Measurement & Analysis (MS-09)'].forEach((text) => assert.ok(control.includes(text), `index lacks ${text}`));
    const resultsLine = fullPages.findIndex((page, i) => i >= 2 && page.includes('6. Results')) + 1;
    assert.ok(fullPages[1].split('\n').includes(String(resultsLine)));
  });

  await check('step & touch note, tower feet and the new basis source text are printed', () => {
    ['Basis of the limits: Earth fault current If = 20 kA', 'Gravel: rho_s = 3000.00', 'Concrete: rho_s = 5000.00', 'Soil: rho_s = 25.68', 'Foot-1', 'Foot-4', 'Foot connection', 'Tower T1', 'the 1 tower assessed is within limit', '1 of 1 tower assessed'].forEach((text) => assert.ok(fullText.replace(/ρ/g, '').includes(text.replace(/ρ/g, '')), `lacks ${text}`));
    // the cell wraps across lines, so spaces are ignored
    assert.ok(fullText.replace(/\s/g, '').includes('IS3043:2018Cl.10.3(limits);IEEE81-2012Cl.9(measurementandIf/Itscaling)'));
    assert.strictEqual(allReport.stepTouchVoltage.length > 0 && require('../src/acceptance').resolveProfile().get('ST-LIMIT').source, 'IS 3043:2018 Cl. 10.3 (limits); IEEE 81-2012 Cl. 9 (measurement and If/It scaling)');
    assert.ok(!fullText.includes('Cl. 10.3 (IS 3043:2018); 9.3'));
  });

  await check('continuity results carry the limit', () => {
    const continuity = fullPages.find((page, i) => i >= 2 && page.includes('Earth Continuity & Bonding Test (MS-03)'));
    assert.ok(continuity.replace(/\s/g, '').includes('Limit(Ω)'));
  });

  await check('loop observation and CAP clauses read as the standard asks', () => {
    const cap = fullPages.filter((page, i) => i >= 2 && page.includes('Corrective action')).join(' ').replace(/\s+/g, ' ');
    ['Zs 1.92 Ω on B-E exceeds 1.8 Ω (ETPL criterion)', '38.4.3', '25.6.3.4', 'ETPL criterion'].forEach((text) => assert.ok(cap.includes(text), `CAP lacks ${text}`));
  });

  await check('unapproved report is marked as a draft on every page after the cover', async () => {
    const draft = await generateFinalReportPdf(report, { status: 'in_review' });
    const draftPages = await textOf(draft.buffer);
    assert.ok(draftPages.slice(1).every((page) => page.includes('DRAFT – NOT APPROVED')));
    assert.ok(draftPages[0].includes('R0') && draftPages[0].includes('First Issue'));
  });

  await check('a deliberately mismatched report is refused before anything is drawn', async () => {
    const data = buildFinalReportData(report);
    data.summaryTable.find((row) => row.testId === 'riserIntegrityTest').unhealthy = 0;
    await assert.rejects(generateFinalReportPdf(report, { data }), (error) => error instanceof ConsistencyError && /Riser/.test(error.message));
  });

  await check('without Claude (no key, API error, refusal, unsupported figures) the report still generates with a "not AI-written" note', async () => {
    const data = buildFinalReportData(report);
    const noKey = await generateNarrative(data, { config: { apiKey: '', model: 'claude-opus-5-5', effort: 'medium' } });
    assert.strictEqual(noKey.aiWritten, false);
    assert.ok(/not configured/.test(noKey.note) && noKey.keyFindings.length >= 3);
    const down = await generateNarrative(data, { client: failingClient, config: { apiKey: 'x', model: 'm', effort: 'medium' } });
    assert.strictEqual(down.aiWritten, false);
    const refused = await generateNarrative(data, { client: { messages: { create: async () => ({ stop_reason: 'refusal', content: [] }) } }, config: { apiKey: 'x', model: 'm', effort: 'medium' } });
    assert.ok(/declined/.test(refused.note));
    const invented = await generateNarrative(data, { client: goodClient({ keyFindings: ['EP-07 reads 9.99 ohm.', 'b', 'c'], conditionStatement: '47 points fail.' }), config: { apiKey: 'x', model: 'm', effort: 'medium' } });
    assert.strictEqual(invented.aiWritten, false);
    assert.ok(/unsupported figures/.test(invented.note));
    const pdf = await generateFinalReportPdf(report, { narrative: down });
    const text = (await textOf(pdf.buffer)).join('\n').replace(/\s+/g, ' ');
    assert.ok(text.includes('Narrative not AI-written'));
    assert.ok(text.includes('5 of 10 earth pits need corrective action'));
  });

  await check('a Claude narrative built from the computed results is used, uses medium effort, and sees results only', async () => {
    const data = buildFinalReportData(report);
    let request;
    const client = { messages: { create: async (body) => { request = body; return { stop_reason: 'end_turn', model: 'claude-opus-5-5', content: [{ type: 'text', text: JSON.stringify({ keyFindings: ['3 P1 pits: EP-03, EP-07, EP-08.', 'Loop impedance: 1 of 4 panels above limit (MCC-2 7F2 Feeder).', 'PFC: APDB-2 Incomer could not be assessed.'], conditionStatement: 'The earthing system needs attention: 3 of 22 assessed points are above their limits.' }) }] }; } } };
    const narrative = await generateNarrative(data, { client, config: { apiKey: 'x', model: 'claude-opus-5-5', effort: 'medium' } });
    assert.strictEqual(narrative.aiWritten, true);
    assert.strictEqual(request.model, 'claude-opus-5-5');
    assert.strictEqual(request.output_config.effort, 'medium');
    assert.ok(!/0\.82|1\.10|raw/i.test(request.messages[0].content), 'raw readings are not sent');
    assert.deepStrictEqual(unsupportedFigures({ keyFindings: ['Electrode pass rate 88.9% (IS 3043:2018 Table 4), EP-07 is P1'], conditionStatement: '' }, narrativeInput(data)), []);
    assert.deepStrictEqual(unsupportedFigures({ keyFindings: ['EP-07 is 9.99 ohm'], conditionStatement: '' }, narrativeInput(data)), ['9.99']);
    const pdf = await generateFinalReportPdf(report, { narrative });
    const text = (await textOf(pdf.buffer)).join('\n').replace(/\s+/g, ' ');
    assert.ok(text.includes('written by the report AI') && !text.includes('Narrative not AI-written'));
  });

  await check('config: medium effort is the default and the stale "keep high" comment is gone', () => {
    const config = require('../src/config');
    const source = require('fs').readFileSync(require.resolve('../src/config'), 'utf8');
    assert.ok(!/keep the report writer at high/.test(source));
    assert.ok(/ANTHROPIC_EFFORT \|\| 'medium'/.test(source), 'default effort is medium');
    assert.ok(typeof config.claude.effort === 'string' && config.claude.effort.length > 0);
  });

  // ---- Daily progress report -----------------------------------------------------------------
  const day = (date, tests, extra = {}) => {
    const payload = buildPayload();
    const base = sampleReport();
    const report = normalizeReportInput({ ...payload, project: { ...base.project, reportDate: date, nextDayPlan: 'Earth pits EP-07 to EP-10.', dataCollectionMode: 'Printed sheets', engineerName: 'H. S. Patil' }, tests: { soilResistivity: tests.soil, electrodeResistance: true, ...(tests.rest ? { loopImpedanceTest: true, prospectiveFaultCurrent: true, riserIntegrityTest: true } : {}) }, ...(extra.electrodeResistance ? { electrodeResistance: extra.electrodeResistance } : {}) });
    report.kind = 'daily';
    report.collectionDate = date;
    return report;
  };
  const payload = buildPayload();
  const day1 = day('2026-09-14', { soil: true, rest: false }, { electrodeResistance: payload.electrodeResistance.slice(0, 6) });
  const day2 = day('2026-09-15', { soil: false, rest: true }, { electrodeResistance: payload.electrodeResistance.slice(6) });

  await check('daily progress report: one page per daily report, progress against planned, preliminary statuses', async () => {
    const daily = await generateDailyProgressPdf([day2, day1]);
    const dailyPages = await textOf(daily.buffer);
    assert.strictEqual(dailyPages.length, 2);
    assert.ok(dailyPages[0].includes('14-09-2026') && dailyPages[0].includes('024_FI_ESHA_DPR_01'));
    assert.ok(dailyPages[1].includes('15-09-2026') && dailyPages[1].includes('024_FI_ESHA_DPR_02'));
    ['Daily Progress Report', 'Done today', 'Cumulative', 'Planned', 'Complete', 'Notable observations today', 'Plan for next day', 'Prepared by (ETPL engineer)', 'Acknowledged by (client representative)', 'preliminary', 'CONFIDENTIAL'].forEach((text) => assert.ok(dailyPages.join(' ').replace(/\s+/g, ' ').toLowerCase().includes(text.toLowerCase()), `daily lacks ${text}`));
    assert.ok(dailyPages[0].includes('EP-03') && dailyPages[0].includes('P1') && !dailyPages[0].includes('Corrective Action Plan'));
    assert.ok(dailyPages[1].includes('100%'), 'cumulative 10 of 10 pits on day 2');
    assert.ok(dailyPages[0].includes('60%'), '6 of 10 pits on day 1');
  });

  await check('daily: every observation carries its reading; the soil footnote does not repeat itself; sign-off has Name / Signature / Date', async () => {
    const daily = await generateDailyProgressPdf([day1, day2]);
    const pages = await textOf(daily.buffer);
    const first = pages[0].replace(/\s+/g, ' ');
    assert.ok(/EP-03 Resistance 1\.45 Ω \(limit 6\.83 Ω\)\. C1: Corroded/.test(first), first);
    assert.ok(/EP-05 Resistance 1\.30 Ω \(limit 6\.83 Ω\)\. C8/.test(first), 'every defect row has its reading');
    assert.ok(!/EP-07 Resistance 7\.12 Ω \(limit[^.]*\)\. Resistance 7\.12/.test(first), 'a reading already quoted is not repeated');
    pages.forEach((page) => ['Name:', 'Signature:', 'Date:'].forEach((label) => assert.ok(page.includes(label), `sign-off lacks ${label}`)));
  });

  await check('a single daily report can be rendered alone while its cumulative counts use the earlier days', async () => {
    day1.id = 'D1';
    day2.id = 'D2';
    const only = await generateDailyProgressPdf([day1, day2], { onlyReportId: 'D2' });
    const text = await textOf(only.buffer);
    assert.strictEqual(text.length, 1);
    assert.ok(text[0].includes('15-09-2026') && text[0].includes('100%'));
  });

  await check('daily electrode limits fall back to soil from the same or EARLIER days only, printed as preliminary', async () => {
    const noSoil = (report) => ({ ...report, project: { ...report.project, assumedSoilResistivity: '' }, soilResistivity: { locations: [] } });
    const early = { ...day1, id: 'E1' };
    const middle = { ...noSoil(day('2026-09-15', { soil: false, rest: false }, { electrodeResistance: payload.electrodeResistance.slice(6) })), kind: 'daily', collectionDate: '2026-09-15', id: 'E2' };
    const late = { ...day('2026-09-20', { soil: true, rest: false }, { electrodeResistance: payload.electrodeResistance.slice(6) }), id: 'E3' };
    const withLaterSoil = await generateDailyProgressPdf([early, middle, late], { onlyReportId: 'E2' });
    assert.ok(withLaterSoil.fallbacks[1] && withLaterSoil.fallbacks[1].date === '2026-09-14', 'uses the 14th, not the later 20th');
    const text = (await textOf(withLaterSoil.buffer)).join('\n').replace(/\s+/g, ' ');
    assert.ok(text.includes('preliminary – ρ from 14-09-2026'), 'daily prints the source day');
    const footnote = text.slice(text.indexOf('Electrode limits are preliminary'), text.indexOf('Statuses are preliminary'));
    assert.ok(!/measured on/.test(footnote) && (footnote.match(/14-09-2026/g) || []).length === 1, `footnote says the date once: ${footnote}`);
    assert.ok(withLaterSoil.datas[1].tests.find((test) => test.id === 'electrodeResistance').points.every((point) => Number.isFinite(point.limit)), 'limits are calculated');
    const onlyLater = await generateDailyProgressPdf([middle, late], { onlyReportId: 'E2' });
    assert.strictEqual(onlyLater.fallbacks[0], null, 'a later day is never used');
    assert.ok(onlyLater.datas[0].tests.find((test) => test.id === 'electrodeResistance').points.every((point) => !Number.isFinite(point.limit)), 'still no limit: soil resistivity required');
    assert.strictEqual(onlyLater.datas[0].basis.rho, null);
    assert.strictEqual(withLaterSoil.fallbacks[0], null, 'a day with its own soil is untouched');
  });

  console.log(`\n${passed} report-pdf checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
