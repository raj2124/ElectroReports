// The secondary AI checks (Session 6I), both through OpenAI and neither ever producing a number or a verdict:
//
//   sanityCheck   the OPENAI_SANITY_MODEL reads the final report's narrative against the results-only JSON it was written
//                 from and reports ONLY a number or verdict that contradicts those results. Any flag whose own text carries a
//                 figure that is in neither the results nor the narrative is dropped (a verifier must not introduce numbers).
//                 The outcome is kept on the report as aiSanityCheck and becomes the "sanity check" line of the review checklist.
//   suggestWording the OPENAI_OBSERVATION_MODEL drafts one or two sentences of remark wording for a flagged point, from its
//                 condition codes only. A draft with any digit or any verdict word is thrown away and the template text is
//                 used. The engineer clicks to insert it; nothing is written automatically.
//
// A provider failure never blocks anything: the caller gets { status: 'error', reason } (sanity) or the template text with the
// reason (wording). The guard (src/aiGuard.js) logs every call and pauses a provider that keeps failing. AI_CHECKS_USE_MOCK
// (default true) runs both on deterministic stand-ins; nothing is sent to OpenAI until it is set to false.

const crypto = require('crypto');
const { buildFinalReportData } = require('./finalReport/data');
const { narrativeInput } = require('./finalReport/narrative');
const { CONDITION_CODE_LEGEND } = require('./sheetDefinitions');
const { ProviderError, classify } = require('./aiGuard');
const { stable } = require('./entryCredit');

const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const figuresIn = (value) => String(value).replace(/\bIS(?:\/IEC)?\s*\d{2,6}(?::\d{4})?/g, ' ').replace(/\b(?:Cl\.?|Clause|Table|Part|Section)\s*[\d.]+/gi, ' ').replace(/\b(?:P|C)[123]\b/g, ' ').match(/\d+(?:\.\d+)?/g) || [];

// Words that state a result. Remark wording describes what was seen and what to do; the verdict is the app's.
const VERDICT_WORDS = /\b(pass(ed|es)?|fail(ed|s|ure)?|compliant|non-?compliant|complies|healthy|unhealthy|critical|unsafe|safe|acceptable|not acceptable|marginal|within (the )?limits?|above (the )?limits?|below (the )?limits?|exceed(s|ed)?|meets?|satisfactory|unsatisfactory|good condition|poor condition)\b/i;

const CODE_PHRASE = {
  C1: 'corrosion is visible', C2: 'the strip or test link is disconnected', C3: 'only a single earth connection is present', C4: 'the chamber is not accessible or is damaged',
  C5: 'the electrode is damaged', C6: 'a connection is loose', C7: 'the joint is not cleaned, tinned or protected', C8: 'no tag or label is fitted', OTH: 'another condition was noted (see the remark)'
};

// The fixed wording used whenever the model's draft is missing or rejected: only the condition codes, no figure, no verdict.
function templateWording(conditionCodes = []) {
  const phrases = conditionCodes.map((code) => CODE_PHRASE[code]).filter(Boolean);
  if (!phrases.length) return 'This point was flagged by the assessment; inspect it and record the cause before the report is issued.';
  const list = phrases.length > 1 ? `${phrases.slice(0, -1).join(', ')} and ${phrases[phrases.length - 1]}` : phrases[0];
  return `Observed: ${list}. Inspect and rectify as per the corrective action plan.`;
}

// A draft is usable only if it is plain wording.
function acceptableWording(draft) {
  const value = text(draft).replace(/\s+/g, ' ');
  return Boolean(value) && value.length <= 280 && !/\d/.test(value) && !VERDICT_WORDS.test(value);
}

// ---- the stand-ins and the real client --------------------------------------------------------------
// A deterministic "verifier": flags a figure in the narrative that the results do not hold, and a good-condition claim when the
// overall status is not Healthy. Only for tests and for running the app without OpenAI.
function createMockChecksClient({ sanityModel = 'mock-sanity', observationModel = 'mock-observation', failWith = null } = {}) {
  return {
    mock: true,
    async sanity({ input, narrative }) {
      if (failWith) throw new ProviderError(failWith, `mock failure ${failWith}`);
      const flags = [];
      // a figure of the narrative that the results do not hold, even rounded to 0, 1 or 2 decimals
      const allowed = new Set([...Array(10).keys()]);
      figuresIn(JSON.stringify(input)).forEach((figure) => [0, 1, 2].forEach((digits) => allowed.add(Number(Number(figure).toFixed(digits)))));
      const extra = figuresIn([...narrative.keyFindings, narrative.conditionStatement].join(' ')).filter((figure) => !allowed.has(Number(figure)));
      [...new Set(extra)].forEach((figure) => flags.push({ area: 'Key findings', severity: 'warning', issue: `The narrative states ${figure}, which is not among the computed results.` }));
      const status = text(input.overall && input.overall.status);
      if (/good condition|healthy/i.test(narrative.conditionStatement) && status && !/^healthy$/i.test(status) && !/minor/i.test(status)) {
        flags.push({ area: 'Condition statement', severity: 'critical', issue: `The condition statement describes a good condition but the computed overall status is "${status}".` });
      }
      return { consistent: flags.length === 0, flags, notes: flags.length ? 'Contradictions found.' : 'No contradiction found.', model: sanityModel };
    },
    async wording({ conditionCodes }) {
      if (failWith) throw new ProviderError(failWith, `mock failure ${failWith}`);
      return { text: templateWording(conditionCodes), model: observationModel };
    }
  };
}

const SANITY_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    consistent: { type: 'boolean' },
    flags: { type: 'array', items: { type: 'object', additionalProperties: false, properties: { area: { type: 'string' }, severity: { type: 'string' }, issue: { type: 'string' } }, required: ['area', 'severity', 'issue'] } },
    notes: { type: 'string' }
  },
  required: ['consistent', 'flags', 'notes']
};
const WORDING_SCHEMA = { type: 'object', additionalProperties: false, properties: { text: { type: 'string' } }, required: ['text'] };

function createOpenAiChecksClient({ config, clientFactory = (key) => new (require('openai'))({ apiKey: key }) }) {
  const ask = async (model, instructions, prompt, name, schema) => {
    if (!config.apiKey) throw new ProviderError('NOT_CONFIGURED', 'OPENAI_API_KEY is not set.', { retryable: false });
    try {
      const response = await clientFactory(config.apiKey).responses.create({
        model,
        instructions,
        input: [{ role: 'user', content: [{ type: 'input_text', text: prompt }] }],
        text: { format: { type: 'json_schema', name, strict: true, schema } }
      });
      return { parsed: JSON.parse(response.output_text), model: response.model || model };
    } catch (error) {
      if (error instanceof SyntaxError) throw new ProviderError('BAD_OUTPUT', 'The checker returned something that is not JSON.');
      throw classify(error);
    }
  };
  return {
    mock: false,
    async sanity({ input, narrative }) {
      const { parsed, model } = await ask(config.sanityModel, 'You are a narrow numeric and verdict verifier for an earthing system health assessment report. You never write, edit or restyle text and you never produce a report number.',
        `Cross-check the drafted narrative against the computed results.\nFlag ONLY a specific number or pass/fail/healthy/critical-type verdict stated in the narrative that DIRECTLY CONTRADICTS the corresponding computed value or status below. Do not flag references to standards or clauses, general engineering context, qualitative advice, or anything simply absent from the results. Quote the narrative claim and the contradicting computed value in "issue". Do not rewrite the narrative and do not introduce any number of your own.\n\nComputed results JSON:\n${JSON.stringify(input, null, 2)}\n\nNarrative JSON:\n${JSON.stringify(narrative, null, 2)}`, 'electroreports_sanity_check', SANITY_SCHEMA);
      return { ...parsed, model };
    },
    async wording({ sheetLabel, pointKind, conditionCodes }) {
      const known = conditionCodes.map((code) => `${code}: ${(CONDITION_CODE_LEGEND.find((entry) => entry.code === code) || {}).text || ''}`).join('; ');
      const { parsed, model } = await ask(config.observationModel, 'You draft short, neutral remark wording for a site engineer to review. You never state a number, a measurement, a limit, a standard or a verdict (pass, fail, healthy, critical, compliant, safe...).',
        `Write one or two plain sentences of remark wording for a ${pointKind || 'point'} on the "${sheetLabel}" sheet that the assessment flagged. Describe only what the condition codes say was observed and suggest inspecting or rectifying it. Do not mention any tag, location, number or measurement, and do not say whether it passed or failed.\nCondition codes noted: ${known || 'none'}.`, 'electroreports_remark_wording', WORDING_SCHEMA);
      return { text: parsed.text, model };
    }
  };
}

function createAiChecks({ store, config, guard, client = null }) {
  const checks = client || (config.scan.aiChecksUseMock ? createMockChecksClient() : createOpenAiChecksClient({ config: config.openai }));
  const narrativeOf = (report) => {
    const narrative = report.reportNarrative;
    return narrative && Array.isArray(narrative.keyFindings) ? { keyFindings: narrative.keyFindings.map(String), conditionStatement: text(narrative.conditionStatement) } : null;
  };
  const hashOf = (input, narrative) => crypto.createHash('sha256').update(stable({ input, narrative })).digest('hex');
  const current = (report) => {
    const narrative = narrativeOf(report);
    if (!narrative) return null;
    const input = narrativeInput(buildFinalReportData(report));
    return { input, narrative, hash: hashOf(input, narrative) };
  };

  // Runs the sanity check on the report's stored narrative. Resolves { status: 'passed' | 'issues' | 'error', ... }; on 'error'
  // the report is left as it was. Never throws for a provider problem.
  async function sanityCheck(report) {
    const now = current(report);
    if (!now) return { status: 'error', reason: 'There is no narrative to check yet: export the PDF once so the narrative is written.', code: 'NO_NARRATIVE' };
    let raw;
    try {
      raw = await guard.run('openai', 'sanity', config.openai.sanityModel, () => checks.sanity({ input: now.input, narrative: now.narrative }));
    } catch (error) {
      return { status: 'error', reason: error.message, code: error.code || 'UPSTREAM' };
    }
    // a verifier must not introduce figures: a flag that carries one that is in neither the results nor the narrative is dropped
    const allowed = new Set([...figuresIn(JSON.stringify(now.input)), ...figuresIn(JSON.stringify(now.narrative))].map(Number));
    const flags = (Array.isArray(raw.flags) ? raw.flags : []).map((entry) => ({ area: text(entry.area).slice(0, 80), severity: ['critical', 'warning', 'info'].includes(text(entry.severity)) ? text(entry.severity) : 'warning', issue: text(entry.issue).slice(0, 400) })).filter((entry) => entry.issue);
    const kept = flags.filter((entry) => figuresIn(entry.issue).every((figure) => allowed.has(Number(figure))));
    const result = { status: kept.length ? 'issues' : 'passed', flags: kept, dropped: flags.length - kept.length, notes: text(raw.notes).slice(0, 200), provider: 'openai', model: raw.model || config.openai.sanityModel, at: new Date().toISOString(), narrativeHash: now.hash };
    store.updateReport(report.id, (stored) => ({ ...stored, aiSanityCheck: result }));
    return result;
  }

  // Wording for one flagged point. point: { sheetKey, label (the sheet's name), kind, conditionCodes[] }. Always returns text.
  async function suggestWording({ sheetLabel = '', pointKind = '', conditionCodes = [] }) {
    const codes = [...new Set((Array.isArray(conditionCodes) ? conditionCodes : []).map(text).filter((code) => CODE_PHRASE[code]))];
    const fallback = templateWording(codes);
    try {
      const draft = await guard.run('openai', 'wording', config.openai.observationModel, () => checks.wording({ sheetLabel, pointKind, conditionCodes: codes }));
      if (!acceptableWording(draft.text)) return { text: fallback, source: 'template', status: { ok: true, note: 'The draft had a number or a verdict, so the standard wording is used.' } };
      return { text: text(draft.text).replace(/\s+/g, ' '), source: 'ai', model: draft.model, status: { ok: true } };
    } catch (error) {
      return { text: fallback, source: 'template', status: { ok: false, code: error.code || 'UPSTREAM', reason: error.message } };
    }
  }

  return { sanityCheck, suggestWording, currentHash: (report) => (current(report) || {}).hash || '', client: checks };
}

// The hash of what a sanity check looked at: the results-only input and the narrative. A narrative that changed since is not checked.
function sanityHashOf(report) {
  const narrative = report.reportNarrative && Array.isArray(report.reportNarrative.keyFindings) ? { keyFindings: report.reportNarrative.keyFindings.map(String), conditionStatement: text(report.reportNarrative.conditionStatement) } : null;
  return narrative ? crypto.createHash('sha256').update(stable({ input: narrativeInput(buildFinalReportData(report)), narrative })).digest('hex') : '';
}

module.exports = { sanityHashOf, createAiChecks, createMockChecksClient, createOpenAiChecksClient, templateWording, acceptableWording, VERDICT_WORDS };
