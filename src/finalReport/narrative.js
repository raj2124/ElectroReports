// Key findings and overall condition for the executive summary (standard 5.1 C and D).
// The deterministic text here is built from the computed results only and is always
// available; the Claude-written version (added separately) replaces it when it succeeds.

const plural = (count, one, many = `${one}s`) => `${count} ${count === 1 ? one : many}`;
const list = (items) => items.join(', ');

function deterministicNarrative(data) {
  const findings = [];
  if (data.soil.average !== null && data.soil.average !== undefined) {
    findings.push(`Soil resistivity averages ${data.soil.average.toFixed(1)} ohm-m, classified ${String(data.soil.category || '').toLowerCase()} per IS 3043:2018 Table 4.`);
  }
  const electrode = data.tests.find((test) => test.id === 'electrodeResistance');
  if (electrode) {
    const { P1, P2 } = data.electrodePriority;
    const notMeasured = electrode.points.filter((point) => !point.counted).map((point) => point.label);
    const parts = [];
    if (P1.count) parts.push(`P1: ${list(P1.tags)}`);
    if (P2.count) parts.push(`P2: ${list(P2.tags)}`);
    findings.push(
      P1.count + P2.count
        ? `${P1.count + P2.count} of ${electrode.points.length} earth pits need corrective action (${parts.join('; ')})${notMeasured.length ? `; ${list(notMeasured)} could not be measured` : ''}.`
        : `All ${electrode.points.length} earth pits are within their calculated limits (priority P3).`
    );
  }
  data.summaryTable.filter((row) => row.testId !== 'electrodeResistance' && row.total > 0).forEach((row) => {
    const test = data.tests.find((entry) => entry.id === row.testId);
    const over = test.points.filter((point) => point.counted && !point.withinLimit).map((point) => point.label);
    findings.push(
      over.length
        ? `${row.test}: ${over.length} of ${plural(row.total, row.unit.replace(/s$/, ''))} above limit (${list(over)}).`
        : `${row.test}: all ${plural(row.total, row.unit.replace(/s$/, ''))} assessed are within limit.`
    );
  });
  data.cap.filter((line) => line.kind === 'gap' && line.gap === 'NOT_ASSESSABLE').forEach((line) => {
    findings.push(`${line.test}: ${list(line.affectedPoints)} could not be assessed because the device breaking capacity was not recorded.`);
  });

  const status = data.overall.label;
  const actions = data.cap.filter((line) => line.kind === 'point').length;
  let condition;
  if (data.overall.total === 0) {
    condition = 'No pass/fail results were available to assess the overall condition of the earthing system.';
  } else if (status === 'Healthy' && !actions) {
    condition = `The earthing system of ${data.project.unit || 'the unit'} is in good condition: all measured points meet the acceptance criteria stated in the Basis of Assessment.`;
  } else if (status === 'Healthy' || status === 'Healthy - Minor Actions') {
    condition = `The earthing system of ${data.project.unit || 'the unit'} is generally healthy, with ${plural(actions, 'point')} needing corrective action as listed in the Corrective Action Plan.`;
  } else {
    condition = `The earthing system of ${data.project.unit || 'the unit'} needs attention: ${data.overall.unhealthy} of ${data.overall.total} assessed points are above their limits and ${plural(actions, 'point')} need corrective action, as listed in the Corrective Action Plan.`;
  }
  return {
    keyFindings: findings,
    conditionStatement: condition,
    aiWritten: false,
    note: 'Narrative not AI-written: this text was generated from the computed results by a fixed template.'
  };
}

// ---- Claude-written narrative ---------------------------------------------------------
// Claude writes the key findings and the condition statement from the computed results only.
// If it is unavailable, or its text cannot be trusted, the deterministic text is used and the
// report says so: generation is never blocked.

const SYSTEM_PROMPT = [
  'You write two parts of an earthing system health assessment report for an industrial client: the key findings (3 to 6 bullets) and the overall condition statement (2 to 3 sentences).',
  'Use ONLY the computed results in the JSON. Quote counts, tags, readings and statuses exactly as given. Do not calculate, round, estimate or add any number, tag or standard clause that is not in the JSON.',
  'Do not change or soften any status. Mention every P1 pit, every point above limit and every data gap. No recommendations beyond pointing to the Corrective Action Plan. Plain, factual engineering English; no marketing language.'
].join(' ');

const NARRATIVE_SCHEMA = {
  type: 'object',
  properties: {
    keyFindings: { type: 'array', items: { type: 'string' }, minItems: 3, maxItems: 6 },
    conditionStatement: { type: 'string' }
  },
  required: ['keyFindings', 'conditionStatement'],
  additionalProperties: false
};

// What Claude is allowed to see: results only, no raw readings.
function narrativeInput(data) {
  return {
    unit: data.project.unit || null,
    soil: { averageOhmMetre: data.soil.average, corrosivityClass: data.soil.category },
    overall: { assessed: data.overall.total, withinLimit: data.overall.healthy, aboveLimit: data.overall.unhealthy, notAssessed: data.overall.notAssessed, status: data.overall.label },
    tests: data.summaryTable.map((row) => ({ test: row.test, unit: row.unit, total: row.total, withinLimit: row.healthy, aboveLimit: row.unhealthy, notAssessed: row.notAssessed, passPercent: row.passPct, status: row.status })),
    electrodePriority: data.electrodePriority,
    correctiveActions: data.cap.map((line) => ({ test: line.test, points: line.kind === 'point' ? [line.point] : line.affectedPoints, issue: line.observation, priorityOrStatus: line.priorityOrStatus }))
  };
}

// Numbers in a text, ignoring clause / standard / table references the writer may legitimately cite.
function figuresIn(value) {
  return String(value)
    .replace(/\bIS(?:\/IEC)?\s*\d{2,6}(?::\d{4})?/g, ' ')
    .replace(/\b(?:Cl\.?|Clause|Table|Part|Section)\s*[\d.]+/gi, ' ')
    .replace(/\b(?:P|C)[123]\b/g, ' ')
    .match(/\d+(?:\.\d+)?/g) || [];
}

function unsupportedFigures(output, input) {
  const allowed = new Set(figuresIn(JSON.stringify(input)).map((figure) => Number(figure)));
  [...Array(10).keys()].forEach((n) => allowed.add(n)); // "one", "2 of" style small words are not figures
  return figuresIn([...output.keyFindings, output.conditionStatement].join(' ')).filter((figure) => !allowed.has(Number(figure)));
}

function describeFailure(error) {
  const message = String((error && error.message) || error || 'unknown error');
  if (/not configured|api key/i.test(message)) return 'the Anthropic API key is not configured';
  if (/declined|refusal/i.test(message)) return 'Claude declined the request';
  if (/unsupported figures/i.test(message)) return message;
  return 'Claude was unavailable or returned an unusable response';
}

async function generateNarrative(data, options = {}) {
  const fallback = deterministicNarrative(data);
  const config = options.config || require('../config').claude;
  try {
    let client = options.client;
    if (!client) {
      if (!config.apiKey) {
        throw new Error('Anthropic API key is not configured.');
      }
      const Anthropic = require('@anthropic-ai/sdk');
      client = new Anthropic({ apiKey: config.apiKey });
    }
    const input = narrativeInput(data);
    const response = await client.messages.create(
      {
        model: config.model,
        max_tokens: 2000,
        system: SYSTEM_PROMPT,
        thinking: { type: 'adaptive' },
        output_config: { effort: config.effort, format: { type: 'json_schema', schema: NARRATIVE_SCHEMA } },
        messages: [{ role: 'user', content: `Computed results (JSON):\n${JSON.stringify(input, null, 2)}` }]
      },
      { timeout: options.timeoutMs || 90000 }
    );
    if (response.stop_reason === 'refusal') {
      throw new Error('Claude declined the request.');
    }
    const textBlock = (response.content || []).filter((block) => block && block.type === 'text').map((block) => block.text).join('');
    const parsed = JSON.parse(textBlock);
    if (!Array.isArray(parsed.keyFindings) || !parsed.keyFindings.length || typeof parsed.conditionStatement !== 'string' || !parsed.conditionStatement.trim()) {
      throw new Error('Claude returned an unexpected response format.');
    }
    const extra = unsupportedFigures(parsed, input);
    if (extra.length) {
      throw new Error(`Claude text rejected: unsupported figures ${[...new Set(extra)].join(', ')}.`);
    }
    return { keyFindings: parsed.keyFindings.map(String), conditionStatement: parsed.conditionStatement.trim(), aiWritten: true, model: response.model || config.model, note: '' };
  } catch (error) {
    return { ...fallback, note: `${fallback.note} (${describeFailure(error)}.)`, failure: describeFailure(error) };
  }
}

module.exports = { deterministicNarrative, generateNarrative, narrativeInput, unsupportedFigures };
