const fs = require('fs');
const os = require('os');
const path = require('path');
const OpenAI = require('openai');
const Anthropic = require('@anthropic-ai/sdk');

const config = require('./config');
const { REPORT_AI_SYSTEM_GUIDE, REPORT_AI_OBSERVATION_GUIDE } = require('./aiReferenceGuide');
const {
  TEST_LIBRARY,
  EQUIPMENT_LIBRARY,
  calculateSoilSummary,
  deriveElectrodeAssessment,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveEarthContinuityAssessment,
  deriveTowerFootingAssessment,
  summarizeTowerGroups,
  buildTowerGroupKey,
  buildExecutiveSnapshot,
  getElectrodeMeasuredValue,
  round
} = require('./reportModel');

let openAiClient = null;
let anthropicClient = null;
const referenceFileCache = new Map();
const AI_HIGHLIGHT_LIMIT = 6;
// Guard so the base64-encoded reference PDFs stay well under the 32 MB
// Anthropic request limit (leaving headroom for the prompt + context).
const CLAUDE_REFERENCE_MAX_TOTAL_BYTES = 24 * 1024 * 1024;

function safeText(value, fallback = '') {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || fallback;
}

function getReportTitle(report) {
  const selected = TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);
  if (!selected.length) {
    return 'Earthing System Health Assessment';
  }
  if (selected.length === 1) {
    return selected[0].label;
  }
  return 'Earthing System Health Assessment';
}

function discoverReferencePdfPaths() {
  const configured = Array.isArray(config.ai.referencePdfPaths) ? config.ai.referencePdfPaths : [];
  if (configured.length) {
    return configured.filter((filePath) => fs.existsSync(filePath));
  }

  const home = os.homedir();
  const downloads = path.join(home, 'Downloads');
  const preferred = [
    path.join(downloads, '01_520_Report_Earthing System Health Assessment_R0.pdf'),
    path.join(downloads, '02_521_Earthing System Health Assessment.pdf')
  ];

  const generatedDir = config.app.generatedDir;
  const generatedPdfs = fs.existsSync(generatedDir)
    ? fs
        .readdirSync(generatedDir)
        .filter((fileName) => fileName.toLowerCase().endsWith('.pdf'))
        .map((fileName) => path.join(generatedDir, fileName))
        .sort((left, right) => fs.statSync(right).mtimeMs - fs.statSync(left).mtimeMs)
    : [];

  return [...preferred, generatedPdfs[0]].filter((filePath) => filePath && fs.existsSync(filePath));
}

function getSelectedEquipmentLabels(report) {
  const selected = new Set(Array.isArray(report?.project?.equipmentSelections) ? report.project.equipmentSelections : []);
  return EQUIPMENT_LIBRARY.filter((item) => selected.has(item.id)).map((item) => item.label);
}

function summarizeStatusRows(rows) {
  const tones = rows.map((row) => row?.status?.tone || 'neutral');
  const total = tones.length;
  const healthy = tones.filter((tone) => tone === 'healthy').length;
  const warning = tones.filter((tone) => tone === 'warning').length;
  const critical = tones.filter((tone) => tone === 'critical').length;
  const pending = tones.filter((tone) => tone === 'neutral').length;
  return { total, healthy, warning, critical, pending };
}

function buildComplianceSummary(report) {
  const rows = [];
  const soil = calculateSoilSummary(report);
  const towerSummaries = summarizeTowerGroups(report.towerFootingResistance);

  if (report.tests.soilResistivity) {
    const tones = soil.locations.map((location) => location.category?.tone || 'neutral');
    const total = tones.length;
    const healthy = tones.filter((tone) => tone === 'healthy').length;
    const warning = tones.filter((tone) => tone === 'warning').length;
    const critical = tones.filter((tone) => tone === 'critical').length;
    const pending = tones.filter((tone) => tone === 'neutral').length;
    rows.push({
      id: 'soilResistivity',
      label: 'Soil Resistivity Test',
      testArea: 'Soil Resistivity',
      total,
      healthy,
      warning,
      critical,
      pending,
      compliance: total ? round(((healthy + warning) / total) * 100, 1) : 0
    });
  }

  const modules = [
    {
      enabled: report.tests.electrodeResistance,
      id: 'electrodeResistance',
      label: 'Earth Electrode Resistance Test',
      testArea: 'Earth Pit Resistance',
      rows: report.electrodeResistance.map((row) => ({ status: deriveElectrodeAssessment(row).status }))
    },
    {
      enabled: report.tests.continuityTest,
      id: 'continuityTest',
      label: 'Earthing Continuity Test',
      testArea: 'Grid Integrity',
      rows: report.continuityTest.map((row) => ({ status: deriveContinuityAssessment(row).status }))
    },
    {
      enabled: report.tests.loopImpedanceTest,
      id: 'loopImpedanceTest',
      label: 'Earth Loop Impedance Test',
      testArea: 'Loop Impedance',
      rows: report.loopImpedanceTest.map((row) => ({ status: deriveLoopImpedanceAssessment(row).status }))
    },
    {
      enabled: report.tests.prospectiveFaultCurrent,
      id: 'prospectiveFaultCurrent',
      label: 'Prospective Fault Current Test',
      testArea: 'Fault Current',
      rows: report.prospectiveFaultCurrent.map((row) => ({ status: deriveFaultCurrentAssessment(row).status }))
    },
    {
      enabled: report.tests.riserIntegrityTest,
      id: 'riserIntegrityTest',
      label: 'Riser / Grid Integrity Test',
      testArea: 'Riser Integrity',
      rows: report.riserIntegrityTest.map((row) => ({ status: deriveRiserAssessment(row).status }))
    },
    {
      enabled: report.tests.earthContinuityTest,
      id: 'earthContinuityTest',
      label: 'Earth Continuity Test',
      testArea: 'Earth Continuity',
      rows: report.earthContinuityTest.map((row) => ({ status: deriveEarthContinuityAssessment(row).status }))
    },
    {
      enabled: report.tests.towerFootingResistance,
      id: 'towerFootingResistance',
      label: 'Tower Footing Resistance Measurement & Analysis',
      testArea: 'Tower Footing',
      rows: report.towerFootingResistance.map((group) => ({
        status: deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group))).status
      }))
    }
  ];

  modules.forEach((module) => {
    if (!module.enabled) {
      return;
    }
    const counts = summarizeStatusRows(module.rows);
    rows.push({
      id: module.id,
      label: module.label,
      testArea: module.testArea,
      ...counts,
      compliance: counts.total ? round((counts.healthy / counts.total) * 100, 1) : 0
    });
  });

  return rows;
}

function buildDetailedModuleFacts(report) {
  const soil = calculateSoilSummary(report);
  const towerSummaries = summarizeTowerGroups(report.towerFootingResistance);

  return {
    soilResistivity: report.tests.soilResistivity
      ? soil.locations.map((location) => ({
          location: location.name,
          averageOhmMeter: location.overallAverage,
          category: location.category?.label || 'Insufficient Data',
          drivenElectrodeDiameterMm: safeText(location.drivenElectrodeDiameter),
          drivenElectrodeLengthM: safeText(location.drivenElectrodeLength),
          notes: safeText(location.notes)
        }))
      : [],
    electrodeResistance: report.tests.electrodeResistance
      ? report.electrodeResistance.map((row) => {
          const assessment = deriveElectrodeAssessment(row);
          return {
            tag: safeText(row.tag),
            location: safeText(row.location),
            measuredWithGridOhm: getElectrodeMeasuredValue(row),
            status: assessment.status.label,
            comment: assessment.comment,
            rowObservation: safeText(row.rowObservation)
          };
        })
      : [],
    continuityTest: report.tests.continuityTest
      ? report.continuityTest.map((row) => {
          const assessment = deriveContinuityAssessment(row);
          return {
            mainLocation: safeText(row.mainLocation),
            measurementPoint: safeText(row.measurementPoint),
            resistanceOhm: safeText(row.resistance),
            impedanceOhm: safeText(row.impedance),
            status: assessment.status.label,
            comment: assessment.comment,
            rowObservation: safeText(row.rowObservation)
          };
        })
      : [],
    loopImpedanceTest: report.tests.loopImpedanceTest
      ? report.loopImpedanceTest.map((row) => {
          const assessment = deriveLoopImpedanceAssessment(row);
          return {
            location: safeText(row.location),
            feederTag: safeText(row.feederTag),
            measuredPoint: safeText(row.measuredPoints),
            deviceType: safeText(row.deviceType),
            deviceRatingA: safeText(row.deviceRating),
            breakingCapacityKa: safeText(row.breakingCapacity),
            loopImpedanceOhm: safeText(row.loopImpedance),
            voltageV: safeText(row.voltage),
            status: assessment.status.label,
            comment: assessment.comment,
            rowObservation: safeText(row.rowObservation)
          };
        })
      : [],
    prospectiveFaultCurrent: report.tests.prospectiveFaultCurrent
      ? report.prospectiveFaultCurrent.map((row) => {
          const assessment = deriveFaultCurrentAssessment(row);
          return {
            location: safeText(row.location),
            feederTag: safeText(row.feederTag),
            measuredPoint: safeText(row.measuredPoints),
            deviceType: safeText(row.deviceType),
            deviceRatingA: safeText(row.deviceRating),
            breakingCapacityKa: safeText(row.breakingCapacity),
            loopImpedanceOhm: safeText(row.loopImpedance),
            faultCurrentKa: safeText(row.prospectiveFaultCurrent),
            voltageV: safeText(row.voltage),
            status: assessment.status.label,
            comment: assessment.comment,
            rowObservation: safeText(row.rowObservation)
          };
        })
      : [],
    riserIntegrityTest: report.tests.riserIntegrityTest
      ? report.riserIntegrityTest.map((row) => {
          const assessment = deriveRiserAssessment(row);
          return {
            mainLocation: safeText(row.mainLocation),
            measurementPoint: safeText(row.measurementPoint),
            resistanceTowardsEquipmentOhm: safeText(row.resistanceTowardsEquipment),
            resistanceTowardsGridOhm: safeText(row.resistanceTowardsGrid),
            status: assessment.status.label,
            comment: assessment.comment,
            rowObservation: safeText(row.rowObservation)
          };
        })
      : [],
    earthContinuityTest: report.tests.earthContinuityTest
      ? report.earthContinuityTest.map((row) => {
          const assessment = deriveEarthContinuityAssessment(row);
          return {
            tag: safeText(row.tag),
            location: safeText(row.locationBuildingName),
            distance: safeText(row.distance),
            measuredValue: safeText(row.measuredValue),
            status: assessment.status.label,
            comment: assessment.comment,
            rowObservation: safeText(row.rowObservation)
          };
        })
      : [],
    towerFootingResistance: report.tests.towerFootingResistance
      ? report.towerFootingResistance.map((group) => {
          const assessment = deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group)));
          return {
            srNo: safeText(group.srNo),
            mainLocationTower: safeText(group.mainLocationTower),
            totalImpedanceZt: assessment.totalImpedanceZt,
            totalCurrentItotal: assessment.totalCurrentItotal,
            zsat: assessment.zsat,
            status: assessment.status.label,
            comment: assessment.comment,
            readings: (Array.isArray(group.readings) ? group.readings : []).map((reading) => ({
              measurementPointLocation: safeText(reading.measurementPointLocation),
              footToEarthingConnectionStatus: safeText(reading.footToEarthingConnectionStatus),
              measuredCurrentMa: safeText(reading.measuredCurrentMa),
              measuredImpedanceOhm: safeText(reading.measuredImpedance),
              rowObservation: safeText(reading.rowObservation)
            }))
          };
        })
      : []
  };
}

function statusRank(value) {
  const text = safeText(value).toLowerCase();
  if (!text) {
    return 0;
  }
  if (
    text.includes('critical') ||
    text.includes('action required') ||
    text.includes('not acceptable') ||
    text.includes('> permissible') ||
    text.includes('un-healthy') ||
    text.includes('needs attention')
  ) {
    return 5;
  }
  if (text.includes('warning') || text.includes('marginal') || text.includes('monitoring') || text.includes('medium')) {
    return 4;
  }
  if (text.includes('high')) {
    return 3;
  }
  if (text.includes('healthy') || text.includes('compliant') || text.includes('low')) {
    return 2;
  }
  return 1;
}

function compactModuleFacts(moduleFacts) {
  return Object.fromEntries(
    Object.entries(moduleFacts).map(([key, rows]) => {
      const list = Array.isArray(rows) ? rows : [];
      const highlights = [...list]
        .sort((left, right) => {
          const leftRank = Math.max(statusRank(left.status), statusRank(left.comment), statusRank(left.category));
          const rightRank = Math.max(statusRank(right.status), statusRank(right.comment), statusRank(right.category));
          return rightRank - leftRank;
        })
        .slice(0, AI_HIGHLIGHT_LIMIT);

      return [
        key,
        {
          totalRows: list.length,
          highlights
        }
      ];
    })
  );
}

function buildAiContext(report) {
  const snapshot = buildExecutiveSnapshot(report);
  const complianceSummary = buildComplianceSummary(report);
  const detailedModuleFacts = compactModuleFacts(buildDetailedModuleFacts(report));
  const overallCompliance = complianceSummary.length
    ? round(
        complianceSummary.reduce((sum, row) => sum + Number(row.compliance || 0), 0) / complianceSummary.length,
        1
      )
    : 0;

  return {
    reportTitle: getReportTitle(report),
    project: {
      projectNo: report.project.projectNo,
      clientName: report.project.clientName,
      siteLocation: report.project.siteLocation,
      workOrder: report.project.workOrder,
      reportDate: report.project.reportDate,
      engineerName: report.project.engineerName,
      zohoProjectName: safeText(report.project.zohoProjectName),
      zohoProjectOwner: safeText(report.project.zohoProjectOwner),
      zohoProjectStage: safeText(report.project.zohoProjectStage)
    },
    selectedMeasurementSheets: TEST_LIBRARY.filter((test) => report.tests[test.id]).map((test) => ({
      id: test.id,
      label: test.label
    })),
    selectedInstruments: getSelectedEquipmentLabels(report),
    executiveSnapshot: {
      selectedTests: snapshot.selectedTests.map((test) => test.label),
      meanSoilResistivity: snapshot.soil.overallAverage,
      soilCategory: snapshot.soil.category?.label || 'Insufficient Data',
      healthyElectrodes: snapshot.healthyElectrodes,
      totalElectrodes: snapshot.totalElectrodes,
      overallCompliancePercent: overallCompliance
    },
    complianceSummary,
    detailedModuleFacts,
    guardrails: {
      deterministicCalculationsOwnedByApp: true,
      doNotInventMeasurements: true,
      doNotFillEngineerOnlyObservationColumns: true,
      selectedScopeOnly: true
    }
  };
}

function buildNarrativeSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      overallAssessment: { type: 'string' },
      executiveSummary: { type: 'string' },
      closingSummary: { type: 'string' },
      keyFindings: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            testArea: { type: 'string' },
            keyFinding: { type: 'string' },
            priority: { type: 'string' },
            recommendedAction: { type: 'string' },
            status: { type: 'string' }
          },
          required: ['testArea', 'keyFinding', 'priority', 'recommendedAction', 'status']
        }
      },
      moduleSummaries: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            testId: { type: 'string' },
            label: { type: 'string' },
            summary: { type: 'string' }
          },
          required: ['testId', 'label', 'summary']
        }
      }
    },
    required: ['overallAssessment', 'executiveSummary', 'closingSummary', 'keyFindings', 'moduleSummaries']
  };
}

function buildNarrativePrompt(context) {
  return `
Generate ElectroReports narrative content from the supplied report context.

Output requirements:
- overallAssessment: one concise paragraph for the Executive Summary "Overall Assessment" section.
- executiveSummary: one polished paragraph that reads like a final report introduction for the executive-summary page.
- closingSummary: one concise closing paragraph for the report conclusion, written like a final submission close-out.
- keyFindings: 3 to 8 rows, ordered by severity first, using only P1 — Critical, P2 — High, P3 — Moderate, or P4 — Normal.
- moduleSummaries: one concise polished paragraph per selected measurement sheet for insertion into that module chapter.

Content rules:
- Do not invent measurements or standards.
- Do not repeat every row; summarize patterns.
- Mention only measurement sheets selected in this report.
- Prioritize critical and warning areas over compliant ones.
- Use "Action Required", "Monitoring", or "Compliant" for status.
- Keep wording suitable for final client submission.
- Keep wording aligned with Elegrow's finalized booklet style.
- For module summaries, explain the outcome of that measurement sheet in a clean paragraph that can sit under the section heading without needing edits.
- Recommended actions should sound field-practical and client-ready, not generic.
- Do not mention the application, automation, uploaded context, JSON, or software-generated behavior.
- Use compact, polished engineering prose with clear compliance phrasing.
- Avoid filler sentences and avoid repeating the same clause structure in every section.

Report context JSON:
${JSON.stringify(context, null, 2)}
`.trim();
}

// ---------------------------------------------------------------------------
// Provider configuration + clients
//   Claude (Anthropic /v1/messages) = PRIMARY narrative brain.
//   OpenAI (Responses API)          = SECONDARY: observations + sanity-check,
//                                     and the narrative fallback if Claude fails.
// GUARDRAIL: neither provider may emit a number or pass/fail band that appears
// in a report. All values come from reportModel.js via buildAiContext().
// ---------------------------------------------------------------------------

function isOpenAiConfigured() {
  return Boolean(config.openai.apiKey);
}

// Backward-compatible alias (older callers used isAiConfigured()).
function isAiConfigured() {
  return isOpenAiConfigured();
}

function isClaudeConfigured() {
  return Boolean(config.claude.apiKey);
}

function getOpenAiClient() {
  if (!isOpenAiConfigured()) {
    throw new Error('OpenAI API key is not configured for ElectroReports.');
  }
  if (!openAiClient) {
    openAiClient = new OpenAI({ apiKey: config.openai.apiKey });
  }
  return openAiClient;
}

function getClaudeClient() {
  if (!isClaudeConfigured()) {
    throw new Error('Anthropic (Claude) API key is not configured for ElectroReports.');
  }
  if (!anthropicClient) {
    anthropicClient = new Anthropic({ apiKey: config.claude.apiKey });
  }
  return anthropicClient;
}

// Read the reference booklet PDFs discovered by discoverReferencePdfPaths() and
// return them as base64 document blocks for Claude. This is what makes grounding
// actually work — the old getReferenceFileIds() uploader was never called.
function readReferenceDocuments() {
  const documents = [];
  let totalBytes = 0;
  for (const filePath of discoverReferencePdfPaths()) {
    try {
      const buffer = fs.readFileSync(filePath);
      if (totalBytes + buffer.length > CLAUDE_REFERENCE_MAX_TOTAL_BYTES) {
        continue;
      }
      totalBytes += buffer.length;
      documents.push({ name: path.basename(filePath), base64: buffer.toString('base64') });
    } catch (_error) {
      // Skip unreadable / missing reference files silently.
    }
  }
  return documents;
}

function extractClaudeText(response) {
  const blocks = Array.isArray(response?.content) ? response.content : [];
  return blocks
    .filter((block) => block && block.type === 'text' && typeof block.text === 'string')
    .map((block) => block.text)
    .join('')
    .trim();
}

function parseStructuredJson(text, providerLabel) {
  try {
    return JSON.parse(text || '{}');
  } catch (_error) {
    throw new Error(`${providerLabel} returned an unexpected structured-output format.`);
  }
}

// PRIMARY narrative path — Claude via the Anthropic Messages API. Feeds the ETPL
// system guide AND the reference booklet PDFs as real document blocks, and
// reproduces the exact same JSON schema the OpenAI path produced.
async function generateNarrativeWithClaude(report) {
  const client = getClaudeClient();
  const context = buildAiContext(report);
  const referenceDocs = readReferenceDocuments();
  const referenceDocuments = referenceDocs.map((doc) => doc.name);

  const content = [];
  for (const doc of referenceDocs) {
    content.push({
      type: 'document',
      source: { type: 'base64', media_type: 'application/pdf', data: doc.base64 }
    });
  }
  content.push({
    type: 'text',
    text: buildNarrativePrompt({ ...context, referenceDocuments })
  });

  const response = await client.messages.create({
    model: config.claude.model,
    max_tokens: 8000,
    system: REPORT_AI_SYSTEM_GUIDE,
    thinking: { type: 'adaptive' },
    output_config: {
      format: { type: 'json_schema', schema: buildNarrativeSchema() }
    },
    messages: [{ role: 'user', content }]
  });

  const parsed = parseStructuredJson(extractClaudeText(response), 'Claude');
  return {
    generatedAt: new Date().toISOString(),
    provider: 'claude',
    model: response.model || config.claude.model,
    referenceDocuments,
    ...parsed
  };
}

// SECONDARY / FALLBACK narrative path — the original OpenAI Responses call,
// preserved verbatim so the report can still be produced if Claude fails.
async function generateNarrativeWithOpenAI(report) {
  const client = getOpenAiClient();
  const context = buildAiContext(report);
  const referenceDocuments = discoverReferencePdfPaths().map((filePath) => path.basename(filePath));
  const content = [
    {
      type: 'input_text',
      text: buildNarrativePrompt({
        ...context,
        referenceDocuments
      })
    }
  ];

  const response = await client.responses.create({
    model: config.openai.model,
    instructions: REPORT_AI_SYSTEM_GUIDE,
    input: [
      {
        role: 'user',
        content
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'electroreports_narrative',
        strict: true,
        schema: buildNarrativeSchema()
      }
    }
  });

  const parsed = parseStructuredJson(response.output_text, 'OpenAI');
  return {
    generatedAt: new Date().toISOString(),
    provider: 'openai',
    model: response.model || config.openai.model,
    referenceDocuments,
    ...parsed
  };
}

// Public entry point. Claude is primary; OpenAI is the fallback if the Claude
// call is unavailable or errors. Same return contract in both cases, so the
// server + PDF code downstream is unchanged.
async function generateReportNarrative(report) {
  if (isClaudeConfigured()) {
    try {
      return await generateNarrativeWithClaude(report);
    } catch (error) {
      console.error('Claude narrative generation failed; falling back to OpenAI.', error);
    }
  }
  return generateNarrativeWithOpenAI(report);
}

// ---------------------------------------------------------------------------
// OpenAI SECONDARY role: measurement-sheet observations / recommendations.
// Only the two modules that expose engineer observation+recommendation columns
// (Earth Electrode Resistance, Riser / Grid Integrity) are covered — those are
// the cells that print. OpenAI DESCRIBES results in prose; it must not emit any
// number or pass/fail band (those come from reportModel.js).
// ---------------------------------------------------------------------------

function buildObservationSchema() {
  const rowItems = (extraKey) => ({
    type: 'array',
    items: {
      type: 'object',
      additionalProperties: false,
      properties: {
        index: { type: 'integer' },
        [extraKey]: { type: 'string' },
        observation: { type: 'string' },
        recommendation: { type: 'string' }
      },
      required: ['index', extraKey, 'observation', 'recommendation']
    }
  });

  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      electrodeResistance: rowItems('tag'),
      riserIntegrityTest: rowItems('measurementPoint')
    },
    required: ['electrodeResistance', 'riserIntegrityTest']
  };
}

function buildObservationPrompt(context) {
  return `
You are completing the engineer observation and recommendation columns for the
Earth Electrode Resistance and Riser / Grid Integrity measurement sheets.

For each row (identified by its zero-based "index", matching row order in the
supplied facts) write:
- observation: one concise sentence describing the recorded condition of that row
  in field-engineer language, grounded in the already-computed status/comment.
- recommendation: one concise, field-practical next action for that row.

Hard rules:
- Describe results; never compute or restate any numeric value, resistance,
  average, percentage, or pass/fail band. Those belong to the deterministic
  application output only.
- Base observations strictly on the provided per-row status/comment facts; do not
  invent conditions, standards, or measurements.
- If a module is not selected or has no rows, return an empty array for it.
- Keep wording aligned with Elegrow's finalized booklet style and suitable for a
  final client submission.

Report context JSON:
${JSON.stringify(context, null, 2)}
`.trim();
}

// Runs only when at least one of the two supported modules is selected.
async function generateMeasurementObservations(report) {
  const wantsElectrode = Boolean(report?.tests?.electrodeResistance);
  const wantsRiser = Boolean(report?.tests?.riserIntegrityTest);
  if (!wantsElectrode && !wantsRiser) {
    return null;
  }

  const client = getOpenAiClient();
  const context = buildAiContext(report);
  const response = await client.responses.create({
    model: config.openai.model,
    instructions: REPORT_AI_OBSERVATION_GUIDE,
    input: [
      {
        role: 'user',
        content: [{ type: 'input_text', text: buildObservationPrompt(context) }]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'electroreports_observations',
        strict: true,
        schema: buildObservationSchema()
      }
    }
  });

  const parsed = parseStructuredJson(response.output_text, 'OpenAI');
  return {
    generatedAt: new Date().toISOString(),
    provider: 'openai',
    model: response.model || config.openai.model,
    modules: parsed
  };
}

// ---------------------------------------------------------------------------
// OpenAI SECONDARY role: sanity-check Claude's narrative against the
// deterministic results. Verify/describe only — no numbers emitted.
// ---------------------------------------------------------------------------

function buildSanityCheckSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      consistent: { type: 'boolean' },
      flags: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            area: { type: 'string' },
            severity: { type: 'string' },
            issue: { type: 'string' }
          },
          required: ['area', 'severity', 'issue']
        }
      },
      notes: { type: 'string' }
    },
    required: ['consistent', 'flags', 'notes']
  };
}

function buildSanityCheckPrompt(context, narrative) {
  return `
Cross-check the drafted report narrative against the deterministic application
results. You are a NARROW numeric/verdict verifier — not an author, editor, or
style reviewer.

Flag an inconsistency ONLY when the narrative states a specific numeric value or
a pass/fail verdict (e.g. compliant/non-compliant, healthy/critical, safe/unsafe)
that DIRECTLY CONTRADICTS the corresponding value or status in the deterministic
results below. Example: the narrative calls an electrode "healthy" when its
deterministic status is "critical", or cites a mean soil resistivity that differs
from the deterministic figure.

Do NOT flag any of the following. These are always acceptable and are NOT
inconsistencies:
- Citations or references to standards (e.g. IS 3043, IEEE 80, IEEE 81) or their
  named limits/clauses. These are correct and expected for an Earthing System
  Health Assessment.
- Engineering context, methodology, background, or explanatory prose.
- Narrative, recommendations, or qualitative descriptions that do not state a
  number or verdict contradicting the deterministic results.
- A value, standard, or fact simply being ABSENT from the deterministic JSON.
  Only a direct contradiction counts — never treat "not present" or
  "not verifiable" as an inconsistency.

Output:
- consistent: false ONLY if you find at least one number or verdict that
  contradicts the deterministic results; otherwise true.
- flags: one entry per contradiction (area = section/module, severity =
  "critical" | "warning" | "info", issue = quote the narrative claim AND the
  contradicting deterministic value). Empty if none.
- notes: a one-line overall verdict.

Do not rewrite the narrative and do not emit any numeric report value yourself.

Deterministic results JSON:
${JSON.stringify(context, null, 2)}

Drafted narrative JSON:
${JSON.stringify(narrative, null, 2)}
`.trim();
}

async function sanityCheckNarrative(report, narrative) {
  if (!narrative) {
    return null;
  }
  const client = getOpenAiClient();
  const context = buildAiContext(report);
  const response = await client.responses.create({
    model: config.openai.model,
    instructions: REPORT_AI_SYSTEM_GUIDE,
    input: [
      {
        role: 'user',
        content: [{ type: 'input_text', text: buildSanityCheckPrompt(context, narrative) }]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'electroreports_sanity_check',
        strict: true,
        schema: buildSanityCheckSchema()
      }
    }
  });

  const parsed = parseStructuredJson(response.output_text, 'OpenAI');
  return {
    generatedAt: new Date().toISOString(),
    provider: 'openai',
    model: response.model || config.openai.model,
    ...parsed
  };
}

function getAiStatus() {
  return {
    narrativeProvider: isClaudeConfigured() ? 'claude' : 'openai',
    claude: {
      configured: isClaudeConfigured(),
      model: config.claude.model,
      role: config.claude.role
    },
    openai: {
      configured: isOpenAiConfigured(),
      model: config.openai.model,
      role: config.openai.role
    },
    referenceDocuments: discoverReferencePdfPaths().map((filePath) => path.basename(filePath))
  };
}

module.exports = {
  getAiStatus,
  generateReportNarrative,
  generateMeasurementObservations,
  sanityCheckNarrative,
  isAiConfigured,
  isClaudeConfigured,
  isOpenAiConfigured
};
