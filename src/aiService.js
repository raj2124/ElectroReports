const fs = require('fs');
const os = require('os');
const path = require('path');
const OpenAI = require('openai');

const config = require('./config');
const { REPORT_AI_SYSTEM_GUIDE } = require('./aiReferenceGuide');
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
const referenceFileCache = new Map();
const AI_HIGHLIGHT_LIMIT = 6;

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

function isAiConfigured() {
  return Boolean(config.ai.apiKey);
}

function getClient() {
  if (!isAiConfigured()) {
    throw new Error('OpenAI API key is not configured for ElectroReports.');
  }
  if (!openAiClient) {
    openAiClient = new OpenAI({ apiKey: config.ai.apiKey });
  }
  return openAiClient;
}

async function getReferenceFileIds(client) {
  const ids = [];
  const referencePaths = discoverReferencePdfPaths();

  for (const filePath of referencePaths) {
    const stats = fs.statSync(filePath);
    const cached = referenceFileCache.get(filePath);
    if (cached && cached.size === stats.size && cached.mtimeMs === stats.mtimeMs) {
      ids.push(cached.fileId);
      continue;
    }
    const uploaded = await client.files.create({
      file: fs.createReadStream(filePath),
      purpose: 'user_data'
    });
    referenceFileCache.set(filePath, {
      fileId: uploaded.id,
      size: stats.size,
      mtimeMs: stats.mtimeMs
    });
    ids.push(uploaded.id);
  }

  return ids;
}

async function generateReportNarrative(report) {
  const client = getClient();
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
    model: config.ai.model,
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

  let parsed;
  try {
    parsed = JSON.parse(response.output_text || '{}');
  } catch (_error) {
    throw new Error('OpenAI returned an unexpected AI narrative format.');
  }

  return {
    generatedAt: new Date().toISOString(),
    model: response.model || config.ai.model,
    referenceDocuments,
    ...parsed
  };
}

function getAiStatus() {
  return {
    configured: isAiConfigured(),
    model: config.ai.model,
    referenceDocuments: discoverReferencePdfPaths().map((filePath) => path.basename(filePath))
  };
}

module.exports = {
  getAiStatus,
  generateReportNarrative,
  isAiConfigured
};
