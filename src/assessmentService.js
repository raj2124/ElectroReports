// Computed assessments for a report or an in-progress draft: the single source of
// truth the browser reads. Every status, priority, limit, basis, remark, flag,
// classification and validation message shown in the UI is produced here by
// reportModel's rules, from the project's acceptance profile and the sheet
// definitions. The browser holds no copy of any rule.
//
// Unlike normalizeReportInput this is lenient: a draft is incomplete by nature, so
// missing sections, blank rows and absent project fields are fine. It only throws
// for something that cannot be assessed at all, such as an unknown acceptance profile.

const { describeCriterion } = require('./acceptance');
const { engineerRemark, REMARK_FIELDS } = require('./remarks');
const { getSheet } = require('./sheetDefinitions');
const { validateReport } = require('./sheetValidation');
const {
  TEST_LIBRARY,
  buildAssessmentContext,
  buildOverallStatus,
  buildTestSummaries,
  calculateSoilSummary,
  getCorrosionAllowancePct,
  getStandardGuidance,
  summarizePanels,
  summarizeTowerGroups,
  buildTowerGroupKey,
  deriveElectrodeAssessment,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveEarthContinuityAssessment,
  deriveGroundResistanceAssessment,
  deriveStepTouchAssessment,
  deriveTowerFootingAssessment
} = require('./reportModel');

const ROW_DERIVERS = {
  continuityTest: deriveContinuityAssessment,
  loopImpedanceTest: deriveLoopImpedanceAssessment,
  prospectiveFaultCurrent: deriveFaultCurrentAssessment,
  riserIntegrityTest: deriveRiserAssessment,
  groundResistance: deriveGroundResistanceAssessment,
  stepTouchVoltage: deriveStepTouchAssessment,
  earthContinuityTest: deriveEarthContinuityAssessment
};

// Criteria that make up each test's Basis of Assessment.
const TEST_CRITERIA = {
  soilResistivity: ['SOIL-METHOD', 'SOIL-CLASS'],
  electrodeResistance: ['PIT-ROD', 'PIT-STRIP', 'PIT-PLATE', 'PIT-METHOD', 'PIT-BAND', 'PIT-PRIORITY', 'PIT-CONN'],
  continuityTest: ['CONT-LIMIT'],
  loopImpedanceTest: ['LOOP-LIMIT'],
  prospectiveFaultCurrent: ['PFC-RULE'],
  riserIntegrityTest: ['RISER-LIMIT'],
  groundResistance: ['GND-LIMIT', 'PIT-METHOD'],
  stepTouchVoltage: ['ST-LIMIT', 'ST-METHOD'],
  earthContinuityTest: ['CONT-LIMIT'],
  towerFootingResistance: ['TOWER-ZSAT']
};

// Tests with a pass/fail unit that feeds the action counts (soil has none).
const ACTION_TESTS = ['continuityTest', 'loopImpedanceTest', 'prospectiveFaultCurrent', 'riserIntegrityTest', 'groundResistance', 'stepTouchVoltage', 'earthContinuityTest', 'towerFootingResistance'];

const arrayOf = (value) => (Array.isArray(value) ? value : []);
const ISO_DATE = /^\d{4}-\d{2}-\d{2}$/;

function citation(criterion) {
  const clause = criterion.clause && criterion.clause !== '-' ? `, Cl. ${criterion.clause}` : '';
  return `${criterion.source}${clause} (${describeCriterion(criterion).basisLabel})`;
}

function limitBasis(criterionId, profile) {
  return criterionId ? describeCriterion(profile.get(criterionId)) : null;
}

function sheetLabel(testId) {
  const sheet = getSheet(testId);
  return sheet && sheet.code !== 'legacy' ? sheet.code : (TEST_LIBRARY.find((test) => test.id === testId) || {}).shortLabel || testId;
}

function buildModules({ report, soil, summaries, assessments }) {
  const modules = [];
  TEST_LIBRARY.filter((test) => report?.tests?.[test.id]).forEach((test) => {
    const base = { id: test.id, label: test.label, shortLabel: test.shortLabel || test.label };
    if (test.id === 'soilResistivity') {
      const count = soil.locations.length;
      modules.push({
        ...base,
        tone: soil.category.tone,
        badge: soil.category.label,
        detail:
          soil.overallAverage === null
            ? soil.pendingLocations && soil.pendingLocations.length
              ? `Pending: reading unit not selected for ${soil.pendingLocations.join(', ')}.`
              : `Awaiting numeric readings across ${count} soil location${count === 1 ? '' : 's'}.`
            : `${soil.overallAverage} ohm-m mean across ${count} location${count === 1 ? '' : 's'}${soil.pendingLocations && soil.pendingLocations.length ? `; ${soil.pendingLocations.join(', ')} excluded (reading unit not selected)` : ''}.`,
        complete: soil.overallAverage !== null
      });
      return;
    }

    const summary = summaries[test.id] || { total: 0, healthy: 0, unhealthy: 0, notAssessed: 0 };
    const parts = summary.total ? [`${summary.healthy}/${summary.total} within limit`] : [];
    if (summary.unhealthy) {
      parts.push(`${summary.unhealthy} above limit`);
    }
    if (test.id === 'electrodeResistance') {
      const rows = assessments.electrodeResistance.rows;
      const p1 = rows.filter((row) => row.priority === 'P1').length;
      const p2 = rows.filter((row) => row.priority === 'P2').length;
      if (p1 || p2) {
        parts.push(`P1 ${p1} / P2 ${p2}`);
      }
    }
    if (summary.notAssessed) {
      parts.push(`${summary.notAssessed} not assessed`);
    }
    const module = {
      ...base,
      tone: summary.total ? summary.tone : 'neutral',
      badge: summary.total ? summary.label : 'Pending',
      detail: parts.length ? parts.join(' • ') : 'Awaiting measurement values.',
      complete: summary.total > 0
    };
    if (test.id === 'towerFootingResistance') {
      const rows = assessments.towerFootingResistance.rows;
      module.groupCount = rows.length;
      module.completeGroups = rows.filter((row) => Number.isFinite(row.totalImpedanceZt) && Number.isFinite(row.totalCurrentItotal)).length;
    }
    modules.push(module);
  });
  return modules;
}

function buildActions(assessments) {
  const attention = (rows) => rows.filter((row) => !['healthy', 'neutral'].includes(row.status.tone)).length;
  const critical = (rows) => rows.filter((row) => row.status.tone === 'critical').length;
  const electrodes = assessments.electrodeResistance.rows;
  const electrodeP1 = electrodes.filter((row) => row.priority === 'P1').length;
  const electrodeP2 = electrodes.filter((row) => row.priority === 'P2').length;
  const counts = {};
  ACTION_TESTS.forEach((id) => {
    const rows = (assessments[id] && assessments[id].rows) || [];
    counts[id] = { attention: attention(rows), critical: critical(rows) };
  });
  return {
    electrodeP1,
    electrodeP2,
    continuityAttention: counts.continuityTest.attention,
    loopAttention: counts.loopImpedanceTest.attention,
    faultAttention: counts.prospectiveFaultCurrent.attention,
    riserAttention: counts.riserIntegrityTest.attention,
    groundAttention: counts.groundResistance.attention,
    stepTouchAttention: counts.stepTouchVoltage.attention,
    earthContinuityAttention: counts.earthContinuityTest.attention,
    towerFootingAttention: counts.towerFootingResistance.attention,
    criticalTotal: electrodeP1 + ACTION_TESTS.reduce((sum, id) => sum + counts[id].critical, 0),
    total: electrodeP1 + electrodeP2 + ACTION_TESTS.reduce((sum, id) => sum + counts[id].attention, 0)
  };
}

// Banner above the electrode table: where the per-pit limits come from.
function buildElectrodeBanner(context, profile) {
  const rod = profile.get('PIT-ROD');
  const clauses = ['PIT-PLATE', 'PIT-ROD', 'PIT-STRIP'].map((id) => profile.get(id).clause).join(' / ');
  const basis = `${rod.source} Cl. ${clauses}`;
  if (context.rho === null) {
    return `Permissible limits are calculated per pit (${basis}). Enter soil resistivity, or an assumed value with its source, to calculate them.`;
  }
  const origin = context.rhoAssumed ? `assumed: ${context.rhoSource}` : context.rhoSource.toLowerCase();
  return `Permissible limit is calculated per pit from site rho = ${context.rho} ohm-m (${origin}). Basis: ${basis}.`;
}

// Calibration recorded per day on a FINAL report (assembled from its dailies).
function calibrationItems(report) {
  const seen = new Set();
  const items = [];
  arrayOf(report?.calibrationRecords).forEach((record) => {
    const { calibrationDue, workDate } = record || {};
    if (!ISO_DATE.test(String(calibrationDue || '')) || !ISO_DATE.test(String(workDate || '')) || calibrationDue >= workDate) {
      return;
    }
    const key = [record.sheet, record.instrument, record.serialNo, calibrationDue, workDate].join('|');
    if (seen.has(key)) {
      return;
    }
    seen.add(key);
    items.push({
      id: `CAL_EXPIRED:${key}`,
      code: 'CAL_EXPIRED',
      sheet: record.sheet,
      blocksApproval: true,
      text: `${sheetLabel(record.sheet)} (${record.collectionDate || workDate}): instrument calibration expired on ${calibrationDue}, before the work date ${workDate}.`
    });
  });
  return items;
}

// What the reviewer must look at before approval (standard 6.4). Items with
// blocksApproval stop a FINAL report being approved until acknowledged with a note.
function buildReviewItems(source, validation, assessments) {
  const items = [];
  validation.warnings
    .filter((warning) => warning.review)
    .forEach((warning) => items.push({ id: `${warning.code}:${warning.path}`, code: warning.code, sheet: warning.sheet, blocksApproval: Boolean(warning.blocksApproval), text: warning.message }));
  Object.entries(assessments).forEach(([testId, entry]) => {
    ((entry && entry.rows) || []).forEach((row, index) => {
      (row.flags || []).filter((flag) => flag.review).forEach((flag) => {
        items.push({ id: `${flag.code}:${testId}.${index}`, code: flag.code, sheet: testId, blocksApproval: false, text: `${sheetLabel(testId)} row ${index + 1}: ${flag.message}.` });
      });
    });
  });
  const known = new Set(items.map((item) => item.id));
  calibrationItems(source).forEach((item) => {
    if (!known.has(item.id)) {
      items.push(item);
    }
  });
  return items;
}

function assessReport(report) {
  const source = report && typeof report === 'object' ? report : {};
  const context = buildAssessmentContext(source);
  const { profile } = context;
  const soilSummary = calculateSoilSummary(source);
  const soil = { ...soilSummary, corrosionAllowancePct: getCorrosionAllowancePct(soilSummary.overallAverage, profile) };
  const guidance = getStandardGuidance(profile);

  const assessments = {};

  assessments.electrodeResistance = {
    rows: arrayOf(source.electrodeResistance).map((row) => {
      const result = deriveElectrodeAssessment(row, context);
      return {
        status: result.status,
        standard: result.standard,
        priority: result.priority,
        limit: result.limit,
        limitBasis: limitBasis(result.limitCriterionId, profile),
        measured: result.measured,
        measuredBasis: result.measuredBasis,
        conditionCodes: result.conditionCodes,
        fop: result.fop,
        flags: result.flags,
        comment: result.comment,
        remark: result.comment
      };
    })
  };

  Object.entries(ROW_DERIVERS).forEach(([id, derive]) => {
    const rows = arrayOf(source[id]);
    const field = REMARK_FIELDS[id];
    assessments[id] = {
      rows: rows.map((row) => {
        const result = derive(row, context);
        const typed = field ? engineerRemark(row?.[field]) : '';
        return {
          ...result,
          measured: result.measured ?? null,
          typedRemark: typed,
          // The engineer's own words win; otherwise the remark follows the live status.
          remark: typed || result.comment
        };
      })
    };
    if (id === 'loopImpedanceTest' || id === 'prospectiveFaultCurrent') {
      assessments[id].panels = summarizePanels(rows, derive, context).map((panel) => ({ startIndex: panel.startIndex, status: panel.status }));
    }
  });

  const groups = arrayOf(source.towerFootingResistance);
  const towerSummaries = summarizeTowerGroups(groups);
  assessments.towerFootingResistance = {
    rows: groups.map((group) => {
      const result = deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group)), context);
      const typed = engineerRemark(group?.remarks);
      return {
        status: result.status,
        standard: result.standard,
        totalImpedanceZt: result.totalImpedanceZt,
        totalCurrentItotal: result.totalCurrentItotal,
        zsat: result.zsat,
        flags: result.flags,
        conditionCodes: result.conditionCodes,
        defectText: result.defectText,
        comment: result.comment,
        typedRemark: typed,
        remark: typed || [result.comment, result.defectText].filter(Boolean).join(' – ')
      };
    })
  };

  const summaries = buildTestSummaries(source, context);
  const overall = buildOverallStatus(source, context);
  const modules = buildModules({ report: source, soil, summaries, assessments });
  const actions = buildActions(assessments);
  const completed = modules.filter((module) => module.complete).length;
  const measurementCoverage = {
    total: modules.length,
    completed,
    score: modules.length ? Math.round((completed / modules.length) * 100) : 0
  };
  const healthySignals = modules.filter((module) => module.id !== 'soilResistivity' && module.complete && module.tone === 'healthy').length;

  let tone = 'healthy';
  let label = 'Healthy Snapshot';
  if (!modules.length) {
    tone = 'neutral';
    label = 'Awaiting Selection';
  } else if (actions.criticalTotal > 0) {
    tone = 'critical';
    label = 'Action Required';
  } else if (actions.total > 0) {
    tone = 'warning';
    label = 'Review Needed';
  }

  const selectedIds = modules.map((module) => module.id);
  const criteriaIds = [...new Set(selectedIds.flatMap((id) => TEST_CRITERIA[id] || []))];
  const validation = validateReport(source, context);

  return {
    profile: {
      id: profile.profileId,
      description: profile.description,
      clientSpecDocument: profile.clientSpecDocument,
      register: profile.register
    },
    basis: {
      rho: context.rho,
      rhoAssumed: context.rhoAssumed,
      rhoSource: context.rhoSource
    },
    soil,
    tests: assessments,
    summaries,
    overall,
    modules,
    actions,
    measurementCoverage,
    healthySignals,
    tone,
    label,
    validation,
    reviewItems: buildReviewItems(source, validation, assessments),
    banners: { electrode: buildElectrodeBanner(context, profile) },
    standards: guidance,
    references: Object.fromEntries(selectedIds.map((id) => [id, (TEST_CRITERIA[id] || []).map((criterionId) => citation(profile.get(criterionId)))])),
    basisOfAssessment: criteriaIds.map((criterionId) => describeCriterion(profile.get(criterionId)))
  };
}

module.exports = { assessReport, TEST_CRITERIA };
