// Computed results for the final report (standard Part 5.1).
//
// Every number, status, priority and CAP line in the PDF comes from here, built once
// from assessReport(): nothing in the renderer recomputes a count. Several views of the
// same records are kept separately (summary table, per-point result tables, electrode
// priority block, scope table, CAP) so checkConsistency can compare them and block
// release if any of them differ.

const { assessReport, TEST_CRITERIA } = require('../assessmentService');
const { TEST_LIBRARY, buildTestSummaries, buildOverallStatus, buildAssessmentContext, asLooseNumber } = require('../reportModel');
const { CONDITION_CODE_LEGEND } = require('../sheetDefinitions');

const CODE_TEXT = Object.fromEntries(CONDITION_CODE_LEGEND.map((entry) => [entry.code, entry.text]));

// Per test: how a pass/fail unit is named, and how a point is labelled from its source row.
const TEST_META = {
  electrodeResistance: { unit: 'pits', label: (row) => row.tag || row.location || '-' },
  continuityTest: { unit: 'points', label: (row) => [row.mainLocation, row.measurementPoint].filter(Boolean).join(' / ') || '-' },
  loopImpedanceTest: { unit: 'panels', panel: true, label: (row) => [row.location, row.feederTag].filter(Boolean).join(' ') || '-' },
  prospectiveFaultCurrent: { unit: 'panels', panel: true, label: (row) => [row.location, row.feederTag].filter(Boolean).join(' ') || '-' },
  riserIntegrityTest: { unit: 'points', label: (row) => [row.mainLocation, row.measurementPoint].filter(Boolean).join(' / ') || '-' },
  groundResistance: { unit: 'points', label: (row) => row.tag || row.location || '-' },
  stepTouchVoltage: { unit: 'points', label: (row) => [row.location, row.measurementType].filter(Boolean).join(' ') || '-' },
  towerFootingResistance: { unit: 'towers', label: (row) => row.mainLocationTower || '-' },
  earthContinuityTest: { unit: 'points', label: (row) => [row.tag, row.locationBuildingName].filter(Boolean).join(' ') || '-' }
};

// Retired legacy sheet stays reportable only when an old report still carries data for it.
const REPORT_TESTS = TEST_LIBRARY.filter((test) => test.id !== 'soilResistivity');

const CODE_ACTION = {
  C1: 'Repair or replace the corroded electrode/strip (IS 3043:2018 Cl. 14.2)',
  C2: 'Reconnect the strip / restore the test link and bonding (IS 3043:2018 Cl. 25.6.3.4)',
  C3: 'Provide a second earth connection (IS 3043:2018 Cl. 4.2)',
  C4: 'Expose and restore the pit chamber with a removable cover; measure and report',
  C5: 'Repair or replace the damaged electrode',
  C6: 'Clean and re-tighten the loose connection; re-test',
  C7: 'Clean, tin and protect the joint (IS 3043:2018 Cl. 25.6.3.4)',
  C8: 'Fix a permanent tag/label at the point',
  OTH: 'Review the engineer remark and assign action'
};

const arrayOf = (value) => (Array.isArray(value) ? value : []);
const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const fmt2 = (value) => (Number.isFinite(value) ? value.toFixed(2) : '-');
const isMeasured = (tone) => ['healthy', 'warning', 'critical'].includes(tone);

function codesOf(value) {
  return arrayOf(value).map(text).filter(Boolean);
}

function pointFromRow({ testId, row, index, status, extra = {} }) {
  const meta = TEST_META[testId];
  return {
    index,
    label: meta.label(row || {}),
    statusLabel: status.label,
    tone: status.tone,
    counted: isMeasured(status.tone),
    withinLimit: status.tone === 'healthy' || status.withinLimit === true,
    codes: [],
    flags: [],
    ...extra
  };
}

// Pass/fail points of one test, from the assessment rows the UI already reads.
function collectPoints(testId, report, assessment) {
  const rows = arrayOf(report[testId]);
  const assessed = arrayOf(assessment.tests[testId] && assessment.tests[testId].rows);
  const meta = TEST_META[testId];
  if (meta.panel) {
    return arrayOf(assessment.tests[testId].panels).map((panel) => {
      const phaseRows = assessed.slice(panel.startIndex, panel.startIndex + 3);
      const source = rows[panel.startIndex] || {};
      return pointFromRow({
        testId,
        row: source,
        index: panel.startIndex / 3,
        status: panel.status,
        extra: {
          device: [source.deviceType, source.deviceRating].filter(Boolean).join(' '),
          breakingCapacity: asLooseNumber(source.breakingCapacity),
          phases: phaseRows.map((entry, i) => ({ phase: text(rows[panel.startIndex + i] && rows[panel.startIndex + i].measuredPoints) || ['R-E', 'Y-E', 'B-E'][i], measured: entry.measured, status: entry.status.label })),
          pscNotRecorded: phaseRows.some((entry) => entry.pscNotRecorded)
        }
      });
    });
  }
  return assessed.map((entry, index) => {
    const row = rows[index] || {};
    return pointFromRow({
      testId,
      row,
      index,
      status: entry.status,
      extra: {
        measured: entry.measured ?? null,
        limit: entry.limit ?? null,
        priority: entry.priority || null,
        codes: codesOf(entry.conditionCodes),
        flags: arrayOf(entry.flags),
        comment: entry.typedRemark || entry.comment || '',
        defectText: entry.defectText || '',
        zt: entry.totalImpedanceZt ?? null,
        zsat: entry.zsat ?? null,
        pitType: row.electrodeType || '',
        location: row.location || ''
      }
    });
  });
}

// Real defects of one point (condition codes, over-limit, electrode P1/P2, tower "Not given"
// feet). One CAP line per point carries all of them, so a point is "in the CAP" exactly
// when this returns something. Data gaps are not defects: see gapsOf.
function defectsOf(testId, point) {
  const lines = [];
  if (testId === 'electrodeResistance') {
    if (point.priority === 'P1' || point.priority === 'P2') {
      lines.push(point.comment);
    }
  } else if (point.tone === 'critical' && point.phases) {
    const over = point.phases.filter((phase) => phase.status === '>Permissible Limit');
    lines.push(`Above limit on ${over.map((phase) => `${phase.phase} (${fmt2(phase.measured)})`).join(', ') || 'at least one phase'}.`);
  } else if (point.tone === 'critical') {
    lines.push(point.comment || `${point.statusLabel}.`);
  } else if (point.statusLabel === 'Marginal') {
    lines.push('Total footing impedance is marginal (above Zsat, within 1.2 x Zsat).');
  }
  point.codes.forEach((code) => lines.push(`${code}: ${CODE_TEXT[code] || code}`));
  point.flags.forEach((flag) => {
    if (flag.code === 'FOOT_NOT_GIVEN') {
      lines.push(flag.message);
    }
  });
  return [...new Set(lines.filter(Boolean))];
}

// Data gaps of one point: the record is incomplete, so the point could not be fully assessed.
// Grouped into one CAP line per test and gap.
const GAPS = {
  PSC_NOT_RECORDED: { text: 'PSC (L-N) not recorded; PFC (L-PE) assessed alone.', action: 'Record PSC (L-N) at these panels and re-assess' },
  NOT_ASSESSABLE: { text: 'Device breaking capacity not recorded, so the fault current cannot be assessed.', action: 'Record the device nameplate breaking capacity (kA) and re-assess' },
  FOP_UNCONFIRMED: { text: '61.8% reading not confirmed by the 50% / 70% readings.', action: 'Re-test with a corrected probe layout' }
};

function gapsOf(testId, point) {
  const gaps = [];
  if (point.statusLabel === 'Not Assessable') {
    gaps.push('NOT_ASSESSABLE');
  }
  if (point.pscNotRecorded) {
    gaps.push('PSC_NOT_RECORDED');
  }
  if (point.flags.some((flag) => flag.code === 'FOP_DISAGREE')) {
    gaps.push('FOP_UNCONFIRMED');
  }
  return gaps;
}

function actionFor(testId, point) {
  const actions = point.codes.map((code) => CODE_ACTION[code]).filter(Boolean);
  if (point.flags.some((flag) => flag.code === 'FOOT_NOT_GIVEN')) {
    actions.push('Provide the missing foot-to-earth connection and re-test the tower');
  }
  if (testId === 'electrodeResistance' && point.statusLabel === '>Permissible Limit') {
    actions.push('Rectify within 30 days: repair/replace electrode and strip, restore test link and bonding; re-test');
  } else if (testId === 'electrodeResistance' && point.statusLabel === 'Not Measurable') {
    actions.push('Expose the pit, measure and report within 30 days');
  } else if (testId === 'electrodeResistance' && point.priority === 'P2' && !actions.length) {
    actions.push('Rectify in the next maintenance cycle; monitor at the next annual test');
  } else if (point.tone === 'critical' && !actions.length) {
    actions.push('Inspect terminations and bonding; rectify and re-test');
  } else if (point.statusLabel === 'Marginal' && !actions.length) {
    actions.push('Improve the footing earth and re-test');
  }
  return [...new Set(actions)].join('; ') || 'Review and rectify';
}

function buildCap(tests) {
  const lines = [];
  tests.forEach((test) => {
    test.points.forEach((point) => {
      const defects = defectsOf(test.id, point);
      if (!defects.length) {
        return;
      }
      lines.push({
        kind: 'point',
        testId: test.id,
        test: test.label,
        point: point.label,
        observation: defects.join(' '),
        action: actionFor(test.id, point),
        priorityOrStatus: test.id === 'electrodeResistance' ? point.priority || point.statusLabel : point.statusLabel,
        clause: test.id === 'electrodeResistance' ? 'IS 3043:2018 Cl. 14.2' : test.basisClause
      });
    });
  });
  tests.forEach((test) => {
    Object.keys(GAPS).forEach((gap) => {
      const affected = test.points.filter((point) => gapsOf(test.id, point).includes(gap)).map((point) => point.label);
      if (!affected.length) {
        return;
      }
      lines.push({
        kind: 'gap',
        gap,
        testId: test.id,
        test: test.label,
        point: affected.join(', '),
        affectedPoints: affected,
        observation: GAPS[gap].text,
        action: GAPS[gap].action,
        priorityOrStatus: 'Data gap',
        clause: test.basisClause
      });
    });
  });
  return lines;
}

function buildInstruments(report) {
  const records = arrayOf(report.calibrationRecords).length
    ? arrayOf(report.calibrationRecords).map((record) => ({ instrument: record.instrument, serialNo: record.serialNo, calibrationDue: record.calibrationDue, workDate: record.workDate || record.collectionDate }))
    : Object.values(report.sheetHeaders || {}).map((header) => ({ instrument: header.instrument, serialNo: header.serialNo, calibrationDue: header.calibrationDue, workDate: text(report.project && report.project.reportDate) }));
  const seen = new Set();
  const rows = [];
  records.forEach((record) => {
    if (!text(record.instrument) && !text(record.calibrationDue)) {
      return;
    }
    const key = [record.instrument, record.serialNo, record.calibrationDue, record.workDate].map(text).join('|');
    if (!seen.has(key)) {
      seen.add(key);
      rows.push({ instrument: text(record.instrument), serialNo: text(record.serialNo), calibrationDue: text(record.calibrationDue), workDate: text(record.workDate) });
    }
  });
  rows.sort((a, b) => a.workDate.localeCompare(b.workDate) || a.instrument.localeCompare(b.instrument));
  rows.forEach((row) => {
    row.expired = /^\d{4}-\d{2}-\d{2}$/.test(row.calibrationDue) && /^\d{4}-\d{2}-\d{2}$/.test(row.workDate) && row.calibrationDue < row.workDate;
  });
  return rows;
}

function documentNumber(project) {
  const seq = text(project.projectNo).split(/[-_ ]/)[0] || '000';
  const unit = text(project.unitCode || project.unit).replace(/[^A-Za-z0-9]/g, '').toUpperCase().slice(0, 8) || 'SITE';
  return { seq, unit, documentNo: `${seq}_${unit}_ESHA` };
}

// Planned quantities are not stored by the app yet: read project.plannedQuantities
// ({ testId: n }) when present. A missing figure is reported as "not set" and warned about.
function buildScope(report, tests) {
  const planned = (report.project && report.project.plannedQuantities) || {};
  return tests.map((test) => {
    const value = Number(planned[test.id]);
    return {
      testId: test.id,
      test: test.label,
      unit: test.unit,
      planned: Number.isFinite(value) ? value : null,
      completed: test.points.length,
      notAssessed: test.points.filter((point) => !point.counted).length
    };
  });
}

function surveyDates(report) {
  const manifest = report.collectionManifest;
  const dates = manifest ? arrayOf(manifest.includedDates) : [];
  const own = text(report.project && report.project.reportDate);
  return dates.length ? dates : own ? [own] : [];
}

function buildFinalReportData(report) {
  const assessment = assessReport(report);
  const project = report.project || {};
  const selected = REPORT_TESTS.filter((test) => report.tests && report.tests[test.id]);

  const tests = selected.map((test) => {
    const meta = TEST_META[test.id];
    const clauseCriterion = (assessment.basisOfAssessment || []).find((entry) => (TEST_CRITERIA[test.id] || []).includes(entry.id));
    return {
      id: test.id,
      label: test.label,
      sheetCode: test.sheetCode,
      unit: meta.unit,
      basisClause: clauseCriterion ? `${clauseCriterion.source}${clauseCriterion.clause && clauseCriterion.clause !== '-' ? `, Cl. ${clauseCriterion.clause}` : ''}` : '',
      points: collectPoints(test.id, report, assessment)
    };
  });

  const modelSummaries = buildTestSummaries(report, buildAssessmentContext(report));
  const summaryTable = tests.map((test) => {
    const summary = modelSummaries[test.id] || { total: 0, healthy: 0, unhealthy: 0, notAssessed: 0, passRate: null, label: 'Pending' };
    return {
      testId: test.id,
      test: test.label,
      unit: test.unit,
      total: summary.total,
      healthy: summary.healthy,
      unhealthy: summary.unhealthy,
      notAssessed: summary.notAssessed,
      passPct: summary.passRate === null ? null : Math.round(summary.passRate * 1000) / 10,
      status: summary.label
    };
  });

  const electrode = tests.find((test) => test.id === 'electrodeResistance');
  const priorities = { P1: [], P2: [], P3: [] };
  if (electrode) {
    electrode.points.forEach((point) => point.priority && priorities[point.priority].push(point.label));
  }

  const soil = assessment.soil;
  const cap = buildCap(tests);
  const instruments = buildInstruments(report);
  const expired = instruments.filter((row) => row.expired);
  if (expired.length) {
    const names = expired.map((row) => `${`${row.instrument} ${row.serialNo}`.trim()} (due ${row.calibrationDue}, used ${row.workDate})`);
    cap.push({
      kind: 'calibration',
      gap: 'CAL_EXPIRED',
      testId: 'instruments',
      test: 'Instruments',
      point: names.join(', '),
      affectedPoints: names,
      observation: 'Instrument calibration had expired on the work date.',
      action: 'Reviewer to confirm validity of the readings taken with these instruments; re-test if in doubt',
      priorityOrStatus: 'Data gap',
      clause: 'ET-D03-REG-002'
    });
  }

  const scope = buildScope(report, tests);
  const warnings = [];
  scope.filter((row) => row.planned === null).forEach((row) => warnings.push(`Planned quantity not set for ${row.test}.`));
  if (!text(project.unitCode)) {
    warnings.push('Unit code not set: the document number uses "SITE". Set the unit code on the project step.');
  }

  return {
    project: { ...project, ...documentNumber(project) },
    kind: report.kind || 'standalone',
    surveyDates: surveyDates(report),
    soil: { average: soil.overallAverage, category: soil.category && soil.category.label, locations: arrayOf(soil.locations), corrosionAllowancePct: soil.corrosionAllowancePct },
    basis: { rho: assessment.basis.rho, rhoAssumed: assessment.basis.rhoAssumed, rhoSource: assessment.basis.rhoSource, criteria: assessment.basisOfAssessment || [] },
    tests,
    summaryTable,
    overall: buildOverallStatus(report, buildAssessmentContext(report)),
    electrodePriority: { P1: { count: priorities.P1.length, tags: priorities.P1 }, P2: { count: priorities.P2.length, tags: priorities.P2 }, P3: { count: priorities.P3.length, tags: priorities.P3 } },
    scope,
    instruments,
    cap,
    reviewItems: assessment.reviewItems || [],
    warnings,
    fmt2
  };
}

// ---- Consistency check ------------------------------------------------------------
class ConsistencyError extends Error {
  constructor(problems) {
    super(`Report counts are inconsistent: ${problems.join(' | ')}`);
    this.name = 'ConsistencyError';
    this.code = 'REPORT_INCONSISTENT';
    this.problems = problems;
  }
}

// Recounts every view from the per-point tables and compares: summary table, results
// tables, electrode priority block, scope table and CAP. Returns { ok, problems, warnings }.
function checkConsistency(data) {
  const problems = [];
  const same = (what, a, b) => {
    if (a !== b) {
      problems.push(`${what}: ${a} vs ${b}`);
    }
  };

  data.tests.forEach((test) => {
    const row = data.summaryTable.find((entry) => entry.testId === test.id);
    if (!row) {
      problems.push(`${test.label}: missing from the summary table`);
      return;
    }
    const counted = test.points.filter((point) => point.counted);
    const within = counted.filter((point) => point.withinLimit).length;
    same(`${test.label} total (summary vs results table)`, row.total, counted.length);
    same(`${test.label} within limit (summary vs results table)`, row.healthy, within);
    same(`${test.label} above limit (summary vs results table)`, row.unhealthy, counted.length - within);
    same(`${test.label} not assessed (summary vs results table)`, row.notAssessed, test.points.length - counted.length);
    const scope = data.scope.find((entry) => entry.testId === test.id);
    same(`${test.label} completed (scope vs results table)`, scope ? scope.completed : undefined, test.points.length);
    same(`${test.label} CAP defect lines (CAP vs defect points)`, data.cap.filter((line) => line.kind === 'point' && line.testId === test.id).length, test.points.filter((point) => defectsOf(test.id, point).length).length);
    Object.keys(GAPS).forEach((gap) => {
      const expected = test.points.filter((point) => gapsOf(test.id, point).includes(gap)).length;
      const lines = data.cap.filter((line) => line.kind === 'gap' && line.testId === test.id && line.gap === gap);
      same(`${test.label} ${gap} points (CAP vs results table)`, lines.reduce((sum, line) => sum + line.affectedPoints.length, 0), expected);
      same(`${test.label} ${gap} CAP lines (one per test)`, lines.length, expected ? 1 : 0);
    });
  });
  data.summaryTable.forEach((row) => {
    if (!data.tests.some((test) => test.id === row.testId)) {
      problems.push(`${row.test}: in the summary table but has no results section`);
    }
  });

  const electrode = data.tests.find((test) => test.id === 'electrodeResistance');
  if (electrode) {
    ['P1', 'P2', 'P3'].forEach((priority) => {
      same(`Electrode ${priority} (priority block vs results table)`, data.electrodePriority[priority].count, electrode.points.filter((point) => point.priority === priority).length);
    });
  }

  const calLines = data.cap.filter((line) => line.kind === 'calibration');
  const expiredCount = (data.instruments || []).filter((row) => row.expired).length;
  same('Expired calibrations (CAP vs instruments table)', calLines.reduce((sum, line) => sum + line.affectedPoints.length, 0), expiredCount);
  same('Expired calibration CAP lines (one)', calLines.length, expiredCount ? 1 : 0);

  const assessedTotal = data.summaryTable.reduce((sum, row) => sum + row.total, 0);
  same('Overall assessed total (overall vs summary table)', data.overall.total, assessedTotal);
  same('Overall within limit (overall vs summary table)', data.overall.healthy, data.summaryTable.reduce((sum, row) => sum + row.healthy, 0));

  return { ok: !problems.length, problems, warnings: data.warnings || [] };
}

function assertConsistent(data) {
  const result = checkConsistency(data);
  if (!result.ok) {
    throw new ConsistencyError(result.problems);
  }
  return result;
}

module.exports = { buildFinalReportData, checkConsistency, assertConsistent, ConsistencyError, defectsOf, gapsOf, REPORT_TESTS, TEST_META };
