// Review and approval of daily and final reports (standard 6.4): draft -> in_review ->
// (marked as checked) -> approved. This module holds the rules; reportStore applies them inside
// its transactions and records every step in report_transitions.
//
// Cover initials come from those transitions: Prepared = submit to review, Checked = mark as
// checked, Approved = approve. A report approved once is never edited: re-issue copies it as
// the next revision (R1, R2, ...) and the revision table lists every issue.

const { assessReport } = require('./assessmentService');
const { buildFinalReportData, checkConsistency } = require('./finalReport/data');
const { getReportKind, getReviewStatus, REPORT_KINDS } = require('./reportCollectionModel');

// Which step a stored transition row was. Rows written before the `action` column existed
// are classified from their statuses.
function transitionAction(row) {
  if (row.action) {
    return row.action;
  }
  if (row.toStatus === 'approved') return 'approve';
  if (row.toStatus === 'in_review') return 'submit';
  return 'return';
}

function initialsOf(name) {
  const text = String(name || '').trim();
  if (!text) {
    return '';
  }
  if (/^[A-Z]{2,4}$/.test(text)) {
    return text;
  }
  return text
    .split(/[^A-Za-z]+/)
    .filter(Boolean)
    .map((word) => word[0].toUpperCase())
    .join('')
    .slice(0, 4);
}

// Transitions of the current review round: everything after the latest "submit".
function currentRound(transitions) {
  const rows = transitions.map((row) => ({ ...row, action: transitionAction(row) }));
  let start = 0;
  rows.forEach((row, index) => {
    if (row.action === 'submit') {
      start = index;
    }
  });
  return rows.slice(start);
}

function lastOf(rows, action) {
  return rows.filter((row) => row.action === action).pop() || null;
}

function p1PitPhotoGaps(report) {
  const assessment = assessReport(report);
  const rows = Array.isArray(report.electrodeResistance) ? report.electrodeResistance : [];
  return assessment.tests.electrodeResistance.rows
    .map((entry, index) => ({ entry, row: rows[index] || {} }))
    .filter(({ entry, row }) => entry.priority === 'P1' && !(Array.isArray(row.rowPhotos) && row.rowPhotos.length) && !String(row.photoRef || '').trim())
    .map(({ row }) => row.tag || row.location || 'pit');
}

// The reviewer checklist. `auto` is what the app verified itself; manual items must be ticked
// by the checker before the report can be marked as checked.
function buildChecklist(report, { consistency, data }) {
  const kind = getReportKind(report);
  const items = [];
  items.push({
    id: 'counts',
    text: 'Counts match across scope, tables, summary and CAP',
    auto: consistency.ok ? 'pass' : 'fail',
    detail: consistency.ok ? 'The app compared every count; the reviewer confirms.' : consistency.problems.join(' | '),
    manual: true,
    blocking: true
  });
  if (kind === REPORT_KINDS.FINAL) {
    const gaps = data.tests.some((test) => test.id === 'electrodeResistance') ? p1PitPhotoGaps(report) : [];
    items.push({
      id: 'p1',
      text: 'Every P1 has a photo and an action',
      auto: gaps.length ? 'warn' : 'pass',
      detail: gaps.length ? `No photo reference for: ${gaps.join(', ')}.` : 'Every P1 pit has a photo reference and a CAP line.',
      manual: true
    });
    items.push({
      id: 'assumptions',
      text: 'Assumptions declared; limits and clauses shown',
      auto: data.basis.rhoAssumed && !String(data.basis.rhoSource || '').trim() ? 'fail' : 'pass',
      detail: data.basis.rhoAssumed ? `Soil resistivity is assumed: ${data.basis.rhoSource || 'no source given'}.` : 'No assumed values; the Basis of Assessment lists every limit.',
      manual: true
    });
    items.push({ id: 'lastyear', text: 'Readings differing > 50% from last year\'s value for the same tag are explained', auto: null, detail: 'The app has no previous-year values; check against the last report.', manual: true });
    items.push({ id: 'otherunit', text: 'No text referring to another unit, client or project', auto: null, detail: '', manual: true });
  }
  return items;
}

// The independent sanity check of a final report's narrative (src/aiChecks.js): 'passed', 'issues', or 'not run' (never run,
// could not run, or the narrative has changed since it ran). A check kept by an older version of the app (no status) is read as
// it was: consistent or not.
function sanityState(report) {
  const check = report.aiSanityCheck;
  if (!check) return { state: 'not run', reason: '', flags: [] };
  if (check.status === 'error') return { state: 'not run', reason: check.reason || 'The check could not run.', flags: [] };
  if (check.narrativeHash) {
    let now = '';
    try { now = require('./aiChecks').sanityHashOf(report); } catch (_error) { now = ''; }
    if (!now || now !== check.narrativeHash) return { state: 'not run', reason: 'The narrative changed after the check ran.', flags: [] };
  }
  const flags = Array.isArray(check.flags) ? check.flags : [];
  const issues = check.status ? check.status === 'issues' : check.consistent === false;
  return { state: issues ? 'issues' : 'passed', reason: '', flags, notes: check.notes || '', model: check.model || '', at: check.generatedAt || check.at || '' };
}

const ACK_LABEL = {
  calibration: 'Calibration acknowledged',
  sanityCheck: 'Sanity check not run - acknowledged',
  sanityIssues: 'Sanity check issues - acknowledged',
  p1Photos: 'P1 photo missing - acknowledged'
};

// Hard stops: no note can clear these, the record itself must be fixed. A FINAL report with no
// instrument recorded cannot be approved, because no reading could then be traced to a calibrated
// instrument.
function approvalBlockers(report) {
  const blockers = [];
  if (getReportKind(report) === REPORT_KINDS.FINAL && !buildFinalReportData(report).instruments.length) {
    blockers.push({ code: 'NO_INSTRUMENTS', text: 'No instrument is recorded on the measurement sheets. Record the instrument, serial number and calibration date before approving.' });
  }
  return blockers;
}

// Notes the approver must write before approval. `provided` is filled by the caller.
function requiredAcknowledgements(report) {
  const items = [];
  const calibration = assessReport(report).reviewItems.filter((item) => item.blocksApproval);
  if (getReportKind(report) === REPORT_KINDS.FINAL) {
    const photoGaps = report.tests && report.tests.electrodeResistance ? p1PitPhotoGaps(report) : [];
    if (photoGaps.length) {
      items.push({ key: 'p1Photos', text: `No photo reference for P1 pit${photoGaps.length === 1 ? '' : 's'} ${photoGaps.join(', ')}: attach the photo, or acknowledge that the report is issued without it.`, required: true });
    }
    if (calibration.length) {
      items.push({ key: 'calibration', text: calibration.map((item) => item.text).join(' '), required: true });
    }
    const sanity = sanityState(report);
    if (sanity.state === 'not run') {
      items.push({ key: 'sanityCheck', text: `Sanity check not run: the independent check of the narrative against the computed results was not performed${sanity.reason ? ` (${sanity.reason})` : ''}.`, required: true });
    } else if (sanity.state === 'issues') {
      items.push({ key: 'sanityIssues', text: `The independent sanity check found ${sanity.flags.length} possible contradiction${sanity.flags.length === 1 ? '' : 's'} between the narrative and the computed results: ${sanity.flags.map((flag) => `${flag.area ? `${flag.area}: ` : ''}${flag.issue}`).join(' | ')} Confirm that you have read them and that the report may be issued.`, required: true });
    }
  }
  return items;
}

function reviewWarnings(report, data) {
  const warnings = [...data.warnings];
  const narrative = report.reportNarrative;
  if (narrative && !narrative.aiWritten) {
    warnings.push(narrative.note || 'Narrative not AI-written.');
  } else if (!narrative && getReportKind(report) === REPORT_KINDS.FINAL) {
    warnings.push('Narrative has not been generated yet; the PDF export generates it.');
  }
  return warnings;
}

// Revision rows for the cover: R0, R1, ... each with the initials of its own review.
function revisionRows(store, report) {
  const chain = [];
  let cursor = report;
  while (cursor) {
    chain.unshift(cursor);
    cursor = cursor.revisionOf ? store.getReport(cursor.revisionOf) : null;
  }
  return chain.map((entry, index) => {
    const round = (store.listTransitions(entry.id) || []).map((row) => ({ ...row, action: transitionAction(row) }));
    return {
      revision: `R${Number.isInteger(entry.revision) ? entry.revision : index}`,
      date: entry.approvedAt ? entry.approvedAt.slice(0, 10) : '',
      issue: index === 0 ? 'First Issue' : entry.revisionReason || 'Re-issue',
      prepared: initialsOf((lastOf(round, 'submit') || {}).actor),
      checked: initialsOf((lastOf(round, 'check') || {}).actor),
      approved: initialsOf((lastOf(round, 'approve') || {}).actor)
    };
  });
}

// Full names, designations and dates of the latest revision's review, for the document control page.
function approvalBlock(store, report) {
  const rows = (store.listTransitions(report.id) || []).map((row) => ({ ...row, action: transitionAction(row) }));
  const entry = (action) => {
    const row = lastOf(rows, action);
    return row ? { name: row.actor || '', designation: row.designation || '', date: String(row.at || '').slice(0, 10) } : {};
  };
  return { prepared: entry('submit'), checked: entry('check'), approved: entry('approve') };
}

function buildReviewState(store, report) {
  const transitions = store.listTransitions(report.id).map((row) => ({ ...row, action: transitionAction(row) }));
  const round = currentRound(transitions);
  const status = getReviewStatus(report);
  let data;
  let consistency;
  try {
    data = buildFinalReportData(report);
    consistency = checkConsistency(data);
  } catch (error) {
    data = { warnings: [], tests: [], basis: {} };
    consistency = { ok: false, problems: [`The report could not be assessed: ${error.message}`], warnings: [] };
  }
  const sanityOf = sanityState(report);
  const checklist = buildChecklist(report, { consistency, data }).map((item) => ({ ...item, ticked: Boolean(report.review && report.review.ticks && report.review.ticks[item.id]) }));
  const checkedRow = lastOf(round, 'check');
  const acknowledgements = requiredAcknowledgements(report);
  const approveRow = lastOf(transitions, 'approve');
  return {
    reportId: report.id,
    kind: getReportKind(report),
    status,
    revision: Number.isInteger(report.revision) ? report.revision : 0,
    documentNo: data.project ? data.project.documentNo : '',
    checked: status === 'in_review' && checkedRow ? { by: checkedRow.actor, at: checkedRow.at, note: checkedRow.note } : null,
    checklist,
    acknowledgements,
    blockers: status === 'approved' ? [] : approvalBlockers(report),
    consistency,
    warnings: reviewWarnings(report, data),
    narrative: report.reportNarrative ? { aiWritten: Boolean(report.reportNarrative.aiWritten), model: report.reportNarrative.model || '', note: report.reportNarrative.note || '' } : null,
    sanityCheck: sanityOf.state,
    sanity: sanityOf,
    transitions,
    initials: {
      prepared: initialsOf((lastOf(round, 'submit') || lastOf(transitions, 'submit') || {}).actor),
      checked: initialsOf((checkedRow || {}).actor),
      approved: initialsOf((approveRow || {}).actor)
    },
    available: {
      submit: status === 'draft',
      check: status === 'in_review' && !checkedRow,
      approve: status === 'in_review' && Boolean(checkedRow),
      return: status === 'in_review',
      reissue: status === 'approved'
    }
  };
}

module.exports = { sanityState, ACK_LABEL, approvalBlockers, approvalBlock, transitionAction, initialsOf, currentRound, lastOf, buildChecklist, requiredAcknowledgements, buildReviewState, revisionRows };
