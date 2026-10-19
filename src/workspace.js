// Project workspace (Session 6C.4): one workspace per Zoho project, listing its site days with
// each daily report's status, the cumulative progress against the planned quantities, and the
// final reports. The rules here decide what a new day carries over from the project, which
// projects are active, and what assembling the final report returns.

const { getReportKind, getReviewStatus, summarizeDailyCollection, getCollectionDate, REPORT_KINDS } = require('./reportCollectionModel');
const { TEST_LIBRARY } = require('./reportModel');
const { buildFinalReportData } = require('./finalReport/data');
const { cumulativeCounts } = require('./finalReport/dailyPdf');
const { currentRound, lastOf, transitionAction } = require('./review');

const text = (value) => String(value === undefined || value === null ? '' : value).trim();

// Project details a new day takes over from the project. Day-specific fields are NOT carried:
// the work date, the plan for the next day (it belongs to the day it was written on) and the
// ground condition, which is observed afresh each day.
const DAY_ONLY_PROJECT_FIELDS = ['reportDate', 'nextDayPlan', 'groundCondition'];

// Fields of the project record kept when a project is created or edited.
const PROJECT_FIELDS = [
  'projectNo', 'clientName', 'siteLocation', 'workOrder', 'engineerName', 'clientRepresentative', 'unit', 'unitCode', 'clientId',
  'acceptanceProfile', 'soilSpacingPreset', 'assumedSoilResistivity', 'assumedSoilResistivitySource', 'plannedQuantities',
  'zohoProjectId', 'zohoProjectName', 'zohoProjectOwner', 'zohoProjectStage', 'dataCollectionMode', 'siteAccess'
];

function pickProjectFields(input) {
  const picked = {};
  PROJECT_FIELDS.forEach((key) => {
    const value = input && input[key];
    if (value !== undefined && value !== null && (typeof value === 'object' || text(value) !== '')) {
      picked[key] = typeof value === 'object' ? { ...value } : text(value);
    }
  });
  return picked;
}

// The key of a workspace: the Zoho project, or the project number for a project not in Zoho.
function findProject(store, input) {
  const zohoId = text(input && input.zohoProjectId);
  const projectNo = text(input && input.projectNo).toLowerCase();
  return (
    store.listProjects().find((project) => (zohoId ? text(project.zohoProjectId) === zohoId : !text(project.zohoProjectId) && text(project.projectNo).toLowerCase() === projectNo)) || null
  );
}

// One workspace per Zoho project: returns the existing one, or creates it.
function ensureProject(store, input) {
  const fields = pickProjectFields(input);
  if (!fields.projectNo || !fields.clientName) {
    const error = new Error('A project needs a project number and a client name.');
    error.code = 'PROJECT_INVALID';
    throw error;
  }
  const existing = findProject(store, fields);
  if (existing) {
    return { project: existing, created: false };
  }
  return { project: store.addProject(fields), created: true };
}

const dayStatusOf = (store, report) => {
  const status = getReviewStatus(report);
  if (status === 'in_review') {
    const checked = lastOf(currentRound((store.listTransitions(report.id) || []).map((row) => ({ ...row, action: transitionAction(row) }))), 'check');
    return checked ? 'checked' : 'in_review';
  }
  return status;
};

const byDate = (a, b) => getCollectionDate(a).localeCompare(getCollectionDate(b)) || String(a.createdAt).localeCompare(String(b.createdAt));

function finalsOf(store, projectId) {
  return store
    .listProjectReports(projectId, { kind: REPORT_KINDS.FINAL })
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))
    .map((report) => ({
      id: report.id,
      revision: Number.isInteger(report.revision) ? report.revision : 0,
      status: dayStatusOf(store, report),
      createdAt: report.createdAt,
      approvedAt: report.approvedAt || '',
      includedDates: (report.collectionManifest && report.collectionManifest.includedDates) || [],
      flaggedDates: (report.collectionManifest && report.collectionManifest.flaggedDates) || []
    }));
}

// Active until the latest final report is approved; then Completed. A day recorded after that
// approval puts the project back among the active ones (the approved report no longer covers it).
function projectState(dailies, finals) {
  const latest = finals[0];
  if (!latest || latest.status !== 'approved') {
    return 'active';
  }
  const newerDay = dailies.some((report) => String(report.createdAt) > String(latest.createdAt));
  return newerDay ? 'active' : 'completed';
}

function progressOf(dailies, planned) {
  const datas = [];
  dailies.forEach((report) => {
    try {
      datas.push(buildFinalReportData(report));
    } catch (_error) {
      datas.push({ tests: [], soil: { locations: [] } });
    }
  });
  const counts = datas.length ? cumulativeCounts(datas, datas.length - 1) : {};
  const meta = new Map();
  meta.set('soilResistivity', { label: 'Soil resistivity', unit: 'locations' });
  TEST_LIBRARY.forEach((test) => meta.set(test.id, { label: test.label, unit: test.unitLabel || 'points' }));
  datas.forEach((data) => data.tests.forEach((test) => meta.set(test.id, { label: test.label, unit: test.unit })));
  const ids = [...new Set([...Object.keys(planned || {}), ...Object.keys(counts)])];
  return ids
    .filter((id) => meta.has(id))
    .map((id) => {
      const plannedQty = Number.isFinite(Number(planned && planned[id])) && planned && planned[id] !== undefined ? Number(planned[id]) : null;
      const done = counts[id] || 0;
      return { testId: id, label: meta.get(id).label, unit: meta.get(id).unit, planned: plannedQty, done, percent: plannedQty ? Math.round((100 * done) / plannedQty) : null };
    });
}

// What the browser shows for the project: days, progress, final reports, and the template a
// new day starts from.
function buildWorkspace(store, projectId) {
  const project = store.getProject(projectId);
  if (!project) {
    return null;
  }
  const dailies = store.listProjectReports(projectId, { kind: REPORT_KINDS.DAILY }).sort(byDate);
  const finals = finalsOf(store, projectId);
  const latestDay = dailies[dailies.length - 1];
  const planned = project.plannedQuantities || (latestDay && latestDay.project && latestDay.project.plannedQuantities) || {};
  const manifest = summarizeDailyCollection(dailies);
  const carried = { ...(latestDay && latestDay.project ? latestDay.project : {}), ...project };
  DAY_ONLY_PROJECT_FIELDS.forEach((key) => delete carried[key]);
  ['id', 'createdAt', 'updatedAt'].forEach((key) => delete carried[key]);
  return {
    project,
    state: projectState(dailies, finals),
    days: dailies.map((report) => ({ reportId: report.id, date: getCollectionDate(report), status: dayStatusOf(store, report), revision: Number.isInteger(report.revision) ? report.revision : 0, updatedAt: report.updatedAt })),
    progress: progressOf(dailies, planned),
    planned,
    finals,
    assembly: { includedDates: manifest.includedDates, flaggedDates: manifest.flaggedDates, hasUnapproved: manifest.hasUnapproved, totalDailyReports: manifest.totalDailyReports },
    dayTemplate: { project: carried, tests: latestDay ? { ...latestDay.tests } : {} }
  };
}

// Home screen row of one project.
function summarizeProject(store, project) {
  const dailies = store.listProjectReports(project.id, { kind: REPORT_KINDS.DAILY }).sort(byDate);
  const finals = finalsOf(store, project.id);
  const days = dailies.length;
  return {
    id: project.id,
    projectNo: project.projectNo,
    clientName: project.clientName,
    siteLocation: project.siteLocation || '',
    zohoProjectId: project.zohoProjectId || '',
    state: projectState(dailies, finals),
    days,
    approvedDays: dailies.filter((report) => getReviewStatus(report) === 'approved').length,
    lastDay: days ? getCollectionDate(dailies[days - 1]) : '',
    finalStatus: finals[0] ? finals[0].status : '',
    finalRevision: finals[0] ? finals[0].revision : null,
    updatedAt: project.updatedAt
  };
}

function listProjectSummaries(store) {
  return store
    .listProjects()
    .map((project) => summarizeProject(store, project))
    .sort((a, b) => String(b.lastDay || b.updatedAt).localeCompare(String(a.lastDay || a.updatedAt)));
}

// One daily per date. Returns the existing report of that date, or null.
function findDay(store, projectId, workDate) {
  return store.listProjectReports(projectId, { kind: REPORT_KINDS.DAILY }).find((report) => getCollectionDate(report) === workDate) || null;
}

// Stitch every daily into a new DRAFT final; an earlier final is kept untouched.
function assembleFor(store, projectId, { actor = '', userId = '' } = {}) {
  const { report, provenance } = store.assembleFinal(projectId, { actor, userId });
  const manifest = report.collectionManifest || {};
  return {
    report,
    includedDates: manifest.includedDates || [],
    flaggedDates: manifest.flaggedDates || [],
    hasUnapproved: Boolean(manifest.hasUnapproved),
    supersededReadings: provenance.length
  };
}

module.exports = { DAY_ONLY_PROJECT_FIELDS, PROJECT_FIELDS, pickProjectFields, findProject, ensureProject, buildWorkspace, summarizeProject, listProjectSummaries, findDay, assembleFor, dayStatusOf };
