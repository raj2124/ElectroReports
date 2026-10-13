// Multi-day collection model (ADDITIVE).
//
// This module introduces the vocabulary and pure logic for a project whose data
// is collected over several days ("daily reports") and later stitched into one
// "final report". It is deliberately standalone: it imports no app state and is
// imported by nothing in the standalone-report path, so that existing single
// reports keep behaving EXACTLY as before.
//
// The invariant behind "preserve the standalone path exactly": any report that
// predates this model — i.e. has no `kind` field — is treated as `standalone`.
// Nothing here mutates a report; every function returns new values.
//
// Locked product decision encoded here: a final report stitches from ALL daily
// reports (not approved-only). It must record which collection dates it included
// and flag any that were still in draft/in_review, so no day can silently drop.
// That manifest is produced by `summarizeDailyCollection` and attached by
// `assembleFinalReport`.
//
// Row-level stitching of the measurement sheets is intentionally NOT done here;
// it lands in a later step because the PDF and AI layers consume those arrays in
// specific shapes. This step delivers the vocabulary + the inclusion manifest +
// the assembly skeleton (project, tests-union, ordered constituents).

const REPORT_KINDS = Object.freeze({
  STANDALONE: 'standalone',
  DAILY: 'daily',
  FINAL: 'final'
});

// Lifecycle of a daily report. Order matters: it defines review progression and
// is used to rank/flag statuses in the manifest.
const DAILY_STATUSES = Object.freeze({
  DRAFT: 'draft',
  IN_REVIEW: 'in_review',
  APPROVED: 'approved'
});

const DAILY_STATUS_ORDER = [DAILY_STATUSES.DRAFT, DAILY_STATUSES.IN_REVIEW, DAILY_STATUSES.APPROVED];

function asTrimmedString(value) {
  return String(value === undefined || value === null ? '' : value).trim();
}

// Default-to-standalone is the load-bearing rule that protects existing reports:
// an unknown or missing kind reads as standalone, never daily/final.
function getReportKind(report) {
  const kind = asTrimmedString(report && report.kind).toLowerCase();
  if (kind === REPORT_KINDS.DAILY || kind === REPORT_KINDS.FINAL) {
    return kind;
  }
  return REPORT_KINDS.STANDALONE;
}

function isStandaloneReport(report) {
  return getReportKind(report) === REPORT_KINDS.STANDALONE;
}

function isDailyReport(report) {
  return getReportKind(report) === REPORT_KINDS.DAILY;
}

function isFinalReport(report) {
  return getReportKind(report) === REPORT_KINDS.FINAL;
}

// A daily report's review status. Anything unrecognized (including a standalone
// report that has no status at all) reads as `draft` — the safest default,
// because draft is what gets flagged in the manifest.
function getDailyStatus(report) {
  const status = asTrimmedString(report && report.dailyStatus).toLowerCase();
  return DAILY_STATUS_ORDER.includes(status) ? status : DAILY_STATUSES.DRAFT;
}

// Daily AND final reports share the lifecycle draft -> in_review -> approved.
// `reviewStatus` is canonical; `dailyStatus` is read as a fallback for reports
// written before the field was generalised.
function getReviewStatus(report) {
  const status = asTrimmedString((report && (report.reviewStatus || report.dailyStatus)) || '').toLowerCase();
  return DAILY_STATUS_ORDER.includes(status) ? status : DAILY_STATUSES.DRAFT;
}

// Forward one step, or send an in_review report back to draft for rework.
const ALLOWED_TRANSITIONS = Object.freeze({
  draft: ['in_review'],
  in_review: ['approved', 'draft'],
  approved: []
});

function canTransition(fromStatus, toStatus) {
  return (ALLOWED_TRANSITIONS[asTrimmedString(fromStatus).toLowerCase()] || []).includes(
    asTrimmedString(toStatus).toLowerCase()
  );
}

function isApprovedStatus(status) {
  return asTrimmedString(status).toLowerCase() === DAILY_STATUSES.APPROVED;
}

// The date the data was collected in the field. Prefer an explicit
// `collectionDate`; fall back to the project's report date (the standalone
// field) so a daily report created before we add a dedicated picker still sorts
// sensibly. Empty string when nothing is known.
function getCollectionDate(report) {
  const explicit = asTrimmedString(report && report.collectionDate);
  if (explicit) {
    return explicit;
  }
  return asTrimmedString(report && report.project && report.project.reportDate);
}

// Stable chronological ordering. Reports with a known date sort ascending by
// date; ties and unknown dates fall back to createdAt, then id, so ordering is
// deterministic regardless of input order.
function compareByCollectionDate(a, b) {
  const dateA = getCollectionDate(a);
  const dateB = getCollectionDate(b);
  if (dateA !== dateB) {
    if (!dateA) return 1;
    if (!dateB) return -1;
    return dateA < dateB ? -1 : 1;
  }
  const createdA = asTrimmedString(a && a.createdAt);
  const createdB = asTrimmedString(b && b.createdAt);
  if (createdA !== createdB) {
    return createdA < createdB ? -1 : 1;
  }
  return asTrimmedString(a && a.id) < asTrimmedString(b && b.id) ? -1 : 1;
}

// Build the inclusion manifest for a set of daily reports.
//
// This is the direct encoding of the locked decision. It lists every daily
// report that will be stitched (ALL of them, in date order) with its collection
// date and status, and flags any that are not yet approved. `flaggedDates` and
// `hasUnapproved` give the final report an at-a-glance "these days were still in
// draft/in_review when assembled" banner.
function summarizeDailyCollection(dailyReports) {
  const reports = (Array.isArray(dailyReports) ? dailyReports : []).slice().sort(compareByCollectionDate);

  const entries = reports.map((report) => {
    const status = getReviewStatus(report);
    return {
      reportId: asTrimmedString(report && report.id),
      collectionDate: getCollectionDate(report),
      dailyStatus: status,
      flagged: !isApprovedStatus(status)
    };
  });

  const includedDates = [];
  const seenDates = new Set();
  entries.forEach((entry) => {
    if (entry.collectionDate && !seenDates.has(entry.collectionDate)) {
      seenDates.add(entry.collectionDate);
      includedDates.push(entry.collectionDate);
    }
  });

  const flaggedEntries = entries.filter((entry) => entry.flagged);
  const flaggedDates = [];
  const seenFlagged = new Set();
  flaggedEntries.forEach((entry) => {
    const label = entry.collectionDate || '(no date)';
    if (!seenFlagged.has(label)) {
      seenFlagged.add(label);
      flaggedDates.push(label);
    }
  });

  const statusCounts = DAILY_STATUS_ORDER.reduce((acc, status) => {
    acc[status] = entries.filter((entry) => entry.dailyStatus === status).length;
    return acc;
  }, {});

  return {
    totalDailyReports: entries.length,
    includedDates,
    entries,
    flaggedDates,
    hasUnapproved: flaggedEntries.length > 0,
    statusCounts
  };
}


// ---- Row-level stitching ----------------------------------------------------
// A "point" is the physical thing re-tested across days. rowIds are unique per
// day, so identity comes from the sheet's own location fields. Rows whose key is
// empty cannot be matched and are always kept (never superseded).
const norm = (value) => asTrimmedString(value).toLowerCase();
const POINT_KEYS = {
  electrodeResistance: (row) => norm(row.tag) || norm(row.location),
  continuityTest: (row) => [norm(row.mainLocation), norm(row.measurementPoint)].filter(Boolean).join('|'),
  loopImpedanceTest: (row) => (norm(row.location) || norm(row.feederTag) ? [norm(row.location), norm(row.feederTag), norm(row.measuredPoints)].join('|') : ''),
  prospectiveFaultCurrent: (row) => (norm(row.location) || norm(row.feederTag) ? [norm(row.location), norm(row.feederTag), norm(row.measuredPoints)].join('|') : ''),
  riserIntegrityTest: (row) => [norm(row.mainLocation), norm(row.measurementPoint)].filter(Boolean).join('|'),
  earthContinuityTest: (row) => [norm(row.tag), norm(row.locationBuildingName)].filter(Boolean).join('|'),
  // groupId is random per day, so identify a tower by its location.
  towerFootingResistance: (row) => norm(row.mainLocationTower),
  soilResistivity: (location) => norm(location.name)
};
const ROW_SHEETS = Object.keys(POINT_KEYS).filter((sheet) => sheet !== 'soilResistivity');

function overrideKey(sheet, pointKey) {
  return `${sheet}::${pointKey}`;
}

// Resolve one sheet. Candidates for the same point are ordered by collection date;
// the latest wins unless the engineer named another day in `overrides`. Every
// non-chosen candidate is returned as a provenance entry, never dropped.
function stitchSheet(sheet, reports, overrides, overriddenBy, provenance) {
  const groups = new Map();
  const order = [];
  reports.forEach((report) => {
    (Array.isArray(report[sheet]) ? report[sheet] : []).forEach((row) => {
      const pointKey = POINT_KEYS[sheet](row || {});
      const mapKey = pointKey || `__unkeyed__${order.length}`;
      if (!groups.has(mapKey)) {
        groups.set(mapKey, { pointKey, candidates: [] });
        order.push(mapKey);
      }
      groups.get(mapKey).candidates.push({ report, row });
    });
  });

  return order.map((mapKey) => {
    const { pointKey, candidates } = groups.get(mapKey);
    let chosen = candidates[candidates.length - 1];
    let resolution = 'latest_day';
    const forced = pointKey ? overrides[overrideKey(sheet, pointKey)] : undefined;
    if (forced !== undefined) {
      const match = candidates.filter((c) => c.report.id === forced).pop();
      if (match) {
        chosen = match;
        resolution = 'engineer_override';
      }
    }
    candidates.forEach((candidate) => {
      if (candidate === chosen) return;
      provenance.push({
        sheet,
        pointKey,
        keptReportId: asTrimmedString(chosen.report.id),
        keptCollectionDate: getCollectionDate(chosen.report),
        supersededReportId: asTrimmedString(candidate.report.id),
        supersededCollectionDate: getCollectionDate(candidate.report),
        supersededRow: candidate.row,
        resolution,
        overriddenBy: resolution === 'engineer_override' ? overriddenBy || '' : ''
      });
    });
    return chosen.row;
  });
}

// Soil resistivity is structured as named locations; a location is the point.
function stitchSoil(reports, overrides, overriddenBy, provenance) {
  const flattened = reports.map((report) => ({
    id: report.id,
    createdAt: report.createdAt,
    collectionDate: report.collectionDate,
    project: report.project,
    soilResistivity: (report.soilResistivity && report.soilResistivity.locations) || []
  }));
  return { locations: stitchSheet('soilResistivity', flattened, overrides, overriddenBy, provenance) };
}

// Assemble a final report from ALL daily reports in a collection.
//
// Carries project context, unions the selected tests, orders constituents
// chronologically, attaches the inclusion manifest, and stitches every
// measurement sheet (latest day wins per point, engineer override, superseded
// readings returned in `assemblyProvenance` for the store to persist).
function assembleFinalReport(dailyReports, options = {}) {
  const reports = (Array.isArray(dailyReports) ? dailyReports : []).slice().sort(compareByCollectionDate);
  const manifest = summarizeDailyCollection(reports);

  // Tests union: a test belongs on the final report if ANY daily report enabled it.
  const testsUnion = {};
  reports.forEach((report) => {
    const tests = (report && report.tests) || {};
    Object.keys(tests).forEach((testId) => {
      testsUnion[testId] = Boolean(testsUnion[testId]) || Boolean(tests[testId]);
    });
  });

  // Project context: prefer an explicit override, else the latest daily report's
  // project (latest wins so the newest client/site corrections carry through).
  const latest = reports.length ? reports[reports.length - 1] : null;
  const project = options.project || (latest && latest.project) || {};

  const overrides = options.overrides || {};
  const overriddenBy = options.overriddenBy || '';
  const assemblyProvenance = [];
  const sheets = {};
  ROW_SHEETS.forEach((sheet) => {
    sheets[sheet] = stitchSheet(sheet, reports, overrides, overriddenBy, assemblyProvenance);
  });
  sheets.soilResistivity = stitchSoil(reports, overrides, overriddenBy, assemblyProvenance);

  return {
    kind: REPORT_KINDS.FINAL,
    reviewStatus: DAILY_STATUSES.DRAFT,
    project,
    tests: testsUnion,
    ...sheets,
    collectionManifest: manifest,
    constituentReportIds: manifest.entries.map((entry) => entry.reportId).filter(Boolean),
    assemblyProvenance
  };
}

module.exports = {
  REPORT_KINDS,
  DAILY_STATUSES,
  DAILY_STATUS_ORDER,
  getReportKind,
  isStandaloneReport,
  isDailyReport,
  isFinalReport,
  getDailyStatus,
  getReviewStatus,
  canTransition,
  ALLOWED_TRANSITIONS,
  POINT_KEYS,
  isApprovedStatus,
  getCollectionDate,
  compareByCollectionDate,
  summarizeDailyCollection,
  assembleFinalReport
};
