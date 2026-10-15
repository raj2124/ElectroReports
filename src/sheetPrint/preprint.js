// Pre-printed rows: the tags and locations already recorded for the same project in earlier
// reports, so the engineer finds them printed on the sheet instead of writing them again.
const asText = (value) => String(value === undefined || value === null ? '' : value).trim();
const norm = (value) => asText(value).toLowerCase();

function reportDate(report) {
  return asText(report && report.collectionDate) || asText(report && report.project && report.project.reportDate);
}

// A report belongs to the project when the Zoho project matches, or the project number does.
function sameProject(report, project) {
  const other = (report && report.project) || {};
  const zoho = asText(project.zohoProjectId);
  if (zoho && asText(other.zohoProjectId) === zoho) {
    return true;
  }
  const number = norm(project.projectNo);
  return Boolean(number) && norm(other.projectNo) === number;
}

// Earlier reports of the project (dated before the work date, or undated), oldest first.
// The report being worked on (`excludeReportId`) is never a source for its own sheets.
function earlierReports(reports, project, workDate, excludeReportId = '') {
  return reports
    .filter((report) => report && report.id !== excludeReportId && sameProject(report, project))
    .filter((report) => !reportDate(report) || reportDate(report) < workDate)
    .sort((a, b) => (reportDate(a) || '9999').localeCompare(reportDate(b) || '9999') || asText(a.createdAt).localeCompare(asText(b.createdAt)));
}

// One pre-printed item per physical point (a pit, a feeder, a tower...), de-duplicated by identity
// across days, in order of first appearance; the latest wording of a point is the one printed.
// Rows with no identity text are skipped.
function collectPreItems(reports, sheetKey, rule) {
  const { keys, identity, identityKeys } = rule;
  if (!keys.length) {
    return [];
  }
  const identityOf = (item) => {
    const parts = identityKeys || keys;
    if (identity === 'first') {
      const key = keys.find((candidate) => item[candidate]);
      return key ? `${key}:${norm(item[key])}` : '';
    }
    return parts.map((key) => norm(item[key])).join('|');
  };
  const items = new Map();
  reports.forEach((report) => {
    const rows = Array.isArray(report[sheetKey]) ? report[sheetKey] : [];
    rows.forEach((row) => {
      const item = {};
      keys.forEach((key) => {
        const value = asText(row && row[key]);
        if (value) {
          item[key] = value;
        }
      });
      const id = identityOf(item);
      if (!id || /^[:|]*$/.test(id)) {
        return;
      }
      items.set(id, { ...(items.get(id) || {}), ...item });
    });
  });
  return [...items.values()];
}

module.exports = { earlierReports, collectPreItems, sameProject, reportDate };
