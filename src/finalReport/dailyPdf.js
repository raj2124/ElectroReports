// ESHA Daily Progress Report (standard 5.2): one page per daily report, to the client.
// No CAP and no final conclusions: statuses are preliminary until the final report is approved.
const { buildFinalReportData, assertConsistent, defectsOf } = require('./data');
const { calculateSoilSummary } = require('../reportModel');
const { ReportDoc, PAGE, COLORS, statusColor, formatDate } = require('./layout');

const pad2 = (value) => String(value).padStart(2, '0');

// Distinct points per test measured up to and including the given daily report.
function cumulativeCounts(datas, upToIndex) {
  const seen = {};
  datas.slice(0, upToIndex + 1).forEach((data) => {
    data.tests.forEach((test) => {
      seen[test.id] = seen[test.id] || new Set();
      test.points.forEach((point) => seen[test.id].add(point.label));
    });
    seen.soilResistivity = seen.soilResistivity || new Set();
    data.soil.locations.forEach((location) => seen.soilResistivity.add(location.name));
  });
  return Object.fromEntries(Object.entries(seen).map(([id, set]) => [id, set.size]));
}

function progressRows(data, cumulative, planned, report) {
  const rows = [];
  if (report.tests && report.tests.soilResistivity) {
    rows.push({ id: 'soilResistivity', test: 'Soil resistivity', unit: 'locations', today: data.soil.locations.length });
  }
  data.tests.forEach((test) => rows.push({ id: test.id, test: test.label, unit: test.unit, today: test.points.length }));
  return rows.map((row) => {
    const plannedQty = planned[row.id];
    const total = cumulative[row.id] || 0;
    return [
      row.test,
      row.unit,
      String(row.today),
      String(total),
      plannedQty === undefined ? { text: 'Not set', color: COLORS.neutral } : String(plannedQty),
      plannedQty ? `${Math.round((100 * total) / plannedQty)}%` : '-'
    ];
  });
}

function notableRows(data) {
  const rows = [];
  data.tests.forEach((test) => {
    test.points.forEach((point) => {
      const defects = defectsOf(test.id, point);
      if (defects.length) {
        const status = test.id === 'electrodeResistance' ? point.priority || point.statusLabel : point.statusLabel;
        rows.push([point.label, defects.join(' '), { text: status, color: statusColor(status), bold: true }]);
      }
    });
  });
  return rows;
}

const dateOf = (report) => String(report.collectionDate || (report.project && report.project.reportDate) || '');

// Electrode limits need the site soil resistivity. A day that measured none uses the project's
// soil data from the same or EARLIER days only (never a later day), merged by location with the
// latest reading winning. The limits are then preliminary, and the report says whose rho they use.
// Returns the report as the day should be assessed, plus { rho, date } when a fallback was used.
function withEarlierSoil(report, sortedReports) {
  if (calculateSoilSummary(report).overallAverage > 0) {
    return { report, fallback: null };
  }
  const assumed = Number(report.project && report.project.assumedSoilResistivity);
  if (assumed > 0) {
    return { report, fallback: null };
  }
  const date = dateOf(report);
  const byLocation = new Map();
  let latestDate = '';
  sortedReports.filter((other) => dateOf(other) <= date).forEach((other) => {
    const locations = (other.soilResistivity && other.soilResistivity.locations) || [];
    locations.forEach((location, index) => {
      const single = { ...other, soilResistivity: { locations: [location] } };
      if (calculateSoilSummary(single).overallAverage > 0) {
        byLocation.set(String(location.name || `Location ${index + 1}`).trim().toLowerCase(), location);
        latestDate = dateOf(other);
      }
    });
  });
  if (!byLocation.size) {
    return { report, fallback: null };
  }
  const merged = { ...report, soilResistivity: { locations: [...byLocation.values()] } };
  const rho = calculateSoilSummary(merged).overallAverage;
  const source = `preliminary – ρ from ${formatDate(latestDate)}`;
  // The borrowed rho enters as the report's assumed value, so every limit, comment and flag
  // that depends on it carries the source; the day's own soil figures stay empty.
  const effective = { ...report, project: { ...report.project, assumedSoilResistivity: rho, assumedSoilResistivitySource: source } };
  return { report: effective, fallback: { rho, date: latestDate, source } };
}

function drawDay(doc, report, data, index, total, cumulative, planned, fallback) {
  const project = data.project;
  const date = report.collectionDate || project.reportDate;
  doc.meta.documentNo = `${project.documentNo}_DPR_${pad2(index + 1)}`;
  doc.meta.revision = formatDate(date);
  doc.newPage();
  doc.heading('Daily Progress Report', 1);
  doc.table({
    columns: [['Item', 25], ['Detail', 75]],
    rows: [
      ['Client / unit', `${project.clientName || '-'}${project.unit ? ` – ${project.unit}` : ''}`],
      ['Project', project.projectNo || '-'],
      ['Survey day', `${formatDate(date)}  (Day ${index + 1} of ${total} recorded)`],
      ['Engineers on site', project.engineerName || '-'],
      ['Data collection mode', project.dataCollectionMode || '-']
    ]
  });
  doc.heading('Progress', 2);
  doc.table({
    columns: [['Test', 32], ['Unit', 12], ['Done today', 14], ['Cumulative', 14], ['Planned', 14], ['Complete', 14]],
    rows: progressRows(data, cumulative, planned, report),
    alignRight: [2, 3, 4, 5]
  });
  doc.heading('Notable observations today', 2);
  const notable = notableRows(data);
  if (notable.length) {
    doc.table({ columns: [['Point', 20], ['Observation', 64], ['Preliminary status', 16]], rows: notable });
  } else {
    doc.paragraph('No defects or points above limit were found today.');
  }
  if (fallback) {
    doc.note(`Electrode limits are preliminary – ρ from ${formatDate(fallback.date)} (${fallback.rho.toFixed(2)} Ω·m, project soil data of the same or an earlier day). No soil resistivity was measured today. Statuses are preliminary until the final report is reviewed and approved.`);
  } else if (data.soil.locations.length && data.soil.average !== null) {
    doc.note(`Soil resistivity measured at ${data.soil.locations.length} location${data.soil.locations.length === 1 ? '' : 's'}; site average ${data.soil.average.toFixed(2)} Ω·m. Statuses are preliminary until the final report is reviewed and approved.`);
  } else {
    doc.note('Statuses are preliminary until the final report is reviewed and approved.');
  }
  doc.heading('Plan for next day', 2);
  doc.paragraph(project.nextDayPlan || 'Not recorded.', { align: 'left' });
  doc.space(10);
  doc.table({
    columns: [['Prepared by (ETPL engineer)', 50], ['Acknowledged by (client representative)', 50]],
    rows: [['\n\n\n', '\n\n\n']]
  });
}

// reports: the project's daily reports in date order (or one report). One page each.
async function generateDailyProgressPdf(reports, options = {}) {
  const list = (Array.isArray(reports) ? reports : [reports]).slice().sort((a, b) => String(a.collectionDate || a.project.reportDate).localeCompare(String(b.collectionDate || b.project.reportDate)));
  const resolved = list.map((report) => withEarlierSoil(report, list));
  const datas = resolved.map((entry) => buildFinalReportData(entry.report));
  datas.forEach((data) => assertConsistent(data));
  const planned = Object.assign({}, ...list.map((report) => report.project.plannedQuantities || {}));
  const first = datas[0];
  const approved = list.every((report) => (report.reviewStatus || report.dailyStatus) === 'approved');
  const doc = new ReportDoc({
    title: 'Earthing System Health Assessment',
    unit: first.project.unit,
    documentNo: `${first.project.documentNo}_DPR_01`,
    revision: formatDate(first.surveyDates[0]),
    watermark: options.watermark !== undefined ? options.watermark : approved ? '' : 'DRAFT – NOT APPROVED',
    info: { Subject: 'ESHA daily progress report' }
  });
  list.forEach((report, index) => {
    if (!options.onlyReportId || report.id === options.onlyReportId) {
      drawDay(doc, report, datas[index], index, list.length, cumulativeCounts(datas, index), planned, resolved[index].fallback);
    }
  });
  const pageCount = doc.pageIndex + 1;
  doc.finish({ coverPages: 0 });
  return { buffer: await doc.toBuffer(), pageCount, datas, fallbacks: resolved.map((entry) => entry.fallback) };
}

module.exports = { generateDailyProgressPdf, withEarlierSoil };
