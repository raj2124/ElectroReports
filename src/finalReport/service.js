// Export of daily and final report PDFs for the server: narrative, revisions, logo, status.
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { getReportKind, getReviewStatus, REPORT_KINDS } = require('../reportCollectionModel');
const { buildFinalReportData } = require('./data');
const { generateNarrative, narrativeInput } = require('./narrative');
const { generateFinalReportPdf } = require('./pdf');
const { generateDailyProgressPdf } = require('./dailyPdf');
const { resolveClientLogo } = require('./logos');
const { revisionRows } = require('../review');

const sanitize = (value) => String(value || 'report').replace(/[^A-Za-z0-9._-]+/g, '_');

// The stored narrative is reused while the computed results are unchanged, so exporting twice
// does not call Claude twice and an approved report keeps the text that was reviewed.
async function narrativeFor(store, report, data, options) {
  const hash = crypto.createHash('sha256').update(JSON.stringify(narrativeInput(data))).digest('hex');
  const stored = report.reportNarrative;
  if (stored && stored.inputHash === hash) {
    return stored;
  }
  if (stored && getReviewStatus(report) === 'approved') {
    return stored;
  }
  const narrative = { ...(await generateNarrative(data, options.narrative || {})), inputHash: hash, generatedAt: new Date().toISOString() };
  // Kept with the report, so what the reviewer approved is what every later export prints.
  store.updateReport(report.id, (current) => ({ ...current, reportNarrative: narrative }));
  return narrative;
}

async function exportReportPdf(store, report, { generatedDir, publicPathPrefix = '/generated-pdfs', ...options } = {}) {
  const kind = getReportKind(report);
  const status = getReviewStatus(report);
  let result;
  let name;
  if (kind === REPORT_KINDS.FINAL) {
    const data = buildFinalReportData(report);
    const narrative = await narrativeFor(store, report, data, options);
    result = await generateFinalReportPdf(report, { narrative, revisions: revisionRows(store, report), clientLogo: resolveClientLogo(store, report), status });
    name = `${result.data.project.documentNo}_R${Number.isInteger(report.revision) ? report.revision : 0}`;
  } else if (kind === REPORT_KINDS.DAILY) {
    const siblings = store.listProjectReports(report.projectId, { kind: REPORT_KINDS.DAILY });
    result = await generateDailyProgressPdf(siblings.length ? siblings : [report], { onlyReportId: report.id });
    name = `${result.datas[0].project.documentNo}_DPR_${sanitize(report.collectionDate || report.project.reportDate)}`;
  } else {
    throw new Error('Standalone reports use the legacy layout.');
  }
  fs.mkdirSync(generatedDir, { recursive: true });
  const fileName = `${sanitize(name)}-${Date.now()}.pdf`;
  const filePath = path.join(generatedDir, fileName);
  fs.writeFileSync(filePath, result.buffer);
  return { fileName, filePath, pdfUrl: `${publicPathPrefix}/${fileName}`, narrative: result.narrative ? { aiWritten: result.narrative.aiWritten, note: result.narrative.note } : undefined };
}

module.exports = { exportReportPdf, narrativeFor };
