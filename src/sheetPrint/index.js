// Printing measurement sheets for a site visit: project + work date + selected tests in, one PDF
// out (A4, portrait or landscape per sheet as in the approved v3 layout). Every printed page is
// issued a unique sheet ID + QR code and recorded in `issued_sheets`, with the pre-printed rows.
const crypto = require('crypto');
const { SPECS, PRINTABLE_SHEET_IDS } = require('./spec');
const { renderSheetsPdf, pageCapacity } = require('./render');
const { newSheetId } = require('./ids');
const { earlierReports, collectPreItems } = require('./preprint');

class SheetPrintError extends Error {
  constructor(message, status = 400) {
    super(message);
    this.name = 'SheetPrintError';
    this.status = status;
  }
}

const asText = (value) => String(value === undefined || value === null ? '' : value).trim();

function validate(body) {
  const input = body || {};
  const project = {
    projectNo: asText(input.project && input.project.projectNo),
    clientName: asText(input.project && input.project.clientName),
    siteLocation: asText(input.project && input.project.siteLocation),
    workOrder: asText(input.project && input.project.workOrder),
    zohoProjectId: asText(input.project && input.project.zohoProjectId)
  };
  const workDate = asText(input.workDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(workDate)) {
    throw new SheetPrintError('Choose the work date to print on the sheets.');
  }
  if (!project.projectNo) {
    throw new SheetPrintError('Project number is required to print sheets.');
  }
  const sheets = [...new Set(Array.isArray(input.sheets) ? input.sheets.map(asText) : [])];
  if (!sheets.length) {
    throw new SheetPrintError('Select at least one test to print.');
  }
  const unknown = sheets.filter((id) => !PRINTABLE_SHEET_IDS.includes(id));
  if (unknown.length) {
    throw new SheetPrintError(`These tests have no printed sheet: ${unknown.join(', ')}.`);
  }
  // Print in the approved order (MS-01..MS-09) whatever order they were ticked in.
  const ordered = PRINTABLE_SHEET_IDS.filter((id) => sheets.includes(id));
  return { project, workDate, sheets: ordered, taskId: asText(input.taskId), engineerName: asText(input.engineerName), reportId: asText(input.reportId) };
}

// The engineer printed on the sheets: the Zoho task owner when a task is chosen, else the
// free-text name. A task that cannot be resolved falls back to the free text, with a warning.
async function resolveEngineer({ taskId, engineerName }, project, getTasks) {
  const warnings = [];
  if (taskId && getTasks && project.zohoProjectId) {
    try {
      const tasks = await getTasks(project.zohoProjectId);
      const task = (tasks || []).find((item) => String(item.id) === taskId);
      if (task && asText(task.ownerName)) {
        return { engineer: asText(task.ownerName), warnings };
      }
      warnings.push(`Zoho task ${taskId} has no owner; used the engineer name entered.`);
    } catch (_error) {
      warnings.push('Zoho tasks could not be read; used the engineer name entered.');
    }
  } else if (taskId) {
    warnings.push('No Zoho project is linked, so the task owner could not be looked up.');
  }
  return { engineer: engineerName, warnings };
}

// Plans the pages: one or more per selected sheet (continuation pages when more pre-printed
// points exist than one page holds), each with its own sheet ID.
function planPages(valid, reports, engineer) {
  const earlier = earlierReports(reports, valid.project, valid.workDate, valid.reportId);
  const batchId = `PB-${crypto.randomBytes(5).toString('hex').toUpperCase()}`;
  const issuedAt = new Date().toISOString();
  const pages = [];
  valid.sheets.forEach((sheetKey) => {
    const spec = SPECS[sheetKey];
    const { rows, items: perPage } = pageCapacity(spec);
    const preItems = collectPreItems(earlier, sheetKey, spec.preprint);
    const count = Math.max(1, Math.ceil(preItems.length / perPage));
    for (let page = 0; page < count; page += 1) {
      const slice = preItems.slice(page * perPage, (page + 1) * perPage);
      pages.push({
        sheetKey,
        sheetId: newSheetId(),
        batchId,
        sheetCode: spec.code,
        docNo: spec.docNo,
        project: valid.project,
        workDate: valid.workDate,
        engineer,
        reportId: valid.reportId,
        page: page + 1,
        pages: count,
        itemOffset: page * perPage,
        rowOffset: page * perPage * spec.group,
        rowCount: rows,
        group: spec.group,
        preItems: slice,
        preRows: slice,
        issuedAt
      });
    }
  });
  return pages;
}

// body: { project, workDate, sheets[], taskId?, engineerName?, reportId? }
// deps: { store, getTasks(zohoProjectId) }
async function printSheets(body, { store, getTasks }) {
  const valid = validate(body);
  const { engineer, warnings } = await resolveEngineer(valid, valid.project, getTasks);
  const pages = planPages(valid, store.listReports(), engineer);
  const { buffer } = await renderSheetsPdf(pages, {
    title: `ESHA measurement sheets – ${valid.project.projectNo} – ${valid.workDate}`
  });
  // Recorded only once the PDF rendered, so a failed print never leaves IDs behind.
  store.addIssuedSheets(pages);
  return { buffer, pages, batchId: pages[0].batchId, warnings, engineer };
}

module.exports = { printSheets, planPages, validate, SheetPrintError };
