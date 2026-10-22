// ESHA final report (standard Part 5.1), rendered with pdfkit from the computed results in
// ./data.js. Nothing here recomputes a count; the consistency check runs before anything is drawn
// and blocks release if the views of the records differ.
const fs = require('fs');
const { buildFinalReportData, assertConsistent } = require('./data');
const { deterministicNarrative } = require('./narrative');
const { ReportDoc, PAGE, COLORS, statusColor, formatDate, LOGO_PATH } = require('./layout');

const CONSULTANT = ['Elegrow Technology Pvt. Ltd.', '302 Sangini Aspire, Pal, Near Pal RTO,', 'Surat, Gujarat 395009'];

const METHODOLOGY = {
  soilResistivity: ['Soil resistivity', 'Wenner four-probe method: four probes at equal spacing a in a straight line, driven 10–15 cm; current is injected through the outer pair and the potential measured across the inner pair. rho = 2πaR is calculated for each spacing and direction (IS 3043:2018 Cl. 40; IEEE 81-2012 Cl. 7).'],
  electrodeResistance: ['Earth electrode resistance', 'Fall-of-potential method with the electrode isolated from the grid, the potential probe at 61.8% of the current-probe distance and checked at 50% and 70% (IS 3043:2018 Cl. 41; IEEE 81-2012 Cl. 8). The permissible limit is calculated per electrode from the site soil resistivity (Cl. 14.2).'],
  continuityTest: ['Earth continuity and bonding', 'Low-resistance measurement between the reference earth point and each tested point, with the test-lead resistance subtracted (IEEE 81-2012 Cl. 10; IS 3043:2018 Cl. 17.2.3).'],
  loopImpedanceTest: ['Earth fault loop impedance', 'Live phase-to-earth loop impedance (R-E, Y-E, B-E) measured at each panel. A panel is healthy only if all three phases meet the criterion.'],
  prospectiveFaultCurrent: ['Prospective fault current', 'PFC (L-PE) and PSC (L-N) measured at each panel and compared with the breaking capacity of the protective device.'],
  riserIntegrityTest: ['Riser integrity', 'Low-resistance measurement from each riser point towards the equipment and towards the earth grid, net of test-lead resistance.'],
  groundResistance: ['Ground (grid) resistance', 'Fall-of-potential measurement of the in-service grid, compared with the limit of the associated pit or the client value.'],
  stepTouchVoltage: ['Step and touch voltage', 'Injected-current method: measured voltages are scaled by If/It (fault current ÷ test current) before comparison with the tolerable limits (IEEE 81-2012 Cl. 9.3(g)).'],
  towerFootingResistance: ['Tower footing', 'Four feet per tower; the total impedance Zt is the sum of the foot impedances and is compared with the tower acceptance value Zsat. A foot without an earth connection is reported as a defect.'],
  earthContinuityTest: ['Earth continuity (legacy sheet)', 'Point-to-point earth path resistance, as recorded on the legacy continuity sheet.']
};

const PRIORITY_MEANING = {
  P1: 'Exceeds limit, not measurable, or safety defect – rectify within 30 days and re-test',
  P2: 'Warning band or maintenance defect – rectify in the maintenance cycle',
  P3: 'Healthy – routine annual testing'
};

const ANNEXURES = [
  ['A', 'Signed field measurement sheets (scans)'],
  ['B', 'Photographs of P1 points, by pit tag'],
  ['C', 'Instrument calibration certificates'],
  ['D', 'Daily progress reports issued during the survey']
];

const STANDARDS = [
  'IS 3043:2018 – Code of practice for earthing',
  'IEEE 80-2013 – Guide for safety in AC substation grounding',
  'IEEE 81-2012 – Guide for measuring earth resistivity, ground impedance and earth surface potentials of a grounding system',
  'IS/IEC 60364 – Electrical installations of buildings (protection against electric shock; fault-loop and protective-device requirements)',
  'CEA (Measures relating to Safety and Electric Supply) Regulations, 2023'
];

const APPROVAL_ROLES = [['prepared', 'Prepared by'], ['checked', 'Checked by'], ['approved', 'Approved by']];

function resolveRevisions(options, status) {
  if (Array.isArray(options.revisions) && options.revisions.length) {
    return options.revisions;
  }
  return [{ revision: 'R0', date: '', issue: 'First Issue', prepared: '', checked: '', approved: '' }];
}

function drawCover(doc, data, revisions, clientLogo) {
  const pdf = doc.doc;
  doc.newPage();
  const { margin, contentWidth: cw } = PAGE;
  pdf.rect(margin, 60, cw, PAGE.height - 110).lineWidth(1.2).strokeColor(COLORS.navy).stroke();
  const boxW = 200;
  const top = 100;
  const boxes = [
    { x: margin + 20, head: 'Consultant', lines: CONSULTANT, logo: LOGO_PATH },
    { x: PAGE.width - margin - 20 - boxW, head: 'Client', lines: [data.project.clientName || '-', data.project.siteLocation || ''].filter(Boolean), logoBuffer: clientLogo }
  ];
  boxes.forEach((box) => {
    pdf.rect(box.x, top, boxW, 150).lineWidth(0.6).strokeColor(COLORS.grey).stroke();
    try {
      if (box.logo && fs.existsSync(box.logo)) {
        pdf.image(box.logo, box.x + 10, top + 10, { fit: [30, 30] });
      } else if (box.logoBuffer) {
        pdf.image(box.logoBuffer, box.x + 10, top + 10, { fit: [120, 40] });
      }
    } catch (_error) {
      // an unreadable logo leaves the box empty
    }
    pdf.font('Bold').fontSize(10).fillColor(COLORS.navy).text(box.head, box.x + 10, top + 62, { lineBreak: false });
    pdf.font('Body').fontSize(8.5).fillColor(COLORS.ink).text(box.lines.join('\n'), box.x + 10, top + 78, { width: boxW - 20, lineGap: 2 });
  });
  const titleTop = top + 190;
  pdf.rect(margin + 20, titleTop, cw - 40, 70).lineWidth(1).strokeColor(COLORS.navy).stroke();
  pdf.font('Bold').fontSize(16).fillColor(COLORS.navy).text('Earthing System Health Assessment', margin + 20, titleTop + 18, { width: cw - 40, align: 'center' });
  pdf.font('Body').fontSize(11).text(`${data.project.unit ? `${data.project.unit}  ·  ` : ''}Final Report`, margin + 20, titleTop + 42, { width: cw - 40, align: 'center' });

  const columns = ['Revision', 'Date of Submission', 'Issue', 'Prepared By', 'Checked By', 'Approved By'];
  const weights = [10, 17, 29, 14, 14, 16];
  const tableW = cw - 40;
  const widths = weights.map((weight) => (weight / 100) * tableW);
  let y = titleTop + 120;
  const drawRow = (cells, header) => {
    pdf.font(header ? 'Bold' : 'Body').fontSize(header ? 7.5 : 8);
    const height = Math.max(18, ...cells.map((value, i) => pdf.heightOfString(String(value), { width: widths[i] - 6 }) + 10));
    let x = margin + 20;
    cells.forEach((value, i) => {
      if (header) {
        pdf.rect(x, y, widths[i], height).fillColor(COLORS.light).fill();
      }
      pdf.rect(x, y, widths[i], height).lineWidth(0.5).strokeColor(COLORS.grey).stroke();
      pdf.font(header ? 'Bold' : 'Body').fontSize(header ? 7.5 : 8).fillColor(header ? COLORS.navy : COLORS.ink).text(String(value), x + 3, y + 5, { width: widths[i] - 6, align: 'center' });
      x += widths[i];
    });
    y += height;
  };
  drawRow(columns, true);
  revisions.forEach((revision) => drawRow([revision.revision, revision.date ? formatDate(revision.date) : '-', revision.issue, revision.prepared || '-', revision.checked || '-', revision.approved || '-'], false));

  y += 30;
  [
    ['Internal Project No.', data.project.seq],
    ['Zoho Project', data.project.projectNo || '-'],
    ['Document No.', data.project.documentNo],
    // optional: printed only when it was filled in
    ...(String(data.project.workOrder || '').trim() ? [['Work Order No.', data.project.workOrder]] : [])
  ].forEach(([key, value]) => {
    pdf.rect(margin + 20, y, 170, 18).lineWidth(0.5).strokeColor(COLORS.grey).stroke();
    pdf.rect(margin + 190, y, cw - 210, 18).stroke();
    pdf.font('Bold').fontSize(8).fillColor(COLORS.navy).text(key, margin + 26, y + 5, { lineBreak: false });
    pdf.font('Body').fontSize(8.5).fillColor(COLORS.ink).text(String(value), margin + 196, y + 5, { lineBreak: false });
    y += 18;
  });
}

function summaryRows(data) {
  return data.summaryTable.map((row, i) => [
    String(i + 1),
    row.test,
    row.unit,
    String(row.total),
    String(row.healthy),
    String(row.unhealthy),
    String(row.notAssessed),
    row.passPct === null ? '-' : `${row.passPct.toFixed(1)}%`,
    { text: row.notAssessed > 0 ? `${row.status} (${row.notAssessed} not assessed)` : row.status, color: statusColor(row.status), bold: true }
  ]);
}

function drawExecutiveSummary(doc, data, narrative) {
  doc.section('Executive Summary');
  doc.heading('A. Overall summary of tests', 2);
  if (data.summaryTable.length) {
    doc.table({
      columns: [['Sr.', 4], ['Test', 20], ['Unit', 7], ['Assessed', 8], ['Within limit', 7], ['Above limit', 7], ['Not assessed', 8], ['Pass %', 7], ['Overall status', 17]],
      rows: summaryRows(data),
      statusColumns: [8],
      alignRight: [3, 4, 5, 6, 7]
    });
    doc.note('Assessed = points with a reading that could be compared with a limit. Pass % = points within limit ÷ assessed points (Warning-band points are within limit but carry priority P2). Not measurable / not assessable points are excluded and counted separately. Overall status: 0 unhealthy = Healthy; pass rate ≥ 90% = Healthy – Minor Actions; otherwise >Attention Required (ET-D03-REG-002, OVERALL).');
  } else {
    doc.paragraph('No pass/fail tests were selected for this report.');
  }

  if (data.tests.some((test) => test.id === 'electrodeResistance')) {
    doc.heading('B. Earth electrode priority summary', 2);
    doc.table({
      columns: [['Priority', 8], ['Meaning', 45], ['Qty', 6], ['Pits', 41]],
      rows: ['P1', 'P2', 'P3'].map((priority) => [{ text: priority, color: statusColor(priority), bold: true }, PRIORITY_MEANING[priority], String(data.electrodePriority[priority].count), data.electrodePriority[priority].tags.join(', ') || '-']),
      alignRight: [2]
    });
  }

  doc.heading('C. Key findings', 2);
  doc.bullets(narrative.keyFindings.length ? narrative.keyFindings : ['No findings could be derived from the recorded data.']);
  doc.note(
    narrative.aiWritten
      ? 'Findings text is written by the report AI from the computed results above; it cannot change any value or status.'
      : narrative.note
  );
  doc.heading('D. Overall condition', 2);
  doc.paragraph(narrative.conditionStatement);
}

function drawIntroduction(doc, data, report) {
  doc.section('Introduction & Scope of Work');
  const project = data.project;
  const tests = data.tests.map((test) => test.label.toLowerCase());
  if (report.tests && report.tests.soilResistivity) {
    tests.unshift('soil resistivity');
  }
  doc.heading('Purpose', 2);
  doc.paragraph(
    `Elegrow Technology Pvt. Ltd. (ETPL) was engaged by ${project.clientName || 'the client'} to assess the health of the earthing system${project.unit ? ` of ${project.unit}` : ''}${project.siteLocation ? `, ${project.siteLocation}` : ''}. The survey measured ${tests.join(', ') || 'the selected tests'} and compared each result with the acceptance criteria listed in the Basis of Assessment, so that the condition of the earthing system is known and every deficiency has a corrective action and a priority.`
  );
  const team = [...new Set([project.engineerName, project.zohoProjectOwner].map((name) => String(name || '').trim()).filter(Boolean))];
  doc.table({
    columns: [['Item', 25], ['Detail', 75]],
    rows: [
      ['Client', project.clientName || '-'],
      ['Site location', project.siteLocation || '-'],
      ['Unit', project.unit || '-'],
      ['Survey dates', data.surveyDates.length ? `${data.surveyDates.map(formatDate).join(', ')}  (${data.surveyDates.length} site day${data.surveyDates.length === 1 ? '' : 's'})` : '-'],
      ['ETPL team', team.join(', ') || '-'],
      ['Client representative', project.clientRepresentative || '-']
    ],
    boldFirstColumn: true
  });
  doc.heading('Applicable standards', 2);
  doc.bullets(STANDARDS);
  doc.heading('Scope – planned vs completed', 2);
  const rows = data.scope.map((row) => [
    row.test,
    row.unit,
    row.planned === null ? { text: 'Not set', color: COLORS.neutral } : String(row.planned),
    String(row.completed),
    row.notAssessed ? `${row.notAssessed} not measured / not assessed` : ''
  ]);
  if (data.soil.locations.length && report.tests && report.tests.soilResistivity) {
    const planned = project.plannedQuantities && project.plannedQuantities.soilResistivity;
    rows.unshift(['Soil resistivity', 'locations', planned === undefined ? { text: 'Not set', color: COLORS.neutral } : String(planned), String(data.soil.locations.length), '']);
  }
  doc.table({ columns: [['Test', 34], ['Unit', 10], ['Planned', 14], ['Completed', 14], ['Remark', 28]], rows, alignRight: [2, 3] });
}

// "Cl." is added only to a bare clause number; a clause that already names its own standard or
// clause ("IS 3043:2018 Cl. 42 (method)") is printed as it stands.
function citeSource(criterion) {
  const clause = String(criterion.clause || '').trim();
  if (!clause || clause === '-') {
    return criterion.source;
  }
  return /^(?:\d|Table)/.test(clause) ? `${criterion.source}, Cl. ${clause}` : `${criterion.source}; ${clause}`;
}

function drawBasis(doc, data) {
  doc.section('Basis of Assessment');
  doc.paragraph("Every limit, formula and classification used in this report is listed below with its source. Values are taken from the project's acceptance profile (ET-D03-REG-002); where a value is an ETPL acceptance criterion or a client specification rather than a clause of the standard, it is labelled as such.");
  doc.table({
    columns: [['Criterion', 20], ['Rule / value used', 40], ['Basis', 14], ['Source & clause', 26]],
    rows: data.basis.criteria.map((criterion) => [criterion.parameter, criterion.ruleUsed, criterion.basisLabel, citeSource(criterion)])
  });
  doc.heading('Declared assumptions', 2);
  const assumptions = [];
  if (data.basis.rho !== null && data.basis.rho !== undefined) {
    assumptions.push(
      data.basis.rhoAssumed
        ? `Soil resistivity ρ = ${data.basis.rho} Ω·m was ASSUMED for the electrode limits (${data.basis.rhoSource}).`
        : `Soil resistivity used for the electrode limits: ρ = ${data.basis.rho} Ω·m (${data.basis.rhoSource || 'measured'}). No assumed values were used.`
    );
  }
  assumptions.push('Statuses are computed by the application from the recorded readings against the criteria above; no value is typed into this report by hand.');
  doc.bullets(assumptions);
}

function drawInstruments(doc, data, annexures) {
  doc.section('Instruments Used');
  doc.table({
    columns: [['Work date', 10], ['Instrument', 22], ['Serial No.', 12], ['Range', 20], ['Accuracy', 14], ['Calibration due', 11], ['Validity', 11]],
    rows: data.instruments.length
      ? data.instruments.map((row) => [formatDate(row.workDate) || '-', row.instrument || '-', row.serialNo || '-', row.range || '-', row.accuracy || '-', formatDate(row.calibrationDue) || '-', row.expired ? { text: 'Expired on work date', color: COLORS.red, bold: true } : row.calibrationDue ? { text: 'Valid', color: COLORS.green, bold: true } : { text: 'Not recorded', color: COLORS.neutral, bold: true }])
      : [['-', 'No instrument details were recorded on the sheets', '-', '-', '-', '-', { text: 'Not recorded', color: COLORS.red, bold: true }]]
  });
  doc.note(`Range and accuracy are the manufacturer's figures for the instrument model. Calibration certificates: ${annexures.has('C') ? 'attached as Annexure C' : 'provided separately (Annexure C)'}.`);
}

function drawMethodology(doc, report, data) {
  doc.section('Testing Methodology');
  const ids = [...(report.tests && report.tests.soilResistivity ? ['soilResistivity'] : []), ...data.tests.map((test) => test.id)];
  ids.forEach((id) => {
    const [head, body] = METHODOLOGY[id];
    doc.heading(head, 2);
    doc.paragraph(body);
  });
  doc.heading('Current-probe distance, lead nulling and safety', 2);
  doc.bullets([
    'Current-probe distance: for fall-of-potential tests the current probe is placed far enough from the electrode (or grid) under test for the 61.8% potential reading to lie on the flat part of the resistance curve, typically not less than 5 times the longest dimension of the electrode system. The distance used is recorded on each sheet; readings at 50% and 70% confirm the 61.8% value within the tolerance in the Basis of Assessment, otherwise the point is re-tested.',
    'Test-lead nulling: before each series of low-resistance measurements the leads are shorted and the instrument nulled. The lead resistance recorded in the sheet header is subtracted from every reading before it is compared with a limit (IEEE 81-2012 Cl. 10.2).',
    'Safety: ETPL procedure requires a valid client work permit, isolation and lock-out/tag-out (LOTO) of any equipment whose earth connection is opened for a test, PPE (safety helmet, safety shoes, insulated gloves) for every member of the team, and the earth connection restored and checked before leaving the point.'
  ]);
}

function drawSoil(doc, data) {
  doc.heading('Soil Resistivity', 2, { indexed: true });
  const ground = String(data.project.groundCondition || '').trim();
  data.soil.locations.forEach((location) => {
    doc.heading(`Location: ${location.name}`, 2);
    if (location.pending) {
      doc.paragraph(`Not assessed: ${location.pendingReason || 'reading unit not selected'}.`);
      return;
    }
    const conditions = [location.weather ? `weather: ${location.weather}` : '', ground ? `ground condition: ${ground}` : '', location.probeDepthCm ? `probe depth: ${location.probeDepthCm} cm` : ''].filter(Boolean);
    if (conditions.length) {
      doc.note(`Test conditions – ${conditions.join('; ')}.`);
    }
    const showR = location.inputUnit === 'ohm';
    const columns = [['Spacing a (m)', 12]];
    location.directionLabels.forEach((label) => {
      if (showR) columns.push([`R ${label} (Ω)`, 11]);
      columns.push([`ρ ${label} (Ω·m)`, 12]);
    });
    columns.push(['Mean ρ (Ω·m)', 12]);
    const cellsOf = (row) => [row.spacing === null ? '-' : row.spacing.toFixed(1), ...row.rhos.flatMap((value, d) => [...(showR ? [Number.isFinite(row.rs[d]) ? row.rs[d].toFixed(2) : '-'] : []), Number.isFinite(value) ? value.toFixed(2) : '-']), row.mean === null ? '-' : row.mean.toFixed(2)];
    const filler = columns.length - 2;
    doc.table({
      columns,
      rows: [
        ...location.rows.map(cellsOf),
        [{ text: 'Location average', bold: true }, ...Array(filler).fill(''), { text: location.average === null ? '-' : location.average.toFixed(2), bold: true }]
      ],
      alignRight: columns.map((_column, i) => i).slice(1)
    });
    if (location.series.length && location.series[0].points.length > 1) {
      doc.lineChart({ title: `Apparent resistivity ρ against probe spacing – ${location.name}`, series: location.series, xLabel: 'Spacing a (m)', yLabel: 'ρ (Ω·m)' });
    }
  });
  doc.table({
    columns: [['Site average ρ (Ω·m)', 25], ['Corrosivity class (IS 3043:2018 Table 4)', 40], ['Steel corrosion allowance', 35]],
    rows: [[data.soil.average === null ? '-' : data.soil.average.toFixed(2), data.soil.category || '-', data.soil.corrosionAllowancePct === null || data.soil.corrosionAllowancePct === undefined ? '-' : `${data.soil.corrosionAllowancePct}% in conductor cross-section (IS 3043:2018 Table 5)`]]
  });
}

function drawResults(doc, data, report) {
  doc.section('Results');
  if (report.tests && report.tests.soilResistivity) {
    drawSoil(doc, data);
  }
  data.tests.forEach((test) => {
    const row = data.summaryTable.find((entry) => entry.testId === test.id);
    doc.heading(`${test.label}${test.sheetCode && test.sheetCode !== 'legacy' ? ` (${test.sheetCode})` : ''}`, 2, { indexed: true });
    doc.paragraph(
      `${row.total} of ${test.points.length} ${test.points.length === 1 ? test.unit.replace(/s$/, '') : test.unit} assessed: ${row.healthy} within limit, ${row.unhealthy} above limit${row.notAssessed ? `, ${row.notAssessed} not measured / not assessed` : ''}. Overall status: ${row.status}.`,
      { align: 'left', gap: 4 }
    );
    if (!test.points.length) {
      doc.paragraph('No readings were recorded for this test.');
      return;
    }
    doc.table({
      columns: test.columns,
      rows: test.points.flatMap((point) => [...(point.subRows || []), point.cells]),
      statusColumns: test.statusColumns || [],
      fontSize: test.columns.length > 9 ? 7.2 : 7.8
    });
    if (test.note) {
      doc.note(test.note);
    }
    if (test.id === 'electrodeResistance') {
      doc.note('Limits: rod R = (100ρ/2πl)·ln(2l/d) (IS 3043:2018 Cl. 14.2.2); strip R = (100ρ/2πl)·ln(4l/d), d = strip width (Cl. 14.2.3); plate per Cl. 14.2.1. Status compares the resistance with the limit only (Within limit; Warning above 90% of the limit; Above limit; Not measurable); Priority carries the verdict. Test condition: Isolated = electrode disconnected from the grid, Grid-connected = bonded. 50/70% check: the 61.8% value is confirmed when the 50% and 70% readings agree within tolerance. Priority: P1 above limit, not measurable or safety defect; P2 warning band or maintenance defect; P3 healthy.');
      doc.electrodeChart(test.points);
    }
  });
}

function drawCap(doc, data) {
  doc.section('Corrective Action Plan');
  if (!data.cap.length) {
    doc.paragraph('No corrective actions are required: no defects or data gaps were recorded.');
    return;
  }
  doc.table({
    columns: [['Sr', 4], ['Test / point', 14], ['Observation', 28], ['Corrective action', 25], ['Priority / status', 12], ['Clause', 17]],
    rows: data.cap.map((line, i) => [String(i + 1), `${line.test}\n${line.point}`, line.observation, line.action, { text: line.priorityOrStatus, color: statusColor(line.priorityOrStatus), bold: true }, line.clause || '-'])
  });
  doc.note('Generated from the same per-point records as the summary and the results sections; counts cannot differ between them. Data-gap lines list the affected points of one test together.');
}

function drawMaintenance(doc) {
  doc.section('Recommendations for Maintenance & Re-testing');
  doc.bullets([
    'Measure earth resistance annually and in any case not less than once in two years (IS 3043:2018 Cl. 38.4.2).',
    'Carry out periodic visual inspection of connections, joints and test links (IS 3043:2018 Cl. 38.4.3).',
    'Keep records of the initial and all subsequent tests (IS 3043:2018 Cl. 38.4.1).',
    'Re-test every point listed in the Corrective Action Plan after rectification, and in any case within 30 days for priority P1 pits.'
  ]);
}

function drawAnnexures(doc, annexures) {
  doc.section('Annexures');
  doc.table({
    columns: [['Annexure', 13], ['Content', 67], ['Status', 20]],
    rows: ANNEXURES.map(([letter, content]) => [letter, content, annexures.has(letter) ? { text: 'Attached', color: COLORS.green, bold: true } : { text: 'Provided separately', color: COLORS.neutral, bold: true }]),
    boldFirstColumn: true
  });
}

// References picked by the reviewer (6H.5): title, where it is kept and its date, nothing from the documents'
// text. They are pointers for the reader and play no part in any figure or verdict of this report.
function drawReferences(doc, references) {
  doc.section('References');
  doc.note('Documents the reviewer selected as further reading. They were not used in any calculation, result or finding in this report.');
  doc.table({
    columns: [['Document', 38], ['Location', 50], ['Dated', 12]],
    rows: references.map((reference) => [String(reference.title || ''), `${reference.source || ''}: ${reference.path || ''}`, String(reference.modified || '').slice(0, 10)])
  });
}

function drawControlAndIndex(doc, data, revisions, sections, approvals) {
  const pdf = doc.doc;
  pdf.switchToPage(1);
  doc.y = PAGE.top;
  doc.heading('Document Control & Distribution', 1);
  doc.table({
    columns: [['Item', 28], ['Detail', 72]],
    rows: [
      ['Document No.', data.project.documentNo],
      ['Current revision', revisions[revisions.length - 1].revision],
      ['Classification', doc.meta.classification],
      ['Distribution', 'Client (controlled copy); Elegrow Technology Pvt. Ltd. project file']
    ],
    boldFirstColumn: true
  });
  doc.heading('Approval', 2);
  doc.table({
    columns: [['Role', 16], ['Name', 34], ['Designation', 30], ['Date', 20]],
    rows: APPROVAL_ROLES.map(([key, label]) => {
      const entry = (approvals && approvals[key]) || {};
      return [label, entry.name || '-', entry.designation || '-', entry.date ? formatDate(entry.date) : '-'];
    }),
    boldFirstColumn: true
  });
  doc.heading('Index', 1);
  doc.table({
    columns: [['Section', 88], ['Page', 12]],
    rows: sections.map((section) => [section.sub ? { text: `      ${section.title}` } : section.title, String(section.page + 1)]),
    alignRight: [1]
  });
  doc.note('The index lists only the sections and tests present in this report; it is generated from the site records.');
}

// Builds the final report. options: { narrative, revisions, clientLogo (Buffer), status,
// approvals: { prepared|checked|approved: { name, designation, date } }, annexures: ['A', ...] attached }.
async function generateFinalReportPdf(report, options = {}) {
  const data = options.data || buildFinalReportData(report);
  const consistency = assertConsistent(data);
  const narrative = options.narrative || deterministicNarrative(data);
  const revisions = resolveRevisions(options);
  const status = options.status || 'draft';
  const annexures = new Set(options.annexures || []);

  const doc = new ReportDoc({
    title: 'Earthing System Health Assessment',
    unit: data.project.unit,
    documentNo: data.project.documentNo,
    revision: revisions[revisions.length - 1].revision,
    watermark: status === 'approved' ? '' : 'DRAFT – NOT APPROVED',
    info: { Subject: `ESHA final report ${data.project.documentNo}` }
  });

  drawCover(doc, data, revisions, options.clientLogo || null);
  doc.newPage(); // page 2: document control and index, filled once the section pages are known
  doc.sections.length = 0;
  drawExecutiveSummary(doc, data, narrative);
  drawIntroduction(doc, data, report);
  drawBasis(doc, data);
  drawInstruments(doc, data, annexures);
  drawMethodology(doc, report, data);
  drawResults(doc, data, report);
  drawCap(doc, data);
  drawMaintenance(doc);
  drawAnnexures(doc, annexures);
  if (Array.isArray(report.references) && report.references.length) drawReferences(doc, report.references);
  drawControlAndIndex(doc, data, revisions, doc.sections, options.approvals);
  const pageCount = doc.pageIndex + 1;
  doc.finish({ coverPages: 1 });
  const buffer = await doc.toBuffer();
  return { buffer, data, consistency, narrative, pageCount };
}

module.exports = { generateFinalReportPdf };
