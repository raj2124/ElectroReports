const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');

const {
  TEST_LIBRARY,
  calculateSoilSummary,
  getElectrodeMeasuredValue,
  deriveElectrodeAssessment,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveEarthContinuityAssessment,
  deriveTowerFootingAssessment,
  summarizeTowerGroups,
  buildTowerGroupKey,
  buildExecutiveSnapshot,
  calculateDrivenElectrodeResistance,
  EQUIPMENT_LIBRARY,
  round
} = require('./reportModel');

const PAGE = {
  margin: 42,
  width: 595.28,
  height: 841.89
};

const COLORS = {
  brand: '#1f5fbf',
  brandDark: '#123a77',
  brandSoft: '#edf4ff',
  border: '#d7e2f2',
  ink: '#10243f',
  muted: '#61748d',
  white: '#ffffff',
  healthy: '#1f8f5f',
  healthySoft: '#eaf8f1',
  warning: '#c08700',
  warningSoft: '#fff6dd',
  critical: '#c63a3a',
  criticalSoft: '#ffebeb',
  neutral: '#718096',
  neutralSoft: '#edf2f7',
  rowAlt: '#f8fbff'
};

const LOGO_PATH = path.join(__dirname, '..', 'public', 'assets', 'elegrow-logo-full.png');
const SYMBOL_PATH = path.join(__dirname, '..', 'public', 'assets', 'elegrow-symbol.png');
const UNICODE_FONT_PATH = '/System/Library/Fonts/Supplemental/Arial Unicode.ttf';

const STANDARD_LIBRARY = {
  soilResistivity: ['IEEE 81-2012', 'IS 3043:2018 Clause 9.2'],
  electrodeResistance: ['IS 3043:2018', 'IS 3043:2018 Clause 9.2.2', 'CEA Regulations 2010 Clause 41'],
  continuityTest: ['IS 3043 continuity guidance', 'IEC 60364-6 continuity verification'],
  loopImpedanceTest: ['IEC 60364-6 loop verification', 'Protective device disconnection criteria'],
  prospectiveFaultCurrent: ['IEC 60909 fault current principles', 'Protective device breaking-capacity checks'],
  riserIntegrityTest: ['IEEE 81-2012 Clause 10', 'IS 3043 bonding and continuity guidance'],
  earthContinuityTest: ['IS 3043:2018', 'IEC 60364 continuity verification'],
  towerFootingResistance: ['Project tower footing grouped Zsat limit 10 ohm']
};

const EQUIPMENT_DETAILS = {
  mi3152: {
    label: 'MI 3152 EurotestXC',
    description:
      'A multifunctional installation tester used for earthing and installation verification activities, including continuity and related low-resistance measurements.',
    details: [
      ['Voltage Range', '90-260 V AC'],
      ['Current Range', '7-200 mA'],
      ['Voltage Accuracy', '±(2 % of reading + 2 digits)'],
      ['Current Accuracy', '±(3 % of reading + 5 digits)'],
      ['Over-voltage Category', '300 V'],
      ['Frequency Band', '55 Hz - 15 kHz'],
      ['Protection', 'IP 65 case closed, IP 54 case open'],
      ['Battery Supply', '14.4 V DC (4.4 Ah Li-ion)']
    ]
  },
  mi3290: {
    label: 'MI 3290 GF Earth Analyser',
    description:
      'A dedicated grounding-system analyser for soil resistivity, earth electrode testing, and related earthing-system measurements.',
    details: [
      ['Open Terminal Test Voltage Range', '20-40 V AC / less than 50 VAC'],
      ['Current Range', '7-200 mA'],
      ['Short Circuit Test Current', 'Greater than 220 mA'],
      ['Voltage Accuracy', '±(2 % of reading + 2 digits)'],
      ['Current Accuracy', '±(3 % of reading + 5 digits)'],
      ['Protection', 'IP 65 case closed, IP 54 case open'],
      ['Measurement Range', '0.010 ohm - 199.9 ohm'],
      ['Battery Supply', '14.4 V DC (4.4 Ah Li-ion)']
    ]
  },
  kyoritsu4118a: {
    label: 'Kyoritsu Digital PSC Loop Tester 4118A',
    description:
      'A loop-impedance and prospective short-circuit current tester used to verify fault loop performance and protective-device operation.',
    details: [
      ['Functional Range', '0.02 ohm to 2000 ohm'],
      ['Open Terminal Test Voltage', '20-40 V AC'],
      ['Current Measuring Clamps', 'A1018 / A1019'],
      ['Current Range', '0.0 mA - 99.9 mA'],
      ['Current Resolution', '0.1 mA'],
      ['Current Accuracy', '±(5 % of reading + 5 digits) / Indicative by clamp model'],
      ['Ambient Temperature', '23 °C ± 5 °C'],
      ['Use Case', 'Loop impedance, PSC/PFC verification, and feeder fault-path analysis']
    ]
  }
};

const ABBREVIATIONS = [
  ['GND', 'Grounding'],
  ['PFC', 'Prospective Fault Current'],
  ['DLRO', 'Digital Low Resistance Ohmmeter'],
  ['MET', 'Main Earthing Terminal'],
  ['PPE', 'Personal Protective Equipment'],
  ['LOTO', 'Lock-Out Tag-Out']
];

const SAFETY_MEASURES = [
  'All personnel must wear appropriate PPE during testing activities.',
  'For live testing such as loop impedance, use approved testers and follow safe work methods.',
  'Barricade and sign the test area before measurements begin.',
  'Check environmental and access conditions to ensure safe testing operations.'
];

const REPORTING_POINTS = [
  'Executive Summary and overall health classification.',
  'Detailed analysis and results for each selected test conducted.',
  'Observation sheets with health status bands such as Healthy, Warning, Not Measurable, and > Permissible Limit.',
  'Comparison of measured values against standard or project reference limits.',
  'Photographs and row-level observations where recorded in the app.'
];

const ELECTRODE_SUMMARY_MAJOR_POINTS = [
  'Measured Earth Electrode Resistance values are compared with the calculated standard earth electrode resistance and the engineer can complete the final gap-analysis recommendation in the summary table.',
  'Healthy means the measured earth electrode resistance value or grid resistance value is within the standard resistance limit. Physical condition of the earth electrode and pit must still be reviewed before final observation and recommendation are issued.',
  'Warning means the measured earth electrode resistance value or grid resistance value is close to the calculated standard resistance value and should be monitored carefully in the next inspection cycle.',
  'Permissible Limit means the measured earth electrode resistance value or grid resistance value is higher than the calculated standard earth electrode resistance value and requires engineering review.',
  'Visual inspection should be carried out for the full earthing electrode and earthing pit system at least once every quarter for rigidity, deterioration, loose contacts, and corrosion.',
  'All earthing systems should additionally be tested for resistance on a dry day during the dry season not less than once every year.',
  'As per the applicable CEA regulation, the earthing system must be tested annually and rectifications should be completed wherever required to maintain overall system health.',
  'A record of every earthing-system test and its result should be maintained by the client for inspection and trend review.'
];

const LOOP_NOTE =
  'NOTE: As per IS/IEC 60364-6:2016, earth loop impedance must ensure that the protective device operates within the safe disconnection time applicable to the tested circuit.';

const FAULT_NOTE =
  'NOTE: As per IS/IEC 60364-6:2016, prospective fault current (PFC) at each point must remain less than the breaking-capacity rating of the installed protective device.';

const RISER_NOTE =
  'NOTE: As per IEEE 81-2012, Clause 10, riser bonding resistance values should remain less than or equal to 1 ohm for a healthy condition.';

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function addReportPage(doc, drawChrome, options = {}) {
  const currentLayout = doc?.page?.layout || 'portrait';
  doc.addPage({
    size: 'A4',
    margin: PAGE.margin,
    layout: options.layout || currentLayout
  });
  if (drawChrome) {
    drawChrome(doc);
  }
}

function safeText(value, fallback = '-') {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || fallback;
}

function sanitizeFilePart(value) {
  return safeText(value, 'report').replace(/[^a-z0-9_-]+/gi, '-');
}

function getSelectedTests(report) {
  return TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);
}

function getApplicableStandards(report) {
  const selected = getSelectedTests(report);
  const seen = new Set();
  return selected.flatMap((test) => {
    return (STANDARD_LIBRARY[test.id] || []).filter((entry) => {
      if (seen.has(entry)) {
        return false;
      }
      seen.add(entry);
      return true;
    });
  });
}

function getSelectedEquipment(report) {
  const selectedIds = Array.isArray(report?.project?.equipmentSelections) ? report.project.equipmentSelections : [];
  const selected = EQUIPMENT_LIBRARY.filter((equipment) => selectedIds.includes(equipment.id));
  return selected.length ? selected : [...EQUIPMENT_LIBRARY];
}

function getSoilLocations(report) {
  if (Array.isArray(report?.soilResistivity?.locations) && report.soilResistivity.locations.length) {
    return report.soilResistivity.locations;
  }
  return [
    {
      locationId: 'soil-location-1',
      name: 'Location 1',
      direction1: report?.soilResistivity?.direction1 || [],
      direction2: report?.soilResistivity?.direction2 || [],
      drivenElectrodeDiameter: report?.soilResistivity?.drivenElectrodeDiameter || '',
      drivenElectrodeLength: report?.soilResistivity?.drivenElectrodeLength || '',
      notes: report?.soilResistivity?.notes || ''
    }
  ];
}

function tonePalette(tone) {
  if (tone === 'healthy') {
    return { fill: COLORS.healthySoft, text: COLORS.healthy };
  }
  if (tone === 'warning') {
    return { fill: COLORS.warningSoft, text: COLORS.warning };
  }
  if (tone === 'critical') {
    return { fill: COLORS.criticalSoft, text: COLORS.critical };
  }
  return { fill: COLORS.neutralSoft, text: COLORS.neutral };
}

function getReportTitle(report) {
  const selected = TEST_LIBRARY.filter((test) => report?.tests?.[test.id]);
  if (!selected.length) {
    return 'ElectroReports Assessment';
  }
  if (selected.length === 1) {
    return selected[0].label;
  }
  if (selected.length > 1) {
    return 'Earthing System Health Assessment';
  }
  return 'Electrical Measurement Assessment';
}

function ensureSpace(doc, requiredHeight, drawChrome) {
  const bottom = doc.page.height - PAGE.margin - 34;
  if (doc.y + requiredHeight <= bottom) {
    return;
  }
  addReportPage(doc, drawChrome);
}

function drawBodyChrome(doc, report) {
  const pageWidth = doc.page.width;
  const pageHeight = doc.page.height;
  const topY = 20;
  const footerLineY = pageHeight - PAGE.margin + 2;
  const footerTextY = pageHeight - PAGE.margin - 12;

  if (fs.existsSync(SYMBOL_PATH)) {
    const watermarkSize = Math.min(pageWidth, pageHeight) * 0.34;
    doc.save();
    doc.opacity(0.05);
    doc.image(SYMBOL_PATH, (pageWidth - watermarkSize) / 2, (pageHeight - watermarkSize) / 2 - 24, {
      fit: [watermarkSize, watermarkSize],
      align: 'center',
      valign: 'center'
    });
    doc.restore();
  }

  doc
    .moveTo(PAGE.margin, topY + 44)
    .lineTo(pageWidth - PAGE.margin, topY + 44)
    .lineWidth(1.5)
    .strokeColor(COLORS.brand)
    .stroke();

  if (fs.existsSync(LOGO_PATH)) {
    doc.image(LOGO_PATH, PAGE.margin, topY, { fit: [180, 34] });
  }

  doc
    .font('Helvetica')
    .fontSize(8)
    .fillColor(COLORS.muted)
    .text(`Project No: ${safeText(report.project.projectNo)}`, pageWidth - PAGE.margin - 180, topY + 6, {
      width: 180,
      align: 'right'
    })
    .text(`Client: ${safeText(report.project.clientName)}`, pageWidth - PAGE.margin - 180, topY + 18, {
      width: 180,
      align: 'right'
    });

  doc
    .moveTo(PAGE.margin, footerLineY)
    .lineTo(pageWidth - PAGE.margin, footerLineY)
    .lineWidth(0.8)
    .strokeColor(COLORS.border)
    .stroke();

  doc
    .font('Helvetica')
    .fontSize(8)
    .fillColor(COLORS.muted)
    .text('ElectroReports | Elegrow Technology', PAGE.margin, footerTextY, {
      width: 220
    })
    .text(`Date of Testing: ${safeText(report.project.reportDate)}`, pageWidth - PAGE.margin - 180, footerTextY, {
      width: 180,
      align: 'right'
    });

  doc.y = 86;
}

function drawCoverPage(doc, report, snapshot) {
  const title = getReportTitle(report);

  if (fs.existsSync(SYMBOL_PATH)) {
    const watermarkSize = Math.min(doc.page.width, doc.page.height) * 0.5;
    const watermarkX = (doc.page.width - watermarkSize) / 2;
    const watermarkY = (doc.page.height - watermarkSize) / 2;
    doc.save();
    doc.opacity(0.08);
    doc.image(SYMBOL_PATH, watermarkX, watermarkY, {
      fit: [watermarkSize, watermarkSize],
      align: 'center',
      valign: 'center'
    });
    doc.restore();
  }

  if (fs.existsSync(LOGO_PATH)) {
    doc.image(LOGO_PATH, PAGE.margin + 50, 36, { fit: [360, 78], align: 'center' });
  }

  doc
    .roundedRect(PAGE.margin, 150, doc.page.width - PAGE.margin * 2, 108, 18)
    .fillAndStroke(COLORS.brandSoft, COLORS.border);

  doc
    .font('Helvetica-Bold')
    .fontSize(24)
    .fillColor(COLORS.brandDark)
    .text(title, PAGE.margin + 24, 186, {
      width: doc.page.width - PAGE.margin * 2 - 48,
      align: 'center'
    });

  doc
    .font('Helvetica')
    .fontSize(10)
    .fillColor(COLORS.muted)
    .text('Generated from field measurement entries recorded in ElectroReports.', PAGE.margin + 24, 220, {
      width: doc.page.width - PAGE.margin * 2 - 48,
      align: 'center'
    });

  const cardY = 310;
  const cardW = (doc.page.width - PAGE.margin * 2 - 16) / 2;
  const cards = [
    { label: 'Client Name', value: report.project.clientName },
    { label: 'Project Number', value: report.project.projectNo },
    { label: 'Site Location', value: report.project.siteLocation },
    { label: 'Work Order / Ref', value: report.project.workOrder },
    { label: 'Date of Testing', value: report.project.reportDate },
    { label: 'Testing Engineer', value: report.project.engineerName }
  ];

  cards.forEach((card, index) => {
    const column = index % 2;
    const row = Math.floor(index / 2);
    const x = PAGE.margin + column * (cardW + 16);
    const y = cardY + row * 74;
    doc.roundedRect(x, y, cardW, 58, 12).fillAndStroke(COLORS.white, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(card.label.toUpperCase(), x + 14, y + 10, { width: cardW - 28 });
    doc
      .font('Helvetica-Bold')
      .fontSize(12)
      .fillColor(COLORS.ink)
      .text(safeText(card.value), x + 14, y + 26, { width: cardW - 28 });
  });

  doc
    .roundedRect(PAGE.margin, 560, doc.page.width - PAGE.margin * 2, 118, 12)
    .fillAndStroke(COLORS.white, COLORS.border);

  doc
    .font('Helvetica-Bold')
    .fontSize(11)
    .fillColor(COLORS.ink)
    .text('Selected Measurement Sections', PAGE.margin + 16, 578);

  const selectedLabels = snapshot.selectedTests.map((test) => test.label);
  doc
    .font('Helvetica')
    .fontSize(10)
    .fillColor(COLORS.muted)
    .text(selectedLabels.join(' | '), PAGE.margin + 16, 600, {
      width: doc.page.width - PAGE.margin * 2 - 32
    });

  doc
    .font('Helvetica')
    .fontSize(9)
    .fillColor(COLORS.muted)
    .text('This report includes only the measurement types selected during data entry.', PAGE.margin + 16, 640, {
      width: doc.page.width - PAGE.margin * 2 - 32
    });
}

function drawSectionTitle(doc, title, subtitle, drawChrome) {
  const boxHeight = 52;
  ensureSpace(doc, boxHeight + 22, drawChrome);
  const titleY = doc.y;
  const titleHeight = doc.heightOfString(title, {
    width: doc.page.width - PAGE.margin * 2 - 32
  });

  doc.roundedRect(PAGE.margin, titleY, doc.page.width - PAGE.margin * 2, boxHeight, 18).fillAndStroke(COLORS.brandSoft, COLORS.border);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(COLORS.brandDark)
    .text(title, PAGE.margin + 22, titleY + (boxHeight - titleHeight) / 2, {
      width: doc.page.width - PAGE.margin * 2 - 44,
      align: 'left'
    });

  doc.y = titleY + boxHeight + 10;
  if (subtitle) {
    doc
      .font('Helvetica')
      .fontSize(8.5)
      .fillColor(COLORS.muted)
      .text(subtitle, PAGE.margin + 10, doc.y, {
        width: doc.page.width - PAGE.margin * 2 - 20
      });
    doc.moveDown(1.2);
  }
}

function measureTextHeight(doc, text, width, options = {}) {
  doc.save();
  doc.font(options.font || 'Helvetica').fontSize(options.size || 8.5);
  const height = doc.heightOfString(safeText(text), {
    width,
    align: options.align || 'left'
  });
  doc.restore();
  return height;
}

function drawKeyValueGrid(doc, rows, drawChrome) {
  const gridWidth = doc.page.width - PAGE.margin * 2;
  const colWidth = (gridWidth - 14) / 2;
  const rowHeight = 58;
  const gapY = 10;
  const startY = doc.y;
  const rowPairs = [];
  for (let index = 0; index < rows.length; index += 2) {
    rowPairs.push([rows[index], rows[index + 1] || null]);
  }

  ensureSpace(doc, rowPairs.length * rowHeight + Math.max(0, rowPairs.length - 1) * gapY + 8, drawChrome);

  rowPairs.forEach((pair, pairIndex) => {
    const baseY = startY + pairIndex * (rowHeight + gapY);
    pair.forEach((row, columnIndex) => {
      if (!row) {
        return;
      }
      const x = PAGE.margin + columnIndex * (colWidth + 14);
      doc.roundedRect(x, baseY, colWidth, rowHeight, 14).fillAndStroke(COLORS.white, COLORS.border);
      doc
        .font('Helvetica-Bold')
        .fontSize(8)
        .fillColor(COLORS.muted)
        .text(row.label.toUpperCase(), x + 14, baseY + 10, {
          width: colWidth - 28
        });
      doc
        .font('Helvetica-Bold')
        .fontSize(11)
        .fillColor(COLORS.ink)
        .text(safeText(row.value), x + 14, baseY + 28, {
          width: colWidth - 28
        });
    });
  });

  doc.y = startY + rowPairs.length * rowHeight + Math.max(0, rowPairs.length - 1) * gapY + 6;
}

function drawStatusPill(doc, text, tone, x, y) {
  const palette = tonePalette(tone);
  const width = doc.widthOfString(text, { font: 'Helvetica-Bold', size: 8 }) + 16;
  doc.roundedRect(x, y, width, 16, 8).fill(palette.fill);
  doc
    .font('Helvetica-Bold')
    .fontSize(8)
    .fillColor(palette.text)
    .text(text, x + 8, y + 4, { width: width - 16, align: 'center' });
  return width;
}

function drawMetricCards(doc, metrics, drawChrome) {
  const totalWidth = doc.page.width - PAGE.margin * 2;
  const cardWidth = (totalWidth - 16) / Math.max(metrics.length, 1);
  const height = 70;

  ensureSpace(doc, height + 12, drawChrome);

  metrics.forEach((metric, index) => {
    const x = PAGE.margin + index * (cardWidth + 8);
    const y = doc.y;
    const palette = tonePalette(metric.tone || 'neutral');
    doc.roundedRect(x, y, cardWidth, height, 12).fillAndStroke(COLORS.white, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(metric.label.toUpperCase(), x + 12, y + 12, { width: cardWidth - 24 });
    doc
      .font('Helvetica-Bold')
      .fontSize(22)
      .fillColor(COLORS.ink)
      .text(safeText(metric.value), x + 12, y + 28, { width: cardWidth - 24 });
    if (metric.meta) {
      drawStatusPill(doc, metric.meta, metric.tone || 'neutral', x + 12, y + 50);
    } else {
      doc.roundedRect(x + 12, y + 51, cardWidth - 24, 4, 2).fill(palette.fill);
    }
  });

  doc.moveDown(4.5);
}

function drawParagraph(doc, text, drawChrome) {
  ensureSpace(doc, 48, drawChrome);
  doc
    .font('Helvetica')
    .fontSize(10)
    .fillColor(COLORS.ink)
    .text(text, PAGE.margin, doc.y, {
      width: doc.page.width - PAGE.margin * 2,
      align: 'justify'
    });
  doc.moveDown(1.35);
}

function drawBulletList(doc, items, drawChrome) {
  (Array.isArray(items) ? items : []).forEach((item) => {
    ensureSpace(doc, 20, drawChrome);
    doc
      .font('Helvetica')
      .fontSize(10)
      .fillColor(COLORS.ink)
      .text(`• ${item}`, PAGE.margin, doc.y, {
        width: doc.page.width - PAGE.margin * 2,
        align: 'left'
      });
    doc.moveDown(0.5);
  });
}

function drawSimpleTable(doc, headers, rows, drawChrome, columnWidths = []) {
  const totalWidth = doc.page.width - PAGE.margin * 2;
  const normalizedWidths =
    columnWidths.length === headers.length ? columnWidths : headers.map(() => 1 / headers.length);
  drawTable(
    doc,
    headers.map((header, index) => ({
      key: `c${index}`,
      label: header,
      width: normalizedWidths[index]
    })),
    rows.map((row) => {
      const normalized = {};
      headers.forEach((_, index) => {
        normalized[`c${index}`] = row[index] ?? '';
      });
      return normalized;
    }),
    drawChrome
  );
}

function drawNoteBox(doc, text, drawChrome) {
  const content = safeText(text, '');
  if (!content) {
    return;
  }

  doc.font('Helvetica').fontSize(8.5);
  const width = doc.page.width - PAGE.margin * 2;
  const textHeight = doc.heightOfString(content, { width: width - 64 });
  const height = Math.max(42, textHeight + 22);
  ensureSpace(doc, height + 10, drawChrome);

  const x = PAGE.margin;
  const y = doc.y;
  doc.roundedRect(x, y, width, height, 10).fillAndStroke(COLORS.warningSoft, COLORS.border);
  doc.font('Helvetica-Bold').fontSize(8.5).fillColor(COLORS.warning).text('NOTE', x + 12, y + 10);
  doc.font('Helvetica').fontSize(8.5).fillColor(COLORS.ink).text(content.replace(/^NOTE:\s*/i, ''), x + 52, y + 10, {
    width: width - 64
  });
  doc.y = y + height + 8;
}

function toReportBandLabel(status) {
  if (status?.tone === 'healthy') {
    return 'Healthy';
  }
  if (status?.tone === 'warning') {
    return 'Warning';
  }
  if (status?.tone === 'critical') {
    return '> Permissible Limit';
  }
  return 'Pending';
}

function toRiserCommentLabel(status) {
  if (status?.tone === 'healthy') {
    return 'Healthy';
  }
  if (status?.tone === 'neutral') {
    return 'Pending';
  }
  return 'Un-Healthy';
}

function drawSubsectionParagraph(doc, title, text, drawChrome) {
  doc.moveDown(0.35);
  ensureSpace(doc, 52, drawChrome);
  doc
    .font('Helvetica-Bold')
    .fontSize(9)
    .fillColor(COLORS.brandDark)
    .text(title.toUpperCase(), PAGE.margin, doc.y, {
      width: doc.page.width - PAGE.margin * 2
    });
  doc.moveDown(0.35);
  drawParagraph(doc, text, drawChrome);
}

function drawBarChart(doc, title, entries, drawChrome) {
  const rows = (Array.isArray(entries) ? entries : []).filter((entry) => Number.isFinite(entry?.value));
  if (!rows.length) {
    return;
  }

  const chartHeight = rows.length * 30 + 56;
  ensureSpace(doc, chartHeight, drawChrome);

  const x = PAGE.margin;
  const width = doc.page.width - PAGE.margin * 2;
  const y = doc.y;
  const innerPadding = 14;
  const labelWidth = Math.min(190, width * 0.34);
  const valueWidth = 38;
  const barAreaWidth = width - innerPadding * 2 - labelWidth - valueWidth - 18;
  const maxValue = Math.max(...rows.map((entry) => entry.value), 1);

  doc.roundedRect(x, y, width, chartHeight, 12).fillAndStroke(COLORS.white, COLORS.border);
  doc
    .font('Helvetica-Bold')
    .fontSize(9)
    .fillColor(COLORS.brandDark)
    .text(title, x + 12, y + 10, { width: width - 24 });

  rows.forEach((entry, index) => {
    const rowY = y + 34 + index * 30;
    const barX = x + innerPadding + labelWidth + 10;
    const barY = rowY + 6;
    const barWidth = Math.max(8, (barAreaWidth * entry.value) / maxValue);
    const fill =
      entry.tone === 'critical' ? COLORS.critical : entry.tone === 'warning' ? COLORS.warning : entry.tone === 'healthy' ? COLORS.healthy : COLORS.brand;

    doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor(COLORS.ink)
      .text(entry.label, x + innerPadding, rowY, { width: labelWidth - 6 });
    doc.roundedRect(barX, barY, barAreaWidth, 12, 6).fill(COLORS.neutralSoft);
    doc.roundedRect(barX, barY, barWidth, 12, 6).fill(fill);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.ink)
      .text(String(entry.value), barX + barAreaWidth + 6, rowY, {
        width: valueWidth,
        align: 'right'
      });
  });

  doc.y = y + chartHeight + 12;
}

function drawSoilLocationChart(doc, title, categories, direction1Values, direction2Values, drawChrome) {
  const points = Array.isArray(categories) ? categories.filter(Boolean) : [];
  const seriesOne = (Array.isArray(direction1Values) ? direction1Values : []).map((value) =>
    Number.isFinite(value) ? value : null
  );
  const seriesTwo = (Array.isArray(direction2Values) ? direction2Values : []).map((value) =>
    Number.isFinite(value) ? value : null
  );
  const hasData = seriesOne.some((value) => Number.isFinite(value)) || seriesTwo.some((value) => Number.isFinite(value));
  if (!points.length || !hasData) {
    return;
  }

  const width = doc.page.width - PAGE.margin * 2;
  const height = 320;
  ensureSpace(doc, height + 16, drawChrome);

  const x = PAGE.margin;
  const y = doc.y;
  const chartX = x + 34;
  const chartY = y + 42;
  const chartWidth = width - 68;
  const plotX = chartX + 42;
  const plotY = chartY + 18;
  const plotWidth = chartWidth - 84;
  const plotHeight = height - 146;
  const leftMax = Math.max(...seriesOne.filter((value) => Number.isFinite(value)), 1);
  const rightMax = Math.max(...seriesTwo.filter((value) => Number.isFinite(value)), 1);
  const pointGap = points.length > 1 ? plotWidth / (points.length - 1) : 0;

  doc.roundedRect(x, y, width, height, 14).fillAndStroke(COLORS.white, COLORS.border);
  doc
    .font('Helvetica-Bold')
    .fontSize(10)
    .fillColor('#4f5865')
    .text(title, x + 18, y + 14, {
      width: width - 36,
      align: 'center'
    });

  for (let step = 0; step <= 5; step += 1) {
    const guideY = plotY + (plotHeight / 5) * step;
    const leftValue = round((leftMax / 5) * (5 - step), 1);
    const rightValue = round((rightMax / 5) * (5 - step), 1);
    doc.moveTo(plotX, guideY).lineTo(plotX + plotWidth, guideY).lineWidth(0.8).strokeColor(COLORS.border).stroke();
    doc
      .font('Helvetica')
      .fontSize(7)
      .fillColor(COLORS.muted)
      .text(String(leftValue), chartX, guideY - 4, { width: 34, align: 'right' })
      .text(String(rightValue), plotX + plotWidth + 8, guideY - 4, { width: 34, align: 'left' });
  }

  doc.moveTo(plotX, plotY).lineTo(plotX, plotY + plotHeight).lineTo(plotX + plotWidth, plotY + plotHeight).lineWidth(1).strokeColor('#c8d5e8').stroke();
  doc.moveTo(plotX + plotWidth, plotY).lineTo(plotX + plotWidth, plotY + plotHeight).lineWidth(1).strokeColor('#c8d5e8').stroke();

  points.forEach((point, index) => {
    const pointX = plotX + pointGap * index;
    doc
      .font('Helvetica')
      .fontSize(7.5)
      .fillColor(COLORS.muted)
      .text(String(point), pointX - 16, plotY + plotHeight + 8, {
        width: 32,
        align: 'center'
      });
  });

  const leftPoints = [];
  const rightPoints = [];

  const plotSeries = (values, maxValue, color, targetPoints) => {
    let previous = null;
    values.forEach((value, index) => {
      if (!Number.isFinite(value)) {
        previous = null;
        targetPoints[index] = null;
        return;
      }
      const pointX = plotX + pointGap * index;
      const pointY = plotY + plotHeight - (value / maxValue) * plotHeight;
      if (previous) {
        doc.moveTo(previous.x, previous.y).lineTo(pointX, pointY).lineWidth(2.3).strokeColor(color).stroke();
      }
      doc.circle(pointX, pointY, 2.6).fill(color);
      targetPoints[index] = { x: pointX, y: pointY, value };
      previous = { x: pointX, y: pointY };
    });
  };

  plotSeries(seriesOne, leftMax, '#4a8fdc', leftPoints);
  plotSeries(seriesTwo, rightMax, '#f18728', rightPoints);

  points.forEach((_, index) => {
    const first = leftPoints[index];
    const second = rightPoints[index];

    if (first) {
      const closeToSecond = second && Math.abs(first.y - second.y) < 18;
      const firstYOffset = closeToSecond ? 10 : -16;
      doc
        .font('Helvetica-Bold')
        .fontSize(8)
        .fillColor('#243243')
        .text(String(round(first.value, 1)), first.x - 18, first.y + firstYOffset, {
          width: 36,
          align: 'center'
        });
    }

    if (second) {
      const closeToFirst = first && Math.abs(second.y - first.y) < 18;
      const secondYOffset = closeToFirst ? -18 : -16;
      doc
        .font('Helvetica-Bold')
        .fontSize(8)
        .fillColor('#243243')
        .text(String(round(second.value, 1)), second.x - 18, second.y + secondYOffset, {
          width: 36,
          align: 'center'
        });
    }
  });

  doc.save();
  doc.rotate(-90, { origin: [chartX - 2, plotY + plotHeight / 2] });
  doc
    .font('Helvetica')
    .fontSize(8.5)
    .fillColor('#4f5865')
    .text('Resistivity ρ in Ohm. Meter', chartX - plotHeight / 2 - 60, plotY - 14, {
      width: 120,
      align: 'center'
    });
  doc.restore();
  doc.save();
  doc.rotate(-90, { origin: [plotX + plotWidth + 34, plotY + plotHeight / 2] });
  doc
    .font('Helvetica')
    .fontSize(8.5)
    .fillColor('#4f5865')
    .text('Resistivity ρ in Ohm. Meter', plotX + plotWidth - plotHeight / 2 - 24, plotY + plotWidth + 8, {
      width: 120,
      align: 'center'
    });
  doc.restore();
  doc
    .font('Helvetica')
    .fontSize(8.5)
    .fillColor('#4f5865')
    .text('Spacing of Probes in Meters', x, y + height - 54, {
      width: width,
      align: 'center'
    });

  const legendY = y + height - 28;
  const legendStartX = x + width / 2 - 130;
  [
    { color: '#4a8fdc', label: 'Direction - 1 (Left Hand Side)' },
    { color: '#f18728', label: 'Direction - 2 (Right Hand Side)' }
  ].forEach((entry, index) => {
    const legendX = legendStartX + index * 155;
    doc.roundedRect(legendX, legendY + 4, 22, 4, 2).fill(entry.color);
    doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor('#4f5865')
      .text(entry.label, legendX + 28, legendY, {
        width: 122
      });
  });

  doc.y = y + height + 12;
}

function drawLineChart(doc, title, categories, series, drawChrome) {
  const points = Array.isArray(categories) ? categories.filter(Boolean) : [];
  const activeSeries = (Array.isArray(series) ? series : []).filter((entry) => Array.isArray(entry?.values) && entry.values.some((value) => Number.isFinite(value)));
  if (!points.length || !activeSeries.length) {
    return;
  }

  const width = doc.page.width - PAGE.margin * 2;
  const height = 250;
  ensureSpace(doc, height + 16, drawChrome);

  const x = PAGE.margin;
  const y = doc.y;
  const plotX = x + 46;
  const plotY = y + 34;
  const plotWidth = width - 74;
  const plotHeight = height - 78;
  const allValues = activeSeries.flatMap((entry) => entry.values).filter((value) => Number.isFinite(value));
  const maxValue = Math.max(...allValues, 1);

  doc.roundedRect(x, y, width, height, 12).fillAndStroke(COLORS.white, COLORS.border);
  doc.font('Helvetica-Bold').fontSize(10).fillColor(COLORS.brandDark).text(title, x + 12, y + 10, { width: width - 24 });

  doc.moveTo(plotX, plotY).lineTo(plotX, plotY + plotHeight).lineTo(plotX + plotWidth, plotY + plotHeight).lineWidth(1).strokeColor(COLORS.border).stroke();

  for (let step = 0; step <= 4; step += 1) {
    const value = round((maxValue / 4) * (4 - step), 1);
    const guideY = plotY + (plotHeight / 4) * step;
    doc.moveTo(plotX, guideY).lineTo(plotX + plotWidth, guideY).lineWidth(0.6).strokeColor(COLORS.border).stroke();
    doc.font('Helvetica').fontSize(7).fillColor(COLORS.muted).text(String(value), x + 4, guideY - 4, { width: 36, align: 'right' });
  }

  const pointGap = points.length > 1 ? plotWidth / (points.length - 1) : 0;
  points.forEach((point, index) => {
    const pointX = plotX + pointGap * index;
    doc.font('Helvetica').fontSize(7.5).fillColor(COLORS.muted).text(String(point), pointX - 16, plotY + plotHeight + 8, {
      width: 32,
      align: 'center'
    });
  });

  activeSeries.forEach((entry) => {
    let previous = null;
    entry.values.forEach((value, index) => {
      if (!Number.isFinite(value)) {
        previous = null;
        return;
      }
      const pointX = plotX + pointGap * index;
      const pointY = plotY + plotHeight - (value / maxValue) * plotHeight;
      if (previous) {
        doc.moveTo(previous.x, previous.y).lineTo(pointX, pointY).lineWidth(2).strokeColor(entry.color).stroke();
      }
      doc.circle(pointX, pointY, 3).fill(entry.color);
      previous = { x: pointX, y: pointY };
    });
  });

  activeSeries.forEach((entry, index) => {
    const legendX = x + 12 + index * 138;
    const legendY = y + height - 22;
    doc.roundedRect(legendX, legendY + 4, 12, 4, 2).fill(entry.color);
    doc.font('Helvetica').fontSize(8).fillColor(COLORS.ink).text(entry.label, legendX + 18, legendY, { width: 120 });
  });

  doc.y = y + height + 8;
}

function drawTable(doc, columns, rows, drawChrome) {
  const totalWidth = doc.page.width - PAGE.margin * 2;
  const x = PAGE.margin;
  const resolveCellText = (row, column) => {
    const allowEmpty = Array.isArray(row.__allowEmpty) && row.__allowEmpty.includes(column.key);
    return allowEmpty ? String(row[column.key] ?? '').trim() : safeText(row[column.key]);
  };
  const drawHeader = () => {
    const headerHeights = columns.map((column) => {
      return measureTextHeight(doc, column.label, totalWidth * column.width - 12, {
        font: 'Helvetica-Bold',
        size: column.headerSize || 7.7,
        align: column.align || 'left'
      });
    });
    const headerHeight = Math.max(24, ...headerHeights.map((height) => height + 12));
    ensureSpace(doc, headerHeight + 4, drawChrome);
    const headerY = doc.y;
    doc.roundedRect(x, headerY, totalWidth, headerHeight, 8).fill(COLORS.brand);
    let cursorX = x;

    columns.forEach((column) => {
      const colWidth = totalWidth * column.width;
      const textHeight = measureTextHeight(doc, column.label, colWidth - 12, {
        font: 'Helvetica-Bold',
        size: column.headerSize || 7.7,
        align: column.align || 'left'
      });
      doc
        .font('Helvetica-Bold')
        .fontSize(column.headerSize || 7.7)
        .fillColor(COLORS.white)
        .text(column.label, cursorX + 6, headerY + (headerHeight - textHeight) / 2, {
          width: colWidth - 12,
          align: column.align || 'left'
        });
      cursorX += colWidth;
    });

    doc.y = headerY + headerHeight + 4;
  };

  drawHeader();

  rows.forEach((row, rowIndex) => {
    const observationBlocks = Array.isArray(row.__observationBlocks) ? row.__observationBlocks.filter(Boolean) : [];
    const cellHeights = columns.map((column) => {
      return measureTextHeight(doc, resolveCellText(row, column), totalWidth * column.width - 12, {
        font: row.__cellStyles?.[column.key]?.bold ? 'Helvetica-Bold' : 'Helvetica',
        size: 8.5,
        align: column.align || 'left'
      });
    });
    const rowHeight = Math.max(24, ...cellHeights.map((height) => height + 10));
    const observationHeight = estimateObservationBlocksHeight(doc, observationBlocks, totalWidth - 20);
    const bottom = doc.page.height - PAGE.margin - 34;
    if (doc.y + rowHeight + 6 + observationHeight > bottom) {
      addReportPage(doc, drawChrome);
      drawHeader();
    }

    const rowY = doc.y;
    doc
      .roundedRect(x, rowY, totalWidth, rowHeight, 6)
      .fillAndStroke(row.__rowFill || (rowIndex % 2 === 0 ? COLORS.white : COLORS.rowAlt), COLORS.border);

    let rowX = x;
    columns.forEach((column) => {
      const colWidth = totalWidth * column.width;
      const cellStyle = row.__cellStyles?.[column.key];
      if (cellStyle?.fill) {
        doc.save();
        doc.rect(rowX + 0.5, rowY + 0.5, colWidth - 1, rowHeight - 1).fill(cellStyle.fill);
        doc.restore();
      }
      const text = resolveCellText(row, column);
      const textHeight = measureTextHeight(doc, text, colWidth - 12, {
        font: cellStyle?.bold ? 'Helvetica-Bold' : 'Helvetica',
        size: 8.5,
        align: column.align || 'left'
      });
      const textY = rowY + Math.max(6, (rowHeight - textHeight) / 2);
      doc
        .font(cellStyle?.bold ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(8.5)
        .fillColor(cellStyle?.text || COLORS.ink)
        .text(text, rowX + 6, textY, {
          width: colWidth - 12,
          align: column.align || 'left'
        });
      rowX += colWidth;
    });

    doc.y = rowY + rowHeight + 4;
    if (observationBlocks.length) {
      drawObservationBlocks(doc, observationBlocks, x + 10, totalWidth - 20, drawChrome);
    }
  });

  doc.y += 6;
}

function drawClassicIndexPage(doc, sectionStarts) {
  const x = PAGE.margin + 8;
  const width = doc.page.width - PAGE.margin * 2 - 16;
  const titleHeight = 40;
  const headerHeight = 34;
  const footerReserve = 102;
  const availableHeight = doc.page.height - doc.y - footerReserve - titleHeight - headerHeight;
  const rowCount = Math.max(sectionStarts.length, 1);
  const rowHeight = Math.max(24, Math.min(32, Math.floor(availableHeight / rowCount)));

  doc.rect(x, doc.y, width, titleHeight).fillAndStroke('#2f78bf', '#111111');
  doc
    .font('Helvetica-Bold')
    .fontSize(17)
    .fillColor(COLORS.white)
    .text('INDEX', x, doc.y + 10, {
      width,
      align: 'center'
    });

  let cursorY = doc.y + titleHeight;
  const columns = [
    { key: 'srNo', label: 'Sr. No.', width: 0.1, align: 'center' },
    { key: 'particular', label: 'Particular', width: 0.62 },
    { key: 'revNo', label: 'Rv. No.', width: 0.14, align: 'center' },
    { key: 'page', label: 'Page', width: 0.14, align: 'center' }
  ];
  const columnXs = [];
  let cursorX = x;
  columns.forEach((column) => {
    columnXs.push(cursorX);
    cursorX += width * column.width;
  });

  doc.rect(x, cursorY, width, headerHeight).fillAndStroke('#d1d1d1', '#111111');
  columns.forEach((column, index) => {
    const colX = columnXs[index];
    const colWidth = width * column.width;
    const textHeight = measureTextHeight(doc, column.label, colWidth - 8, {
      font: 'Helvetica-Bold',
      size: 11,
      align: column.align
    });
    doc
      .font('Helvetica-Bold')
      .fontSize(11)
      .fillColor('#111111')
      .text(column.label, colX + 4, cursorY + Math.max(4, (headerHeight - textHeight) / 2), {
        width: colWidth - 8,
        align: column.align
      });
    doc.rect(colX, cursorY, colWidth, headerHeight).stroke('#111111');
  });
  cursorY += headerHeight;

  sectionStarts.forEach((section, index) => {
    const row = {
      srNo: String(index + 1),
      particular: section.title,
      revNo: 'R0',
      page: String(section.page)
    };
    doc.rect(x, cursorY, width, rowHeight).stroke('#111111');
    columns.forEach((column, columnIndex) => {
      const colX = columnXs[columnIndex];
      const colWidth = width * column.width;
      const textHeight = measureTextHeight(doc, row[column.key], colWidth - 8, {
        font: column.key === 'srNo' ? 'Helvetica-Bold' : 'Helvetica',
        size: 10.5,
        align: column.align
      });
      doc.rect(colX, cursorY, colWidth, rowHeight).stroke('#111111');
      doc
        .font(column.key === 'srNo' ? 'Helvetica-Bold' : 'Helvetica')
        .fontSize(10.5)
        .fillColor('#111111')
        .text(row[column.key], colX + 4, cursorY + Math.max(4, (rowHeight - textHeight) / 2), {
          width: colWidth - 8,
          align: column.align
        });
    });
    cursorY += rowHeight;
  });

  doc.y = cursorY + 6;
}

function drawSurveyOverviewBlock(doc, title, rows, drawChrome) {
  ensureSpace(doc, 40 + rows.length * 32, drawChrome);
  doc
    .font('Helvetica-Bold')
    .fontSize(11)
    .fillColor(COLORS.brandDark)
    .text(title, PAGE.margin, doc.y, {
      width: doc.page.width - PAGE.margin * 2
    });
  doc.moveDown(0.5);
  drawTable(
    doc,
    [
      { key: 'label', label: 'Field', width: 0.28 },
      { key: 'value', label: 'Details', width: 0.72 }
    ],
    rows,
    drawChrome
  );
}

function drawSoilAverageBand(doc, label, value, category, drawChrome) {
  const width = doc.page.width - PAGE.margin * 2;
  const x = PAGE.margin;
  const y = doc.y;
  const height = 48;
  ensureSpace(doc, height + 10, drawChrome);
  const valueWidth = 120;
  const categoryWidth = 120;
  const labelWidth = width - valueWidth - categoryWidth - 8;
  const palette = tonePalette(category?.tone || 'neutral');

  doc.roundedRect(x, y, width, height, 10).fillAndStroke(COLORS.white, COLORS.border);
  doc.roundedRect(x + labelWidth + 4, y + 4, valueWidth - 8, height - 8, 8).fillAndStroke('#dff0d8', COLORS.border);
  doc.roundedRect(x + labelWidth + valueWidth + 4, y + 4, categoryWidth - 8, height - 8, 8).fillAndStroke(palette.fill, COLORS.border);

  const fontName = fs.existsSync(UNICODE_FONT_PATH) ? 'Unicode' : 'Helvetica-Bold';
  doc
    .font(fontName)
    .fontSize(10.5)
    .fillColor(COLORS.ink)
    .text(label, x + 14, y + 14, {
      width: labelWidth - 20,
      align: 'center'
    });
  doc
    .font('Helvetica-Bold')
    .fontSize(13)
    .fillColor(COLORS.ink)
    .text(String(value), x + labelWidth + 4, y + 14, {
      width: valueWidth - 8,
      align: 'center'
    });
  doc
    .font('Helvetica-Bold')
    .fontSize(12)
    .fillColor(palette.text)
    .text(safeText(category?.label, '-'), x + labelWidth + valueWidth + 4, y + 14, {
      width: categoryWidth - 8,
      align: 'center'
    });
  doc.y = y + height + 10;
}

function getStatusCellStyle(tone) {
  if (tone === 'healthy') {
    return { fill: COLORS.healthySoft, text: COLORS.healthy, bold: true };
  }
  if (tone === 'warning') {
    return { fill: COLORS.warningSoft, text: COLORS.warning, bold: true };
  }
  if (tone === 'critical') {
    return { fill: COLORS.criticalSoft, text: COLORS.critical, bold: true };
  }
  return null;
}

function drawSoilAnalysisTable(doc, location, locationSummary, locationIndex, drawChrome) {
  const direction1 = Array.isArray(location?.direction1) ? location.direction1 : [];
  const direction2 = Array.isArray(location?.direction2) ? location.direction2 : [];
  const rowCount = Math.max(direction1.length, direction2.length, 1);
  const width = doc.page.width - PAGE.margin * 2;
  const x = PAGE.margin;
  const y = doc.y;
  const containerPadding = 16;
  const innerX = x + containerPadding;
  const innerWidth = width - containerPadding * 2;
  const titleHeight = 36;
  const locationHeight = 30;
  const groupHeaderHeight = 36;
  const subHeaderHeight = 42;
  const rowHeight = 26;
  const averageHeight = 36;
  const totalHeight = containerPadding * 2 + titleHeight + 10 + locationHeight + 12 + groupHeaderHeight + subHeaderHeight + rowCount * rowHeight + averageHeight;
  const stroke = COLORS.border;
  const unicodeFontName = fs.existsSync(UNICODE_FONT_PATH) ? 'Unicode' : 'Helvetica';
  const categoryPalette = tonePalette(locationSummary?.category?.tone || 'neutral');
  const columnWidths = [0.08, 0.16, 0.17, 0.16, 0.17, 0.14, 0.12];
  const columnXs = [];
  let cursorX = innerX;

  ensureSpace(doc, totalHeight + 18, drawChrome);

  columnWidths.forEach((fraction) => {
    columnXs.push(cursorX);
    cursorX += innerWidth * fraction;
  });

  const titleY = y + containerPadding;
  const locationY = titleY + titleHeight + 10;
  const gridStartY = locationY + locationHeight + 12;
  const tableBottom = y + totalHeight;
  const bodyStartY = gridStartY + groupHeaderHeight + subHeaderHeight;
  const categoryHeight = rowCount * rowHeight + averageHeight;

  const drawCell = (cellX, cellY, cellWidth, cellHeight, text, options = {}) => {
    if (options.fill) {
      doc.save();
      doc.rect(cellX, cellY, cellWidth, cellHeight).fill(options.fill);
      doc.restore();
    }
    doc.rect(cellX, cellY, cellWidth, cellHeight).lineWidth(0.8).strokeColor(stroke).stroke();
    const fontName = options.font || 'Helvetica-Bold';
    const fontSize = options.fontSize || 10;
    const align = options.align || 'center';
    const safe = options.allowEmpty ? String(text ?? '').trim() : safeText(text, '');
    const textWidth = cellWidth - 10;
    const textHeight = measureTextHeight(doc, safe, textWidth, {
      font: fontName,
      size: fontSize,
      align
    });
    doc
      .font(fontName)
      .fontSize(fontSize)
      .fillColor(options.textColor || COLORS.white)
      .text(safe, cellX + 5, cellY + Math.max(5, (cellHeight - textHeight) / 2), {
        width: textWidth,
        align
      });
  };

  const meanForRow = (index) => {
    const first = Number.parseFloat(direction1[index]?.resistivity);
    const second = Number.parseFloat(direction2[index]?.resistivity);
    const values = [first, second].filter((value) => Number.isFinite(value));
    if (!values.length) {
      return '';
    }
    return String(round(values.reduce((sum, value) => sum + value, 0) / values.length, 2));
  };

  doc.roundedRect(x, y, width, totalHeight, 18).fillAndStroke(COLORS.white, COLORS.border);
  doc.roundedRect(innerX, titleY, innerWidth, titleHeight, 12).fill(COLORS.brand);
  doc
    .font('Helvetica-Bold')
    .fontSize(19)
    .fillColor(COLORS.white)
    .text(`Result of Soil Resistivity Test - ${locationIndex + 1}`, innerX, titleY + 6, {
      width: innerWidth,
      align: 'center'
    });

  doc.roundedRect(innerX, locationY, innerWidth, locationHeight, 10).fill('#2d86c9');
  doc
    .font('Helvetica-Bold')
    .fontSize(16)
    .fillColor(COLORS.white)
    .text(`Location – ${safeText(locationSummary?.name || location?.name, `Location ${locationIndex + 1}`)}`, innerX, locationY + 5, {
      width: innerWidth,
      align: 'center'
    });

  drawCell(columnXs[0], gridStartY, innerWidth * columnWidths[0], groupHeaderHeight + subHeaderHeight, 'Sr. No.', {
    fill: COLORS.brand,
    fontSize: 11
  });
  drawCell(columnXs[1], gridStartY, innerWidth * (columnWidths[1] + columnWidths[2]), groupHeaderHeight, 'Direction - 1 (Left Hand Side)', {
    fill: COLORS.brand,
    fontSize: 10.5
  });
  drawCell(columnXs[3], gridStartY, innerWidth * (columnWidths[3] + columnWidths[4]), groupHeaderHeight, 'Direction - 2 (Right Hand Side)', {
    fill: COLORS.brand,
    fontSize: 10.5
  });
  drawCell(columnXs[5], gridStartY, innerWidth * columnWidths[5], groupHeaderHeight + subHeaderHeight, 'Mean Resistivity ρ in Ohm. Meter', {
    fill: COLORS.brand,
    font: unicodeFontName,
    fontSize: 10.5
  });
  drawCell(columnXs[6], gridStartY, innerWidth * columnWidths[6], groupHeaderHeight + subHeaderHeight, 'Category as per IS - 3043 & IEEE - 81', {
    fill: COLORS.brand,
    fontSize: 10.5
  });

  drawCell(columnXs[1], gridStartY + groupHeaderHeight, innerWidth * columnWidths[1], subHeaderHeight, 'Spacing of Probes in Meters', {
    fill: COLORS.brand,
    fontSize: 10
  });
  drawCell(columnXs[2], gridStartY + groupHeaderHeight, innerWidth * columnWidths[2], subHeaderHeight, 'Resistivity ρ in Ohm. Meter', {
    fill: COLORS.brand,
    font: unicodeFontName,
    fontSize: 10
  });
  drawCell(columnXs[3], gridStartY + groupHeaderHeight, innerWidth * columnWidths[3], subHeaderHeight, 'Spacing of Probes in Meters', {
    fill: COLORS.brand,
    fontSize: 10
  });
  drawCell(columnXs[4], gridStartY + groupHeaderHeight, innerWidth * columnWidths[4], subHeaderHeight, 'Resistivity ρ in Ohm. Meter', {
    fill: COLORS.brand,
    font: unicodeFontName,
    fontSize: 10
  });

  for (let index = 0; index < rowCount; index += 1) {
    const rowY = bodyStartY + index * rowHeight;
    drawCell(columnXs[0], rowY, innerWidth * columnWidths[0], rowHeight, String(index + 1), {
      fill: COLORS.white,
      textColor: COLORS.ink,
      font: 'Helvetica',
      fontSize: 10.5
    });
    drawCell(columnXs[1], rowY, innerWidth * columnWidths[1], rowHeight, direction1[index]?.spacing || '', {
      fill: COLORS.white,
      textColor: COLORS.ink,
      font: 'Helvetica',
      fontSize: 10.5
    });
    drawCell(columnXs[2], rowY, innerWidth * columnWidths[2], rowHeight, direction1[index]?.resistivity || '', {
      fill: COLORS.white,
      textColor: COLORS.ink,
      font: 'Helvetica',
      fontSize: 10.5
    });
    drawCell(columnXs[3], rowY, innerWidth * columnWidths[3], rowHeight, direction2[index]?.spacing || '', {
      fill: COLORS.white,
      textColor: COLORS.ink,
      font: 'Helvetica',
      fontSize: 10.5
    });
    drawCell(columnXs[4], rowY, innerWidth * columnWidths[4], rowHeight, direction2[index]?.resistivity || '', {
      fill: COLORS.white,
      textColor: COLORS.ink,
      font: 'Helvetica',
      fontSize: 10.5
    });
    drawCell(columnXs[5], rowY, innerWidth * columnWidths[5], rowHeight, meanForRow(index), {
      fill: COLORS.white,
      textColor: COLORS.ink,
      font: 'Helvetica',
      fontSize: 10.5
    });
  }

  drawCell(columnXs[6], bodyStartY, innerWidth * columnWidths[6], categoryHeight, safeText(locationSummary?.category?.label, '-'), {
    fill: categoryPalette.fill,
    textColor: categoryPalette.text,
    fontSize: 17
  });

  drawCell(columnXs[0], bodyStartY + rowCount * rowHeight, innerWidth * (columnWidths[0] + columnWidths[1] + columnWidths[2] + columnWidths[3] + columnWidths[4]), averageHeight, 'Average Value of Soil Resistivity ρ in Ohm. Meter', {
    fill: COLORS.white,
    textColor: COLORS.ink,
    font: unicodeFontName,
    fontSize: 11.5
  });
  drawCell(columnXs[5], bodyStartY + rowCount * rowHeight, innerWidth * columnWidths[5], averageHeight, locationSummary?.overallAverage === null ? '-' : String(locationSummary.overallAverage), {
    fill: categoryPalette.fill,
    textColor: COLORS.ink,
    fontSize: 15
  });

  doc.y = tableBottom + 10;
}

function buildObservationBlock(row, label = 'Row Observation') {
  const text = safeText(row?.rowObservation, '');
  const photos = (Array.isArray(row?.rowPhotos) ? row.rowPhotos : []).filter((photo) => safeText(photo?.dataUrl, ''));
  if (!text && !photos.length) {
    return null;
  }
  return {
    label,
    text,
    photos
  };
}

function dataUrlToBuffer(dataUrl) {
  const match = String(dataUrl || '').match(/^data:image\/[a-zA-Z0-9.+-]+;base64,(.+)$/);
  if (!match) {
    return null;
  }
  try {
    return Buffer.from(match[1], 'base64');
  } catch (_error) {
    return null;
  }
}

function estimateObservationBlockHeight(doc, block, width) {
  let height = 38;
  if (block.text) {
    doc.font('Helvetica').fontSize(8.5);
    height += doc.heightOfString(block.text, {
      width: width - 24
    });
    height += 8;
  }
  if (block.photos.length) {
    const thumbSize = 74;
    const gap = 8;
    const perRow = Math.max(1, Math.floor((width - 24 + gap) / (thumbSize + gap)));
    const rowCount = Math.ceil(block.photos.length / perRow);
    height += rowCount * thumbSize;
    height += Math.max(0, rowCount - 1) * gap;
    height += 8;
  }
  return height;
}

function estimateObservationBlocksHeight(doc, blocks, width) {
  return blocks.reduce((total, block, index) => {
    return total + estimateObservationBlockHeight(doc, block, width) + (index ? 6 : 0);
  }, 0);
}

function drawObservationBlocks(doc, blocks, x, width, drawChrome) {
  blocks.forEach((block, index) => {
    const height = estimateObservationBlockHeight(doc, block, width);
    ensureSpace(doc, height + 4, drawChrome);

    const y = doc.y;
    doc.roundedRect(x, y, width, height, 10).fillAndStroke(COLORS.brandSoft, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.brandDark)
      .text(block.label.toUpperCase(), x + 12, y + 10, {
        width: width - 24
      });

    let cursorY = y + 24;
    if (block.text) {
      doc
        .font('Helvetica')
        .fontSize(8.5)
        .fillColor(COLORS.ink)
        .text(block.text, x + 12, cursorY, {
          width: width - 24
        });
      cursorY = doc.y + 8;
    }

    if (block.photos.length) {
      const thumbSize = 74;
      const gap = 8;
      const perRow = Math.max(1, Math.floor((width - 24 + gap) / (thumbSize + gap)));

      block.photos.forEach((photo, photoIndex) => {
        const imageBuffer = dataUrlToBuffer(photo.dataUrl);
        if (!imageBuffer) {
          return;
        }
        const column = photoIndex % perRow;
        const rowIndex = Math.floor(photoIndex / perRow);
        const imageX = x + 12 + column * (thumbSize + gap);
        const imageY = cursorY + rowIndex * (thumbSize + gap);
        doc.roundedRect(imageX, imageY, thumbSize, thumbSize, 8).fillAndStroke(COLORS.white, COLORS.border);
        try {
          doc.image(imageBuffer, imageX + 4, imageY + 4, {
            fit: [thumbSize - 8, thumbSize - 8],
            align: 'center',
            valign: 'center'
          });
        } catch (_error) {
          doc
            .font('Helvetica')
            .fontSize(7.5)
            .fillColor(COLORS.muted)
            .text('Image unavailable', imageX + 8, imageY + 30, {
              width: thumbSize - 16,
              align: 'center'
            });
        }
      });
    }

    doc.y = y + height + (index === blocks.length - 1 ? 4 : 6);
  });
}

function classifyObservationTone(tone) {
  if (tone === 'healthy') {
    return 'Healthy';
  }
  if (tone === 'warning') {
    return 'Warning';
  }
  if (tone === 'critical') {
    return '> Permissible Limit';
  }
  return 'Not Measurable';
}

function buildObservationSheetRows(report) {
  const rows = [];
  const pushRow = (title, entries) => {
    const normalized = Array.isArray(entries) ? entries : [];
    const quantity = normalized.length;
    const summary = { healthy: 0, warning: 0, critical: 0, pending: 0 };
    normalized.forEach((entry) => {
      if (!entry || entry.tone === 'neutral') {
        summary.pending += 1;
      } else if (entry.tone === 'healthy') {
        summary.healthy += 1;
      } else if (entry.tone === 'warning') {
        summary.warning += 1;
      } else if (entry.tone === 'critical') {
        summary.critical += 1;
      }
    });
    rows.push([
      title,
      String(quantity),
      String(summary.healthy),
      String(summary.warning),
      String(summary.pending),
      String(summary.critical)
    ]);
  };

  if (report.tests.soilResistivity) {
    const soil = calculateSoilSummary(report);
    const tone = soil.category?.tone || 'neutral';
    pushRow('Soil Resistivity Test', soil.locations.map((location) => ({ tone: location.category?.tone || 'neutral' })));
  }
  if (report.tests.electrodeResistance) {
    pushRow('Earth Electrode Resistance Test', report.electrodeResistance.map((row) => deriveElectrodeAssessment(row).status));
  }
  if (report.tests.continuityTest) {
    pushRow('Earthing Continuity Test', report.continuityTest.map((row) => deriveContinuityAssessment(row).status));
  }
  if (report.tests.loopImpedanceTest) {
    pushRow('Earth Loop Impedance Test', report.loopImpedanceTest.map((row) => deriveLoopImpedanceAssessment(row).status));
  }
  if (report.tests.prospectiveFaultCurrent) {
    pushRow('Prospective Fault Current Test', report.prospectiveFaultCurrent.map((row) => deriveFaultCurrentAssessment(row).status));
  }
  if (report.tests.riserIntegrityTest) {
    pushRow('Riser / Grid Integrity Test', report.riserIntegrityTest.map((row) => deriveRiserAssessment(row).status));
  }
  if (report.tests.earthContinuityTest) {
    pushRow('Ground Resistance / Earth Continuity Test', report.earthContinuityTest.map((row) => deriveEarthContinuityAssessment(row).status));
  }
  if (report.tests.towerFootingResistance) {
    const summaries = summarizeTowerGroups(report.towerFootingResistance);
    pushRow(
      'Tower Footing Test',
      report.towerFootingResistance.map((group) => deriveTowerFootingAssessment(group, summaries.get(buildTowerGroupKey(group))).status)
    );
  }
  return rows;
}

function buildExecutiveSummaryRows(report) {
  const rows = [];
  const selected = getSelectedTests(report);
  selected.forEach((test) => {
    rows.push([test.label, 'To be filled by engineer', 'To be filled by engineer']);
  });
  return rows;
}

function buildComplianceSummaryRows(report) {
  const rows = [];
  const soil = calculateSoilSummary(report);
  const towerSummaries = summarizeTowerGroups(report.towerFootingResistance);
  const push = (config) => rows.push(config);

  if (report.tests.soilResistivity) {
    const entries = soil.locations.map((location) => location.category?.tone || 'neutral');
    const total = entries.length;
    const healthy = entries.filter((tone) => tone === 'healthy').length;
    const warning = entries.filter((tone) => tone === 'warning').length;
    const critical = entries.filter((tone) => tone === 'critical').length;
    const pending = entries.filter((tone) => tone === 'neutral').length;
    const compliance = total ? round(((healthy + warning) / total) * 100, 1) : 0;
    push({
      id: 'soilResistivity',
      label: 'Soil Resistivity Test',
      testArea: 'Soil Resistivity',
      total,
      healthy,
      warning,
      critical,
      pending,
      compliance,
      healthyLabel: `${healthy} (Low)`,
      warningLabel: `${warning} (Medium)`,
      criticalLabel: `${critical} (High)`,
      pendingLabel: pending ? String(pending) : '-'
    });
  }

  const complianceConfigs = [
    {
      enabled: report.tests.electrodeResistance,
      id: 'electrodeResistance',
      label: 'Earth Electrode Resistance Test',
      testArea: 'Earth Pit Resistance',
      rows: report.electrodeResistance.map((row) => deriveElectrodeAssessment(row).status)
    },
    {
      enabled: report.tests.continuityTest,
      id: 'continuityTest',
      label: 'Earthing Continuity Test',
      testArea: 'Grid Integrity',
      rows: report.continuityTest.map((row) => deriveContinuityAssessment(row).status)
    },
    {
      enabled: report.tests.loopImpedanceTest,
      id: 'loopImpedanceTest',
      label: 'Earth Loop Impedance Test',
      testArea: 'Loop Impedance',
      rows: report.loopImpedanceTest.map((row) => deriveLoopImpedanceAssessment(row).status)
    },
    {
      enabled: report.tests.prospectiveFaultCurrent,
      id: 'prospectiveFaultCurrent',
      label: 'Prospective Fault Current Test',
      testArea: 'Fault Current',
      rows: report.prospectiveFaultCurrent.map((row) => deriveFaultCurrentAssessment(row).status)
    },
    {
      enabled: report.tests.riserIntegrityTest,
      id: 'riserIntegrityTest',
      label: 'Riser Integrity & Bonding Test',
      testArea: 'Riser Integrity',
      rows: report.riserIntegrityTest.map((row) => deriveRiserAssessment(row).status)
    },
    {
      enabled: report.tests.earthContinuityTest,
      id: 'earthContinuityTest',
      label: 'Ground Resistance / Earth Continuity Test',
      testArea: 'Earth Continuity',
      rows: report.earthContinuityTest.map((row) => deriveEarthContinuityAssessment(row).status)
    },
    {
      enabled: report.tests.towerFootingResistance,
      id: 'towerFootingResistance',
      label: 'Tower Footing Test',
      testArea: 'Tower Footing',
      rows: report.towerFootingResistance.map((group) => deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group))).status)
    }
  ];

  complianceConfigs.forEach((config) => {
    if (!config.enabled) {
      return;
    }
    const entries = config.rows.map((status) => status?.tone || 'neutral');
    const total = entries.length;
    const healthy = entries.filter((tone) => tone === 'healthy').length;
    const warning = entries.filter((tone) => tone === 'warning').length;
    const critical = entries.filter((tone) => tone === 'critical').length;
    const pending = entries.filter((tone) => tone === 'neutral').length;
    const compliance = total ? round((healthy / total) * 100, 1) : 0;
    push({
      id: config.id,
      label: config.label,
      testArea: config.testArea,
      total,
      healthy,
      warning,
      critical,
      pending,
      compliance,
      healthyLabel: String(healthy),
      warningLabel: String(warning),
      criticalLabel: String(critical),
      pendingLabel: pending ? String(pending) : '-'
    });
  });

  return rows;
}

function getStoredAiNarrative(report) {
  return report && typeof report.aiNarrative === 'object' && report.aiNarrative ? report.aiNarrative : null;
}

function getAiModuleSummary(report, testId) {
  const narrative = getStoredAiNarrative(report);
  if (!Array.isArray(narrative?.moduleSummaries)) {
    return '';
  }
  const match = narrative.moduleSummaries.find((entry) => String(entry?.testId) === String(testId));
  return safeText(match?.summary, '');
}

function renderAiModuleSummary(doc, report, testId, drawChrome, heading = 'Assessment Summary') {
  const summary = getAiModuleSummary(report, testId);
  if (!summary) {
    return;
  }
  drawSubsectionParagraph(doc, heading, summary, drawChrome);
}

function buildOverallAssessmentText(report, summaryRows) {
  const narrative = getStoredAiNarrative(report);
  if (safeText(narrative?.overallAssessment, '')) {
    return safeText(narrative.overallAssessment, '');
  }

  if (!summaryRows.length) {
    return 'No measurement sections were selected for this report.';
  }

  const fullyCompliant = summaryRows.filter((row) => row.compliance >= 99.9).length;
  const criticalRows = summaryRows.filter((row) => row.critical > 0);
  const warningRows = summaryRows.filter((row) => row.warning > 0 && row.critical === 0);

  let overallBand = 'HEALTHY';
  if (criticalRows.length) {
    overallBand = 'ATTENTION REQUIRED';
  } else if (warningRows.length) {
    overallBand = 'MONITORING REQUIRED';
  }

  const averageCompliance = round(
    summaryRows.reduce((sum, row) => sum + row.compliance, 0) / Math.max(summaryRows.length, 1),
    1
  );
  const parts = [
    `The earthing system is presently in ${overallBand} condition with an average compliance of ${averageCompliance}%.`,
    `${fullyCompliant} out of ${summaryRows.length} selected test categories achieved full compliance based on the recorded measurements.`
  ];

  if (criticalRows.length) {
    parts.push(
      `Primary areas requiring corrective action are ${criticalRows.map((row) => row.testArea).join(', ')}.`
    );
  }

  if (warningRows.length) {
    parts.push(
      `The following sections remain in a monitoring / warning band: ${warningRows.map((row) => row.testArea).join(', ')}.`
    );
  }

  return parts.join(' ');
}

function getExecutiveFindingConfig(row) {
  const configs = {
    soilResistivity: {
      action: 'Review the soil classification and consider soil conditioning or earthing design optimization where high-resistivity locations are recorded.',
      compliantAction: 'Continue periodic monitoring of soil resistivity trends during the dry season and retain the current earthing design basis.',
      warningAction: 'Maintain periodic monitoring and review whether preventive treatment is required at medium soil resistivity locations.'
    },
    electrodeResistance: {
      action: 'Inspect earth pits, improve electrode bonding, clean terminations, and plan soil/electrode improvement before the next verification cycle.',
      compliantAction: 'Continue routine testing and quarterly visual inspection of earth pits, covers, and bonding connections.',
      warningAction: 'Schedule enhanced monitoring, inspect physical condition of pits and connections, and plan preventive maintenance before deterioration increases.'
    },
    continuityTest: {
      action: 'Inspect joints, lugs, bonding links, and tray continuity points; re-test after rectification.',
      compliantAction: 'Maintain periodic bonding continuity checks as part of routine earthing-system inspection.',
      warningAction: 'Inspect suspect joints and bonding links, and schedule follow-up continuity testing.'
    },
    loopImpedanceTest: {
      action: 'Review circuit protective-device selection and confirm the measured loop impedance against the device-specific maximum Zs values.',
      compliantAction: 'Maintain routine verification of loop impedance and protective-device coordination.',
      warningAction: 'Review the relevant circuit/device-specific maximum Zs and keep the affected feeder under monitoring.'
    },
    prospectiveFaultCurrent: {
      action: 'Review the feeder protective device and confirm that breaking capacity remains adequate for the measured fault level.',
      compliantAction: 'Maintain routine protective-device verification and feeder fault-level checks.',
      warningAction: 'Confirm device margin against measured fault current and monitor the affected feeder closely.'
    },
    riserIntegrityTest: {
      action: 'Inspect riser clamps, joints, and bonding interfaces; restore low-resistance continuity and re-test after correction.',
      compliantAction: 'Continue periodic riser integrity and bonding verification as part of routine annual inspection.',
      warningAction: 'Inspect the affected riser/bonding points and schedule a follow-up measurement after maintenance.'
    },
    earthContinuityTest: {
      action: 'Inspect the earth path, terminations, and bonding continuity at the affected points and re-test after rectification.',
      compliantAction: 'Continue periodic verification of earth continuity at the tested points.',
      warningAction: 'Inspect the affected earth path and keep the location under enhanced monitoring.'
    },
    towerFootingResistance: {
      action: 'Review the affected tower footing group, inspect footing connections, and plan grounding improvement where Zt exceeds the grouped tolerable limit.',
      compliantAction: 'Continue annual grouped tower footing verification against the fixed Zsat reference.',
      warningAction: 'Keep the tower group under enhanced monitoring and review footing condition before the next dry-season test.'
    }
  };
  return configs[row.id] || {
    action: 'Review the measurement results and take corrective action where the values exceed the acceptable band.',
    compliantAction: 'Continue routine periodic verification.',
    warningAction: 'Keep the affected section under monitoring and review during the next inspection cycle.'
  };
}

function buildExecutiveFindingRows(report, summaryRows) {
  const narrative = getStoredAiNarrative(report);
  if (Array.isArray(narrative?.keyFindings) && narrative.keyFindings.length) {
    return narrative.keyFindings.map((row, index) => ({
      srNo: String(index + 1),
      testArea: safeText(row.testArea, 'Selected Scope'),
      keyFinding: safeText(row.keyFinding, '-'),
      priority: safeText(row.priority, 'P4 — Normal'),
      recommendedAction: safeText(row.recommendedAction, '-'),
      status: safeText(row.status, 'Compliant')
    }));
  }

  const rows = [];
  let srNo = 1;

  [...summaryRows]
    .sort((left, right) => {
      const leftRank = left.critical > 0 ? 0 : left.warning > 0 ? 1 : 2;
      const rightRank = right.critical > 0 ? 0 : right.warning > 0 ? 1 : 2;
      if (leftRank !== rightRank) {
        return leftRank - rightRank;
      }
      return left.compliance - right.compliance;
    })
    .forEach((row) => {
      const config = getExecutiveFindingConfig(row);
      if (row.critical > 0) {
        const priority = row.critical >= 3 || row.compliance < 70 ? 'P1 — Critical' : 'P2 — High';
      rows.push({
        srNo: String(srNo++),
        testArea: row.testArea,
        keyFinding: `${row.critical} reading(s) exceed the permissible limit within ${row.testArea}.`,
        priority,
        recommendedAction: config.action,
        status: 'Action Required'
      });
    } else if (row.warning > 0) {
      rows.push({
        srNo: String(srNo++),
        testArea: row.testArea,
        keyFinding: `${row.warning} reading(s) are in the warning / marginal band within ${row.testArea}.`,
        priority: 'P3 — Moderate',
        recommendedAction: config.warningAction,
        status: 'Monitoring'
      });
    }
  });

  const compliantRows = summaryRows.filter((row) => row.critical === 0 && row.warning === 0);
  if (compliantRows.length) {
    rows.push({
      srNo: String(srNo++),
      testArea: 'All Other Tests',
      keyFinding: `${compliantRows.length} selected test area(s) are fully compliant with no warning or critical observations in the recorded dataset.`,
      priority: 'P4 — Normal',
      recommendedAction: 'Continue routine annual testing and quarterly visual inspections as per the applicable standards and site maintenance plan.',
      status: 'Compliant'
    });
  }

  return rows;
}

function buildSurveyScopeRows(report) {
  return getSelectedTests(report).map((test) => [test.label, '']);
}

function buildRepresentativeRows(count = 2) {
  return Array.from({ length: count }, () => ['', '', '', '']);
}

function renderSurveyOverviewSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Survey Overview', 'Project snapshot, executed scope, instruments, and applicable standards.', drawChrome);
  drawSurveyOverviewBlock(
    doc,
    'Project Information',
    [
      { label: 'Name of the Client', value: report.project.clientName },
      { label: 'Address', value: '' },
      { label: 'Scope of Work', value: getReportTitle(report) },
      { label: 'Date of Survey', value: report.project.reportDate }
    ],
    drawChrome
  );

  doc.moveDown(0.5);
  doc
    .font('Helvetica-Bold')
    .fontSize(9)
    .fillColor(COLORS.brandDark)
    .text('DETAILS OF PROJECT REPRESENTATIVE', PAGE.margin, doc.y, {
      width: doc.page.width - PAGE.margin * 2
    });
  doc.moveDown(0.6);
  drawTable(
    doc,
    [
      { key: 'organization', label: 'Organization', width: 0.26 },
      { key: 'representativeName', label: 'Representative Name', width: 0.22 },
      { key: 'designation', label: 'Designation / Department', width: 0.2 },
      { key: 'email', label: 'Email', width: 0.18 },
      { key: 'mobile', label: 'Mobile', width: 0.14 }
    ],
    [
      ...buildRepresentativeRows(2).map(() => ({
        organization: 'Client-side Representatives',
        representativeName: '',
        designation: '',
        email: '',
        mobile: ''
      })),
      ...buildRepresentativeRows(3).map(() => ({
        organization: 'Elegrow-side Representatives',
        representativeName: '',
        designation: '',
        email: '',
        mobile: ''
      }))
    ],
    drawChrome
  );

  drawSubsectionParagraph(
    doc,
    'Scope of Work',
    'Only the measurement sheets selected in the report builder are listed below. Quantity can be filled by the field engineer after export if required for the final client submission.',
    drawChrome
  );
  drawSimpleTable(doc, ['Test Carried Out', 'Measured Qty (Nos.)'], buildSurveyScopeRows(report), drawChrome, [0.8, 0.2]);

  drawSubsectionParagraph(doc, 'Details of Instruments', 'The following approved instruments are available for earthing-system health assessment reports.', drawChrome);
  getSelectedEquipment(report).forEach((equipment) => {
    const details = EQUIPMENT_DETAILS[equipment.id];
    if (!details) {
      return;
    }
    drawSubsectionParagraph(doc, details.label, details.description, drawChrome);
    drawSimpleTable(doc, ['Parameter', 'Value'], details.details, drawChrome, [0.34, 0.66]);
  });

  drawSubsectionParagraph(
    doc,
    'Applicable Standards & Regulations',
    'The following standards are referenced based on the selected scope of work for this report.',
    drawChrome
  );
  drawBulletList(doc, getApplicableStandards(report), drawChrome);
}

function renderExecutiveSummarySection(doc, report, drawChrome) {
  const summaryRows = buildComplianceSummaryRows(report);
  const findings = buildExecutiveFindingRows(report, summaryRows);
  const narrative = getStoredAiNarrative(report);

  drawSectionTitle(doc, 'Executive Summary of Earthing System Health Assessment', 'Overall assessment, compliance summary, and key recommended actions.', drawChrome);

  drawSubsectionParagraph(doc, 'Overall Assessment', buildOverallAssessmentText(report, summaryRows), drawChrome);

  if (safeText(narrative?.executiveSummary, '')) {
    drawParagraph(doc, safeText(narrative.executiveSummary, ''), drawChrome);
  }

  drawBarChart(
    doc,
    'Test-Wise Compliance (%)',
    summaryRows.map((row) => ({
      label: row.testArea,
      value: row.compliance,
      tone: row.compliance >= 99.9 ? 'healthy' : row.compliance >= 80 ? 'warning' : 'critical'
    })),
    drawChrome
  );

  drawSubsectionParagraph(doc, 'Test-Wise Compliance Summary', 'Compliance roll-up based on the recorded measurements in the selected test sheets.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr.', width: 0.05, align: 'center', headerSize: 7.2 },
      { key: 'testCategory', label: 'Test Category', width: 0.22, headerSize: 7.2 },
      { key: 'totalTested', label: 'Total Tested', width: 0.11, align: 'center', headerSize: 7.2 },
      { key: 'healthy', label: 'Healthy / Within Limit', width: 0.17, align: 'center', headerSize: 7.2 },
      { key: 'warning', label: 'Warning / Marginal', width: 0.14, align: 'center', headerSize: 7.2 },
      { key: 'critical', label: 'Exceeds Limit / Issues', width: 0.14, align: 'center', headerSize: 7.2 },
      { key: 'pending', label: 'Not Measurable', width: 0.09, align: 'center', headerSize: 7.2 },
      { key: 'compliance', label: 'Compliance (%)', width: 0.08, align: 'center', headerSize: 7.2 }
    ],
    summaryRows.map((row, index) => ({
      srNo: String(index + 1),
      testCategory: row.label,
      totalTested: String(row.total),
      healthy: row.healthyLabel,
      warning: row.warningLabel,
      critical: row.criticalLabel,
      pending: row.pendingLabel,
      compliance: String(row.compliance),
      __cellStyles: {
        compliance:
          row.compliance >= 99.9
            ? { fill: COLORS.healthySoft, text: COLORS.healthy, bold: true }
            : row.compliance >= 80
              ? { fill: COLORS.warningSoft, text: COLORS.warning, bold: true }
              : { fill: COLORS.criticalSoft, text: COLORS.critical, bold: true }
      }
    })),
    drawChrome
  );

  drawSubsectionParagraph(doc, 'Key Findings & Recommended Actions', 'Priority-based actions derived from the measured compliance bands and issue severity.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr.', width: 0.05, align: 'center', headerSize: 7.2 },
      { key: 'testArea', label: 'Test Area', width: 0.16, headerSize: 7.2 },
      { key: 'keyFinding', label: 'Key Finding', width: 0.24, headerSize: 7.2 },
      { key: 'priority', label: 'Priority', width: 0.14, align: 'center', headerSize: 7.2 },
      { key: 'recommendedAction', label: 'Recommended Action', width: 0.28, headerSize: 7.2 },
      { key: 'status', label: 'Status', width: 0.13, align: 'center', headerSize: 7.2 }
    ],
    findings.map((row) => ({
      ...row,
      __cellStyles: {
        priority: row.priority.startsWith('P1')
          ? { fill: '#f7dce2', text: COLORS.critical, bold: true }
          : row.priority.startsWith('P2')
            ? { fill: '#fdecc8', text: '#e46a00', bold: true }
            : row.priority.startsWith('P3')
              ? { fill: '#fff3c9', text: '#9a7400', bold: true }
              : { fill: '#e2f1db', text: '#406b34', bold: true },
        status: row.status === 'Action Required'
          ? { fill: '#f7dce2', text: COLORS.critical, bold: true }
          : row.status === 'Monitoring'
            ? { fill: '#fff3c9', text: '#9a7400', bold: true }
            : { fill: '#e2f1db', text: '#406b34', bold: true }
      }
    })),
    drawChrome
  );
}

function renderElectrodeSummarySection(doc, drawChrome) {
  drawSectionTitle(
    doc,
    'Summary of Earth Electrode Resistance Test',
    'Engineer-completed summary worksheet for observations, recommendations, and priority bands.',
    drawChrome
  );
  drawSimpleTable(
    doc,
    ['Sr. No.', 'Priority of Action', 'Observation', 'Recommendation', 'Quantity (Nos.)'],
    [
      ['1', 'P1', '', '', ''],
      ['2', 'P2', '', '', ''],
      ['3', 'P3', '', '', ''],
      ['4', 'P4', '', '', '']
    ],
    drawChrome,
    [0.08, 0.16, 0.28, 0.32, 0.16]
  );
  drawSubsectionParagraph(
    doc,
    'Major Points',
    'The following general guidance remains applicable to the earth electrode resistance summary and can be retained in the final client booklet.',
    drawChrome
  );
  drawBulletList(doc, ELECTRODE_SUMMARY_MAJOR_POINTS, drawChrome);
}

function renderRiserSummarySection(doc, report, drawChrome) {
  const counts = { Healthy: 0, 'Un-Healthy': 0 };
  report.riserIntegrityTest.forEach((row) => {
    const label = toRiserCommentLabel(deriveRiserAssessment(row).status) === 'Healthy' ? 'Healthy' : 'Un-Healthy';
    counts[label] += 1;
  });

  drawSectionTitle(
    doc,
    'Summary of Riser / Grid Integrity Test',
    'Status quantity roll-up with engineer-completed observation and recommendation fields.',
    drawChrome
  );
  drawSimpleTable(
    doc,
    ['Sr.', 'Classification', 'Observation', 'Recommendation', 'Qty'],
    [
      ['1', 'Healthy', '', '', String(counts.Healthy)],
      ['2', 'Un-Healthy', '', '', String(counts['Un-Healthy'])]
    ],
    drawChrome,
    [0.08, 0.2, 0.3, 0.3, 0.12]
  );
}

function renderObservationSheetSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Observation Sheet', 'Measurement-sheet-wise quantity and health classification roll-up.', drawChrome);
  drawSimpleTable(
    doc,
    ['Test Carried Out', 'Measured Qty', 'Healthy', 'Warnings', 'Not Measurable', '> Permissible Limits'],
    buildObservationSheetRows(report),
    drawChrome,
    [0.32, 0.12, 0.12, 0.12, 0.14, 0.18]
  );
}

function renderMethodologySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Earthing System Health Assessment - Methodology', 'Purpose, scope, abbreviations, standards, workflow, and site methodology.', drawChrome);
  drawSubsectionParagraph(
    doc,
    'Purpose',
    'The primary purpose of the assessment is to verify personnel safety, confirm the fault-current return path, protect equipment and infrastructure, and establish a documented baseline of the earthing system health for compliance and maintenance planning. The report is intended to be direct, technically useful, and suitable for final client submission.',
    drawChrome
  );
  drawSubsectionParagraph(
    doc,
    'Scope',
    'The scope covers the selected earthing-system tests recorded in the report builder, including field measurements, comparison against the relevant standards or project reference limits, interpretation of results, and capture of row-level observations where available. The report is limited to the sections selected during execution and export.',
    drawChrome
  );
  drawSubsectionParagraph(doc, 'Abbreviations', 'Common abbreviations used in the report are listed below.', drawChrome);
  drawSimpleTable(doc, ['Abbreviation', 'Meaning'], ABBREVIATIONS, drawChrome, [0.2, 0.8]);
  drawSubsectionParagraph(doc, 'Applicable Standards & Legislation', 'The following standards and legislation apply to this report scope.', drawChrome);
  drawBulletList(doc, getApplicableStandards(report), drawChrome);
  drawSubsectionParagraph(
    doc,
    'Overall Workflow Process',
    'The report workflow remains consistent for all earthing system health assessment reports.',
    drawChrome
  );
  drawBulletList(
    doc,
    [
      'Phase 1: Pre-assessment data collection, document review, and location planning.',
      'Phase 2: On-site testing activities for the selected measurement sheets using calibrated instruments.',
      'Phase 3: Reporting, classification, recommendation capture, and final client-ready compilation.'
    ],
    drawChrome
  );
}

function getMethodologyBlocks(report) {
  const blocks = [];
  if (report.tests.soilResistivity) {
    blocks.push({
      title: 'Soil Resistivity Testing',
      objective: 'To determine the electrical resistivity of the soil at multiple probe spacings, which is fundamental to validating the performance of the earthing system and the standard resistance of earthing electrodes.',
      methodology: [
        'Four electrodes are placed in a straight line at equal intervals using the Wenner method.',
        'The outer probes inject current while the inner probes measure the voltage drop.',
        'The test is repeated across multiple probe spacings and in two directions to account for soil inhomogeneity.',
        'Recorded resistivity values are averaged to produce a report-level mean soil resistivity.'
      ]
    });
  }
  if (report.tests.electrodeResistance) {
    blocks.push({
      title: 'Earth Electrode Resistance Testing',
      objective: 'To measure the resistance of individual earth electrodes and compare the measured values against the standard resistance derived from soil conditions and electrode geometry.',
      methodology: [
        'The electrode is measured with the required field test setup and the entered with-grid value is treated as the primary acceptance value.',
        'Measured values are compared with the applicable project or standard reference limit.',
        'Observations and engineer remarks may be captured at row level for each electrode location.'
      ]
    });
  }
  if (report.tests.continuityTest) {
    blocks.push({
      title: 'Earthing Continuity Testing',
      objective: 'To verify the integrity of protective bonding conductors and confirm a low-resistance continuity path between connected points.',
      methodology: [
        'Point-to-point resistance readings are recorded between the selected main location and measurement point.',
        'Measured values are compared with the continuity reference band to derive the automatic status and comment.',
        'Any abnormal connection, joint, or bonding condition can be added as a row observation.'
      ]
    });
  }
  if (report.tests.loopImpedanceTest || report.tests.prospectiveFaultCurrent) {
    blocks.push({
      title: 'Earth Loop Impedance and Prospective Fault Current Testing',
      objective: 'To verify the earth fault loop path and confirm that the available fault current is compatible with the protective device capability for timely operation.',
      methodology: [
        'Loop impedance measurements are recorded at the selected panel, feeder, or equipment points.',
        'Prospective fault current is recorded together with device type, rating, and breaking capacity where applicable.',
        'Results are interpreted against the relevant device and protection design criteria.'
      ]
    });
  }
  if (report.tests.riserIntegrityTest) {
    blocks.push({
      title: 'Riser / Grid Integrity Testing',
      objective: 'To verify continuity and bonding integrity from the equipment side toward the grid side of the riser or bonded path.',
      methodology: [
        'Resistance is measured toward equipment and toward the grid for each selected point.',
        'Both values are reviewed together to determine continuity health and bonding quality.',
        'Any discrepancy is flagged for physical inspection of lugs, joints, and riser path integrity.'
      ]
    });
  }
  if (report.tests.earthContinuityTest) {
    blocks.push({
      title: 'Earth Continuity / Ground Resistance Testing',
      objective: 'To confirm the condition of the earth path at the measured points and compare the recorded values against the continuity reference band.',
      methodology: [
        'The selected tag, location, and measured value are recorded for each point.',
        'Values are automatically classified into healthy, warning, or > permissible limit bands.',
        'Engineer observations can be appended row-wise where field conditions require further note.'
      ]
    });
  }
  if (report.tests.towerFootingResistance) {
    blocks.push({
      title: 'Tower Footing Resistance Testing',
      objective: 'To evaluate each tower location using four fixed footing points and determine the grouped impedance condition against the fixed Zsat limit.',
      methodology: [
        'Each tower group contains fixed sub-rows Foot-1 to Foot-4.',
        'Grouped total impedance Zt is the sum of the four measured impedance values.',
        'Grouped total current is the sum of the four measured current readings as recorded for the tower location.'
      ]
    });
  }
  return blocks;
}

function renderSelectedMethodologies(doc, report, drawChrome) {
  drawSubsectionParagraph(doc, 'On-Site Testing Methodologies', 'Selected methodologies are listed in the same serial order as the report scope.', drawChrome);
  getMethodologyBlocks(report).forEach((block, index) => {
    drawSubsectionParagraph(doc, `${index + 1}. ${block.title} - Objective`, block.objective, drawChrome);
    drawBulletList(doc, block.methodology, drawChrome);
  });
  drawSubsectionParagraph(doc, 'Equipment Required', 'Only the selected earthing-system instruments are listed below.', drawChrome);
  drawBulletList(doc, getSelectedEquipment(report).map((equipment) => equipment.label), drawChrome);
  drawSubsectionParagraph(doc, 'Safety Measures', 'The following safety measures apply to all reports.', drawChrome);
  drawBulletList(doc, SAFETY_MEASURES, drawChrome);
  drawSubsectionParagraph(doc, 'Reporting', 'The final report package includes the following standard elements.', drawChrome);
  drawBulletList(doc, REPORTING_POINTS, drawChrome);
}

function renderHealthClassificationSection(doc, report, drawChrome) {
  const rows = [];
  const pushRows = (no, test, standard, entries) => {
    entries.forEach((entry, index) => {
      const style =
        entry.tone === 'healthy'
          ? { fill: COLORS.healthySoft, text: COLORS.healthy, bold: true }
          : entry.tone === 'warning'
            ? { fill: COLORS.warningSoft, text: COLORS.warning, bold: true }
            : { fill: COLORS.criticalSoft, text: COLORS.critical, bold: true };
      rows.push({
        no: index === 0 ? String(no) : '',
        test,
        data: entry.data,
        standard,
        indication: entry.indication,
        __allowEmpty: ['no'],
        __cellStyles: {
          indication: style
        }
      });
    });
  };

  let srNo = 1;
  if (report.tests.soilResistivity) {
    pushRows(srNo++, 'Soil Resistivity Test', 'IEEE 81-2012, Table 1 · IS 3043:2018', [
      { data: 'Soil Resistivity < Standard Limit of Soil Resistivity', indication: 'Low', tone: 'healthy' },
      { data: 'Soil Resistivity within 10% of Standard Limit of Soil Resistivity', indication: 'Medium', tone: 'warning' },
      { data: 'Soil Resistivity > Standard Limit of Soil Resistivity', indication: 'High', tone: 'critical' }
    ]);
  }
  if (report.tests.electrodeResistance) {
    pushRows(srNo++, 'Earth Electrode Resistance Test', 'IS 3043:2018, Clause 9.2 – Resistance of Common Types of Earth Electrodes', [
      { data: 'Earth Pit Resistance < Standard Limit of Earth Electrode Resistance', indication: 'Healthy', tone: 'healthy' },
      { data: 'Earth Pit Resistance within 10% of Standard Limit of Earth Electrode Resistance', indication: 'Warning', tone: 'warning' },
      { data: 'Earth Pit Resistance > Standard Limit of Earth Electrode Resistance', indication: '> Permissible Limit', tone: 'critical' }
    ]);
  }
  if (report.tests.continuityTest) {
    pushRows(srNo++, 'Earthing Continuity Test', 'IS 3043 continuity guidance · IEC 60364-6 continuity verification', [
      { data: 'Grid / bonding continuity resistance in the healthy band', indication: 'Healthy', tone: 'healthy' },
      { data: 'Grid / bonding continuity resistance within warning band', indication: 'Warning', tone: 'warning' },
      { data: 'Grid / bonding continuity resistance above acceptable band', indication: '> Permissible Limit', tone: 'critical' }
    ]);
  }
  if (report.tests.loopImpedanceTest) {
    pushRows(srNo++, 'Earth Loop Impedance Test', 'IS/IEC 60364-6:2016', [
      { data: 'Earth Loop Impedance < Standard Limit of Earth Loop Impedance', indication: 'Healthy', tone: 'healthy' },
      { data: 'Earth Loop Impedance within 10% of Standard Limit of Earth Loop Impedance', indication: 'Warning', tone: 'warning' },
      { data: 'Earth Loop Impedance > Standard Limit of Earth Loop Impedance', indication: '> Permissible Limit', tone: 'critical' }
    ]);
  }
  if (report.tests.prospectiveFaultCurrent) {
    pushRows(srNo++, 'Prospective Fault Current Test', 'IS/IEC 60364-6:2016', [
      { data: 'Prospective Fault Current > Breaking Capacity of Protective Device', indication: '> Permissible Limit', tone: 'critical' }
    ]);
  }
  if (report.tests.riserIntegrityTest) {
    pushRows(srNo++, 'Riser Integrity and Bonding Test', 'IEEE 81-2012, Clause 10', [
      { data: 'Riser Bonding Resistance <= 1 ohm; continuity confirmed between equipment and earth grid', indication: 'Healthy', tone: 'healthy' },
      { data: 'Riser Bonding Resistance > 1 ohm or open circuit detected', indication: 'Un-Healthy', tone: 'critical' }
    ]);
  }
  if (report.tests.earthContinuityTest) {
    pushRows(srNo++, 'Ground Resistance / Earth Continuity Test', 'IS 3043:2018 · IEC 60364 continuity verification', [
      { data: 'Measured value within reference band', indication: 'Healthy', tone: 'healthy' },
      { data: 'Measured value within 10% of limiting band', indication: 'Warning', tone: 'warning' },
      { data: 'Measured value above acceptable continuity band', indication: '> Permissible Limit', tone: 'critical' }
    ]);
  }
  if (report.tests.towerFootingResistance) {
    pushRows(srNo++, 'Tower Footing Test', 'Project grouped Zsat limit = 10 ohm', [
      { data: 'Total Impedance Zt ≤ Zsat', indication: 'Healthy', tone: 'healthy' },
      { data: 'Total Impedance Zt within 20% above Zsat', indication: 'Marginal', tone: 'warning' },
      { data: 'Total Impedance Zt > grouped tolerable limit', indication: '> Permissible Limit', tone: 'critical' }
    ]);
  }

  drawSectionTitle(doc, 'Healthiness Classifications', 'Classification rules applied to the selected measurement sheets.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'no', label: 'No', width: 0.05, align: 'center' },
      { key: 'test', label: 'Test', width: 0.2 },
      { key: 'data', label: 'Data', width: 0.36 },
      { key: 'standard', label: 'Standard', width: 0.29 },
      { key: 'indication', label: 'Indication', width: 0.1, align: 'center' }
    ],
    rows,
    drawChrome
  );
}

function renderOverviewSection(doc, report, drawChrome) {
  const snapshot = buildExecutiveSnapshot(report);
  const soilSummary = snapshot.soil;
  const tone = soilSummary.category?.tone || 'neutral';
  const electrodeText =
    snapshot.totalElectrodes > 0 ? `${snapshot.healthyElectrodes}/${snapshot.totalElectrodes}` : '0';
  const overviewRows = [
    { label: 'Client Name', value: report.project.clientName },
    { label: 'Project Number', value: report.project.projectNo },
    { label: 'Site Location', value: report.project.siteLocation },
    { label: 'Work Order / Ref', value: report.project.workOrder },
    { label: 'Date of Testing', value: report.project.reportDate },
    { label: 'Testing Engineer', value: report.project.engineerName }
  ];

  if (report.project.zohoProjectName) {
    overviewRows.push({ label: 'Zoho Project', value: report.project.zohoProjectName });
  }
  if (report.project.zohoProjectStage) {
    overviewRows.push({ label: 'Zoho Stage', value: report.project.zohoProjectStage });
  }

  drawSectionTitle(doc, 'Project Overview', 'Core project details and measurement scope.', drawChrome);
  drawKeyValueGrid(doc, overviewRows, drawChrome);

  drawMetricCards(
    doc,
    [
      {
        label: 'Selected Tests',
        value: String(snapshot.selectedTests.length),
        meta: 'Included',
        tone: 'neutral'
      },
      {
        label: 'Mean Soil Resistivity',
        value: soilSummary.overallAverage === null ? '-' : `${soilSummary.overallAverage}`,
        meta: soilSummary.category.label,
        tone
      },
      {
        label: 'Healthy Electrodes',
        value: electrodeText,
        meta: snapshot.totalElectrodes ? 'Within limit' : 'No data',
        tone: snapshot.totalElectrodes ? 'healthy' : 'neutral'
      }
    ],
    drawChrome
  );

  drawParagraph(
    doc,
    `This report was generated from field entries captured in ElectroReports for ${safeText(report.project.clientName)} at ${safeText(
      report.project.siteLocation
    )}. Only the selected test modules are included, which keeps the report aligned with the executed scope of work.`,
    drawChrome
  );
}

function renderSoilTheorySection(doc, report, drawChrome) {
  const summary = calculateSoilSummary(report);

  drawSectionTitle(
    doc,
    'Soil Resistivity Test - Theory & Methodology',
    'Soil classification logic, field method, and report-level overview before the analysis pages.',
    drawChrome
  );

  drawMetricCards(
    doc,
    [
      {
        label: 'Direction 1 Average',
        value: summary.direction1Average === null ? '-' : `${summary.direction1Average}`,
        meta: 'ohm-m',
        tone: summary.category.tone
      },
      {
        label: 'Direction 2 Average',
        value: summary.direction2Average === null ? '-' : `${summary.direction2Average}`,
        meta: 'ohm-m',
        tone: summary.category.tone
      },
      {
        label: 'Overall Mean',
        value: summary.overallAverage === null ? '-' : `${summary.overallAverage}`,
        meta: summary.category.label,
        tone: summary.category.tone
      }
    ],
    drawChrome
  );

  renderAiModuleSummary(doc, report, 'soilResistivity', drawChrome);

  drawSubsectionParagraph(
    doc,
    'Theory',
    'Soil resistivity is measured using the Wenner four-pin method to understand the resistive behavior of the soil at different probe spacings. Readings are recorded location-wise in two directions so each tested location can be assessed individually and then rolled up into a report-level mean.',
    drawChrome
  );

  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Four probes are placed at equal spacing in a straight line. The outer probes inject current and the inner probes measure the voltage drop. The same spacing sequence is repeated in two directions for each location, and the resulting resistivity values are analyzed location-wise before deriving the report-level mean.',
    drawChrome
  );
}

function renderSoilGraphSection(doc, report, drawChrome) {
  const locations = getSoilLocations(report);
  let chartsOnPage = 0;

  const startChartPage = (needsNewPage) => {
    if (needsNewPage) {
      addReportPage(doc, drawChrome, { layout: 'portrait' });
    }
    drawSectionTitle(
      doc,
      'Soil Resistivity Test - Analysis',
      'Location-wise soil resistivity graphs for both test directions.',
      drawChrome
    );
    chartsOnPage = 0;
  };

  startChartPage(false);

  locations.forEach((location, locationIndex) => {
    if (chartsOnPage === 2) {
      startChartPage(true);
    }
    const categories = (location.direction1 || []).map((row) => safeText(row.spacing, '-'));
    drawSoilLocationChart(
      doc,
      `Location – ${safeText(location.name, `Location ${locationIndex + 1}`)}`,
      categories,
      (location.direction1 || []).map((row) => Number.parseFloat(row.resistivity)),
      (location.direction2 || []).map((row) => Number.parseFloat(row.resistivity)),
      drawChrome
    );
    chartsOnPage += 1;
  });
}

function renderSoilResultSection(doc, report, drawChrome) {
  const summary = calculateSoilSummary(report);
  const locations = getSoilLocations(report);
  locations.forEach((location, locationIndex) => {
    const locationSummary = summary.locations[locationIndex] || {};

    if (locationIndex > 0) {
      addReportPage(doc, drawChrome, { layout: 'portrait' });
    }
    drawSectionTitle(
      doc,
      'Soil Resistivity Test - Analysis',
      `Measured results for ${safeText(location.name, `Location ${locationIndex + 1}`)} with location-wise mean and category.`,
      drawChrome
    );
    drawSoilAnalysisTable(doc, location, locationSummary, locationIndex, drawChrome);

    drawParagraph(
      doc,
      locationSummary.overallAverage === null
        ? 'Summary: This location is awaiting enough numeric data in both directions to produce a mean soil resistivity value.'
        : `Summary: The average soil resistivity for ${safeText(location.name, `Location ${locationIndex + 1}`)} is ${locationSummary.overallAverage} ohm-meter and the soil category is ${locationSummary.category.label}.`,
      drawChrome
    );

    if (location.notes) {
      drawParagraph(doc, `Site note: ${location.notes}`, drawChrome);
    }
  });

  drawParagraph(
    doc,
    summary.overallAverage === null
      ? 'Summary: The report does not yet contain enough numeric data across the tested locations to produce a complete mean soil resistivity value.'
      : `Summary: The mean soil resistivity for this report is ${summary.overallAverage} ohm-m and the overall soil category is ${summary.category.label}. This report-level mean is derived from the tested soil locations and is used for the electrode standard-limit calculations where geometry has been provided.`,
    drawChrome
  );
}

function renderSoilStandardLimitSection(doc, report, drawChrome) {
  if (!report.tests.soilResistivity) {
    return;
  }

  const summary = calculateSoilSummary(report);
  const locations = getSoilLocations(report);

  drawSectionTitle(
    doc,
    'Calculation of Earth Electrode & Grid Resistance (Standard Limit)',
    'Derived location-wise from mean soil resistivity and the entered driven earth electrode geometry.',
    drawChrome
  );

  drawParagraph(
    doc,
    'The standard resistance value of a driven rod / pipe electrode is derived using each location mean soil resistivity and the entered electrode dimensions. The calculation uses R = (ρ / 2πl) × ln(2l / d), where ρ is soil resistivity in ohm-meter, l is electrode length in meters, and d is electrode diameter in meters.',
    drawChrome
  );

  locations.forEach((location, locationIndex) => {
    const locationSummary = summary.locations[locationIndex] || {};
    const diameter = Number.parseFloat(location.drivenElectrodeDiameter);
    const length = Number.parseFloat(location.drivenElectrodeLength);
    const calculated = calculateDrivenElectrodeResistance(locationSummary.overallAverage, length, diameter);

    drawSubsectionParagraph(
      doc,
      `Calculation for ${safeText(location.name, `Location ${locationIndex + 1}`)}`,
      calculated === null
        ? 'Enter the driven earth electrode diameter and length for this location to generate the standard resistance value automatically.'
        : `The calculated standard resistance value for this location is ${calculated} ohm based on the mean soil resistivity and the entered driven electrode dimensions.`,
      drawChrome
    );

    drawSimpleTable(
      doc,
      ['Description', 'Data', 'Units', 'Remarks / Source'],
      [
        ['Average Soil Resistivity ρ', locationSummary.overallAverage === null ? '-' : String(locationSummary.overallAverage), 'ohm-meter', 'Measured from the location soil resistivity sheet'],
        ['Driven Earth Electrode - Diameter', Number.isFinite(diameter) ? String(diameter) : '-', 'millimeter', 'Entered in the soil location measurement sheet'],
        ['Driven Earth Electrode - Length', Number.isFinite(length) ? String(length) : '-', 'meter', 'Entered in the soil location measurement sheet'],
        ['Calculated Standard Resistance', calculated === null ? '-' : String(calculated), 'ohm', 'Derived using the standard electrode formula']
      ],
      drawChrome,
      [0.34, 0.18, 0.12, 0.36]
    );
  });
}

function renderElectrodeSection(doc, report, drawChrome) {
  drawSectionTitle(
    doc,
    'Earthing Electrode Resistance Test - Measurement & Analysis',
    'Measured resistance checked against the standard earth electrode resistance value and the entered with-grid result.',
    drawChrome
  );

  renderAiModuleSummary(doc, report, 'electrodeResistance', drawChrome);

  drawSubsectionParagraph(
    doc,
    'Theory',
    'Measured earth electrode resistance values are compared against the applicable limit to evaluate the effectiveness of the electrode/grid path. The with-grid value is treated as the main acceptance value in the report.',
    drawChrome
  );

  drawTable(
    doc,
    [
      { key: 'tag', label: 'Pit Tag', width: 0.06 },
      { key: 'location', label: 'Location', width: 0.09 },
      { key: 'electrodeType', label: 'Type', width: 0.07 },
      { key: 'materialType', label: 'Material', width: 0.07 },
      { key: 'length', label: 'Length', width: 0.05 },
      { key: 'diameter', label: 'Dia', width: 0.05 },
      { key: 'resistanceWithoutGrid', label: 'Without Grid', width: 0.08 },
      { key: 'resistanceWithGrid', label: 'With Grid', width: 0.08 },
      { key: 'standard', label: 'Standard', width: 0.09 },
      { key: 'comment', label: 'Comment', width: 0.08 },
      { key: 'observation', label: 'Observation', width: 0.11 },
      { key: 'recommendation', label: 'Recommendation', width: 0.12 },
      { key: 'priority', label: 'Priority', width: 0.05 }
    ],
    report.electrodeResistance.map((row) => ({
      ...row,
      resistanceWithoutGrid: safeText(row.resistanceWithoutGrid, ''),
      resistanceWithGrid: safeText(row.resistanceWithGrid || row.measuredResistance, ''),
      standard: 'IS 3043 / 4.60 ohm',
      comment: toReportBandLabel(deriveElectrodeAssessment(row).status),
      observation: '',
      recommendation: '',
      priority: '',
      __cellStyles: {
        comment: getStatusCellStyle(deriveElectrodeAssessment(row).status.tone)
      },
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );

  renderElectrodeSummarySection(doc, drawChrome);
}

function renderContinuitySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Continuity Test', 'Point-to-point continuity readings.', drawChrome);
  renderAiModuleSummary(doc, report, 'continuityTest', drawChrome);
  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Point-to-point continuity measurements are compared against the continuity reference band to confirm whether the bonded path remains electrically sound.',
    drawChrome
  );
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'mainLocation', label: 'Main Location', width: 0.18 },
      { key: 'measurementPoint', label: 'Measurement Point', width: 0.2 },
      { key: 'resistance', label: 'Resistance (ohm)', width: 0.12 },
      { key: 'impedance', label: 'Impedance (ohm)', width: 0.12 },
      { key: 'status', label: 'Status', width: 0.12 },
      { key: 'comment', label: 'Comment', width: 0.18 }
    ],
    report.continuityTest.map((row) => ({
      ...row,
      status: deriveContinuityAssessment(row).status.label,
      comment: safeText(row.comment, deriveContinuityAssessment(row).comment),
      __cellStyles: {
        status: getStatusCellStyle(deriveContinuityAssessment(row).status.tone),
        comment: getStatusCellStyle(deriveContinuityAssessment(row).status.tone)
      },
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderLoopSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Earth Loop Impedance Test - Measurement & Analysis', 'Measured Zs values across selected panels or equipment.', drawChrome);
  renderAiModuleSummary(doc, report, 'loopImpedanceTest', drawChrome);
  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Measured Zs values are reviewed against the expected protection design band. Final compliance remains device-specific, so the report highlights measured values and attention points clearly.',
    drawChrome
  );
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.06 },
      { key: 'location', label: 'Location of Panel / Equipment', width: 0.16 },
      { key: 'feederTag', label: 'Name of Feeder & Tag No.', width: 0.13 },
      { key: 'deviceType', label: 'Type', width: 0.08 },
      { key: 'deviceRating', label: 'Rating (A)', width: 0.08 },
      { key: 'breakingCapacity', label: 'Breaking (kA)', width: 0.08 },
      { key: 'measuredPoints', label: 'Measured Points', width: 0.09 },
      { key: 'loopImpedance', label: 'Loop Impedance (Z)', width: 0.1 },
      { key: 'voltage', label: 'Voltage', width: 0.08 },
      { key: 'remarks', label: 'Comment', width: 0.14 }
    ],
    report.loopImpedanceTest.map((row) => ({
      ...row,
      loopImpedance: safeText(row.loopImpedance || row.measuredZs, ''),
      remarks: safeText(row.remarks, toReportBandLabel(deriveLoopImpedanceAssessment(row).status)),
      __cellStyles: {
        remarks: getStatusCellStyle(deriveLoopImpedanceAssessment(row).status.tone)
      },
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
  drawNoteBox(doc, LOOP_NOTE, drawChrome);
}

function renderFaultCurrentSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Fault Current Test - Measurement & Analysis', 'Feeder details with loop impedance and fault current values.', drawChrome);
  renderAiModuleSummary(doc, report, 'prospectiveFaultCurrent', drawChrome);
  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Prospective fault current is compared against the entered protective-device breaking capacity to identify whether adequate interruption margin is available.',
    drawChrome
  );
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.06 },
      { key: 'location', label: 'Location of Panel / Equipment', width: 0.14 },
      { key: 'feederTag', label: 'Name of Feeder & Tag No.', width: 0.12 },
      { key: 'deviceType', label: 'Device', width: 0.08 },
      { key: 'deviceRating', label: 'Rating (A)', width: 0.08 },
      { key: 'breakingCapacity', label: 'Breaking (kA)', width: 0.09 },
      { key: 'measuredPoints', label: 'Measured Points', width: 0.1 },
      { key: 'loopImpedance', label: 'Loop Z (ohm)', width: 0.08 },
      { key: 'prospectiveFaultCurrent', label: 'PFC', width: 0.08 },
      { key: 'voltage', label: 'Voltage', width: 0.08 },
      { key: 'remark', label: 'Comment', width: 0.09 }
    ],
    report.prospectiveFaultCurrent.map((row) => ({
      ...row,
      remark: safeText(row.comment, toReportBandLabel(deriveFaultCurrentAssessment(row).status)),
      __cellStyles: {
        remark: getStatusCellStyle(deriveFaultCurrentAssessment(row).status.tone)
      },
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
  drawNoteBox(doc, FAULT_NOTE, drawChrome);
}

function renderRiserSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Riser / Grid Integrity Test', 'Resistance verification towards equipment and grid.', drawChrome);
  renderAiModuleSummary(doc, report, 'riserIntegrityTest', drawChrome);
  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Resistance is recorded toward equipment and toward the grid to confirm integrity of the riser path and bonding continuity across the measured point.',
    drawChrome
  );
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'mainLocation', label: 'Main Location', width: 0.16 },
      { key: 'measurementPoint', label: 'Measurement Point', width: 0.2 },
      { key: 'resistanceTowardsEquipment', label: 'Towards Equipment', width: 0.12 },
      { key: 'resistanceTowardsGrid', label: 'Towards Grid', width: 0.12 },
      { key: 'comment', label: 'Comment', width: 0.1 },
      { key: 'observation', label: 'Observation', width: 0.12 },
      { key: 'recommendation', label: 'Recommendation', width: 0.1 }
    ],
    report.riserIntegrityTest.map((row) => ({
      ...row,
      comment: toRiserCommentLabel(deriveRiserAssessment(row).status),
      observation: '',
      recommendation: '',
      __cellStyles: {
        comment: getStatusCellStyle(deriveRiserAssessment(row).status.tone)
      },
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
  drawNoteBox(doc, RISER_NOTE, drawChrome);
  renderRiserSummarySection(doc, report, drawChrome);
}

function renderEarthContinuitySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Earth Continuity Test', 'Earth path verification by location and measured value.', drawChrome);
  renderAiModuleSummary(doc, report, 'earthContinuityTest', drawChrome);
  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Earth continuity values are assessed against the continuity reference band to identify whether the earth path at each recorded location remains acceptable.',
    drawChrome
  );
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'tag', label: 'Tag', width: 0.12 },
      { key: 'locationBuildingName', label: 'Location / Building', width: 0.2 },
      { key: 'distance', label: 'Distance', width: 0.1 },
      { key: 'measuredValue', label: 'Measured Value', width: 0.12 },
      { key: 'status', label: 'Status', width: 0.12 },
      { key: 'remark', label: 'Remark', width: 0.26 }
    ],
    report.earthContinuityTest.map((row) => ({
      ...row,
      status: deriveEarthContinuityAssessment(row).status.label,
      remark: safeText(row.remark, deriveEarthContinuityAssessment(row).comment),
      __cellStyles: {
        status: getStatusCellStyle(deriveEarthContinuityAssessment(row).status.tone),
        remark: getStatusCellStyle(deriveEarthContinuityAssessment(row).status.tone)
      },
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderTowerFootingSection(doc, report, drawChrome) {
  const towerSummaries = summarizeTowerGroups(report.towerFootingResistance);
  const towerGroups = Array.isArray(report.towerFootingResistance) ? report.towerFootingResistance : [];

  drawSectionTitle(
    doc,
    'Tower Footing Resistance Measurement & Analysis',
    'Each tower location is evaluated using 4 fixed footing rows: Foot-1, Foot-2, Foot-3, and Foot-4.',
    drawChrome
  );

  renderAiModuleSummary(doc, report, 'towerFootingResistance', drawChrome);

  drawSubsectionParagraph(
    doc,
    'Objective',
    'Assess each tower location by summing the 4 footing impedance readings to obtain Zt and summing the 4 footing current readings to obtain the tower current total.',
    drawChrome
  );
  drawSubsectionParagraph(
    doc,
    'Scope',
    'This section covers one main tower location per grouped block, with fixed measurement point locations Foot-1 through Foot-4 and shared totals/remarks across the 4 footing rows.',
    drawChrome
  );
  drawSubsectionParagraph(
    doc,
    'Standards',
    'Project default tolerable impedance limit Zsat is taken as 10 ohm unless edited. Measurements are analyzed using the grouped comparison logic requested for this report.',
    drawChrome
  );
  drawSubsectionParagraph(
    doc,
    'Theory',
    'For each tower location, the total impedance Zt is obtained by summing the measured impedance values of Foot-1, Foot-2, Foot-3, and Foot-4. The total current is obtained by summing the 4 measured current values exactly as recorded in the footing rows.',
    drawChrome
  );
  drawSubsectionParagraph(
    doc,
    'Methodology',
    'Each tower location retains 4 fixed footing rows. Shared values are shown once per group in the Excel-style structure. Totals remain blank until all 4 required values for that calculation are entered. Remarks are assigned by comparing Zt against Zsat.',
    drawChrome
  );

  const tableRows = towerGroups.flatMap((group) => {
    const assessment = deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group)));
    const readings = Array.isArray(group.readings) ? group.readings : [];
    return readings.map((reading, readingIndex) => ({
      srNo: safeText(group.srNo, '-'),
      mainLocationTower: safeText(group.mainLocationTower, '-'),
      measurementPointLocation: safeText(reading.measurementPointLocation, `Foot-${readingIndex + 1}`),
      footToEarthingConnectionStatus: safeText(reading.footToEarthingConnectionStatus, 'Given'),
      measuredCurrentMa: safeText(reading.measuredCurrentMa, ''),
      measuredImpedance: safeText(reading.measuredImpedance, ''),
      standardTolerableImpedanceZsat: '10',
      totalImpedanceZt: assessment.totalImpedanceZt === null ? '-' : String(assessment.totalImpedanceZt),
      totalCurrentItotal: assessment.totalCurrentItotal === null ? '-' : String(assessment.totalCurrentItotal),
      remarks: assessment.comment,
      __cellStyles: {
        remarks: getStatusCellStyle(assessment.status.tone)
      },
      __observationBlocks: [buildObservationBlock(reading, safeText(reading.measurementPointLocation, 'Foot'))].filter(Boolean)
    }));
  });

  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.06 },
      { key: 'mainLocationTower', label: 'Main Location – Tower', width: 0.13 },
      { key: 'measurementPointLocation', label: 'Measurement Point Location', width: 0.12 },
      { key: 'footToEarthingConnectionStatus', label: 'Connection Status', width: 0.1 },
      { key: 'measuredCurrentMa', label: 'Measured Current I (mA)', width: 0.09 },
      { key: 'measuredImpedance', label: 'Measured Impedance (ohm)', width: 0.09 },
      { key: 'totalImpedanceZt', label: 'Total Impedance Zt (ohm)', width: 0.1 },
      { key: 'totalCurrentItotal', label: 'Total Current | Total (A)', width: 0.09 },
      { key: 'standardTolerableImpedanceZsat', label: 'Zsat', width: 0.08 },
      { key: 'remarks', label: 'Remarks', width: 0.14 }
    ],
    tableRows,
    drawChrome
  );

  drawSubsectionParagraph(
    doc,
    'Interpretation',
    towerGroups.length
      ? towerGroups
          .map((group) => {
            const assessment = deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group)));
            return `${safeText(group.mainLocationTower, 'Tower')} is rated ${assessment.status.label.toLowerCase()} with grouped Zt ${
              assessment.totalImpedanceZt === null ? '-' : assessment.totalImpedanceZt
            } ohm against Zsat ${assessment.zsat === null ? '-' : assessment.zsat} ohm.`;
          })
          .join(' ')
      : 'No tower footing groups were available for interpretation.',
    drawChrome
  );
}

function renderConclusion(doc, report, drawChrome) {
  const narrative = getStoredAiNarrative(report);
  const soil = calculateSoilSummary(report);
  const electrodeOverLimit = report.electrodeResistance.filter((row) => {
    return deriveElectrodeAssessment(row).status.tone === 'critical';
  }).length;
  const continuityAttention = report.continuityTest.filter((row) => {
    return deriveContinuityAssessment(row).status.tone !== 'healthy';
  }).length;
  const loopAttention = report.loopImpedanceTest.filter((row) => {
    return deriveLoopImpedanceAssessment(row).status.tone !== 'healthy';
  }).length;
  const faultAttention = report.prospectiveFaultCurrent.filter((row) => {
    return !['healthy', 'neutral'].includes(deriveFaultCurrentAssessment(row).status.tone);
  }).length;
  const riserAttention = report.riserIntegrityTest.filter((row) => {
    return !['healthy', 'neutral'].includes(deriveRiserAssessment(row).status.tone);
  }).length;
  const earthContinuityAttention = report.earthContinuityTest.filter((row) => {
    return !['healthy', 'neutral'].includes(deriveEarthContinuityAssessment(row).status.tone);
  }).length;
  const towerSummaries = summarizeTowerGroups(report.towerFootingResistance);
  const towerGroups = Array.isArray(report.towerFootingResistance) ? report.towerFootingResistance : [];
  const towerFootingAttention = towerGroups.filter((row) => {
    return !['healthy', 'neutral'].includes(deriveTowerFootingAssessment(row, towerSummaries.get(buildTowerGroupKey(row))).status.tone);
  }).length;

  drawSectionTitle(doc, 'Conclusion', 'High-level outcome of the selected measurement sections.', drawChrome);

  if (safeText(narrative?.closingSummary, '')) {
    drawParagraph(doc, safeText(narrative.closingSummary, ''), drawChrome);
    return;
  }

  const parts = [];
  if (report.tests.soilResistivity) {
    parts.push(
      soil.overallAverage === null
        ? 'Soil resistivity data was captured, but there is not enough numeric data to calculate a mean value.'
        : `Mean soil resistivity is ${soil.overallAverage} ohm-m and falls under the ${soil.category.label.toLowerCase()} category.`
    );
  }
  if (report.tests.electrodeResistance) {
    parts.push(
      electrodeOverLimit
        ? `${electrodeOverLimit} electrode reading(s) exceed the 4.60 ohm permissible limit and should be reviewed.`
        : 'All entered electrode resistance readings are within the 4.60 ohm permissible limit.'
    );
  }
  if (report.tests.continuityTest && report.continuityTest.length) {
    parts.push(
      continuityAttention
        ? `${continuityAttention} continuity point(s) need attention based on recorded resistance values.`
        : 'Continuity readings are within the healthy band for the entered rows.'
    );
  }
  if (report.tests.loopImpedanceTest && report.loopImpedanceTest.length) {
    parts.push(
      loopAttention
        ? `${loopAttention} loop impedance point(s) are above the healthy band and should be checked with the protection design.`
        : 'Loop impedance values are within the healthy band for the entered rows.'
    );
  }
  if (report.tests.prospectiveFaultCurrent && report.prospectiveFaultCurrent.length) {
    parts.push(
      faultAttention
        ? `${faultAttention} fault current row(s) need device-capacity review against the entered breaking capacity.`
        : 'Prospective fault current values are within the entered device breaking capacities.'
    );
  }
  if (report.tests.riserIntegrityTest && report.riserIntegrityTest.length) {
    parts.push(
      riserAttention
        ? `${riserAttention} riser integrity row(s) exceed the continuity reference band and should be investigated.`
        : 'Riser integrity readings are within the continuity reference band for the entered rows.'
    );
  }
  if (report.tests.earthContinuityTest && report.earthContinuityTest.length) {
    parts.push(
      earthContinuityAttention
        ? `${earthContinuityAttention} earth continuity point(s) need attention based on the recorded measured values.`
        : 'Earth continuity readings are within the reference band for the entered rows.'
    );
  }
  if (report.tests.towerFootingResistance && report.towerFootingResistance.length) {
    parts.push(
      towerFootingAttention
        ? `${towerFootingAttention} tower footing row(s) need review against the grouped Zsat limit.`
        : 'Tower footing groups are within the entered Zsat limits for the entered rows.'
    );
  }

  drawParagraph(doc, parts.join(' '), drawChrome);
}

function renderDetailAnalysisIntro(doc, drawChrome) {
  drawSectionTitle(doc, 'Detail Analysis', 'Module-wise measurement sheets, graphs, interpretations, and tables.', drawChrome);
  drawParagraph(
    doc,
    'The following chapters present the selected measurement sheets in detail. Each section includes the recorded data, visual interpretation aids, and the automatically derived status based on the entered values and configured references.',
    drawChrome
  );
}

function renderIndexPage(doc, sectionStarts) {
  drawClassicIndexPage(doc, sectionStarts);
}

function startSectionPage(doc, drawChrome, sectionStarts, title, renderer, options = {}) {
  addReportPage(doc, drawChrome, { layout: options.layout || 'portrait' });
  sectionStarts.push({ title, page: doc.bufferedPageRange().start + doc.bufferedPageRange().count });
  renderer();
}

async function generateElectroReportPdf(report, options) {
  const generatedDir = options.generatedDir;
  const publicPathPrefix = options.publicPathPrefix || '/generated-pdfs';

  ensureDir(generatedDir);

  const fileName = `${sanitizeFilePart(report.project.projectNo)}-${Date.now()}.pdf`;
  const filePath = path.join(generatedDir, fileName);
  const doc = new PDFDocument({
    size: 'A4',
    margin: PAGE.margin,
    bufferPages: true
  });
  const stream = fs.createWriteStream(filePath);
  let currentPageNumber = 1;
  const sectionStarts = [];

  doc.pipe(stream);
  if (fs.existsSync(UNICODE_FONT_PATH)) {
    doc.registerFont('Unicode', UNICODE_FONT_PATH);
  }
  doc.on('pageAdded', () => {
    currentPageNumber += 1;
  });

  const drawChrome = (activeDoc) => drawBodyChrome(activeDoc, report);

  const snapshot = buildExecutiveSnapshot(report);
  drawCoverPage(doc, report, snapshot);

  doc.addPage({
    size: 'A4',
    margin: PAGE.margin
  });
  const indexPageBufferIndex = 1;
  startSectionPage(doc, drawChrome, sectionStarts, 'Survey Overview', () => {
    renderSurveyOverviewSection(doc, report, drawChrome);
  });

  startSectionPage(doc, drawChrome, sectionStarts, 'Executive Summary', () => {
    renderExecutiveSummarySection(doc, report, drawChrome);
  });

  startSectionPage(doc, drawChrome, sectionStarts, 'Healthiness Classifications', () => {
    renderHealthClassificationSection(doc, report, drawChrome);
  });

  startSectionPage(doc, drawChrome, sectionStarts, 'On-Site Testing Methodologies', () => {
    renderMethodologySection(doc, report, drawChrome);
    renderSelectedMethodologies(doc, report, drawChrome);
  });

  if (report.tests.soilResistivity) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Soil Resistivity Test - Theory & Methodology', () => {
      renderSoilTheorySection(doc, report, drawChrome);
    });
    startSectionPage(doc, drawChrome, sectionStarts, 'Soil Resistivity Test - Analysis (Graphs)', () => {
      renderSoilGraphSection(doc, report, drawChrome);
    });
    startSectionPage(doc, drawChrome, sectionStarts, 'Soil Resistivity Test - Analysis (Results)', () => {
      renderSoilResultSection(doc, report, drawChrome);
    });
    startSectionPage(doc, drawChrome, sectionStarts, 'Calculation of Earth Electrode & Grid Resistance (Standard Limit)', () => {
      renderSoilStandardLimitSection(doc, report, drawChrome);
    });
  }
  if (report.tests.electrodeResistance) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Earth Electrode Resistance Test - Measurement & Analysis', () => {
      renderElectrodeSection(doc, report, drawChrome);
    });
  }
  if (report.tests.continuityTest) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Earthing Continuity Test - Measurement & Analysis', () => {
      renderContinuitySection(doc, report, drawChrome);
    });
  }
  if (report.tests.loopImpedanceTest) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Earth Loop Impedance Test - Measurement & Analysis', () => {
      renderLoopSection(doc, report, drawChrome);
    });
  }
  if (report.tests.prospectiveFaultCurrent) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Prospective Fault Current Test - Measurement & Analysis', () => {
      renderFaultCurrentSection(doc, report, drawChrome);
    });
  }
  if (report.tests.riserIntegrityTest) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Riser / Grid Integrity Test - Measurement & Analysis', () => {
      renderRiserSection(doc, report, drawChrome);
    });
  }
  if (report.tests.earthContinuityTest) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Earth Continuity Test - Measurement & Analysis', () => {
      renderEarthContinuitySection(doc, report, drawChrome);
    });
  }
  if (report.tests.towerFootingResistance) {
    startSectionPage(doc, drawChrome, sectionStarts, 'Tower Footing Resistance Measurement & Analysis', () => {
      renderTowerFootingSection(doc, report, drawChrome);
    });
  }

  startSectionPage(doc, drawChrome, sectionStarts, 'Observation Sheet', () => {
    renderObservationSheetSection(doc, report, drawChrome);
  });

  startSectionPage(doc, drawChrome, sectionStarts, 'Conclusion', () => {
    renderConclusion(doc, report, drawChrome);
  });

  const range = doc.bufferedPageRange();
  doc.switchToPage(indexPageBufferIndex);
  drawChrome(doc);
  doc.y = 86;
  renderIndexPage(doc, sectionStarts);

  for (let i = 0; i < range.count; i += 1) {
    doc.switchToPage(i);
    if (i === 0) {
      continue;
    }
    doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(`Page ${i + 1} of ${range.count}`, PAGE.margin, doc.page.height - PAGE.margin - 12, {
        width: doc.page.width - PAGE.margin * 2,
        align: 'center'
      });
  }

  doc.end();

  await new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });

  return {
    fileName,
    filePath,
    pdfUrl: `${publicPathPrefix}/${fileName}`
  };
}

module.exports = {
  generateElectroReportPdf,
  getReportTitle
};
