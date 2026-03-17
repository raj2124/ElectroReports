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

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function safeText(value, fallback = '-') {
  const text = String(value === undefined || value === null ? '' : value).trim();
  return text || fallback;
}

function sanitizeFilePart(value) {
  return safeText(value, 'report').replace(/[^a-z0-9_-]+/gi, '-');
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
  if (
    selected.every((test) =>
      ['soilResistivity', 'electrodeResistance', 'continuityTest', 'loopImpedanceTest', 'prospectiveFaultCurrent', 'riserIntegrityTest', 'earthContinuityTest', 'towerFootingResistance'].includes(
        test.id
      )
    )
  ) {
    return 'Earthing System Health Assessment';
  }
  return 'Electrical Measurement Assessment';
}

function ensureSpace(doc, requiredHeight, drawChrome) {
  const bottom = doc.page.height - PAGE.margin - 34;
  if (doc.y + requiredHeight <= bottom) {
    return;
  }
  doc.addPage();
  if (drawChrome) {
    drawChrome(doc);
  }
}

function drawBodyChrome(doc, report) {
  const pageWidth = doc.page.width;
  const pageHeight = doc.page.height;
  const topY = 20;

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
    .moveTo(PAGE.margin, pageHeight - 30)
    .lineTo(pageWidth - PAGE.margin, pageHeight - 30)
    .lineWidth(0.8)
    .strokeColor(COLORS.border)
    .stroke();

  doc
    .font('Helvetica')
    .fontSize(8)
    .fillColor(COLORS.muted)
    .text('ElectroReports | Elegrow Technology', PAGE.margin, pageHeight - 22, {
      width: 220
    })
    .text(`Date of Testing: ${safeText(report.project.reportDate)}`, pageWidth - PAGE.margin - 180, pageHeight - 22, {
      width: 180,
      align: 'right'
    });

  doc.y = 86;
}

function drawCoverPage(doc, report, snapshot) {
  const title = getReportTitle(report);

  if (fs.existsSync(SYMBOL_PATH)) {
    doc.save();
    doc.opacity(0.08);
    doc.image(SYMBOL_PATH, PAGE.margin + 90, 170, { fit: [280, 280], align: 'center' });
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
  ensureSpace(doc, 52, drawChrome);
  doc.roundedRect(PAGE.margin, doc.y, doc.page.width - PAGE.margin * 2, 40, 12).fillAndStroke(COLORS.brandSoft, COLORS.border);
  doc
    .font('Helvetica-Bold')
    .fontSize(14)
    .fillColor(COLORS.brandDark)
    .text(title, PAGE.margin + 16, doc.y + 10);
  if (subtitle) {
    doc
      .font('Helvetica')
      .fontSize(8.5)
      .fillColor(COLORS.muted)
      .text(subtitle, PAGE.margin + 16, doc.y + 26);
  }
  doc.moveDown(2.4);
}

function drawKeyValueGrid(doc, rows, drawChrome) {
  const gridWidth = doc.page.width - PAGE.margin * 2;
  const colWidth = (gridWidth - 14) / 2;
  const rowHeight = 48;

  rows.forEach((row, index) => {
    if (index % 2 === 0) {
      ensureSpace(doc, rowHeight + 10, drawChrome);
    }
    const isLeft = index % 2 === 0;
    const x = PAGE.margin + (isLeft ? 0 : colWidth + 14);
    const y = doc.y + (isLeft ? 0 : -rowHeight);
    doc.roundedRect(x, y, colWidth, rowHeight, 10).fillAndStroke(COLORS.white, COLORS.border);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(row.label.toUpperCase(), x + 12, y + 10, { width: colWidth - 24 });
    doc
      .font('Helvetica-Bold')
      .fontSize(11)
      .fillColor(COLORS.ink)
      .text(safeText(row.value), x + 12, y + 24, { width: colWidth - 24 });
    if (!isLeft) {
      doc.moveDown(1.8);
    }
  });

  if (rows.length % 2 === 1) {
    doc.moveDown(1.8);
  }
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
  doc.moveDown(1);
}

function drawSubsectionParagraph(doc, title, text, drawChrome) {
  ensureSpace(doc, 40, drawChrome);
  doc
    .font('Helvetica-Bold')
    .fontSize(9)
    .fillColor(COLORS.brandDark)
    .text(title.toUpperCase(), PAGE.margin, doc.y, {
      width: doc.page.width - PAGE.margin * 2
    });
  doc.moveDown(0.2);
  drawParagraph(doc, text, drawChrome);
}

function drawBarChart(doc, title, entries, drawChrome) {
  const rows = (Array.isArray(entries) ? entries : []).filter((entry) => Number.isFinite(entry?.value));
  if (!rows.length) {
    return;
  }

  const chartHeight = rows.length * 28 + 54;
  ensureSpace(doc, chartHeight, drawChrome);

  const x = PAGE.margin;
  const width = doc.page.width - PAGE.margin * 2;
  const y = doc.y;
  const labelWidth = Math.min(180, width * 0.38);
  const barAreaWidth = width - labelWidth - 24;
  const maxValue = Math.max(...rows.map((entry) => entry.value), 1);

  doc.roundedRect(x, y, width, chartHeight, 12).fillAndStroke(COLORS.white, COLORS.border);
  doc
    .font('Helvetica-Bold')
    .fontSize(9)
    .fillColor(COLORS.brandDark)
    .text(title, x + 12, y + 10, { width: width - 24 });

  rows.forEach((entry, index) => {
    const rowY = y + 32 + index * 28;
    const barX = x + 12 + labelWidth;
    const barY = rowY + 6;
    const barWidth = Math.max(8, (barAreaWidth * entry.value) / maxValue);
    const fill =
      entry.tone === 'critical' ? COLORS.critical : entry.tone === 'warning' ? COLORS.warning : entry.tone === 'healthy' ? COLORS.healthy : COLORS.brand;

    doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor(COLORS.ink)
      .text(entry.label, x + 12, rowY, { width: labelWidth - 8 });
    doc.roundedRect(barX, barY, barAreaWidth, 12, 6).fill(COLORS.neutralSoft);
    doc.roundedRect(barX, barY, barWidth, 12, 6).fill(fill);
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.ink)
      .text(String(entry.value), barX + barAreaWidth + 6, rowY, { width: 40 });
  });

  doc.y = y + chartHeight + 8;
}

function drawTable(doc, columns, rows, drawChrome) {
  const totalWidth = doc.page.width - PAGE.margin * 2;
  const x = PAGE.margin;
  const headerY = doc.y;

  ensureSpace(doc, 28, drawChrome);

  doc.roundedRect(x, headerY, totalWidth, 24, 8).fill(COLORS.brand);
  let cursorX = x;

  columns.forEach((column) => {
    const colWidth = totalWidth * column.width;
    doc
      .font('Helvetica-Bold')
      .fontSize(8)
      .fillColor(COLORS.white)
      .text(column.label, cursorX + 6, headerY + 8, { width: colWidth - 12 });
    cursorX += colWidth;
  });

  doc.y = headerY + 28;

  rows.forEach((row, rowIndex) => {
    const observationBlocks = Array.isArray(row.__observationBlocks) ? row.__observationBlocks.filter(Boolean) : [];
    const cellHeights = columns.map((column) => {
      return doc.heightOfString(safeText(row[column.key]), {
        width: totalWidth * column.width - 12,
        align: column.align || 'left'
      });
    });
    const rowHeight = Math.max(24, ...cellHeights.map((height) => height + 10));
    const observationHeight = estimateObservationBlocksHeight(doc, observationBlocks, totalWidth - 20);
    ensureSpace(doc, rowHeight + 6 + observationHeight, drawChrome);

    const rowY = doc.y;
    doc
      .roundedRect(x, rowY, totalWidth, rowHeight, 6)
      .fillAndStroke(rowIndex % 2 === 0 ? COLORS.white : COLORS.rowAlt, COLORS.border);

    let rowX = x;
    columns.forEach((column) => {
      const colWidth = totalWidth * column.width;
      doc
        .font('Helvetica')
        .fontSize(8.5)
        .fillColor(COLORS.ink)
        .text(safeText(row[column.key]), rowX + 6, rowY + 7, {
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

function renderSoilSection(doc, report, drawChrome) {
  const soil = report.soilResistivity;
  const summary = calculateSoilSummary(report);

  drawSectionTitle(
    doc,
    'Soil Resistivity Test',
    'Automatic direction-wise averaging with classification as per IS 3043 and IEEE 81.',
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

  drawTable(
    doc,
    [
      { key: 'spacing', label: 'Spacing of Probes (m)', width: 0.32 },
      { key: 'resistivity', label: 'Direction 1 Resistivity (ohm-m)', width: 0.34 },
      { key: 'direction2Resistivity', label: 'Direction 2 Resistivity (ohm-m)', width: 0.34 }
    ],
    soil.direction1.map((row, index) => ({
      spacing: row.spacing,
      resistivity: row.resistivity,
      direction2Resistivity: soil.direction2[index] ? soil.direction2[index].resistivity : '-',
      __observationBlocks: [
        buildObservationBlock(row, 'Direction 1 Observation'),
        buildObservationBlock(soil.direction2[index], 'Direction 2 Observation')
      ].filter(Boolean)
    })),
    drawChrome
  );

  if (soil.notes) {
    drawParagraph(doc, `Site note: ${soil.notes}`, drawChrome);
  }
}

function renderElectrodeSection(doc, report, drawChrome) {
  drawSectionTitle(
    doc,
    'Earth Electrode Resistance Test',
    'Measured resistance checked against the 4.60 ohm permissible limit.',
    drawChrome
  );

  drawTable(
    doc,
    [
      { key: 'tag', label: 'Pit Tag', width: 0.08 },
      { key: 'location', label: 'Location', width: 0.1 },
      { key: 'electrodeType', label: 'Type', width: 0.08 },
      { key: 'materialType', label: 'Material', width: 0.08 },
      { key: 'length', label: 'Length', width: 0.06 },
      { key: 'diameter', label: 'Dia', width: 0.06 },
      { key: 'resistanceWithoutGrid', label: 'Without Grid', width: 0.09 },
      { key: 'resistanceWithGrid', label: 'With Grid', width: 0.09 },
      { key: 'standard', label: 'Standard', width: 0.1 },
      { key: 'status', label: 'Status', width: 0.12 },
      { key: 'comment', label: 'Comment', width: 0.14 }
    ],
    report.electrodeResistance.map((row) => ({
      ...row,
      resistanceWithoutGrid: safeText(row.resistanceWithoutGrid, ''),
      resistanceWithGrid: safeText(row.resistanceWithGrid || row.measuredResistance, ''),
      standard: 'IS 3043 / 4.60 ohm',
      status: deriveElectrodeAssessment(row).status.label,
      comment: safeText(row.observation, deriveElectrodeAssessment(row).comment),
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderContinuitySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Continuity Test', 'Point-to-point continuity readings.', drawChrome);
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
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderLoopSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Loop Impedance Test', 'Measured Zs values across selected panels or equipment.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.1 },
      { key: 'mainLocation', label: 'Main Location', width: 0.18 },
      { key: 'panelEquipment', label: 'Panel / Equipment', width: 0.26 },
      { key: 'measuredZs', label: 'Measured Zs (ohm)', width: 0.14 },
      { key: 'status', label: 'Status', width: 0.14 },
      { key: 'remarks', label: 'Remarks', width: 0.18 }
    ],
    report.loopImpedanceTest.map((row) => ({
      ...row,
      status: deriveLoopImpedanceAssessment(row).status.label,
      remarks: safeText(row.remarks, deriveLoopImpedanceAssessment(row).comment),
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderFaultCurrentSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Prospective Fault Current', 'Feeder details with loop impedance and fault current values.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.06 },
      { key: 'location', label: 'Location', width: 0.1 },
      { key: 'feederTag', label: 'Feeder & Tag', width: 0.12 },
      { key: 'deviceType', label: 'Device', width: 0.08 },
      { key: 'deviceRating', label: 'Rating (A)', width: 0.08 },
      { key: 'breakingCapacity', label: 'Breaking (kA)', width: 0.09 },
      { key: 'measuredPoints', label: 'Measured Points', width: 0.1 },
      { key: 'loopImpedance', label: 'Loop Z (ohm)', width: 0.08 },
      { key: 'prospectiveFaultCurrent', label: 'PFC', width: 0.08 },
      { key: 'status', label: 'Status', width: 0.1 },
      { key: 'comment', label: 'Comment', width: 0.11 }
    ],
    report.prospectiveFaultCurrent.map((row) => ({
      ...row,
      status: deriveFaultCurrentAssessment(row).status.label,
      comment: safeText(row.comment, deriveFaultCurrentAssessment(row).comment),
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderRiserSection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Riser / Grid Integrity Test', 'Resistance verification towards equipment and grid.', drawChrome);
  drawTable(
    doc,
    [
      { key: 'srNo', label: 'Sr. No.', width: 0.08 },
      { key: 'mainLocation', label: 'Main Location', width: 0.16 },
      { key: 'measurementPoint', label: 'Measurement Point', width: 0.2 },
      { key: 'resistanceTowardsEquipment', label: 'Towards Equipment', width: 0.12 },
      { key: 'resistanceTowardsGrid', label: 'Towards Grid', width: 0.12 },
      { key: 'status', label: 'Status', width: 0.12 },
      { key: 'comment', label: 'Comment', width: 0.2 }
    ],
    report.riserIntegrityTest.map((row) => ({
      ...row,
      status: deriveRiserAssessment(row).status.label,
      comment: safeText(row.comment, deriveRiserAssessment(row).comment),
      __observationBlocks: [buildObservationBlock(row)].filter(Boolean)
    })),
    drawChrome
  );
}

function renderEarthContinuitySection(doc, report, drawChrome) {
  drawSectionTitle(doc, 'Earth Continuity Test', 'Earth path verification by location and measured value.', drawChrome);
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

  drawSubsectionParagraph(
    doc,
    'Objective',
    'Assess each tower location by summing the 4 footing impedance readings to obtain Zt and summing the 4 footing current readings to obtain Itotal in amperes.',
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
    'For each tower location, the total impedance Zt is obtained by summing the measured impedance values of Foot-1, Foot-2, Foot-3, and Foot-4. The total current Itotal is obtained by summing all measured current values in mA and converting the total to amperes by dividing by 1000.',
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
      srNo: readingIndex === 0 ? safeText(group.srNo, '-') : '',
      mainLocationTower: readingIndex === 0 ? safeText(group.mainLocationTower, '-') : '',
      measurementPointLocation: safeText(reading.measurementPointLocation, `Foot-${readingIndex + 1}`),
      footToEarthingConnectionStatus: safeText(reading.footToEarthingConnectionStatus, 'Given'),
      measuredCurrentMa: safeText(reading.measuredCurrentMa, ''),
      measuredImpedance: safeText(reading.measuredImpedance, ''),
      totalImpedanceZt: readingIndex === 0 ? (assessment.totalImpedanceZt === null ? '-' : String(assessment.totalImpedanceZt)) : '',
      totalCurrentItotal: readingIndex === 0 ? (assessment.totalCurrentItotal === null ? '-' : String(assessment.totalCurrentItotal)) : '',
      standardTolerableImpedanceZsat: readingIndex === 0 ? String(assessment.zsat) : '',
      remarks: readingIndex === 0 ? assessment.comment : '',
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
      { key: 'totalCurrentItotal', label: 'Total Current Itotal (A)', width: 0.09 },
      { key: 'standardTolerableImpedanceZsat', label: 'Zsat', width: 0.08 },
      { key: 'remarks', label: 'Remarks', width: 0.14 }
    ],
    tableRows,
    drawChrome
  );

  drawBarChart(
    doc,
    'Tower Footing Graph: Zt vs Zsat',
    towerGroups.flatMap((group) => {
      const assessment = deriveTowerFootingAssessment(group, towerSummaries.get(buildTowerGroupKey(group)));
      if (!Number.isFinite(assessment.totalImpedanceZt)) {
        return [];
      }
      return [
        {
          label: `${safeText(group.mainLocationTower, 'Tower')} Zt`,
          value: assessment.totalImpedanceZt,
          tone: assessment.status.tone
        },
        {
          label: `${safeText(group.mainLocationTower, 'Tower')} Zsat`,
          value: assessment.zsat,
          tone: 'neutral'
        }
      ];
    }),
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

  doc.pipe(stream);

  const drawChrome = (activeDoc) => drawBodyChrome(activeDoc, report);

  const snapshot = buildExecutiveSnapshot(report);
  drawCoverPage(doc, report, snapshot);

  doc.addPage();
  drawChrome(doc);
  renderOverviewSection(doc, report, drawChrome);

  if (report.tests.soilResistivity) {
    renderSoilSection(doc, report, drawChrome);
  }
  if (report.tests.electrodeResistance) {
    renderElectrodeSection(doc, report, drawChrome);
  }
  if (report.tests.continuityTest) {
    renderContinuitySection(doc, report, drawChrome);
  }
  if (report.tests.loopImpedanceTest) {
    renderLoopSection(doc, report, drawChrome);
  }
  if (report.tests.prospectiveFaultCurrent) {
    renderFaultCurrentSection(doc, report, drawChrome);
  }
  if (report.tests.riserIntegrityTest) {
    renderRiserSection(doc, report, drawChrome);
  }
  if (report.tests.earthContinuityTest) {
    renderEarthContinuitySection(doc, report, drawChrome);
  }
  if (report.tests.towerFootingResistance) {
    renderTowerFootingSection(doc, report, drawChrome);
  }

  renderConclusion(doc, report, drawChrome);

  const range = doc.bufferedPageRange();
  for (let i = 0; i < range.count; i += 1) {
    doc.switchToPage(i);
    if (i === 0) {
      continue;
    }
    doc
      .font('Helvetica')
      .fontSize(8)
      .fillColor(COLORS.muted)
      .text(`Page ${i + 1} of ${range.count}`, PAGE.margin, doc.page.height - 22, {
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
