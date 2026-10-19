// Small pdfkit layout toolkit for the ESHA reports: ETPL brand, header/footer chrome,
// headings, paragraphs, auto-paging tables and the electrode chart. The look follows
// reference/samples (navy #0B2A67, Arial, CONFIDENTIAL footer). Arial is used when installed,
// otherwise the bundled Liberation Sans.
const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');
const { resolveFonts } = require('../sheetPrint/fonts');

const COLORS = {
  navy: '#0B2A67',
  grey: '#9AA3B2',
  light: '#EEF1F6',
  green: '#1E7B34',
  amber: '#B8860B',
  red: '#B42318',
  neutral: '#667085',
  note: '#475467',
  ink: '#101828'
};
const PAGE = { width: 595.28, height: 841.89, margin: 42, top: 55, bottom: 50 };
PAGE.contentWidth = PAGE.width - 2 * PAGE.margin;
const LOGO_PATH = path.join(__dirname, '..', '..', 'public', 'assets', 'elegrow-symbol.png');

// Symbols the PDF fonts carry; model text writes units in words.
function pretty(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/\bohm-m\b/g, 'Ω·m')
    .replace(/\bohms?\b/g, 'Ω')
    .replace(/\brho\b/g, 'ρ');
}

function statusColor(label) {
  const value = String(label || '');
  if (['Healthy', 'P3', 'Within limit', 'Agree'].includes(value)) return COLORS.green;
  if (['Warning', 'Marginal', 'Healthy - Minor Actions', 'Healthy – Minor Actions', 'P2', 'Data gap', 'Review'].includes(value)) return COLORS.amber;
  if (['Not Assessable', 'Pending', 'Incomplete', 'Not set'].includes(value)) return COLORS.neutral;
  return COLORS.red;
}

class ReportDoc {
  constructor({ title, unit, documentNo, revision, classification = 'CONFIDENTIAL', watermark = '', info = {} }) {
    this.meta = { title, unit, documentNo, revision, classification, watermark };
    this.fonts = resolveFonts();
    this.doc = new PDFDocument({
      size: 'A4',
      bufferPages: true,
      autoFirstPage: false,
      margins: { top: PAGE.top, bottom: PAGE.bottom, left: PAGE.margin, right: PAGE.margin },
      info: { Title: title, Author: 'Elegrow Technology Pvt. Ltd.', ...info }
    });
    this.doc.registerFont('Body', this.fonts.regular);
    this.doc.registerFont('Bold', this.fonts.bold);
    this.sections = [];
    this.pageMetas = [];
    this.y = PAGE.top;
    this.coverPages = 0;
    this.hasLogo = fs.existsSync(LOGO_PATH);
    this.doc.on('pageAdded', () => {
      this.y = PAGE.top;
    });
  }

  newPage() {
    this.doc.addPage();
    this.y = PAGE.top;
    this.pageMetas[this.doc.bufferedPageRange().count - 1] = { ...this.meta };
  }

  get pageIndex() {
    return this.doc.bufferedPageRange().count - 1;
  }

  ensure(height) {
    if (this.y + height > PAGE.height - PAGE.bottom) {
      this.newPage();
      return true;
    }
    return false;
  }

  space(points) {
    this.y += points;
  }

  heightOf(value, { font = 'Body', size = 9.5, width = PAGE.contentWidth, lineGap = 2 } = {}) {
    this.doc.font(font).fontSize(size);
    return this.doc.heightOfString(pretty(value), { width, lineGap });
  }

  // A section heading. Level 1 starts on a new page and is recorded for the index.
  section(title, { newPage = true } = {}) {
    if (newPage) {
      this.newPage();
    } else {
      this.ensure(60);
    }
    const numbered = `${this.sections.filter((entry) => !entry.sub).length + 1}. ${title}`;
    this.sections.push({ title: numbered, page: this.pageIndex });
    this.heading(numbered, 1);
  }

  // `indexed` adds a level-2 heading to the document index, under the section above it.
  heading(value, level = 1, { indexed = false } = {}) {
    const size = level === 1 ? 14 : 11;
    this.ensure(size + 40);
    if (indexed) {
      this.sections.push({ title: pretty(value), page: this.pageIndex, sub: true });
    }
    this.y += level === 1 ? 2 : 6;
    this.doc.font('Bold').fontSize(size).fillColor(COLORS.navy).text(pretty(value), PAGE.margin, this.y, { width: PAGE.contentWidth });
    this.y = this.doc.y + (level === 1 ? 8 : 5);
  }

  paragraph(value, { size = 9.5, color = COLORS.ink, align = 'justify', font = 'Body', gap = 6, indent = 0 } = {}) {
    const height = this.heightOf(value, { size, width: PAGE.contentWidth - indent, font });
    this.ensure(Math.min(height, 120) + 4);
    this.doc.font(font).fontSize(size).fillColor(color).text(pretty(value), PAGE.margin + indent, this.y, { width: PAGE.contentWidth - indent, align, lineGap: 2 });
    this.y = this.doc.y + gap;
  }

  bullets(items, options = {}) {
    items.forEach((item) => this.paragraph(`•  ${item}`, { align: 'left', ...options, gap: 3, indent: 0 }));
    this.y += 3;
  }

  note(value) {
    this.paragraph(value, { size: 8, color: COLORS.note, align: 'left', gap: 5 });
  }

  // A table that pages. columns: [[heading, weight], ...]; cells are strings or { text, color, bold };
  // statusColumns get the status colour; alignRight lists right-aligned (numeric) columns.
  table({ columns, rows, statusColumns = [], alignRight = [], fontSize = 7.8, boldFirstColumn = false }) {
    const total = columns.reduce((sum, column) => sum + column[1], 0);
    const widths = columns.map((column) => (column[1] / total) * PAGE.contentWidth);
    const pad = 4;
    const cellText = (cell) => pretty(cell && typeof cell === 'object' ? cell.text : cell);
    const rowHeight = (cells, font) => {
      let max = 0;
      cells.forEach((cell, i) => {
        this.doc.font(font).fontSize(fontSize);
        max = Math.max(max, this.doc.heightOfString(cellText(cell), { width: widths[i] - 2 * pad, lineGap: 1 }));
      });
      return max + 2 * pad - 1;
    };
    const drawRow = (cells, { header = false } = {}) => {
      const height = rowHeight(cells, header || boldFirstColumn ? 'Bold' : 'Body');
      let x = PAGE.margin;
      cells.forEach((cell, i) => {
        const object = cell && typeof cell === 'object' ? cell : { text: cell };
        if (header) {
          this.doc.rect(x, this.y, widths[i], height).fillColor(COLORS.light).fill();
        }
        this.doc.rect(x, this.y, widths[i], height).lineWidth(0.4).strokeColor(COLORS.grey).stroke();
        const colored = !header && (object.color || statusColumns.includes(i));
        const font = header || object.bold || statusColumns.includes(i) || (boldFirstColumn && i === 0) ? 'Bold' : 'Body';
        this.doc
          .font(font)
          .fontSize(fontSize)
          .fillColor(header ? COLORS.navy : colored ? object.color || statusColor(object.text) : COLORS.ink)
          .text(cellText(object), x + pad, this.y + pad, { width: widths[i] - 2 * pad, lineGap: 1, align: alignRight.includes(i) && !header ? 'right' : 'left' });
        x += widths[i];
      });
      this.y += height;
    };
    const headerCells = columns.map((column) => column[0]);
    const headerHeight = rowHeight(headerCells, 'Bold');
    this.ensure(headerHeight + rowHeight(rows[0] || [''], 'Body'));
    drawRow(headerCells, { header: true });
    rows.forEach((cells) => {
      const height = rowHeight(cells, 'Body');
      if (this.y + height > PAGE.height - PAGE.bottom) {
        this.newPage();
        drawRow(headerCells, { header: true });
      }
      drawRow(cells);
    });
    this.y += 6;
  }

  // Measured vs calculated limit, one pair of bars per pit. Pits without a reading are labelled.
  electrodeChart(points) {
    const height = 170;
    this.ensure(height + 20);
    const x0 = PAGE.margin + 34;
    const width = PAGE.contentWidth - 44;
    const top = this.y + 16;
    const plotH = 112;
    const values = points.flatMap((point) => [point.measured, point.limit]).filter((value) => Number.isFinite(value));
    const max = Math.max(1, ...values) * 1.1;
    const doc = this.doc;
    doc.font('Bold').fontSize(8).fillColor(COLORS.navy).text('Measured R (navy) vs calculated limit (grey), Ω', PAGE.margin, this.y, { width: PAGE.contentWidth });
    const step = niceStep(max);
    for (let tick = 0; tick <= max; tick += step) {
      const y = top + plotH - (tick / max) * plotH;
      doc.moveTo(x0, y).lineTo(x0 + width, y).lineWidth(0.3).strokeColor('#D0D5DD').stroke();
      doc.font('Body').fontSize(6.5).fillColor(COLORS.note).text(String(tick), PAGE.margin, y - 3, { width: 28, align: 'right' });
    }
    const group = width / Math.max(points.length, 1);
    points.forEach((point, i) => {
      const gx = x0 + i * group + group * 0.12;
      const bar = group * 0.34;
      const draw = (value, offset, color) => {
        if (!Number.isFinite(value)) return;
        const h = (value / max) * plotH;
        doc.rect(gx + offset, top + plotH - h, bar, h).fillColor(color).fill();
      };
      draw(point.measured, 0, COLORS.navy);
      draw(point.limit, bar + 1, '#C7CDD8');
      const label = point.label.replace(/^.*-(?=EP)/, '').slice(-8);
      doc.font('Body').fontSize(6.5).fillColor(COLORS.ink).text(label, x0 + i * group, top + plotH + 4, { width: group, align: 'center' });
      if (!Number.isFinite(point.measured)) {
        doc.font('Bold').fontSize(6).fillColor(COLORS.red).text('not measured', x0 + i * group, top + plotH + 13, { width: group, align: 'center' });
      }
    });
    this.y = top + plotH + 26;
  }

  // Line chart with markers: series = [{ label, points: [[x, y], ...] }]. Used for rho against spacing.
  lineChart({ title, series, xLabel, yLabel }) {
    const height = 190;
    this.ensure(height + 10);
    const doc = this.doc;
    const x0 = PAGE.margin + 40;
    const width = PAGE.contentWidth - 60;
    const top = this.y + 16;
    const plotH = 122;
    const xs = series.flatMap((entry) => entry.points.map((point) => point[0]));
    const ys = series.flatMap((entry) => entry.points.map((point) => point[1]));
    const xMin = Math.min(...xs);
    const xMax = Math.max(...xs);
    const yMax = Math.max(1, ...ys) * 1.1;
    const xAt = (x) => x0 + (xMax === xMin ? width / 2 : ((x - xMin) / (xMax - xMin)) * width);
    const yAt = (y) => top + plotH - (y / yMax) * plotH;
    const colors = [COLORS.navy, '#D97706'];
    doc.font('Bold').fontSize(8).fillColor(COLORS.navy).text(pretty(title), PAGE.margin, this.y, { width: PAGE.contentWidth });
    const step = niceStep(yMax);
    for (let tick = 0; tick <= yMax; tick += step) {
      doc.moveTo(x0, yAt(tick)).lineTo(x0 + width, yAt(tick)).lineWidth(0.3).strokeColor('#D0D5DD').stroke();
      doc.font('Body').fontSize(6.5).fillColor(COLORS.note).text(String(Math.round(tick * 100) / 100), PAGE.margin, yAt(tick) - 3, { width: 34, align: 'right' });
    }
    doc.moveTo(x0, top + plotH).lineTo(x0 + width, top + plotH).lineWidth(0.6).strokeColor(COLORS.grey).stroke();
    const uniqueX = [...new Set(xs)].sort((a, b) => a - b);
    uniqueX.forEach((x) => doc.font('Body').fontSize(6.5).fillColor(COLORS.ink).text(String(x), xAt(x) - 12, top + plotH + 3, { width: 24, align: 'center' }));
    series.forEach((entry, index) => {
      const color = colors[index % colors.length];
      entry.points.forEach(([x, y], i) => {
        if (i === 0) doc.moveTo(xAt(x), yAt(y)); else doc.lineTo(xAt(x), yAt(y));
      });
      doc.lineWidth(1.1).strokeColor(color).stroke();
      entry.points.forEach(([x, y]) => doc.circle(xAt(x), yAt(y), 1.9).fillColor(color).fill());
      const lx = x0 + 6 + index * 110;
      doc.rect(lx, top + plotH + 17, 8, 3).fillColor(color).fill();
      doc.font('Body').fontSize(7).fillColor(COLORS.ink).text(entry.label, lx + 12, top + plotH + 14, { lineBreak: false });
    });
    doc.font('Body').fontSize(6.5).fillColor(COLORS.note).text(pretty(`${xLabel}  ·  ${yLabel}`), x0 + width - 160, top + plotH + 14, { width: 160, align: 'right', lineBreak: false });
    this.y = top + plotH + 32;
  }

  // Header and footer on every page except the cover, and "Page x of y".
  finish({ coverPages = 1 } = {}) {
    const range = this.doc.bufferedPageRange();
    for (let i = 0; i < range.count; i += 1) {
      this.doc.switchToPage(i);
      const page = this.doc.page;
      const margins = { ...page.margins };
      page.margins.bottom = 0;
      page.margins.top = 0;
      const { documentNo, revision, classification, title, unit, watermark } = this.pageMetas[i] || this.meta;
      if (i >= coverPages) {
        if (this.hasLogo) {
          this.doc.image(LOGO_PATH, PAGE.margin, 20, { width: 16, height: 16 });
        }
        this.doc.font('Bold').fontSize(8.5).fillColor(COLORS.navy).text('Elegrow Technology Pvt. Ltd.', PAGE.margin + 21, 24, { lineBreak: false });
        this.doc.font('Body').fontSize(8).fillColor(COLORS.ink).text(`${documentNo} | ${revision}`, PAGE.margin, 24, { width: PAGE.contentWidth, align: 'right', lineBreak: false });
        if (watermark) {
          this.doc.font('Bold').fontSize(8).fillColor('#C7CDD8').text(watermark, PAGE.margin, 24, { width: PAGE.contentWidth, align: 'center', lineBreak: false });
        }
        this.doc.moveTo(PAGE.margin, 40).lineTo(PAGE.width - PAGE.margin, 40).lineWidth(1).strokeColor(COLORS.navy).stroke();
        this.doc.moveTo(PAGE.margin, PAGE.height - 38).lineTo(PAGE.width - PAGE.margin, PAGE.height - 38).lineWidth(0.5).strokeColor(COLORS.grey).stroke();
        this.doc.font('Bold').fontSize(7.5).fillColor(COLORS.red).text(classification, PAGE.margin, PAGE.height - 30, { lineBreak: false });
        this.doc
          .font('Body')
          .fontSize(7.5)
          .fillColor(COLORS.navy)
          .text(`${title}${unit ? ` – ${unit}` : ''} | Page ${i + 1} of ${range.count}`, PAGE.margin, PAGE.height - 30, { width: PAGE.contentWidth, align: 'right', lineBreak: false });
      } else {
        this.doc.font('Bold').fontSize(7.5).fillColor(COLORS.red).text(classification, PAGE.margin, PAGE.height - 46, { lineBreak: false });
      }
      Object.assign(page.margins, margins);
    }
  }

  async toBuffer() {
    const chunks = [];
    this.doc.on('data', (chunk) => chunks.push(chunk));
    const done = new Promise((resolve, reject) => {
      this.doc.on('end', () => resolve(Buffer.concat(chunks)));
      this.doc.on('error', reject);
    });
    this.doc.end();
    return done;
  }
}

function niceStep(max) {
  const raw = max / 5;
  const magnitude = 10 ** Math.floor(Math.log10(raw));
  const normalized = raw / magnitude;
  return (normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10) * magnitude;
}

function formatDate(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value || ''));
  return match ? `${match[3]}-${match[2]}-${match[1]}` : String(value || '');
}

module.exports = { ReportDoc, PAGE, COLORS, statusColor, pretty, formatDate, LOGO_PATH };
