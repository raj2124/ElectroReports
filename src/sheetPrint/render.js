// Renders the printed measurement sheets (MS-01..MS-09, ET-D03-FRM-001..009 Rev 00).
// A port of reference/generators/make_sheets_v3.py onto the pdfkit canvas, with the same
// coordinates, row-count rule and column proportions. Every page carries its sheet ID and a
// QR code (`ER1|<id>`); a page can also be rendered "filled" (values printed into the cells)
// for tests.
const path = require('path');
const QRCode = require('qrcode');
const { Canvas } = require('./canvas');
const { FIGURES, NAVY, GREY } = require('./diagrams');
const { SPECS, CODE_LEGEND, GROUND_CONDITIONS, soilSpacings } = require('./spec');

const TINT = '#EEF1F6';
const BOX = '#5F6878';
const BLACK = '#000000';
const WHITE = '#FFFFFF';
const M = 26;
const BAND = 74;
const BAND_LANDSCAPE = 61;
const A4 = { w: 595.2755905511812, h: 841.8897637795277 };
const LOGO = path.join(__dirname, '..', '..', 'public', 'assets', 'elegrow-symbol.png');

const pageSize = (land) => (land ? { W: A4.h, H: A4.w } : { W: A4.w, H: A4.h });

function formatDate(iso) {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return match ? `${match[3]} / ${match[2]} / ${match[1]}` : '____ / ____ / 20____';
}

// ---- the page grid: how many rows fit, and how tall they are (the generator's make()) ---------
function infoRowCount(spec) {
  return 3 + spec.headerRows.length;
}

function layoutFor(c, spec) {
  const { W, H } = pageSize(spec.landscape);
  const CW = W - 2 * M;
  const fits = c.stringWidth(spec.title, 'Arial-Bold', 12) + c.stringWidth(spec.reference, 'Arial', 7.5) + 30 < CW;
  const bandTitle = fits ? 22 : 30;
  const infoTop = H - M - 25 - (bandTitle - 22) - 26;
  const gy = infoTop - infoRowCount(spec) * 17;
  const top = gy - 4 - 16 - 6;
  const footerTop = M + 44 + (spec.legend ? 38 : 26);
  const bandHeight = spec.landscape ? BAND_LANDSCAPE : BAND;
  const space = top - 30 - (footerTop + bandHeight + 10);
  let rows = Math.max(spec.baseRows, spec.minRows);
  const fit = Math.floor(Math.floor(space / spec.rowHeight) / spec.group) * spec.group;
  if (spec.fill && fit > rows) {
    rows = fit;
  }
  const rowHeight = Math.min(space / rows, 34);
  if (rowHeight < 17.5) {
    throw new Error(`${spec.title}: rows too small (${rowHeight.toFixed(1)}pt)`);
  }
  return { W, H, CW, bandTitle, infoTop, top, bandHeight, rows, rowHeight, fits };
}

// Rows per printed page for a sheet, and how many pre-printed items (rows, or row groups) fit.
function pageCapacity(spec) {
  const probe = new Canvas();
  const { rows } = layoutFor(probe, spec);
  return { rows, items: rows / spec.group };
}

// ---- drawing helpers --------------------------------------------------------------------------
function tickBox(c, x, y, size, marked) {
  c.setStrokeColor(BOX);
  c.setLineWidth(0.5);
  c.rect(x, y, size, size);
  if (marked) {
    c.setStrokeColor(BLACK);
    c.setLineWidth(0.9);
    c.line(x + 1, y + 1, x + size - 1, y + size - 1);
    c.line(x + 1, y + size - 1, x + size - 1, y + 1);
  }
}

function fitText(c, x, y, value, width, size = 8, minSize = 5.5) {
  let fs = size;
  while (fs > minSize && c.stringWidth(value, 'Arial', fs) > width) {
    fs -= 0.25;
  }
  c.setFont('Arial', fs);
  c.setFillColor(BLACK);
  c.text(x, y, value);
}

function selected(filledValue, optionValue) {
  if (Array.isArray(filledValue)) {
    return filledValue.includes(optionValue);
  }
  return filledValue !== undefined && filledValue !== null && filledValue !== '' && String(filledValue) === String(optionValue);
}

function inlineTicks(c, x, baseline, labels, values, filledValue, cellWidth) {
  const size = 6.2;
  let xx = x;
  labels.forEach((label, index) => {
    tickBox(c, xx, baseline - 1.2, size, filledValue !== undefined && selected(filledValue, values ? values[index] : true));
    c.setFont('Arial', 7);
    c.setFillColor(BLACK);
    c.text(xx + size + 2, baseline, label);
    xx += size + 2 + c.stringWidth(label, 'Arial', 7) + 6;
  });
  return xx - x <= cellWidth;
}

// ---- header ---------------------------------------------------------------------------------
function drawHeader(c, spec, lay, ctx) {
  const { W, H, CW, bandTitle, infoTop } = lay;
  const filled = ctx.filled || null;
  c.image(LOGO, M, H - M - 17, 17, 17);
  c.setFont('Arial-Bold', 9);
  c.setFillColor(NAVY);
  c.text(M + 22, H - M - 12, 'Elegrow Technology Pvt. Ltd.');
  c.setFont('Arial', 8);
  c.text(W - M, H - M - 12, `${spec.docNo} | Rev 00`, { anchor: 'r' });
  c.setStrokeColor(NAVY);
  c.setLineWidth(1.1);
  c.line(M, H - M - 21, W - M, H - M - 21);
  const y = H - M - 25;
  c.setFillColor(NAVY);
  c.rect(M, y - bandTitle, CW, bandTitle, { fill: true, stroke: false });
  c.setFillColor(WHITE);
  c.setFont('Arial-Bold', 12);
  c.text(M + 8, y - 15, spec.title);
  c.setFont('Arial', 7.5);
  if (lay.fits) {
    c.text(W - M - 8, y - 15, spec.reference, { anchor: 'r' });
  } else {
    c.text(M + 8, y - 26, spec.reference);
  }

  const rh = 17;
  const job = (filled && filled.job) || {};
  const rows = [
    [
      { label: 'Client', text: ctx.project.clientName, frac: 0.42 },
      { label: 'Project No.', text: ctx.project.projectNo, frac: 0.33 },
      { label: 'Work Order No.', text: ctx.project.workOrder, frac: 0.25 }
    ],
    [
      { label: 'Site / Unit', text: ctx.project.siteLocation, frac: 0.42 },
      { label: 'Work date', text: formatDate(ctx.workDate), frac: 0.33 },
      { label: 'Sheet', text: `${ctx.page} of ${ctx.pages}`, frac: 0.25 }
    ],
    [
      { label: 'ETPL engineer(s)', text: ctx.engineer || '', frac: 0.42 },
      { label: 'Client representative', text: job.clientRepresentative || '', frac: 0.33 },
      { label: 'Ground condition', ticks: GROUND_CONDITIONS, values: GROUND_CONDITIONS, key: 'groundCondition', frac: 0.25 }
    ],
    ...spec.headerRows
  ];
  const qw = Math.min(66, rows.length * rh);
  const gw = CW - qw - 6;
  const headerValues = (filled && filled.headerValues) || {};
  // cells, labels and values, from the top of the info grid
  rows.forEach((row, r) => {
    let x = M;
    row.forEach((cell) => {
      const w = gw * cell.frac;
      const cellTop = infoTop - r * rh;
      c.setStrokeColor(GREY);
      c.setLineWidth(0.5);
      c.rect(x, cellTop - rh, w, rh);
      c.setFillColor(NAVY);
      c.setFont('Arial-Bold', 6.3);
      c.text(x + 3, cellTop - 6.8, cell.label);
      const baseline = cellTop - 14.6;
      if (cell.ticks) {
        const value = cell.key === 'groundCondition' ? job.groundCondition : headerValues[cell.key];
        inlineTicks(c, x + 3, baseline, cell.ticks, cell.values, filled ? value || '' : undefined, w);
      } else if (cell.subKeys && filled) {
        const text = cell.subKeys
          .map((sub) => `${sub.label} ${headerValues[sub.key] !== undefined ? headerValues[sub.key] : '_____'}${sub.unit}`)
          .join('   ');
        fitText(c, x + 3, baseline, text, w - 6);
      } else if (cell.text !== undefined) {
        const value = filled && cell.key && headerValues[cell.key] !== undefined ? String(headerValues[cell.key]) : cell.text;
        fitText(c, x + 3, baseline, value, w - 6);
      }
      x += w;
    });
  });
  const gy = infoTop - rows.length * rh;
  c.qr(ctx.qr, W - M - qw, infoTop - qw + 1, qw);

  // instrument row
  const iy = gy - 4;
  const ih = 16;
  c.setStrokeColor(GREY);
  c.rect(M, iy - ih, CW, ih);
  c.setFillColor(NAVY);
  c.setFont('Arial-Bold', 6.5);
  c.text(M + 3, iy - 10.5, 'Instrument used:');
  let x = M + 66;
  spec.instruments.forEach((instrument) => {
    tickBox(c, x, iy - 11.5, 7, Boolean(filled && filled.instrument === instrument.id));
    c.setFillColor(BLACK);
    c.setFont('Arial', 7);
    c.text(x + 10, iy - 10.5, instrument.label);
    x += 14 + c.stringWidth(instrument.label, 'Arial', 7) + 8;
  });
  c.setFont('Arial', 7);
  const serialStart = x + 4;
  c.text(serialStart, iy - 10.5, 'Sr. No. __________   Cal. due ___/___/____');
  if (filled) {
    c.setFont('Arial-Bold', 7);
    if (filled.serialNo) {
      c.text(serialStart + c.stringWidth('Sr. No. ', 'Arial', 7) + 1, iy - 9.5, filled.serialNo);
    }
    if (filled.calibrationDue) {
      c.text(serialStart + c.stringWidth('Sr. No. __________   Cal. due ', 'Arial', 7) + 1, iy - 9.5, filled.calibrationDue);
    }
  }
  return iy - ih - 6;
}

// ---- footer ---------------------------------------------------------------------------------
function drawFooter(c, spec, lay, ctx) {
  const { W, CW } = lay;
  const filled = ctx.filled || null;
  const y = M + 44;
  if (spec.legend) {
    const parts = CODE_LEGEND.split(' · ');
    c.setFont('Arial-Bold', 6);
    c.setFillColor(NAVY);
    c.text(M, y + 29, 'Condition codes:');
    c.setFont('Arial', 6);
    c.setFillColor(BLACK);
    c.text(M + 52, y + 29, parts.slice(0, 5).join(' · '));
    c.text(M + 52, y + 21.5, parts.slice(5).join(' · '));
  }
  ['ETPL engineer – name & signature', 'Client representative – name & signature', 'Time start / finish'].forEach((label, i) => {
    const w = CW * (i < 2 ? 0.38 : 0.24);
    const x = M + CW * [0, 0.38, 0.76][i];
    c.setStrokeColor(GREY);
    c.setLineWidth(0.5);
    c.rect(x, y - 6, w, 26);
    c.setFillColor(NAVY);
    c.setFont('Arial-Bold', 6.3);
    c.text(x + 3, y + 13, label);
    if (filled && i === 2 && (filled.timeStart || filled.timeFinish)) {
      c.setFont('Arial-Bold', 8);
      c.setFillColor(BLACK);
      c.text(x + 3, y - 1, `${filled.timeStart || ''} / ${filled.timeFinish || ''}`);
    }
  });
  c.setStrokeColor(GREY);
  c.line(M, M + 14, W - M, M + 14);
  c.setFont('Arial-Bold', 7);
  c.setFillColor(NAVY);
  c.text(M, M + 4, 'INTERNAL');
  c.text(W / 2, M + 4, `Sheet ID ${ctx.sheetId}`, { anchor: 'c' });
  c.setFont('Arial', 7);
  c.text(W - M, M + 4, `${spec.title.split(' – ')[0]} | ${spec.docNo}`, { anchor: 'r' });
}

// ---- table ----------------------------------------------------------------------------------
function ticksCell(c, x, y, w, h, labels, values, filledValue) {
  const sz = 6.6;
  const fs = 6;
  const item = (o) => sz + 2 + c.stringWidth(o, 'Arial', fs);
  let lines = [labels];
  if (labels.length > 1 && labels.reduce((sum, o) => sum + item(o), 0) + 4 * (labels.length - 1) > w - 6) {
    const k = Math.floor((labels.length + 1) / 2);
    lines = [labels.slice(0, k), labels.slice(k)];
  }
  const lh = sz + 3.2;
  const yTop = y + h / 2 + (lines.length * lh) / 2 - lh;
  let index = 0;
  lines.forEach((line, li) => {
    const total = line.reduce((sum, o) => sum + item(o), 0) + 4 * (line.length - 1);
    let xx = x + Math.max(3, (w - total) / 2);
    const yy = yTop - li * lh + 1.6;
    line.forEach((o) => {
      tickBox(c, xx, yy, sz, filledValue !== undefined && selected(filledValue, values[index]));
      c.setFont('Arial', fs);
      c.setFillColor(BLACK);
      c.text(xx + sz + 2, yy + 1.2, o);
      xx += item(o) + 4;
      index += 1;
    });
  });
}

function drawTable(c, spec, lay, top, ctx) {
  const { CW, rows: nRows, rowHeight: rh } = lay;
  const { group } = spec;
  const total = spec.columns.reduce((sum, col) => sum + col.w, 0);
  const ws = spec.columns.map((col) => (col.w / total) * CW);
  const hh = 30;
  let x = M;
  spec.columns.forEach((col, i) => {
    const w = ws[i];
    c.setFillColor(TINT);
    c.rect(x, top - hh, w, hh, { fill: true, stroke: false });
    c.setStrokeColor(GREY);
    c.setLineWidth(0.5);
    c.rect(x, top - hh, w, hh);
    c.setFillColor(NAVY);
    c.setFont('Arial-Bold', 6.5);
    const lines = wrapLabel(c, col.label, w - 3);
    lines.forEach((line, li) => {
      c.text(x + w / 2, top - hh / 2 + (lines.length - 1) * 3.9 - li * 7.8 - 2, line, { anchor: 'c' });
    });
    x += w;
  });
  const y0 = top - hh;
  const itemOffset = ctx.itemOffset || 0;
  const filledRows = (ctx.filled && ctx.filled.rows) || [];
  const spacings = soilSpacings();
  for (let r = 0; r < nRows; r += 1) {
    const y = y0 - (r + 1) * rh;
    x = M;
    const first = r % group === 0;
    const item = Math.floor(r / group);
    const pre = (ctx.preItems && ctx.preItems[item]) || {};
    const fill = filledRows[r] || {};
    spec.columns.forEach((col, i) => {
      const w = ws[i];
      const cellValue = fill[col.key];
      if (col.merge) {
        if (first) {
          c.setStrokeColor(GREY);
          c.rect(x, y0 - (r + group) * rh, w, group * rh);
          const midY = y0 - (r + group / 2) * rh;
          if (col.k === 'sr') {
            c.setFont('Arial', 7.5);
            c.setFillColor(BLACK);
            c.text(x + w / 2, midY - 3, String(itemOffset + item + 1), { anchor: 'c' });
          } else if (col.k === 'text') {
            const value = cellValue !== undefined ? cellValue : pre[col.key];
            if (value) {
              fitText(c, x + 3, midY - 3, String(value), w - 5, 7, 5);
            }
          } else if (col.k === 'digits' && cellValue !== undefined) {
            fitText(c, x + 3, midY - 3, String(cellValue), w - 5, 8, 5);
          } else if (col.k === 'ticks') {
            ticksCell(c, x, y0 - (r + group) * rh, w, group * rh, col.options, col.values, ctx.filled ? cellValue || '' : undefined);
          }
        }
        x += w;
        return;
      }
      c.setStrokeColor(GREY);
      c.setLineWidth(0.5);
      c.rect(x, y, w, rh);
      if (col.k === 'sr') {
        c.setFont('Arial', 7.5);
        c.setFillColor(BLACK);
        c.text(x + w / 2, y + rh / 2 - 3, String((ctx.rowOffset || 0) + r + 1), { anchor: 'c' });
      } else if (col.k === 'fixed') {
        const values = spec.id === 'soilResistivity' ? spacings : col.values;
        c.setFont('Arial', 7.5);
        c.setFillColor(BLACK);
        c.text(x + w / 2, y + rh / 2 - 3, values[r % values.length], { anchor: 'c' });
      } else if (col.k === 'text') {
        const value = cellValue !== undefined ? cellValue : group === 1 ? pre[col.key] : undefined;
        if (value) {
          fitText(c, x + 3, y + rh / 2 - 3, String(value), w - 5, 7, 5);
        }
      } else if (col.k === 'digits') {
        if (cellValue !== undefined && cellValue !== '') {
          fitText(c, x + 3, y + rh / 2 - 3, String(cellValue), w - 5, 8, 5);
        }
      } else if (col.k === 'ticks') {
        ticksCell(c, x, y, w, rh, col.options, col.values, ctx.filled ? cellValue || '' : undefined);
      }
      x += w;
    });
  }
  return y0 - nRows * rh;
}

function wrapLabel(c, label, width) {
  const words = String(label).split(' ');
  const lines = [];
  let current = '';
  words.forEach((word) => {
    const trial = current ? `${current} ${word}` : word;
    if (current && c.stringWidth(trial, 'Arial-Bold', 6.5) > width) {
      lines.push(current);
      current = word;
    } else {
      current = trial;
    }
  });
  if (current) {
    lines.push(current);
  }
  return lines;
}

// ---- SET-UP diagram + how-to-record notes ------------------------------------------------------
function wrapNote(c, text, width, fs) {
  const lines = [];
  let current = '';
  let first = true;
  text.split(' ').forEach((word) => {
    const trial = `${current} ${word}`.trim();
    if (c.stringWidth(`${first ? '• ' : '  '}${trial}`, 'Arial', fs) > width) {
      lines.push(`${first ? '• ' : '  '}${current}`);
      current = word;
      first = false;
    } else {
      current = trial;
    }
  });
  lines.push(`${first ? '• ' : '  '}${current}`);
  return lines;
}

function drawBand(c, spec, lay, yTop) {
  const { CW, bandHeight: bh } = lay;
  const fw = CW * (lay.W < 700 ? 0.5 : 0.4);
  c.setStrokeColor(GREY);
  c.setLineWidth(0.5);
  c.rect(M, yTop - bh, fw, bh);
  c.setFillColor(NAVY);
  c.setFont('Arial-Bold', 6);
  c.text(M + 3, yTop - 7, 'SET-UP');
  FIGURES[spec.id](c, M + 4, yTop - bh + 2, fw - 8, bh - 10);
  const tx = M + fw + 8;
  const tw = CW - fw - 8;
  c.setFont('Arial-Bold', 6.3);
  c.setFillColor(NAVY);
  c.text(tx, yTop - 7, 'HOW TO RECORD');
  // Notes come from the sheet definitions; shrink the type a little if they run long.
  let fs = 6.3;
  let blocks;
  let needed;
  for (;;) {
    c.setFont('Arial', fs);
    const lineHeight = (7.8 * fs) / 6.3;
    blocks = spec.notes.map((note) => wrapNote(c, note, tw, fs));
    needed = 8.5 + blocks.reduce((sum, lines) => sum + lines.length * lineHeight + 0.8, 0);
    if (needed <= bh - 3 || fs <= 5.2) {
      break;
    }
    fs -= 0.1;
  }
  if (needed > bh - 3) {
    throw new Error(`${spec.title}: how-to-record notes overflow the band`);
  }
  c.setFont('Arial', fs);
  c.setFillColor(BLACK);
  let ty = yTop - 7 - 8.5;
  blocks.forEach((lines) => {
    lines.forEach((line) => {
      c.text(tx, ty, line);
      ty -= (7.8 * fs) / 6.3;
    });
    ty -= 0.8;
  });
}

// ---- one page -------------------------------------------------------------------------------
// ctx: { sheetKey, sheetId, project:{clientName,projectNo,workOrder,siteLocation}, workDate,
//        engineer, page, pages, preItems, itemOffset, rowOffset, filled }
function drawPage(c, ctx) {
  const spec = SPECS[ctx.sheetKey];
  if (!spec) {
    throw new Error(`No printable sheet for "${ctx.sheetKey}".`);
  }
  const lay = layoutFor(c, spec);
  c.newPage(lay.W, lay.H);
  const qr = QRCode.create(`ER1|${ctx.sheetId}`, { errorCorrectionLevel: 'M' });
  const full = { ...ctx, qr, rowOffset: (ctx.itemOffset || 0) * spec.group };
  const top = drawHeader(c, spec, lay, full);
  drawFooter(c, spec, lay, full);
  const end = drawTable(c, spec, lay, top, full);
  drawBand(c, spec, lay, end - 6);
  return { rows: lay.rows, rowHeight: lay.rowHeight, landscape: spec.landscape };
}

// pages: array of ctx objects. Returns { buffer, layouts }.
async function renderSheetsPdf(pages, { title = 'ETPL ESHA Measurement Sheets', fonts } = {}) {
  const c = new Canvas({ title, fonts });
  const layouts = pages.map((page) => drawPage(c, page));
  const buffer = await c.toBuffer();
  return { buffer, layouts };
}

module.exports = { renderSheetsPdf, layoutFor, pageCapacity, formatDate, A4 };
