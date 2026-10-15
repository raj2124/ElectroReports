// A small reportlab-style canvas (origin bottom-left, y up) over pdfkit, so the printed-sheet
// layout in reference/generators/make_sheets_v3.py and diagrams.py ports one-to-one with the
// same coordinates and numbers.
const PDFDocument = require('pdfkit');
const { resolveFonts } = require('./fonts');

class Canvas {
  constructor({ title = '', author = 'Elegrow Technology Pvt. Ltd.', fonts } = {}) {
    fonts = fonts || resolveFonts();
    this.doc = new PDFDocument({ autoFirstPage: false, margin: 0, info: { Title: title, Author: author } });
    this.fonts = fonts;
    this.doc.registerFont('Arial', fonts.regular);
    this.doc.registerFont('Arial-Bold', fonts.bold);
    this.W = 0;
    this.H = 0;
    this.fontName = 'Arial';
    this.fontSize = 8;
    this.pageCount = 0;
  }

  newPage(W, H) {
    this.doc.addPage({ size: [W, H], margin: 0 });
    this.W = W;
    this.H = H;
    this.pageCount += 1;
  }

  setFont(name, size) {
    this.fontName = name;
    this.fontSize = size;
    this.doc.font(name).fontSize(size);
  }

  stringWidth(text, name = this.fontName, size = this.fontSize) {
    return this.doc.font(name).fontSize(size).widthOfString(String(text));
  }

  setFillColor(color) { this.doc.fillColor(color); }
  setStrokeColor(color) { this.doc.strokeColor(color); }
  setLineWidth(width) { this.doc.lineWidth(width); }
  setDash(on, off) { if (on === undefined) { this.doc.undash(); } else { this.doc.dash(on, { space: off === undefined ? on : off }); } }

  line(x1, y1, x2, y2) {
    this.doc.moveTo(x1, this.H - y1).lineTo(x2, this.H - y2).stroke();
  }

  rect(x, y, w, h, { fill = false, stroke = true } = {}) {
    this.doc.rect(x, this.H - y - h, w, h);
    this.paint(fill, stroke);
  }

  circle(cx, cy, r, { fill = false, stroke = true } = {}) {
    this.doc.circle(cx, this.H - cy, r);
    this.paint(fill, stroke);
  }

  paint(fill, stroke) {
    if (fill && stroke) { this.doc.fillAndStroke(); } else if (fill) { this.doc.fill(); } else if (stroke) { this.doc.stroke(); }
  }

  polyline(points, { stroke = true, fill = false } = {}) {
    points.forEach(([x, y], index) => (index ? this.doc.lineTo(x, this.H - y) : this.doc.moveTo(x, this.H - y)));
    this.paint(fill, stroke);
  }

  // Text at the baseline (x, y); anchor l / c / r.
  text(x, y, value, { anchor = 'l' } = {}) {
    const string = String(value);
    const width = this.doc.widthOfString(string);
    const left = anchor === 'c' ? x - width / 2 : anchor === 'r' ? x - width : x;
    this.doc.text(string, left, this.H - y, { lineBreak: false, baseline: 'alphabetic' });
  }

  drawString(x, y, value) { this.text(x, y, value); }
  drawCentredString(x, y, value) { this.text(x, y, value, { anchor: 'c' }); }
  drawRightString(x, y, value) { this.text(x, y, value, { anchor: 'r' }); }

  image(source, x, y, w, h) {
    this.doc.image(source, x, this.H - y - h, { width: w, height: h });
  }

  // QR code as vector squares (crisp at any scan resolution). `qr` is a qrcode.create() result.
  qr(qr, x, y, size, quiet = 1) {
    const count = qr.modules.size;
    const unit = size / (count + 2 * quiet);
    this.doc.fillColor('#000000');
    for (let row = 0; row < count; row += 1) {
      for (let col = 0; col < count; col += 1) {
        if (qr.modules.data[row * count + col]) {
          this.doc.rect(x + (col + quiet) * unit, this.H - (y + size - (row + quiet) * unit), unit + 0.05, unit + 0.05).fill();
        }
      }
    }
  }

  toBuffer() {
    return new Promise((resolve, reject) => {
      const chunks = [];
      this.doc.on('data', (chunk) => chunks.push(chunk));
      this.doc.on('end', () => resolve(Buffer.concat(chunks)));
      this.doc.on('error', reject);
      this.doc.end();
    });
  }
}

module.exports = { Canvas };
