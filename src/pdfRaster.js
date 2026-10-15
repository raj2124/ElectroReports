// Reading PDFs with mupdf (WASM, no native dependencies): page count, text lines with
// positions, and rasterising a page to pixels. mupdf is an ES module, so it is loaded lazily.
let mupdfPromise = null;

function loadMupdf() {
  if (!mupdfPromise) {
    mupdfPromise = import('mupdf');
  }
  return mupdfPromise;
}

async function openPdf(buffer) {
  const mupdf = await loadMupdf();
  return { mupdf, doc: mupdf.Document.openDocument(buffer, 'application/pdf') };
}

function pageCount(pdf) {
  return pdf.doc.countPages();
}

function pageSize(pdf, index) {
  const [x0, y0, x1, y1] = pdf.doc.loadPage(index).getBounds();
  return { width: x1 - x0, height: y1 - y0 };
}

// Text fragments of a page: [{ text, x, y, w, h }] in PDF points, y down from the top. A line is
// split where two characters are further apart than a word space, so separately drawn labels
// that happen to share a baseline (neighbouring table headers) stay separate.
function pageTextLines(pdf, index, { gap = 4 } = {}) {
  const fragments = [];
  let current = null;
  const flush = () => {
    if (current && current.text.trim()) {
      fragments.push({ text: current.text, x: current.x0, y: current.y0, w: current.x1 - current.x0, h: current.y1 - current.y0 });
    }
    current = null;
  };
  pdf.doc.loadPage(index).toStructuredText('preserve-whitespace').walk({
    beginLine() { flush(); },
    endLine() { flush(); },
    onChar(char, _origin, _font, _size, quad) {
      const [ulx, uly, urx, , , lly] = quad;
      if (current && ulx - current.x1 > gap) {
        flush();
      }
      if (!current) {
        current = { text: '', x0: ulx, x1: urx, y0: uly, y1: lly };
      }
      current.text += char;
      current.x1 = Math.max(current.x1, urx);
      current.y0 = Math.min(current.y0, uly);
      current.y1 = Math.max(current.y1, lly);
    }
  });
  flush();
  return fragments;
}

// Every character of a page with its box: [{ c, x0, x1, y0, y1 }] (points, y down).
function pageChars(pdf, index) {
  const chars = [];
  pdf.doc.loadPage(index).toStructuredText('preserve-whitespace').walk({
    onChar(char, _origin, _font, _size, quad) {
      chars.push({ c: char, x0: quad[0], x1: quad[2], y0: quad[1], y1: quad[5] });
    }
  });
  return chars;
}

// Rasterises a page to RGBA pixels (opaque white background) at the given dpi.
function renderPageRgba(pdf, index, dpi = 150) {
  const { mupdf } = pdf;
  const scale = dpi / 72;
  const pixmap = pdf.doc.loadPage(index).toPixmap(mupdf.Matrix.scale(scale, scale), mupdf.ColorSpace.DeviceRGB, false, true);
  const width = pixmap.getWidth();
  const height = pixmap.getHeight();
  const channels = pixmap.getNumberOfComponents();
  const source = pixmap.getPixels();
  const rgba = new Uint8ClampedArray(width * height * 4);
  for (let i = 0, j = 0; i < width * height; i += 1, j += channels) {
    rgba[i * 4] = source[j];
    rgba[i * 4 + 1] = source[j + 1];
    rgba[i * 4 + 2] = source[j + 2];
    rgba[i * 4 + 3] = 255;
  }
  return { width, height, data: rgba, png: () => pixmap.asPNG() };
}

module.exports = { openPdf, pageCount, pageSize, pageTextLines, pageChars, renderPageRgba };
