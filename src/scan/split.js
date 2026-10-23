// A scan upload into page images (Session 6I): a PDF is split page by page, a PNG or JPEG is one page. mupdf does
// both (it decodes the images too), so there is no extra image library. Each page comes back as a PNG whose long
// edge is at most MAX_EDGE pixels, and as RGBA pixels for the QR search.

const { openPdf, pageCount, pageSize, renderPageRgba } = require('../pdfRaster');

const MAX_EDGE = 2300;
const ACCEPTED = { 'application/pdf': 'pdf', 'image/png': 'image', 'image/jpeg': 'image', 'image/jpg': 'image' };

class UploadError extends Error {
  constructor(message, code = 'BAD_UPLOAD') {
    super(message);
    this.name = 'UploadError';
    this.code = code;
  }
}

// mimetype is trusted only as a hint: the bytes decide.
function sniff(buffer, mimetype) {
  if (buffer.slice(0, 5).toString('latin1') === '%PDF-') return { kind: 'pdf', magic: 'application/pdf' };
  if (buffer.slice(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))) return { kind: 'image', magic: 'image/png' };
  if (buffer[0] === 0xff && buffer[1] === 0xd8) return { kind: 'image', magic: 'image/jpeg' };
  throw new UploadError(ACCEPTED[mimetype] ? 'That file is not a readable PDF, PNG or JPEG.' : 'Upload a PDF, PNG or JPEG of the printed sheets.');
}

async function splitUpload({ buffer, mimetype = '', maxPages = 30 }) {
  const { kind, magic } = sniff(buffer, mimetype);
  let pdf;
  try {
    pdf = kind === 'pdf' ? await openPdf(buffer) : await openImage(buffer, magic);
  } catch (error) {
    throw new UploadError('That file could not be opened. Is it a damaged or password-protected PDF?');
  }
  const count = pageCount(pdf);
  if (count > maxPages) throw new UploadError(`That PDF has ${count} pages; upload at most ${maxPages} at a time.`, 'TOO_MANY_PAGES');
  const pages = [];
  for (let index = 0; index < count; index += 1) {
    const size = pageSize(pdf, index);
    const dpi = Math.min(220, (MAX_EDGE / Math.max(size.width, size.height)) * 72);
    const raster = renderPageRgba(pdf, index, dpi);
    pages.push({ index, width: raster.width, height: raster.height, png: Buffer.from(raster.png()), rgba: { data: raster.data, width: raster.width, height: raster.height } });
  }
  return pages;
}

async function openImage(buffer, magic) {
  const mupdf = await import('mupdf');
  return { mupdf, doc: mupdf.Document.openDocument(buffer, magic) };
}

module.exports = { splitUpload, UploadError, MAX_EDGE };
