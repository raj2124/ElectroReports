// Test helper (Session 6I): real printed sheet pages, filled in by "hand" with known values.
//   const sheet = await issueFilled(store, { sheetKey, project, workDate, rows, header })
//   sheet.issued   the issued_sheets record (what the QR code leads to)
//   sheet.pdf      the page as a PDF (QR code, printed rows, the values written into the cells)
//   sheet.truth    what a perfect reader would return: { "<cell path>": { value | selected } }
// Values are given as { rows: [ { <column key>: value } ... ] (per printed row; merged columns on the group's first row),
//                       header: { <header key>: value }, instrument, serialNo, calibrationDue, timeStart, timeFinish, job: {...} }.
const { printSheets } = require('../../src/sheetPrint');
const { renderSheetsPdf } = require('../../src/sheetPrint/render');
const { buildForm } = require('../../src/scan/form');

// A page that was already issued (a record of issued_sheets), filled in by "hand": { issued, form, pdf, truth }.
async function fillIssued(issued, { rows = [], header = {}, job = {}, instrument = '', serialNo = '', calibrationDue = '', timeStart = '', timeFinish = '' } = {}) {
  const form = buildForm({ ...issued, sheetId: issued.sheetId });
  const ctx = { sheetKey: issued.sheetKey, sheetId: issued.sheetId, project: issued.project, workDate: issued.workDate, engineer: issued.engineer || '', page: issued.page, pages: issued.pages, itemOffset: issued.itemOffset || 0, preItems: issued.preRows || [], preRows: issued.preRows || [], filled: { job, headerValues: header, instrument, serialNo, calibrationDue, timeStart, timeFinish, rows } };
  const { buffer } = await renderSheetsPdf([ctx]);
  const truth = {};
  const set = (path, value) => { truth[path] = Array.isArray(value) ? { selected: value } : value === true ? { selected: ['true'] } : { value: String(value) }; };
  form.cells.forEach((def) => {
    let value;
    if (def.section === 'header') {
      value = def.key === 'instrument' ? instrument : def.key === 'serialNo' ? serialNo : def.key === 'calibrationDue' ? calibrationDue : def.key === 'timeStart' ? timeStart : def.key === 'timeFinish' ? timeFinish : def.job ? job[def.key] : header[def.key];
    } else if (!def.preprinted) {
      value = (rows[def.row] || {})[def.key];
    }
    if (value === undefined || value === '' || (Array.isArray(value) && !value.length) || value === false) return;
    if (def.kind === 'ticks' && !Array.isArray(value)) value = [value];
    set(def.path, value);
  });
  return { issued, form, pdf: buffer, truth, ctx };
}

async function issueFilled(store, { sheetKey, project, workDate, reportId = '', pageIndex = 0, ...values }) {
  const printed = await printSheets({ project, workDate, sheets: [sheetKey], reportId }, { store, getTasks: async () => [] });
  const page = printed.pages[pageIndex];
  const filled = await fillIssued(store.getIssuedSheet(page.sheetId), values);
  return { ...filled, pageSet: printed.pages };
}

// RGBA pixels -> PNG bytes (no library: stored deflate blocks via zlib).
function encodePng({ data, width, height }) {
  const zlib = require('zlib');
  const crcTable = Array.from({ length: 256 }, (_, n) => { let c = n; for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; return c >>> 0; });
  const crc = (buffer) => { let c = 0xffffffff; for (const byte of buffer) c = crcTable[(c ^ byte) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };
  const chunk = (type, body) => { const out = Buffer.alloc(12 + body.length); out.writeUInt32BE(body.length, 0); out.write(type, 4, 'latin1'); body.copy(out, 8); out.writeUInt32BE(crc(out.subarray(4, 8 + body.length)), 8 + body.length); return out; };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0); header.writeUInt32BE(height, 4); header[8] = 8; header[9] = 6;
  const raw = Buffer.alloc((width * 4 + 1) * height);
  for (let y = 0; y < height; y += 1) { raw[y * (width * 4 + 1)] = 0; Buffer.from(data.buffer, data.byteOffset + y * width * 4, width * 4).copy(raw, y * (width * 4 + 1) + 1); }
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}

// A page image with the QR code painted over (the printed "Sheet ID ..." text stays): a photo whose code will not scan.
function withoutQr(page) {
  const { data, width, height } = page.rgba;
  const copy = new Uint8ClampedArray(data);
  for (let y = Math.floor(height * 0.04); y < Math.floor(height * 0.34); y += 1) {
    for (let x = Math.floor(width * 0.8); x < width; x += 1) {
      const at = (y * width + x) * 4;
      copy[at] = 255; copy[at + 1] = 255; copy[at + 2] = 255; copy[at + 3] = 255;
    }
  }
  return encodePng({ data: copy, width, height });
}

module.exports = { issueFilled, fillIssued, encodePng, withoutQr, renderSheetsPdf };
