// Real PDF / DOCX / XLSX bytes built in memory, for the Graph mock and its tests (nothing on disk).
const PDFDocument = require('pdfkit');
const { zipSync, strToU8 } = require('fflate');

const esc = (value) => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function makePdf(paragraphs) {
  return new Promise((resolve) => {
    const doc = new PDFDocument({ margin: 50 });
    const chunks = [];
    doc.on('data', (chunk) => chunks.push(chunk));
    doc.on('end', () => resolve(Buffer.concat(chunks)));
    (Array.isArray(paragraphs) ? paragraphs : [paragraphs]).forEach((text) => doc.text(String(text)).moveDown());
    doc.end();
  });
}

function makeDocx(paragraphs) {
  const body = (Array.isArray(paragraphs) ? paragraphs : [paragraphs]).map((text) => `<w:p><w:r><w:t xml:space="preserve">${esc(text)}</w:t></w:r></w:p>`).join('');
  return Buffer.from(zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'),
    'word/document.xml': strToU8(`<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}</w:body></w:document>`)
  }));
}

// rows: [[cell, ...], ...]; text cells go through sharedStrings like Excel writes them, numbers inline.
function makeXlsx(rows, sheetName = 'Sheet1') {
  const strings = [];
  const sheetRows = rows.map((row, r) => `<row r="${r + 1}">${row.map((cell, c) => {
    const ref = `${String.fromCharCode(65 + c)}${r + 1}`;
    if (typeof cell === 'number') return `<c r="${ref}"><v>${cell}</v></c>`;
    strings.push(String(cell));
    return `<c r="${ref}" t="s"><v>${strings.length - 1}</v></c>`;
  }).join('')}</row>`).join('');
  return Buffer.from(zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'),
    'xl/workbook.xml': strToU8(`<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="${esc(sheetName)}" sheetId="1"/></sheets></workbook>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${strings.map((text) => `<si><t xml:space="preserve">${esc(text)}</t></si>`).join('')}</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetRows}</sheetData></worksheet>`)
  }));
}

module.exports = { makePdf, makeDocx, makeXlsx };
