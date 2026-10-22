// Text out of the three document types of the knowledge base (Session 6H.3). The file is only ever in memory:
// the caller hands over the bytes, the text comes back, nothing is written anywhere.

const { unzipSync, strFromU8 } = require('fflate');
const { openPdf, pageCount, pageTextLines } = require('./pdfRaster');

const MAX_TEXT = 2000000;
const MAX_ENTRY_BYTES = 100 * 1024 * 1024; // refuse a zip entry that would inflate beyond this

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
const decode = (value) => value.replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, name) => {
  if (name[0] === '#') {
    const code = name[1].toLowerCase() === 'x' ? parseInt(name.slice(2), 16) : parseInt(name.slice(1), 10);
    return Number.isFinite(code) && code > 0 && code < 0x110000 ? String.fromCodePoint(code) : '';
  }
  return ENTITIES[name.toLowerCase()] !== undefined ? ENTITIES[name.toLowerCase()] : whole;
});

function unzipEntries(buffer, wanted) {
  const files = unzipSync(new Uint8Array(buffer), { filter: (file) => file.originalSize <= MAX_ENTRY_BYTES && wanted(file.name) });
  return Object.fromEntries(Object.entries(files).map(([name, bytes]) => [name, strFromU8(bytes)]));
}

async function pdfText(buffer) {
  const pdf = await openPdf(buffer);
  const out = [];
  for (let page = 0; page < pageCount(pdf) && out.join('\n').length < MAX_TEXT; page += 1) {
    const lines = [];
    pageTextLines(pdf, page).sort((a, b) => a.y - b.y || a.x - b.x).forEach((fragment) => {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - fragment.y) < 3) last.text += ` ${fragment.text}`;
      else lines.push({ y: fragment.y, text: fragment.text });
    });
    out.push(lines.map((line) => line.text.trim()).filter(Boolean).join('\n'));
  }
  return out.join('\n\n');
}

function docxText(buffer) {
  const xml = unzipEntries(buffer, (name) => name === 'word/document.xml')['word/document.xml'] || '';
  return decode(xml.replace(/<w:tab\/>/g, '\t').replace(/<w:br[^>]*\/>/g, '\n').replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, ''));
}

function xlsxText(buffer) {
  const files = unzipEntries(buffer, (name) => name === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
  const shared = [...(files['xl/sharedStrings.xml'] || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map((match) => decode([...match[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((part) => part[1]).join('')));
  return Object.keys(files).filter((name) => name.startsWith('xl/worksheets/')).sort().map((name) => [...files[name].matchAll(/<row[^>]*>([\s\S]*?)<\/row>/g)].map((row) => [...row[1].matchAll(/<c([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)].map((cell) => {
    const body = cell[2] || '';
    if (/t="s"/.test(cell[1])) return shared[Number((/<v>(\d+)<\/v>/.exec(body) || [])[1])] || '';
    if (/t="inlineStr"/.test(cell[1])) return decode([...body.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((part) => part[1]).join(''));
    return decode((/<v>([\s\S]*?)<\/v>/.exec(body) || [])[1] || '');
  }).filter((value) => value !== '').join('\t')).filter(Boolean).join('\n')).join('\n\n');
}

// ext: 'pdf' | 'docx' | 'xlsx'. Returns the text (possibly '' for a scanned PDF).
async function extractText(ext, buffer) {
  const text = ext === 'pdf' ? await pdfText(buffer) : ext === 'docx' ? docxText(buffer) : ext === 'xlsx' ? xlsxText(buffer) : '';
  return text.replace(/\r/g, '').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim().slice(0, MAX_TEXT);
}

// About `size` characters per chunk with `overlap` carried over, cut at a space where one is near.
function chunkText(text, size = 1000, overlap = 150) {
  const flat = String(text).replace(/\s+/g, ' ').trim();
  const chunks = [];
  for (let start = 0; start < flat.length; ) {
    let end = Math.min(flat.length, start + size);
    if (end < flat.length) {
      const space = flat.lastIndexOf(' ', end);
      if (space > start + size * 0.6) end = space;
    }
    chunks.push(flat.slice(start, end).trim());
    if (end >= flat.length) break;
    start = Math.max(end - overlap, start + 1);
  }
  return chunks.filter(Boolean);
}

module.exports = { extractText, chunkText };
