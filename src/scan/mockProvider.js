// A deterministic stand-in for the Gemini sheet reader (OCR_USE_MOCK, the default; nothing leaves the machine).
// It never looks at the picture. What it "reads" comes from a truth file (OCR_MOCK_TRUTH_FILE, JSON, re-read on
// every call so a test can change it between calls):
//   { "<sheetId or sheetKey>": { primary: { "<cell path>": { value | selected, confidence?, illegible? } },
//                                escalation: { ... },   // what the stronger model reads (default: the primary)
//                                fail: "QUOTA" },       // simulate a provider failure
//     "__fallbackIds": ["ER-XXXX-XXXXX", ...] }          // what the printed-ID read returns, in turn
// A cell not listed is blank and confident. Without a truth entry the page reads as an empty sheet.

const fs = require('fs');
const { ProviderError } = require('../aiGuard');

function createMockProvider({ truthFile = process.env.OCR_MOCK_TRUTH_FILE || '', truth = null, primaryModel = '' } = {}) {
  let fallbackIndex = 0;
  const load = () => {
    if (truth) return truth;
    if (!truthFile) return {};
    try { return JSON.parse(fs.readFileSync(truthFile, 'utf8')); } catch (_error) { return {}; }
  };
  return {
    mock: true,
    async read({ form, model }) {
      const all = load();
      const entry = all[form.sheetId] || all[form.sheetKey] || {};
      if (entry.fail) throw new ProviderError(entry.fail, `mock failure ${entry.fail}`);
      const escalated = model && primaryModel && model !== primaryModel;
      const given = (escalated && entry.escalation) || entry.primary || {};
      return { cells: form.cells.filter((cell) => !cell.preprinted).map((cell) => ({ path: cell.path, confidence: 0.99, ...(given[cell.path] || {}) })) };
    },
    async readSheetId() {
      const ids = load().__fallbackIds || [];
      const id = ids[fallbackIndex % Math.max(1, ids.length)];
      fallbackIndex += 1;
      return id || '';
    }
  };
}

module.exports = { createMockProvider };
