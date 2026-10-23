// Gemini as a TRANSCRIBER of a printed measurement sheet (Session 6I): it reports what is written and which boxes are
// ticked, cell by cell, with a confidence. It never computes, converts, corrects or guesses; every check and every
// number the report uses comes from the app's own rules. The old per-test OCR (src/geminiOcrService.js) is untouched.

const { GoogleGenerativeAI } = require('@google/generative-ai');
const { ProviderError, classify } = require('../aiGuard');

const SCHEMA = {
  type: 'object',
  properties: {
    cells: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          path: { type: 'string' },
          value: { type: 'string', nullable: true },
          selected: { type: 'array', items: { type: 'string' } },
          confidence: { type: 'number' },
          illegible: { type: 'boolean' }
        },
        required: ['path', 'confidence']
      }
    }
  },
  required: ['cells']
};

function buildPrompt(form) {
  const lines = form.cells.filter((cell) => !cell.preprinted).map((cell) => {
    const options = cell.kind === 'ticks' ? ` [tick boxes: ${cell.options.map((option) => `${option.value} = "${option.label}"`).join('; ')}${cell.multi ? '; several may be ticked' : '; only one should be ticked'}]` : cell.kind === 'tick' ? ' [one tick box]' : cell.kind === 'number' ? ' [a number]' : '';
    return `${cell.path} | ${cell.label}${options}`;
  });
  return [
    `This is a photo or scan of a printed, hand-filled ${form.title} measurement sheet (${form.code}), ${form.landscape ? 'landscape' : 'portrait'}, with ${form.rowCount} table rows.`,
    'Transcribe ONLY what is written or ticked. Do not guess. Do not calculate, convert units, correct, round or complete anything. If a cell is empty, return an empty value with high confidence. If handwriting cannot be read with confidence, set illegible to true and put your best reading (or nothing) in value.',
    'For number cells return the digits exactly as written, with the decimal point or comma as written, without units. For tick-box cells return in "selected" the option codes of the boxes that are marked, exactly as listed; return an empty array when none is marked. For a single tick box return "true" in selected when marked.',
    'Give every cell a confidence between 0 and 1: how sure you are of exactly what you returned. Return one entry for every cell path listed below and no others. Do not read the serial numbers, the printed probe spacings, phase names or foot names: they are not cells.',
    'Cells (path | label):',
    ...lines
  ].join('\n');
}

function createGeminiProvider({ config, clientFactory = (key) => new GoogleGenerativeAI(key) }) {
  const apiKey = String(config.gemini.apiKey || '').trim();
  const generate = async (model, parts, responseSchema) => {
    if (!apiKey) throw new ProviderError('NOT_CONFIGURED', 'GEMINI_API_KEY is not set.', { retryable: false });
    try {
      const result = await clientFactory(apiKey).getGenerativeModel({ model }).generateContent({
        contents: [{ role: 'user', parts }],
        generationConfig: { temperature: 0, responseMimeType: 'application/json', ...(responseSchema ? { responseSchema } : {}) }
      });
      const response = await result.response;
      return typeof response.text === 'function' ? response.text() : '';
    } catch (error) {
      throw classify(error);
    }
  };
  const parse = (raw) => {
    try { return JSON.parse(String(raw).replace(/^```(?:json)?\s*|\s*```$/g, '')); } catch (_error) { throw new ProviderError('BAD_OUTPUT', 'The reader returned something that is not JSON.', { retryable: true }); }
  };
  return {
    mock: false,
    async read({ png, form, model }) {
      const raw = await generate(model, [{ text: buildPrompt(form) }, { inlineData: { mimeType: 'image/png', data: png.toString('base64') } }], SCHEMA);
      const parsed = parse(raw);
      if (!parsed || !Array.isArray(parsed.cells)) throw new ProviderError('BAD_OUTPUT', 'The reader returned no cells.', { retryable: true });
      return parsed;
    },
    // The printed "Sheet ID ER-XXXX-XXXXX" under the page, for a page whose QR code would not scan.
    async readSheetId({ png, model }) {
      const raw = await generate(model, [{ text: 'Find the printed text "Sheet ID" followed by an identifier like ER-ABCD-EFGH2 at the bottom of the page. Return JSON {"sheetId": "<the identifier exactly as printed>"}; return an empty string if you cannot read it. Do not guess.' }, { inlineData: { mimeType: 'image/png', data: png.toString('base64') } }], { type: 'object', properties: { sheetId: { type: 'string' } }, required: ['sheetId'] });
      return String((parse(raw) || {}).sheetId || '');
    }
  };
}

module.exports = { createGeminiProvider, buildPrompt, SCHEMA };
