const { GoogleGenerativeAI } = require('@google/generative-ai');

const {
  SOIL_RESISTIVITY_OCR_SCHEMA,
  ELECTRODE_RESISTANCE_OCR_SCHEMA,
  CONTINUITY_TEST_OCR_SCHEMA,
  LOOP_IMPEDANCE_OCR_SCHEMA,
  PROSPECTIVE_FAULT_CURRENT_OCR_SCHEMA,
  RISER_INTEGRITY_OCR_SCHEMA,
  EARTH_CONTINUITY_OCR_SCHEMA,
  TOWER_FOOTING_OCR_SCHEMA
} = require('./ocrSchemas');
const {
  validateSoilOcrPayload,
  validateElectrodeOcrPayload,
  validateContinuityOcrPayload,
  validateLoopImpedanceOcrPayload,
  validateProspectiveFaultOcrPayload,
  validateRiserIntegrityOcrPayload,
  validateEarthContinuityOcrPayload,
  validateTowerFootingOcrPayload
} = require('./ocrValidation');

class OcrWorkflowError extends Error {
  constructor(message, { status = 400, code = 'OCR_WORKFLOW_ERROR', details = null } = {}) {
    super(message);
    this.name = 'OcrWorkflowError';
    this.status = status;
    this.code = code;
    this.details = details;
  }
}

function workflowError(message, status, code, details = null) {
  return new OcrWorkflowError(message, { status, code, details });
}

function ensureGeminiConfigured(config) {
  if (!config?.gemini?.enabled) {
    throw workflowError('Gemini OCR is disabled.', 503, 'OCR_DISABLED');
  }

  if (!String(config?.gemini?.apiKey || '').trim()) {
    throw workflowError('Gemini API key is not configured.', 503, 'OCR_API_KEY_MISSING');
  }
}

function assertSupportedFile(schema, file) {
  if (!file) {
    throw workflowError('Please upload a sheet image or PDF.', 400, 'OCR_FILE_MISSING');
  }

  const mimeType = String(file.mimetype || '').trim().toLowerCase();
  if (!schema.supportedMimeTypes.has(mimeType)) {
    throw workflowError('Unsupported file type. Please upload PDF, PNG, JPG, JPEG, or WEBP.', 415, 'OCR_UNSUPPORTED_FILE');
  }

  if (!file.buffer || !Buffer.isBuffer(file.buffer) || !file.buffer.length) {
    throw workflowError('Uploaded file is empty.', 400, 'OCR_EMPTY_FILE');
  }
}

function extractJsonPayload(rawText) {
  const text = String(rawText || '').trim();
  if (!text) {
    throw workflowError('Gemini returned an empty OCR response.', 502, 'OCR_EMPTY_RESPONSE');
  }

  const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced ? fenced[1].trim() : text;

  try {
    return JSON.parse(candidate);
  } catch (_error) {
    const firstBrace = candidate.indexOf('{');
    const lastBrace = candidate.lastIndexOf('}');
    if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
      try {
        return JSON.parse(candidate.slice(firstBrace, lastBrace + 1));
      } catch (_nestedError) {
        throw workflowError('Gemini OCR response was not valid JSON.', 502, 'OCR_INVALID_JSON');
      }
    }
    throw workflowError('Gemini OCR response was not valid JSON.', 502, 'OCR_INVALID_JSON');
  }
}

function buildPrompt(schema, instructions) {
  return [
    'You are an OCR extraction assistant for the ElectroReports engineering app.',
    instructions.title,
    'Do not calculate categories, averages, statuses, standards decisions, or recommendations.',
    'Do not invent values. If a value is unreadable, leave it blank and add an uncertainFields item.',
    instructions.structure,
    'Preserve row order exactly as seen on the sheet.',
    'Use the exact JSON shape below:',
    JSON.stringify(schema.template, null, 2),
    'Rules:',
    `- sheetType must be "${schema.id}".`,
    ...instructions.rules,
    '- warnings should capture sheet-level concerns like missing pages, cropped photos, rotation, or low confidence.',
    `- uncertainFields should use a path like ${instructions.uncertainFieldExample}.`,
    '- Return JSON only. No markdown fences.'
  ].join('\n');
}

function createModel(config) {
  const client = new GoogleGenerativeAI(String(config.gemini.apiKey || '').trim());
  return client.getGenerativeModel({
    model: String(config.gemini.model || 'gemini-2.5-pro').trim()
  });
}

async function requestPreview(config, file, schema, prompt, validator) {
  ensureGeminiConfigured(config);
  assertSupportedFile(schema, file);
  const model = createModel(config);

  let result;
  try {
    result = await model.generateContent({
      contents: [
        {
          role: 'user',
          parts: [
            { text: prompt },
            {
              inlineData: {
                mimeType: String(file.mimetype || '').trim(),
                data: file.buffer.toString('base64')
              }
            }
          ]
        }
      ],
      generationConfig: {
        temperature: 0.1,
        responseMimeType: 'application/json'
      }
    });
  } catch (error) {
    if (error instanceof OcrWorkflowError) {
      throw error;
    }
    throw workflowError(
      'Gemini OCR request failed. Please try again with a clearer sheet image or PDF.',
      502,
      'OCR_UPSTREAM_FAILURE'
    );
  }

  const response = await result.response;
  const text = typeof response.text === 'function' ? response.text() : '';
  const parsed = extractJsonPayload(text);
  try {
    return validator(parsed);
  } catch (error) {
    if (error instanceof OcrWorkflowError) {
      throw error;
    }
    throw workflowError(
      error instanceof Error ? error.message : 'The uploaded sheet could not be mapped into the expected format.',
      422,
      'OCR_VALIDATION_FAILED'
    );
  }
}

async function previewSoilSheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    SOIL_RESISTIVITY_OCR_SCHEMA,
    buildPrompt(SOIL_RESISTIVITY_OCR_SCHEMA, {
      title: 'The uploaded file is a Soil Resistivity Test sheet. Extract the handwritten or typed measurement data into valid JSON only.',
      structure: 'Keep locations separate if the document includes multiple soil test locations.',
      uncertainFieldExample: 'locations[0].direction1[2].resistivity',
      rules: [
        '- direction1 and direction2 are arrays of rows with spacing and resistivity.',
        '- spacing values should be copied from the sheet, usually 0.5, 1.0, 1.5, 2.0, 2.5, 3.0 and so on.',
        '- drivenElectrodeDiameter and drivenElectrodeLength are optional and should be blank if not present.',
        '- notes may include any visible soil/location note written on the sheet.'
      ]
    }),
    validateSoilOcrPayload
  );
}

async function previewElectrodeSheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    ELECTRODE_RESISTANCE_OCR_SCHEMA,
    buildPrompt(ELECTRODE_RESISTANCE_OCR_SCHEMA, {
      title: 'The uploaded file is an Earth Electrode Resistance Test sheet. Extract the handwritten or typed row data into valid JSON only.',
      structure: 'Keep each electrode/pit reading as one row in the same order as the sheet.',
      uncertainFieldExample: 'rows[0].resistanceWithGrid',
      rules: [
        '- rows is an array of electrode measurement rows.',
        '- tag is the pit tag or pit identifier if visible.',
        '- location is the row location field if visible.',
        '- electrodeType and materialType should be copied from the sheet when visible; otherwise use common defaults already shown in the template.',
        '- resistanceWithoutGrid and resistanceWithGrid should be copied exactly as written.',
        '- leave optional fields blank if they are not shown on the sheet.'
      ]
    }),
    validateElectrodeOcrPayload
  );
}

async function previewContinuitySheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    CONTINUITY_TEST_OCR_SCHEMA,
    buildPrompt(CONTINUITY_TEST_OCR_SCHEMA, {
      title: 'The uploaded file is a Continuity Test sheet. Extract the handwritten or typed row data into valid JSON only.',
      structure: 'Keep each continuity reading as one row in the same order as the sheet.',
      uncertainFieldExample: 'rows[0].measurementPoint',
      rules: [
        '- rows is an array of continuity measurement rows.',
        '- srNo should be copied if visible; otherwise preserve row order using 1, 2, 3 and so on.',
        '- mainLocation, measurementPoint, resistance, and impedance should be copied exactly as shown.',
        '- do not generate comment or status text.'
      ]
    }),
    validateContinuityOcrPayload
  );
}

async function previewLoopImpedanceSheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    LOOP_IMPEDANCE_OCR_SCHEMA,
    buildPrompt(LOOP_IMPEDANCE_OCR_SCHEMA, {
      title: 'The uploaded file is a Loop Impedance Test sheet. Extract the handwritten or typed feeder-group data into valid JSON only.',
      structure: 'Keep one feeder or panel block as one group. Each group should contain the shared protective-device fields and the R-E, Y-E, B-E measurement rows.',
      uncertainFieldExample: 'groups[0].rows[2].loopImpedance',
      rules: [
        '- groups is an array of feeder groups.',
        '- each group should preserve shared values like location, feederTag, deviceType, deviceRating, and breakingCapacity.',
        '- rows should preserve the measured point labels exactly as R-E, Y-E, and B-E where possible.',
        '- copy loopImpedance and voltage exactly as visible on the sheet.',
        '- do not generate comments or status text.'
      ]
    }),
    validateLoopImpedanceOcrPayload
  );
}

async function previewProspectiveFaultSheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    PROSPECTIVE_FAULT_CURRENT_OCR_SCHEMA,
    buildPrompt(PROSPECTIVE_FAULT_CURRENT_OCR_SCHEMA, {
      title: 'The uploaded file is a Prospective Fault Current Test sheet. Extract the handwritten or typed feeder-group data into valid JSON only.',
      structure: 'Keep one feeder or panel block as one group. Each group should contain the shared protective-device fields and the R-E, Y-E, B-E measurement rows.',
      uncertainFieldExample: 'groups[0].rows[1].prospectiveFaultCurrent',
      rules: [
        '- groups is an array of feeder groups.',
        '- each group should preserve shared values like location, feederTag, deviceType, deviceRating, and breakingCapacity.',
        '- rows should preserve the measured point labels exactly as R-E, Y-E, and B-E where possible.',
        '- copy loopImpedance, prospectiveFaultCurrent, and voltage exactly as visible on the sheet.',
        '- do not generate comments or status text.'
      ]
    }),
    validateProspectiveFaultOcrPayload
  );
}

async function previewRiserIntegritySheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    RISER_INTEGRITY_OCR_SCHEMA,
    buildPrompt(RISER_INTEGRITY_OCR_SCHEMA, {
      title: 'The uploaded file is a Riser / Grid Integrity Test sheet. Extract the handwritten or typed row data into valid JSON only.',
      structure: 'Keep each riser integrity reading as one row in the same order as the sheet.',
      uncertainFieldExample: 'rows[0].measurementPoint',
      rules: [
        '- rows is an array of riser integrity measurement rows.',
        '- srNo should be copied if visible; otherwise preserve row order using 1, 2, 3 and so on.',
        '- mainLocation, measurementPoint, resistanceTowardsEquipment, and resistanceTowardsGrid should be copied exactly as shown.',
        '- do not generate comment or status text.'
      ]
    }),
    validateRiserIntegrityOcrPayload
  );
}

async function previewEarthContinuitySheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    EARTH_CONTINUITY_OCR_SCHEMA,
    buildPrompt(EARTH_CONTINUITY_OCR_SCHEMA, {
      title: 'The uploaded file is an Earth Continuity Test sheet. Extract the handwritten or typed row data into valid JSON only.',
      structure: 'Keep each earth continuity reading as one row in the same order as the sheet.',
      uncertainFieldExample: 'rows[0].measuredValue',
      rules: [
        '- rows is an array of earth continuity measurement rows.',
        '- srNo should be copied if visible; otherwise preserve row order using 1, 2, 3 and so on.',
        '- tag, locationBuildingName, distance, and measuredValue should be copied exactly as shown.',
        '- do not generate remark or status text.'
      ]
    }),
    validateEarthContinuityOcrPayload
  );
}

async function previewTowerFootingSheetWithGemini(config, file) {
  return requestPreview(
    config,
    file,
    TOWER_FOOTING_OCR_SCHEMA,
    buildPrompt(TOWER_FOOTING_OCR_SCHEMA, {
      title: 'The uploaded file is a Tower Footing Resistance Measurement & Analysis sheet. Extract the grouped tower footing data into valid JSON only.',
      structure: 'Keep one main tower location as one group. Each group must contain Foot-1, Foot-2, Foot-3, and Foot-4 readings in order.',
      uncertainFieldExample: 'groups[0].readings[2].measuredImpedance',
      rules: [
        '- groups is an array of tower location groups.',
        '- each group should preserve srNo and mainLocationTower if visible.',
        '- readings should preserve Foot-1 through Foot-4 in order.',
        '- copy footToEarthingConnectionStatus, measuredCurrentMa, and measuredImpedance exactly as shown.',
        '- do not calculate totals or remarks.'
      ]
    }),
    validateTowerFootingOcrPayload
  );
}

module.exports = {
  OcrWorkflowError,
  previewSoilSheetWithGemini,
  previewElectrodeSheetWithGemini,
  previewContinuitySheetWithGemini,
  previewLoopImpedanceSheetWithGemini,
  previewProspectiveFaultSheetWithGemini,
  previewRiserIntegritySheetWithGemini,
  previewEarthContinuitySheetWithGemini,
  previewTowerFootingSheetWithGemini
};
