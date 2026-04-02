const path = require('path');

const { buildPayload } = require('./electroreports-verify');
const { normalizeReportInput } = require('../src/reportModel');
const { generateReportNarrative, isAiConfigured } = require('../src/aiService');
const { generateElectroReportPdf } = require('../src/pdfService');

async function main() {
  if (!isAiConfigured()) {
    throw new Error('OpenAI is not configured. Add OPENAI_API_KEY in .env before generating the AI reference PDF.');
  }

  const payload = buildPayload();
  payload.project.projectNo = 'ER-AI-REFERENCE-20260324';
  payload.project.clientName = 'Reference Client';
  payload.project.siteLocation = 'Elegrow Reference Yard';
  payload.project.workOrder = 'WO-AI-REF-001';
  payload.project.reportDate = '2026-03-24';
  payload.project.engineerName = 'Elegrow Reference Engineer';
  payload.project.equipmentSelections = ['mi3152', 'mi3290', 'kyoritsu4118a'];

  const report = normalizeReportInput(payload);
  const aiNarrative = await generateReportNarrative(report);
  const reportWithNarrative = {
    ...report,
    aiNarrative
  };

  const pdf = await generateElectroReportPdf(reportWithNarrative, {
    generatedDir: path.join(__dirname, '..', 'generated-pdfs'),
    publicPathPrefix: '/generated-pdfs'
  });

  console.log(`[electroreports] AI reference PDF generated`);
  console.log(pdf.filePath);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(`[electroreports] AI reference PDF generation failed: ${error.message}`);
    process.exit(1);
  });
}
