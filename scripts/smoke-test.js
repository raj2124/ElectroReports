const fs = require('fs');
const path = require('path');
const axios = require('axios');
const config = require('../src/config');

const BASE_URL = process.env.SMOKE_BASE_URL || config.app.baseUrl || 'http://127.0.0.1:3000';
const PROJECT_QUERY = process.env.SMOKE_PROJECT_QUERY || '547_02_06_ED_TIPL_LPAS';

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

async function get(url) {
  return axios.get(url, { timeout: 60000 });
}

async function post(url, body) {
  return axios.post(url, body, {
    timeout: 60000,
    headers: { 'Content-Type': 'application/json' }
  });
}

function buildTestPayload(project) {
  const projectName = String(project?.name || project?.projectNumber || 'SMOKE-PROJECT').trim();
  const projectNo = String(project?.projectNumber || projectName).trim();

  return {
    mom: {
      meetingTitle: 'Smoke Test - M.O.M Submission',
      projectName,
      projectNoWorkOrderNo: projectNo,
      clientName: 'Smoke Test Client',
      meetingDate: '2026-02-26',
      meetingTime: '12:30',
      entryTime: '12:25',
      exitTime: '13:10',
      meetingLocation: 'Surat Office',
      meetingCalledBy: 'Elegrow Technology Pvt. Ltd.',
      meetingType: ['Kick-off', 'Status/Review'],
      meetingTypeOther: '',
      facilitatorRepresentative: 'Ami Shah',
      elegrowRepresentative: 'Jignesh Sailor',
      clientRepresentative: 'Client Representative',
      agendaRows: [
        {
          srNo: '1',
          agenda: 'Progress review',
          actionPlan: 'Confirm timelines',
          responsibility: 'Project Manager'
        }
      ],
      attendeeRows: [
        {
          srNo: '1',
          elegrowName: 'Ami Shah',
          clientName: 'Client Representative'
        }
      ],
      organizationAddress:
        '302, Sangini Aspire, Beside Sanskruti Township Near Pal RTO, Pal-Hajira Road, Pal Gam, Surat, Gujarat - 395009'
    },
    options: {
      generatePdf: true,
      printPdf: true,
      sendEmail: true,
      emailTo: '',
      emailCc: '',
      emailSubject: '',
      emailBody: ''
    }
  };
}

async function main() {
  console.log(`[smoke] Base URL: ${BASE_URL}`);

  const health = await get(`${BASE_URL}/api/health`);
  assert(health.data?.success === true, 'Health endpoint failed.');
  console.log(`[smoke] Health OK | zohoMode=${health.data.zohoMode} | emailMode=${health.data.emailMode || 'legacy'}`);

  const projectResp = await get(`${BASE_URL}/api/zoho/projects?query=${encodeURIComponent(PROJECT_QUERY)}`);
  assert(projectResp.data?.success === true, 'Project fetch failed.');
  const projects = Array.isArray(projectResp.data.projects) ? projectResp.data.projects : [];
  assert(projects.length > 0, `No projects found for query: ${PROJECT_QUERY}`);

  const project = projects[0];
  console.log(`[smoke] Project OK | name=${project.name || '-'} | stage=${project.stage || '-'} | id=${project.id || '-'}`);

  if (project.id) {
    const usersResp = await get(`${BASE_URL}/api/zoho/projects/${encodeURIComponent(project.id)}/users`);
    assert(usersResp.data?.success === true, 'Project users fetch failed.');
    const users = Array.isArray(usersResp.data.users) ? usersResp.data.users : [];
    console.log(`[smoke] Users OK | count=${users.length}`);
  }

  const submitPayload = buildTestPayload(project);
  const submitResp = await post(`${BASE_URL}/api/mom/submit`, submitPayload);
  assert(submitResp.data?.success === true, 'M.O.M submit failed.');

  const result = submitResp.data.result || {};
  assert(result.pdfUrl, 'pdfUrl missing in submit result.');
  assert(result.printRequested === true, 'printRequested expected true.');
  assert(result.emailSent === false, 'emailSent expected false in Outlook draft mode.');
  const draftMode = String(result.emailDraft?.mode || '');
  assert(
    draftMode === 'outlook-draft' || draftMode === 'graph-draft',
    'emailDraft.mode expected outlook-draft or graph-draft.'
  );
  if (draftMode === 'graph-draft') {
    assert(result.emailDraft?.outlookDraftWebUrl || result.emailDraft?.outlookComposeUrl, 'Graph draft URL missing.');
  } else {
    assert(result.emailDraft?.outlookComposeUrl, 'Outlook compose URL missing.');
  }

  const pdfAbsoluteUrl = new URL(result.pdfUrl, BASE_URL).toString();
  const pdfResp = await get(pdfAbsoluteUrl);
  assert(pdfResp.status === 200, 'Generated PDF URL not reachable.');
  console.log(`[smoke] Submit OK | pdf=${result.pdfFileName} | print=${result.printRequested} | draft=ready`);

  const localPdfPath = path.join(config.app.generatedDir, result.pdfFileName || '');
  if (result.pdfFileName && fs.existsSync(localPdfPath)) {
    fs.unlinkSync(localPdfPath);
    console.log(`[smoke] Cleanup OK | removed ${result.pdfFileName}`);
  } else {
    console.log('[smoke] Cleanup skipped | test PDF not found locally');
  }

  console.log('[smoke] PASS');
}

main().catch((error) => {
  console.error(`[smoke] FAIL: ${error.message}`);
  process.exit(1);
});
