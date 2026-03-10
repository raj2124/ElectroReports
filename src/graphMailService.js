const fs = require('fs');
const axios = require('axios');

let cachedToken = {
  value: '',
  expiresAt: 0
};

function isGraphDraftConfigured(config) {
  const graph = config?.graph || {};
  return Boolean(
    graph.enabled &&
      graph.tenantId &&
      graph.clientId &&
      graph.clientSecret &&
      graph.mailboxUser
  );
}

function normalizeBaseUrl(url) {
  return String(url || '').replace(/\/+$/, '');
}

function parseAddressList(value) {
  return String(value || '')
    .split(/[;,]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function buildRecipients(addresses = []) {
  return addresses.map((address) => ({
    emailAddress: { address }
  }));
}

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function textToHtmlParagraphs(text) {
  const safe = escapeHtml(text);
  return safe
    .split(/\r?\n\r?\n/)
    .map((segment) => `<p>${segment.replace(/\r?\n/g, '<br>')}</p>`)
    .join('');
}

function buildGraphBodyHtml({ bodyText = '', pdfUrl = '' }) {
  const safePdfUrl = escapeHtml(pdfUrl);
  const clickableLink = safePdfUrl
    ? `<p><a href="${safePdfUrl}" target="_blank" rel="noopener noreferrer">${safePdfUrl}</a></p>`
    : '';
  return `${textToHtmlParagraphs(bodyText)}${clickableLink}`;
}

function resolveGraphErrorMessage(error) {
  const status = error?.response?.status;
  const graphMessage = error?.response?.data?.error?.message;
  if (status && graphMessage) {
    return `Microsoft Graph request failed (${status}): ${graphMessage}`;
  }
  if (status) {
    return `Microsoft Graph request failed (${status}).`;
  }
  return error?.message || 'Microsoft Graph request failed.';
}

async function getGraphAccessToken(config) {
  const now = Date.now();
  if (cachedToken.value && cachedToken.expiresAt - 60_000 > now) {
    return cachedToken.value;
  }

  const graph = config.graph;
  const authorityHost = normalizeBaseUrl(graph.authorityHost);
  const tokenUrl = `${authorityHost}/${encodeURIComponent(graph.tenantId)}/oauth2/v2.0/token`;
  const form = new URLSearchParams();
  form.set('client_id', graph.clientId);
  form.set('client_secret', graph.clientSecret);
  form.set('scope', graph.scope || 'https://graph.microsoft.com/.default');
  form.set('grant_type', 'client_credentials');

  const response = await axios.post(tokenUrl, form.toString(), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    timeout: 30000
  });

  const token = String(response?.data?.access_token || '').trim();
  const expiresIn = Number.parseInt(String(response?.data?.expires_in || '3600'), 10);
  if (!token) {
    throw new Error('Microsoft Graph token response did not include access_token.');
  }

  cachedToken = {
    value: token,
    expiresAt: now + Math.max(300, Number.isNaN(expiresIn) ? 3600 : expiresIn) * 1000
  };

  return token;
}

async function createGraphDraft(config, payload) {
  const token = await getGraphAccessToken(config);
  const graphBaseUrl = normalizeBaseUrl(config.graph.baseUrl || 'https://graph.microsoft.com/v1.0');
  const mailboxUser = encodeURIComponent(config.graph.mailboxUser);

  const toRecipients = buildRecipients(parseAddressList(payload.to));
  const ccRecipients = buildRecipients(parseAddressList(payload.cc));
  const requestBody = {
    subject: String(payload.subject || '').trim(),
    body: {
      contentType: 'HTML',
      content: String(payload.bodyHtml || '').trim()
    },
    toRecipients,
    ccRecipients
  };

  let draftResponse;
  try {
    draftResponse = await axios.post(
      `${graphBaseUrl}/users/${mailboxUser}/messages`,
      requestBody,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json'
        },
        timeout: 45000
      }
    );
  } catch (error) {
    throw new Error(resolveGraphErrorMessage(error));
  }

  const draftId = String(draftResponse?.data?.id || '').trim();
  const webLink = String(draftResponse?.data?.webLink || '').trim();

  if (!draftId) {
    throw new Error('Microsoft Graph draft creation succeeded but draft ID is missing.');
  }

  let attachmentNote = '';
  if (payload.pdfFilePath && fs.existsSync(payload.pdfFilePath)) {
    try {
      const bytes = fs.readFileSync(payload.pdfFilePath);
      const contentBytes = bytes.toString('base64');
      const attachmentPayload = {
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: String(payload.pdfFileName || 'Minutes-of-Meeting.pdf'),
        contentType: 'application/pdf',
        contentBytes
      };

      await axios.post(
        `${graphBaseUrl}/users/${mailboxUser}/messages/${encodeURIComponent(draftId)}/attachments`,
        attachmentPayload,
        {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json'
          },
          timeout: 45000
        }
      );
    } catch (error) {
      attachmentNote = resolveGraphErrorMessage(error);
    }
  } else {
    attachmentNote = 'Generated PDF file was not found on server; attachment step skipped.';
  }

  return {
    mode: 'graph-draft',
    draftId,
    outlookDraftWebUrl: webLink,
    attachmentAutoSupported: true,
    attachmentAdded: !attachmentNote,
    attachmentNote
  };
}

module.exports = {
  isGraphDraftConfigured,
  buildGraphBodyHtml,
  createGraphDraft
};
