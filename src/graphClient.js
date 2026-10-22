// Microsoft Graph client (Session 6H.1): app-only (client credentials) with a certificate or a secret,
// retry with Retry-After on throttling, and errors that carry the status and Graph's error code only (a Graph
// message can contain a user's address, and a token must never reach a log).
//
// HARD RULE: SharePoint is read-only. Any request to /sites or /drives that is not a GET is refused here,
// before it leaves the process (scripts/test-graph.js also scans the knowledge-base code for write calls).
// The only writes this app makes through Graph are events in the Outlook calendars of the scoped group.

const crypto = require('crypto');
const fs = require('fs');
const axios = require('axios');
const defaultConfig = require('./config');

const b64url = (input) => Buffer.from(input).toString('base64url');
const SHAREPOINT_PATH = /^\/(sites|drives)(\/|:|$)/i;
const RETRY_STATUSES = new Set([429, 502, 503, 504]);

class GraphError extends Error {
  constructor(status, code) {
    super(`Microsoft Graph request failed (${status}${code ? `: ${code}` : ''})`);
    this.name = 'GraphError';
    this.status = status;
    this.code = code || '';
  }
}

class GraphConfigError extends Error {}

class SharePointReadOnlyError extends Error {
  constructor(method) {
    super(`SharePoint is read-only: ${method} was refused.`);
    this.name = 'SharePointReadOnlyError';
  }
}

// The path part of a Graph URL (relative or absolute), without the version.
function graphPath(url, baseUrl) {
  const text = String(url);
  const relative = text.startsWith(baseUrl) ? text.slice(baseUrl.length) : text.replace(/^https:\/\/graph\.microsoft\.com\/(v1\.0|beta)/, '');
  return relative.startsWith('/') ? relative : `/${relative}`;
}

function createGraphClient({ http = axios, config = defaultConfig.graph, now = () => Date.now(), sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms)) } = {}) {
  let token = { value: '', expiresAt: 0 };
  let inFlight = null;
  const calls = { read: 0, write: 0 };

  function configured() {
    return Boolean(config.tenantId && config.clientId && (config.clientSecret || (config.certPath && config.certThumbprintSha256)));
  }

  // The body of the token request: a signed assertion (certificate) or the secret.
  function credential(tokenUrl) {
    if (config.certPath && config.certThumbprintSha256) {
      const header = { alg: 'RS256', typ: 'JWT', 'x5t#S256': b64url(Buffer.from(config.certThumbprintSha256, 'hex')) };
      const issued = Math.floor(now() / 1000);
      const payload = { aud: tokenUrl, iss: config.clientId, sub: config.clientId, jti: crypto.randomUUID(), nbf: issued, exp: issued + 600 };
      const unsigned = `${b64url(JSON.stringify(header))}.${b64url(JSON.stringify(payload))}`;
      const signature = crypto.createSign('RSA-SHA256').update(unsigned).sign(fs.readFileSync(config.certPath));
      return { client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer', client_assertion: `${unsigned}.${signature.toString('base64url')}` };
    }
    return { client_secret: config.clientSecret };
  }

  async function accessToken(force = false) {
    if (!force && token.value && now() < token.expiresAt) return token.value;
    if (!configured()) throw new GraphConfigError('Microsoft Graph needs GRAPH_TENANT_ID, GRAPH_CLIENT_ID and a certificate (GRAPH_CERT_PATH + GRAPH_CERT_THUMBPRINT_SHA256) or GRAPH_CLIENT_SECRET.');
    if (!inFlight) {
      inFlight = (async () => {
        const tokenUrl = `${config.authority}/${config.tenantId}/oauth2/v2.0/token`;
        const body = new URLSearchParams({ client_id: config.clientId, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials', ...credential(tokenUrl) });
        try {
          const response = await http.post(tokenUrl, body.toString(), { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, timeout: 15000 });
          const value = response && response.data && response.data.access_token;
          if (!value) throw new Error('no token');
          token = { value, expiresAt: now() + Math.max(60, Number(response.data.expires_in || 3600) - 300) * 1000 };
          return value;
        } catch (error) {
          const status = error && error.response && error.response.status;
          const code = error && error.response && error.response.data && error.response.data.error;
          throw new GraphError(status || 0, typeof code === 'string' ? code : 'token_request_failed');
        }
      })().finally(() => { inFlight = null; });
    }
    return inFlight;
  }

  // method, a path ("/users/x/events") or a full URL (an @odata.nextLink), { params, body, headers, binary }.
  // Resolves { status, data, headers }; rejects with GraphError.
  async function request(method, url, { params, body, headers = {}, binary = false } = {}) {
    const verb = String(method).toUpperCase();
    const path = graphPath(url, config.baseUrl);
    if (SHAREPOINT_PATH.test(path) && verb !== 'GET') throw new SharePointReadOnlyError(verb);
    const target = /^https?:/i.test(String(url)) ? String(url) : `${config.baseUrl}${path}`;
    let refreshed = false;
    for (let attempt = 0; ; attempt += 1) {
      try {
        calls[verb === 'GET' ? 'read' : 'write'] += 1;
        const response = await http.request({
          method: verb,
          url: target,
          params,
          data: body,
          headers: { Authorization: `Bearer ${await accessToken()}`, ...(body !== undefined ? { 'Content-Type': 'application/json' } : {}), ...headers },
          responseType: binary ? 'arraybuffer' : 'json',
          timeout: 60000
        });
        return { status: response.status, data: response.data, headers: response.headers || {} };
      } catch (error) {
        if (error instanceof GraphError || error instanceof GraphConfigError || error instanceof SharePointReadOnlyError) throw error;
        const status = error && error.response && error.response.status;
        if (status === 401 && !refreshed) {
          refreshed = true;
          token = { value: '', expiresAt: 0 };
          continue;
        }
        if (RETRY_STATUSES.has(status) && attempt < 3) {
          const wait = Number.parseInt((error.response.headers || {})['retry-after'], 10);
          await sleep(Math.min(60, Number.isFinite(wait) ? wait : 2 ** attempt) * 1000);
          continue;
        }
        const detail = error && error.response && error.response.data;
        const code = detail && detail.error && typeof detail.error === 'object' ? detail.error.code : '';
        throw new GraphError(status || 0, code || (status ? '' : 'network_error'));
      }
    }
  }

  const get = (url, options) => request('GET', url, options);

  // Every page of a collection (follows @odata.nextLink). Returns { items, deltaLink }.
  async function getAll(url, options) {
    const items = [];
    let next = url;
    let deltaLink = '';
    for (let page = 0; next && page < 1000; page += 1) {
      const { data } = await get(next, page === 0 ? options : undefined);
      items.push(...(Array.isArray(data.value) ? data.value : []));
      deltaLink = data['@odata.deltaLink'] || deltaLink;
      next = data['@odata.nextLink'] || '';
    }
    return { items, deltaLink };
  }

  return { request, get, getAll, configured, calls };
}

module.exports = { createGraphClient, GraphError, SharePointReadOnlyError, SHAREPOINT_PATH, graphPath };
