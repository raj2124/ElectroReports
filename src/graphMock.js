// An in-memory Microsoft Graph for GRAPH_USE_MOCK and the tests: an `http` double that createGraphClient runs
// on, so the real client (auth, retry, the SharePoint read-only guard, paging) is what the code under test
// uses. It holds Outlook calendars (events with extended properties, a mailbox group that gates access) and
// SharePoint sites (document libraries with delta, file content). Nothing here talks to a network.

const { makePdf, makeDocx, makeXlsx } = require('./graphFixtures');

const BASE = 'https://graph.microsoft.com/v1.0';
const fail = (status, code) => Object.assign(new Error(`mock ${status} ${code}`), { response: { status, data: { error: { code, message: 'mock' } }, headers: {} } });

function createMockGraph({ group = [], pageSize = 100 } = {}) {
  const state = {
    group: new Set(group.map((upn) => upn.toLowerCase())),
    mailboxes: new Set(group.map((upn) => upn.toLowerCase())),
    events: new Map(),            // upn -> [event]
    seq: 0,
    pageSize,
    sites: new Map(),             // site name -> { name, url, granted, files: Map(path -> file) }
    version: 1,
    expired: false,
    log: [],                      // [method, path]: every request, to assert on
    sharePointWrites: 0
  };

  const mailbox = (upn) => {
    const key = String(upn).toLowerCase();
    if (!state.mailboxes.has(key) && !state.group.has(key)) throw fail(404, 'ErrorInvalidUser');
    if (!state.group.has(key)) throw fail(403, 'ErrorAccessDenied');
    if (!state.events.has(key)) state.events.set(key, []);
    return state.events.get(key);
  };

  // ---- SharePoint model ------------------------------------------------------------------------
  const siteByUrl = (host, rel) => [...state.sites.values()].find((site) => site.url.toLowerCase() === `https://${host}${rel}`.toLowerCase());
  const siteById = (id) => [...state.sites.values()].find((site) => site.id === id);
  const driveOf = (id) => [...state.sites.values()].find((site) => site.driveId === id);
  const toItem = (site, file) => {
    const parts = file.path.split('/');
    return {
      id: file.id, name: parts[parts.length - 1], size: file.size, file: { mimeType: 'application/octet-stream' }, eTag: file.etag,
      lastModifiedDateTime: file.modified, webUrl: `${site.url}/${parts.map(encodeURIComponent).join('/')}`,
      parentReference: { driveId: site.driveId, path: `/drives/${site.driveId}/root:${parts.length > 2 ? `/${parts.slice(1, -1).join('/')}` : ''}` }
    };
  };
  const sp = {
    addSite(name, url, { granted = true } = {}) {
      const id = `${new URL(url).host},coll-${name},web-${name}`;
      state.sites.set(name, { name, url, id, driveId: `drive-${name}`, granted, files: new Map() });
    },
    // path starts with the library ("Shared Documents/Folder/file.pdf"); content is a Buffer or a function returning one
    addFile(siteName, path, content, { modified = '2026-10-01T09:00:00Z' } = {}) {
      const site = state.sites.get(siteName);
      const previous = site.files.get(path);
      state.version += 1;
      site.files.set(path, { id: previous ? previous.id : `item-${++state.seq}`, path, content, modified, etag: `e${state.version}`, size: Buffer.isBuffer(content) ? content.length : 1000, version: state.version, deleted: false });
    },
    removeFile(siteName, path) {
      const site = state.sites.get(siteName);
      const file = site.files.get(path);
      state.version += 1;
      site.files.set(path, { ...file, deleted: true, version: state.version });
    },
    moveFile(siteName, from, to) {
      const site = state.sites.get(siteName);
      const file = site.files.get(from);
      state.version += 1;
      site.files.delete(from);
      site.files.set(to, { ...file, path: to, version: state.version });
    },
    expireDelta() { state.expired = true; }
  };

  // ---- the http double -----------------------------------------------------------------------------
  async function request(config) {
    const method = String(config.method || 'GET').toUpperCase();
    const url = new URL(config.url);
    const path = decodeURIComponent(url.pathname.replace(/^\/v1\.0/, ''));
    const query = { ...Object.fromEntries(url.searchParams), ...(config.params || {}) };
    state.log.push([method, path]);
    let body = config.data;
    if (typeof body === 'string') body = JSON.parse(body);

    let match;
    // ---- calendar
    if ((match = /^\/users\/([^/]+)\/events(?:\/([^/]+))?$/.exec(path))) {
      const events = mailbox(match[1]);
      const id = match[2];
      const own = (event) => (event.singleValueExtendedProperties || []).length > 0;
      if (!id && method === 'GET') {
        const wanted = /ep\/id eq '([^']+)'/.exec(String(query.$filter || ''));
        const hits = events.filter((event) => !wanted || (event.singleValueExtendedProperties || []).some((prop) => prop.id === wanted[1]));
        const skip = Number(query.$skip || 0);
        const page = hits.slice(skip, skip + state.pageSize);
        const next = skip + state.pageSize < hits.length ? `${BASE}/users/${match[1]}/events?${new URLSearchParams({ ...(query.$filter ? { $filter: query.$filter } : {}), $skip: String(skip + state.pageSize) })}` : undefined;
        return { status: 200, data: { value: page.map((event) => JSON.parse(JSON.stringify(event))), ...(next ? { '@odata.nextLink': next } : {}) }, headers: {} };
      }
      if (!id && method === 'POST') {
        const event = { ...JSON.parse(JSON.stringify(body)), id: `EVT-${++state.seq}` };
        events.push(event);
        return { status: 201, data: JSON.parse(JSON.stringify(event)), headers: {} };
      }
      const event = events.find((entry) => entry.id === id);
      if (!event) throw fail(404, 'ErrorItemNotFound');
      if (method === 'GET') return { status: 200, data: JSON.parse(JSON.stringify(event)), headers: {} };
      if (method === 'PATCH') { Object.assign(event, JSON.parse(JSON.stringify(body))); return { status: 200, data: JSON.parse(JSON.stringify(event)), headers: {} }; }
      if (method === 'DELETE') { events.splice(events.indexOf(event), 1); return { status: 204, data: '', headers: {} }; }
      void own;
    }
    // ---- SharePoint (read only)
    if (/^\/(sites|drives)(\/|:|$)/.test(path)) {
      if (method !== 'GET') { state.sharePointWrites += 1; throw fail(405, 'methodNotAllowed'); }
      if ((match = /^\/sites\/([^:/]+):(\/.+)$/.exec(path))) {
        const site = siteByUrl(match[1], match[2]);
        if (!site) throw fail(404, 'itemNotFound');
        if (!site.granted) throw fail(403, 'accessDenied');
        return { status: 200, data: { id: site.id, name: site.name, webUrl: site.url }, headers: {} };
      }
      if ((match = /^\/sites\/([^/]+)\/drives$/.exec(path))) {
        const site = siteById(match[1]);
        if (!site || !site.granted) throw fail(site ? 403 : 404, site ? 'accessDenied' : 'itemNotFound');
        return { status: 200, data: { value: [{ id: site.driveId, name: 'Documents', webUrl: `${site.url}/Shared%20Documents` }] }, headers: {} };
      }
      if ((match = /^\/drives\/([^/]+)\/root\/delta$/.exec(path))) {
        const site = driveOf(match[1]);
        if (!site || !site.granted) throw fail(site ? 403 : 404, site ? 'accessDenied' : 'itemNotFound');
        const since = query.token ? Number(query.token) : 0;
        if (query.token && state.expired) throw fail(410, 'resyncRequired');
        const changed = [...site.files.values()].filter((file) => file.version > since && !(since === 0 && file.deleted));
        const skip = Number(query.$skiptoken || 0);
        const page = changed.slice(skip, skip + state.pageSize);
        const items = page.map((file) => (file.deleted ? { id: file.id, deleted: { state: 'deleted' } } : toItem(site, file)));
        if (!skip) items.unshift({ id: `folder-${site.name}`, name: 'Shared Documents', folder: { childCount: 0 }, parentReference: { driveId: site.driveId, path: `/drives/${site.driveId}/root:` } });
        const done = skip + state.pageSize >= changed.length;
        return { status: 200, data: { value: items, ...(done ? { '@odata.deltaLink': `${BASE}/drives/${site.driveId}/root/delta?token=${state.version}` } : { '@odata.nextLink': `${BASE}/drives/${site.driveId}/root/delta?${new URLSearchParams({ ...(query.token ? { token: query.token } : {}), $skiptoken: String(skip + state.pageSize) })}` }) }, headers: {} };
      }
      if ((match = /^\/drives\/([^/]+)\/items\/([^/]+)\/content$/.exec(path))) {
        const site = driveOf(match[1]);
        const file = site && [...site.files.values()].find((entry) => entry.id === match[2] && !entry.deleted);
        if (!file) throw fail(404, 'itemNotFound');
        const content = typeof file.content === 'function' ? await file.content() : file.content;
        return { status: 200, data: content, headers: {} };
      }
    }
    throw fail(404, 'mockRouteNotFound');
  }

  const http = {
    post: async (url) => {
      if (/\/oauth2\/v2\.0\/token$/.test(url)) return { data: { access_token: 'mock-token', expires_in: 3600 } };
      throw fail(404, 'mockRouteNotFound');
    },
    request
  };

  const calendar = {
    events: (upn) => (state.events.get(String(upn).toLowerCase()) || []),
    addForeignEvent(upn, event = { subject: 'Dentist' }) { mailbox(upn).push({ ...event, id: `EVT-${++state.seq}` }); },
    addMember(upn) { state.group.add(upn.toLowerCase()); state.mailboxes.add(upn.toLowerCase()); },
    addOutsideMailbox(upn) { state.mailboxes.add(upn.toLowerCase()); }
  };

  return { http, state, sp, calendar };
}

// Demo content for the running app in mock mode: the people of the mock Zoho portal in the scoped group, and the
// three configured SharePoint sites with a few real files.
function seedDemo(mock, config) {
  ['anup', 'karishma', 'nidhi', 'mohana'].forEach((name) => mock.calendar.addMember(`${name}@elegrow.com`));
  mock.calendar.addOutsideMailbox('director@elegrow.com');
  config.kb.sites.forEach((site) => mock.sp.addSite(site.name, site.url));
  const general = config.kb.sites.find((site) => site.scope === 'general');
  const clients = config.kb.sites.filter((site) => site.scope === 'client');
  if (general) {
    const base = `${general.folders[0] || 'Shared Documents'}`;
    mock.sp.addFile(general.name, `${base}/Earthing practice notes.pdf`, () => makePdf(['Earth electrode resistance is verified against the limit calculated from site soil resistivity.', 'Pit resistance readings are taken with the earth analyser after the pit water is stabilised.']));
    mock.sp.addFile(general.name, 'Shared Documents/Other/Not indexed.pdf', () => makePdf('This file is outside the configured folder.'));
  }
  if (clients[0]) {
    mock.sp.addFile(clients[0].name, 'Shared Documents/024-D03_TS-26-GFL/Submitted/GFL ESHA report.pdf', () => makePdf('GFL Dahej earthing system health assessment: pit resistance and continuity results.'));
    mock.sp.addFile(clients[0].name, 'Shared Documents/024-D03_TS-26-GFL/Working/GFL draft.docx', () => makeDocx('Draft notes that must not be indexed.'));
    mock.sp.addFile(clients[0].name, 'Shared Documents/056-D03_TS-26-CLL/Input/CLL site data.xlsx', () => makeXlsx([['Pit', 'Location'], ['P-1', 'Substation yard']]));
  }
  if (clients[1]) {
    mock.sp.addFile(clients[1].name, 'Shared Documents/294_04_23_TS_APL_ESHA/Reference/APL reference.pdf', () => makePdf('APL earlier assessment: soil resistivity and pit resistance reference values.'));
  }
}

module.exports = { createMockGraph, seedDemo, BASE };
