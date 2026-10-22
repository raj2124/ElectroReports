// Live check of the Microsoft Graph setup (docs/graph-setup.md). Run it ONLY after the registration is done and
// GRAPH_USE_MOCK=false in .env. It prints statuses and counts only: never a token, a secret, an address of
// someone else, a file name or any document text.
//
//   node scripts/graph-check.js token
//   node scripts/graph-check.js calendar --member you@domain --outside someone-not-in-the-group@domain [--write-test]
//   node scripts/graph-check.js sites
//
// calendar: reads one event id of the member (must work), tries the outside mailbox (MUST be 403: that is the
//           proof the Exchange restriction holds), and with --write-test creates one clearly named test event in
//           the member's calendar and deletes it again.
// sites:    for each configured site: resolves it, lists its document libraries, reads the first page of the
//           delta feed (counts only) and shows that a write is refused by the client before anything is sent.
require('dotenv').config();
process.env.GRAPH_USE_MOCK = 'false';
const defaultConfig = require('../src/config');
const { createGraphClient, SharePointReadOnlyError } = require('../src/graphClient');
const { PROPERTY_ID } = require('../src/calendarSync');

const args = process.argv.slice(2);
const mode = args[0];
const option = (name) => { const at = args.indexOf(`--${name}`); return at >= 0 ? args[at + 1] : ''; };
const graph = createGraphClient({ config: defaultConfig.graph });
let failed = false;
const report = (ok, line) => { if (!ok) failed = true; console.log(`${ok ? 'PASS' : 'FAIL'}  ${line}`); };
const statusOf = (error) => (error.status === undefined && /^Microsoft Graph needs/.test(error.message) ? error.message : `${error.status || 'error'}${error.code ? ` ${error.code}` : ''}`);

async function token() {
  try {
    await graph.get('/organization', { params: { $select: 'id', $top: 1 } }).catch((error) => { if (error.status === 403) return null; throw error; });
    report(true, 'a token was obtained and Graph answered (a 403 on /organization is expected: the app has no directory permission)');
  } catch (error) {
    report(false, `no token or Graph unreachable (${statusOf(error)})`);
  }
}

async function calendar() {
  const member = option('member');
  const outside = option('outside');
  if (!member || !outside) { console.log('Give --member <mailbox in the group> and --outside <mailbox NOT in the group>.'); process.exit(2); }
  try {
    const { data } = await graph.get(`/users/${encodeURIComponent(member)}/events`, { params: { $top: 1, $select: 'id' } });
    report(true, `member mailbox readable (${Array.isArray(data.value) ? data.value.length : 0} event id returned)`);
  } catch (error) {
    report(false, `member mailbox NOT readable (${statusOf(error)}): is the member in the group, and has the role assignment propagated (up to 30 minutes)?`);
  }
  try {
    await graph.get(`/users/${encodeURIComponent(outside)}/events`, { params: { $top: 1, $select: 'id' } });
    report(false, 'the OUTSIDE mailbox was readable: the Exchange restriction is NOT in force. Stop and fix section 5 of docs/graph-setup.md before switching anything on.');
  } catch (error) {
    report(error.status === 403, `outside mailbox refused (${statusOf(error)}) ${error.status === 403 ? '- the restriction holds' : '- expected 403'}`);
  }
  if (args.includes('--write-test')) {
    let id = '';
    try {
      const start = new Date(Date.now() + 36 * 3600 * 1000).toISOString().slice(0, 10);
      const end = new Date(Date.now() + 60 * 3600 * 1000).toISOString().slice(0, 10);
      const { data } = await graph.request('POST', `/users/${encodeURIComponent(member)}/events`, { body: { subject: 'ElectroReports connectivity test (safe to delete)', isAllDay: true, showAs: 'free', start: { dateTime: `${start}T00:00:00`, timeZone: defaultConfig.graph.timezone }, end: { dateTime: `${end}T00:00:00`, timeZone: defaultConfig.graph.timezone }, singleValueExtendedProperties: [{ id: PROPERTY_ID, value: 'connectivity-test' }] } });
      id = data.id;
      report(true, 'test event created in the member calendar');
    } catch (error) {
      report(false, `test event could not be created (${statusOf(error)})`);
    }
    if (id) {
      try {
        await graph.request('DELETE', `/users/${encodeURIComponent(member)}/events/${encodeURIComponent(id)}`);
        report(true, 'test event deleted again');
      } catch (error) {
        report(false, `test event could not be deleted (${statusOf(error)}): remove "ElectroReports connectivity test" by hand`);
      }
    }
  }
}

async function sites() {
  for (const site of defaultConfig.graph.kb.sites) {
    const label = `${site.name} (${site.scope})`;
    try {
      const url = new URL(site.url);
      const { data: info } = await graph.get(`/sites/${url.host}:${url.pathname}`);
      const { items: drives } = await graph.getAll(`/sites/${encodeURIComponent(info.id)}/drives`);
      report(true, `${label}: site resolved, ${drives.length} document librar${drives.length === 1 ? 'y' : 'ies'}`);
      const { data: delta } = await graph.get(`/drives/${drives[0].id}/root/delta`, { params: { $top: 5 } });
      report(true, `${label}: delta feed readable (${(delta.value || []).length} items on the first page, ${delta['@odata.nextLink'] ? 'more pages' : 'complete'})`);
    } catch (error) {
      report(false, `${label}: ${statusOf(error)}${error.status === 403 ? ' - the app has no "read" grant on this site yet (section 6)' : ''}`);
    }
  }
  try {
    await graph.request('PATCH', '/drives/any/items/any', { body: { name: 'x' } });
    report(false, 'a write to SharePoint was not refused');
  } catch (error) {
    report(error instanceof SharePointReadOnlyError, 'a write to SharePoint is refused by the client before anything is sent');
  }
}

(async () => {
  if (!['token', 'calendar', 'sites'].includes(mode)) { console.log('Usage: node scripts/graph-check.js token | calendar --member <upn> --outside <upn> [--write-test] | sites'); process.exit(2); }
  await ({ token, calendar, sites })[mode]();
  console.log(failed ? '\nSOME CHECKS FAILED' : '\nALL CHECKS PASSED');
  process.exit(failed ? 1 : 0);
})().catch((error) => { console.error(`check stopped: ${error.status || ''} ${error.code || error.message}`.trim()); process.exit(1); });
