// Zoho People. Session 6F is read only: the employee list, approved / pending leave and the holiday
// list (scripts/test-schedule.js counts the calls: only GETs and the OAuth token refresh). Session 6G adds
// the one write the app makes to People: records of the Site Work Log form (insertRecord / updateRecord),
// used only by src/peopleSync.js. The refresh token is the one of the Projects client
// (one Self Client grant carries the Projects and People scopes); the People access token is kept
// here, separate from the Projects one, because the old ZOHO_ACCESS_TOKEN has no People scope.

const axios = require('axios');
const defaultConfig = require('./config');

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
// People's JSON parser rejects a raw "|" inside a string (the Record Key has them): send it as \u007c.
const peopleJson = (value) => JSON.stringify(value).replace(/\|/g, '\\u007c');
const text = (value) => String(value === undefined || value === null ? '' : value).trim();

// "26-Oct-2026" -> "2026-10-26". Anything else -> ''.
function fromPeopleDate(value) {
  const match = /^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/.exec(text(value));
  const month = match ? MONTHS.findIndex((name) => name.toLowerCase() === match[2].toLowerCase()) : -1;
  return match && month >= 0 ? `${match[3]}-${String(month + 1).padStart(2, '0')}-${match[1].padStart(2, '0')}` : '';
}

// "2026-10-26" -> "26-Oct-2026" (the format the leave and holiday calls require).
function toPeopleDate(iso) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text(iso));
  return match ? `${match[3]}-${MONTHS[Number(match[2]) - 1]}-${match[1]}` : '';
}

// Leave status as the planner shows it. Rejected and cancelled leave is not leave at all. An "Absent"
// record (marked from attendance, no approval step) is an approved absence.
function leaveState(approvalStatus, type) {
  const status = text(approvalStatus).toLowerCase();
  if (status === 'approved') return 'approved';
  if (status === 'pending') return 'pending';
  if (!status && /^absent$/i.test(text(type))) return 'approved';
  return '';
}

function normalizeEmployee(record) {
  const first = text(record.FirstName);
  const last = text(record.LastName);
  return {
    employeeId: text(record.EmployeeID),
    recordId: text(record.recordId || record.Zoho_ID),
    zuid: text(record.ZUID),
    email: text(record.EmailID).toLowerCase(),
    name: [first, last].filter(Boolean).join(' '),
    status: text(record.Employeestatus),
    location: text(record.LocationName)
  };
}

// One record of /api/v2/leavetracker/leaves/records. `days` are the days that really count as leave
// (Zoho leaves weekends and holidays out of the day map), so a Sunday inside a leave is not a clash.
function normalizeLeave(id, record) {
  const state = leaveState(record.ApprovalStatus, record.Type);
  if (!state) return null;
  const days = Object.entries(record.Days && typeof record.Days === 'object' ? record.Days : {})
    .map(([day, detail]) => ({ date: fromPeopleDate(day), count: Number.parseFloat(detail && detail.LeaveCount) }))
    .filter((day) => day.date && day.count > 0);
  const from = fromPeopleDate(record.From);
  const to = fromPeopleDate(record.To) || from;
  if (!from) return null;
  return {
    id: text(id),
    employeeId: text(record.EmployeeId),
    zuid: text(record.ZUID),
    name: text(record.Employee),
    from,
    to,
    days: days.length ? days : null,
    state,
    type: text(record.Leavetype) || text(record.Type)
  };
}

// A holiday of the leave tracker. Days are listed one by one (a half day weighs 0.5). `locations` is
// empty for a holiday that applies to everyone. Restricted (optional) holidays are not days off.
function normalizeHoliday(holiday) {
  if (!holiday || holiday.isRestricted) return null;
  const days = [];
  (Array.isArray(holiday.dayDetails) ? holiday.dayDetails : []).forEach((entry) => {
    Object.entries(entry || {}).forEach(([day, detail]) => {
      const date = fromPeopleDate(day);
      if (date) days.push({ date, weight: detail && detail.ishalfday ? 0.5 : 1 });
    });
  });
  if (!days.length) {
    const date = fromPeopleDate(holiday.from);
    if (date) days.push({ date, weight: 1 });
  }
  return { id: text(holiday.id), name: text(holiday.name), locations: (Array.isArray(holiday.locations) ? holiday.locations : []).map((entry) => text(entry && entry.name)).filter(Boolean), days };
}

// ---- Mock portal (ZOHO_USE_MOCK): the same people as the mock Zoho Projects portal ---------------
const mockEmployees = [
  { employeeId: 'ELE-A', recordId: 'P-1001', zuid: '1001', email: 'anup@elegrow.com', name: 'Anup Parekh', status: 'Active', location: 'Surat' },
  { employeeId: 'ELE-K', recordId: 'P-1002', zuid: '1002', email: 'karishma@elegrow.com', name: 'Karishma Prajapati', status: 'Active', location: 'Surat' },
  { employeeId: 'ELE-N', recordId: 'P-1003', zuid: '1003', email: 'nidhi@elegrow.com', name: 'Nidhi Shah', status: 'Active', location: 'Surat' },
  { employeeId: 'ELE-M', recordId: 'P-1004', zuid: '1004', email: 'mohana@elegrow.com', name: 'Mohana Vamsee', status: 'Active', location: 'Work From Home' }
];
const mockLeave = [
  { id: 'L1', employeeId: 'ELE-K', zuid: '1002', name: 'Karishma Prajapati', from: '2026-10-21', to: '2026-10-21', days: [{ date: '2026-10-21', count: 1 }], state: 'approved', type: 'Casual Leave' },
  { id: 'L2', employeeId: 'ELE-A', zuid: '1001', name: 'Anup Parekh', from: '2026-10-22', to: '2026-10-22', days: [{ date: '2026-10-22', count: 1 }], state: 'pending', type: 'Privilege Leave' }
];
const mockHolidays = [
  { id: 'H1', name: 'Dussehra', locations: [], days: [{ date: '2026-10-20', weight: 1 }] },
  { id: 'H2', name: 'Surat Local Day', locations: ['Surat'], days: [{ date: '2026-10-30', weight: 1 }] }
];

// ---- Site Work Log form (6G) --------------------------------------------------------------------
// One field of the form as /people/api/forms/<form>/components lists it (live shape: displayname = label,
// labelname = API name, comptype, Options {Option1: {Value, Id}}, and for a lookup ReferedFormName /
// ReferedFieldName). Older spellings are read too. Only what the app needs is kept.
function normalizeFormField(raw) {
  const label = text(raw.displayname || raw.displayName || raw.labelName || raw.LabelName || raw.label);
  const options = raw.Options && typeof raw.Options === 'object' ? Object.values(raw.Options).map((option) => text(option && typeof option === 'object' ? option.Value : option)).filter(Boolean) : null;
  const lookup = raw.ReferedFieldName || raw.ReferedFormName ? { form: text(raw.ReferedFormName), field: text(raw.ReferedFieldName) } : raw.lookupDetails || raw.LookupDetails || raw.lookup || null;
  return {
    label,
    apiName: text(raw.labelname || raw.componentName || raw.ComponentName || raw.apiName || raw.fieldName || raw.labelName) || label.replace(/\s+/g, '_'),
    type: text(raw.comptype || raw.type || raw.Type || raw.fieldType).toLowerCase(),
    // a dropdown's options are the values to send; a lookup's are the records it can point at ({ value: shown text, id: what to send })
    options: options && !/lookup/i.test(text(raw.comptype)) ? options : null,
    lookupOptions: raw.Options && typeof raw.Options === 'object' && /lookup/i.test(text(raw.comptype)) ? Object.values(raw.Options).filter((option) => option && typeof option === 'object').map((option) => ({ value: text(option.Value), id: text(option.Id) })) : null,
    lookup
  };
}

// Which employee attribute a lookup field expects, read from its metadata: the lookup's display field
// when it names one, otherwise the employee's People record id.
function lookupAttribute(meta) {
  const lookup = meta && meta.lookup;
  const hints = [lookup && typeof lookup === 'object' ? Object.values(lookup).filter((value) => typeof value === 'string').join(' ') : lookup, meta && meta.expects].join(' ').toLowerCase();
  if (/e-?mail/.test(hints)) return 'email';
  if (/employee\s*_?id|emp_?id/.test(hints)) return 'employeeId';
  if (/(^|[^a-z])name/.test(hints)) return 'name';
  return 'recordId';
}

// Site Work Log: the labels the form is expected to carry. The app finds each field by its label when it
// writes (a field's API name is not assumed); ZOHO_PEOPLE_WORKLOG_FIELDS can name one by API name instead.
const WORKLOG_LABELS = {
  recordKey: 'Record Key', engineer: 'Engineer', workDate: 'Work Date', project: 'Project', client: 'Client', test: 'Test', unit: 'Unit',
  plannedQty: 'Planned Qty', completedQty: 'Completed Qty', reportStatus: 'Report Status', returnedCount: 'Returned Count',
  dataGaps: 'Data Gaps', onTime: 'On-time Submission', firstTimeApproved: 'First-time Approved', reportId: 'Report ID'
};

// The mock form has every field; a lookup with no display hint, so the Engineer value is the record id.
const mockLogForm = {
  fields: Object.values(WORKLOG_LABELS).map((label) => ({ label, apiName: label.replace(/[^A-Za-z0-9]+/g, '_'), type: label === 'Engineer' ? 'lookup' : 'text', lookup: null }))
};

const inRange = (from, to) => (entry) => entry.from <= to && entry.to >= from;

// `http` needs get() and post() (axios). `config` is the app config. The factory exists so a test can
// hand in a recording http and look at every call that was made.
function createPeopleClient({ http = axios, config = defaultConfig } = {}) {
  let token = { value: '', expiresAt: 0 };
  let inFlight = null;

  function hasCredentials() {
    return Boolean(config.zoho.refreshToken && config.zoho.clientId && config.zoho.clientSecret);
  }

  async function accessToken(force = false) {
    if (!force && token.value && Date.now() < token.expiresAt) return token.value;
    if (!hasCredentials()) throw new Error('Zoho People needs ZOHO_REFRESH_TOKEN, ZOHO_CLIENT_ID and ZOHO_CLIENT_SECRET.');
    if (!inFlight) {
      inFlight = (async () => {
        const body = new URLSearchParams({ refresh_token: config.zoho.refreshToken, client_id: config.zoho.clientId, client_secret: config.zoho.clientSecret, grant_type: 'refresh_token' });
        const response = await http.post(`${String(config.zoho.accountsBaseUrl).replace(/\/+$/, '')}/oauth/v2/token`, body.toString(), { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, timeout: 15000 });
        const value = response && response.data && response.data.access_token;
        if (!value) throw new Error('Zoho did not return an access token for Zoho People.');
        token = { value, expiresAt: Date.now() + 50 * 60 * 1000 };
        return value;
      })().finally(() => { inFlight = null; });
    }
    return inFlight;
  }

  // GET with one retry on a 401 (expired token). Error messages never carry the token.
  async function get(path, params = {}, retry = true) {
    const url = `${config.zoho.peopleBaseUrl}${path}`;
    try {
      const response = await http.get(url, { headers: { Authorization: `Zoho-oauthtoken ${await accessToken()}` }, params, timeout: 30000 });
      const data = response.data;
      const failure = data && data.response && data.response.status === 1 && data.response.errors;
      if (failure) throw Object.assign(new Error(`Zoho People: ${text(failure.message) || 'error'} (${text(failure.code)})`), { peopleError: true });
      return data;
    } catch (error) {
      const status = error && error.response && error.response.status;
      if (status === 401 && retry) {
        token = { value: '', expiresAt: 0 };
        await accessToken(true);
        return get(path, params, false);
      }
      if (error && error.peopleError) throw error;
      const detail = error && error.response && error.response.data;
      const code = detail && ((detail.error && detail.error.message) || (detail.response && detail.response.errors && detail.response.errors.message));
      throw new Error(status ? `Zoho People request failed (${status})${code ? `: ${text(code)}` : ''}${status === 401 || status === 403 ? '. Check the People scopes of the grant.' : ''}` : `Zoho People request failed: ${text(error && error.message)}`);
    }
  }

  async function listForms() {
    if (config.zoho.useMock) return [];
    const data = await get('/people/api/forms');
    return (data.response && data.response.result ? data.response.result : []).map((form) => ({ displayName: text(form.displayName), formLinkName: text(form.formLinkName) }));
  }

  async function listEmployees() {
    if (config.zoho.useMock) return mockEmployees;
    const employees = [];
    for (let start = 1; start < 5000; start += 200) {
      const data = await get('/people/api/forms/employee/getRecords', { sIndex: start, limit: 200 });
      const rows = (data.response && data.response.result) || [];
      // the key of a row is the id a lookup field takes (the record's Zoho_ID is another number)
      rows.forEach((row) => Object.entries(row).forEach(([rowKey, list]) => (Array.isArray(list) ? list : []).forEach((record) => employees.push(normalizeEmployee({ ...record, recordId: record.recordId || rowKey })))));
      if (rows.length < 200) break;
    }
    return employees;
  }

  // Leave that overlaps [from, to] (ISO dates). The call returns at most 200 records, so it is paged.
  async function listLeave(from, to) {
    if (config.zoho.useMock) return mockLeave.filter(inRange(from, to));
    const seen = new Map();
    for (let start = 1; start < 5000; ) {
      const data = await get('/api/v2/leavetracker/leaves/records', { from: toPeopleDate(from), to: toPeopleDate(to), startIndex: start, limit: 200 });
      const page = Object.entries(data.records || {});
      const before = seen.size;
      page.forEach(([id, record]) => { if (!seen.has(id)) seen.set(id, record); });
      if (page.length < 200 || seen.size === before) break;
      start = seen.size + 1;
    }
    return [...seen.entries()].map(([id, record]) => normalizeLeave(id, record)).filter(Boolean);
  }

  async function listHolidays(from, to) {
    if (config.zoho.useMock) return mockHolidays.filter((holiday) => holiday.days.some((day) => day.date >= from && day.date <= to));
    const data = await get('/api/v2/leavetracker/holidays', { from: toPeopleDate(from), to: toPeopleDate(to) });
    return (Array.isArray(data.holidays) ? data.holidays : []).map(normalizeHoliday).filter(Boolean);
  }

  // ---- Site Work Log (6G) ----------------------------------------------------------------------
  // POST of a form-encoded body, with the same one retry on a 401 as get().
  async function post(path, fields, retry = true) {
    const url = `${config.zoho.peopleBaseUrl}${path}`;
    try {
      const response = await http.post(url, new URLSearchParams(fields).toString(), { headers: { Authorization: `Zoho-oauthtoken ${await accessToken()}`, 'Content-Type': 'application/x-www-form-urlencoded' }, timeout: 30000 });
      const data = response.data;
      const failure = data && data.response && data.response.status === 1 && data.response.errors;
      if (failure) throw Object.assign(new Error(`Zoho People: ${text(failure.message) || 'error'} (${text(failure.code)})`), { peopleError: true });
      return data;
    } catch (error) {
      const status = error && error.response && error.response.status;
      if (status === 401 && retry) {
        token = { value: '', expiresAt: 0 };
        await accessToken(true);
        return post(path, fields, false);
      }
      if (error && error.peopleError) throw error;
      const detail = error && error.response && error.response.data;
      const code = detail && ((detail.error && detail.error.message) || (detail.response && detail.response.errors && detail.response.errors.message));
      throw new Error(status ? `Zoho People request failed (${status})${code ? `: ${text(code)}` : ''}` : `Zoho People request failed: ${text(error && error.message)}`);
    }
  }

  const mockRecords = new Map(); // mock portal: the Site Work Log records the app wrote, by People record id
  let mockSeq = 1000;

  async function hasForm(formLinkName) {
    if (config.zoho.useMock) return true;
    return (await listForms()).some((form) => form.formLinkName === formLinkName);
  }

  async function getFormFields(formLinkName) {
    if (config.zoho.useMock) return mockLogForm.fields;
    const data = await get(`/people/api/forms/${encodeURIComponent(formLinkName)}/components`);
    const out = [];
    const walk = (node) => {
      if (Array.isArray(node)) node.forEach(walk);
      else if (node && typeof node === 'object') {
        if (node.comptype || node.labelname || node.labelName || node.LabelName || node.componentName) out.push(normalizeFormField(node));
        Object.values(node).forEach((value) => typeof value === 'object' && walk(value));
      }
    };
    walk(data.response && data.response.result);
    return out;
  }

  // The record of the form whose `field` equals `value`, or null.
  async function findRecord(formLinkName, field, value) {
    if (config.zoho.useMock) {
      const hit = [...mockRecords.entries()].find(([, record]) => record[field] === value);
      return hit ? hit[0] : null;
    }
    let data;
    try {
      data = await get(`/people/api/forms/${encodeURIComponent(formLinkName)}/getRecords`, { searchParams: peopleJson({ searchField: field, searchOperator: 'Is', searchText: value }) });
    } catch (error) {
      if (/\(7024\)/.test(error.message)) return null; // People answers "No records found" as an error
      throw error;
    }
    const rows = (data.response && data.response.result) || [];
    // the key of a row is the id updateRecord takes (the record's Zoho_ID is another number)
    for (const row of rows) {
      const hit = Object.entries(row).find(([, list]) => Array.isArray(list) && list.length);
      if (hit) return hit[0];
    }
    return null;
  }

  async function insertRecord(formLinkName, data) {
    if (config.zoho.useMock) {
      const id = String(++mockSeq);
      mockRecords.set(id, { ...data });
      return id;
    }
    const result = await post(`/people/api/forms/json/${encodeURIComponent(formLinkName)}/insertRecord`, { inputData: peopleJson(data) });
    const id = text(result.response && result.response.result && result.response.result.pkId);
    if (!id) throw new Error('Zoho People did not return the id of the new record.');
    return id;
  }

  async function updateRecord(formLinkName, recordId, data) {
    if (config.zoho.useMock) {
      if (!mockRecords.has(String(recordId))) throw new Error('Zoho People: record not found.');
      mockRecords.set(String(recordId), { ...mockRecords.get(String(recordId)), ...data });
      return recordId;
    }
    await post(`/people/api/forms/json/${encodeURIComponent(formLinkName)}/updateRecord`, { inputData: peopleJson(data), recordId: String(recordId) });
    return recordId;
  }

  return { listForms, listEmployees, listLeave, listHolidays, hasForm, getFormFields, findRecord, insertRecord, updateRecord, mockRecords: () => [...mockRecords.entries()].map(([id, record]) => ({ id, ...record })) };
}

module.exports = { createPeopleClient, fromPeopleDate, toPeopleDate, leaveState, normalizeEmployee, normalizeLeave, normalizeHoliday, normalizeFormField, lookupAttribute, WORKLOG_LABELS };
