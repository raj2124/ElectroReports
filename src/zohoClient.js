const axios = require('axios');
const config = require('./config');

const mockProjects = [
  {
    id: '1001',
    name: 'Corporate Website Revamp',
    projectNumber: 'PRJ-2026-001',
    clientName: 'Acme Corporation',
    workOrderNo: 'WO-ACME-42',
    ownerName: 'Rakesh Patel',
    stage: 'Execution'
  },
  {
    id: '1002',
    name: 'Mobile App Rollout',
    projectNumber: 'PRJ-2026-002',
    clientName: 'BlueWave Logistics',
    workOrderNo: 'WO-BLUE-88',
    ownerName: 'Nidhi Shah',
    stage: 'Planning'
  }
];

const mockProjectUsers = {
  '1001': [
    { id: 'u101', displayName: 'Rakesh Patel', email: 'rakesh@elegrow.com', role: 'Project Manager' },
    { id: 'u102', displayName: 'Nidhi Shah', email: 'nidhi@elegrow.com', role: 'Engineer' }
  ],
  '1002': [
    { id: 'u201', displayName: 'Mohana Vamsee', email: 'mohana@elegrow.com', role: 'Engineer' },
    { id: 'u202', displayName: 'Arpit Jain', email: 'arpit@elegrow.com', role: 'Coordinator' }
  ]
};

const mockPortalUsers = Array.from(
  Object.values(mockProjectUsers)
    .flat()
    .reduce((map, user) => {
      const normalizedEmail = String(user.email || '').trim().toLowerCase();
      const normalizedName = String(user.displayName || '').trim().toLowerCase();
      const key = `${user.id || ''}::${normalizedEmail}::${normalizedName}`;
      if (!map.has(key)) {
        map.set(key, user);
      }
      return map;
    }, new Map())
    .values()
);

const mockProjectClientUsers = {
  '1001': [
    {
      id: 'cu101',
      displayName: 'Anil Mehta',
      email: 'anil@acme.com',
      role: 'Client User',
      companyName: 'Acme Corporation',
      userType: 'client',
      isClient: true
    }
  ],
  '1002': [
    {
      id: 'cu201',
      displayName: 'Rina Shah',
      email: 'rina@bluewave.com',
      role: 'Client User',
      companyName: 'BlueWave Logistics',
      userType: 'client',
      isClient: true
    }
  ]
};

const mockProjectTasks = {
  '1001': [
    {
      id: 't1001',
      name: 'Finalize MOM template formatting',
      status: 'In Progress',
      ownerName: 'Rakesh Patel',
      dueDate: '2026-03-07',
      percentComplete: '65',
      priority: 'High',
      taskListName: 'Documentation'
    },
    {
      id: 't1002',
      name: 'Review action plan owners',
      status: 'Open',
      ownerName: 'Nidhi Shah',
      dueDate: '2026-03-09',
      percentComplete: '15',
      priority: 'Medium',
      taskListName: 'Review'
    }
  ],
  '1002': [
    {
      id: 't2001',
      name: 'Collect deployment dependencies',
      status: 'Open',
      ownerName: 'Arpit Jain',
      dueDate: '2026-03-08',
      percentComplete: '10',
      priority: 'High',
      taskListName: 'Deployment'
    }
  ]
};

// A D03 project as the live portal returns it (056-D03_TS-26-CLL): the tasks are in the raw v1 shape
// (custom fields as { column_name, label_name, value }, owners under details), so the same
// normalisation runs in the tests and live. Empty custom fields are omitted, as Zoho does.
const rawTask = (id, name, tasklist, { owner = anup, category = 'Offsite', start = '10-26-2026', end = '10-26-2026', extra = [] } = {}) => ({
  id_string: id,
  name,
  status: { name: 'Planned' },
  start_date: start,
  end_date: end,
  tasklist: { id: `tl-${tasklist.replace(/\W+/g, '').toLowerCase()}`, id_string: `tl-${tasklist.replace(/\W+/g, '').toLowerCase()}`, name: tasklist },
  details: { owners: [owner] },
  custom_fields: [
    { column_name: 'UDF_CHAR86', label_name: 'Docs Rev. No.', value: 'R0' },
    { column_name: 'UDF_CHAR81', label_name: 'Work Location', value: category },
    ...extra
  ]
});
function anupOwner() { return { id: '60007088431', zpuid: '27127000000467187', name: 'anup', email: 'anup@elegrow.com' }; }
const anup = anupOwner();
const ONSITE = 'Onsite Survey / Measurement Activities';
const mockRawProjectTasks = {
  '27127000009850033': [
    rawTask('r1', '[IRT] Input Data Receipt &amp; Review (Plant Layout, HT/LV SLD, Feeder Loading Data)', 'Kick-off Meeting & Input Data', { owner: { id: '60007088440', zpuid: '27127000000467199', name: 'Karishma', email: 'karishma@elegrow.com' } }),
    rawTask('r2', '[ESHA] Input Data Receipt & Review (Plot Plan, Earthing Layout, Pit Schedule, Past Reports, Threshold)', 'Kick-off Meeting & Input Data'),
    rawTask('r3', '[ESHA] Soil Resistivity Testing - 4 Locations (Wenner 4-Pin)', ONSITE, { category: 'Onsite', extra: [{ column_name: 'UDF_CHAR90', label_name: 'ESHA Test', value: 'Soil Resistivity' }, { column_name: 'UDF_CHAR91', label_name: 'Planned Quantity', value: '4' }] }),
    rawTask('r4', '[ESHA] Earth Pit Resistance Testing - 170 Pits (Fall-of-Potential / 3-Pole)', ONSITE, { category: 'Onsite', start: '10-19-2026', end: '10-23-2026', extra: [{ column_name: 'UDF_CHAR90', label_name: 'ESHA Test', value: 'Earth Electrode Resistance' }, { column_name: 'UDF_CHAR91', label_name: 'Planned Quantity', value: '170 pits' }] }),
    rawTask('r5', '[ESHA] Pit Resistance Data Tabulation & Pass/Fail vs Acceptance Limit', 'Data Analysis', { start: '10-27-2026', end: '10-27-2026', extra: [{ column_name: 'UDF_CHAR90', label_name: 'ESHA Test', value: 'Earth Electrode Resistance' }] }),
    rawTask('r6', '[ESHA] Draft Earth Pit Resistance Test Report (170 Pits)', 'Draft Report Preparation', { start: '10-28-2026', end: '10-28-2026' })
  ]
};
const mockProjectsD03 = [
  { id: '27127000009850033', name: '056-D03_TS-26-CLL', projectNumber: '056-D03_TS-26-CLL', clientName: 'CLL', workOrderNo: '', ownerName: 'Anup Parekh', ownerEmail: 'anup@elegrow.com', stage: 'Planning', createdAt: String(Date.parse('2026-10-13T08:49:46Z')) },
  { id: '27127000009860001', name: '060-D03_TS-26-XYZ', projectNumber: '060-D03_TS-26-XYZ', clientName: 'XYZ', workOrderNo: '', ownerName: 'Nidhi Shah', ownerEmail: 'nidhi@elegrow.com', stage: 'Active', createdAt: String(Date.parse('2026-10-15T08:00:00Z')) },
  { id: '27127000009860002', name: '024-D03_TS-26-GFL', projectNumber: '024-D03_TS-26-GFL', clientName: 'GFL', workOrderNo: '', ownerName: 'Rakesh Patel', ownerEmail: 'rakesh@elegrow.com', stage: 'Client Review', createdAt: String(Date.parse('2026-07-01T11:19:17Z')) },
  { id: '27127000009860003', name: '011-D03_TS-25-OLD', projectNumber: '011-D03_TS-25-OLD', clientName: 'OLD', workOrderNo: '', ownerName: 'Rakesh Patel', ownerEmail: 'rakesh@elegrow.com', stage: 'Active', createdAt: String(Date.parse('2025-12-20T08:00:00Z')) }
];
mockProjects.push(...mockProjectsD03);
mockProjectUsers['27127000009850033'] = [
  { id: '60007088431', displayName: 'Anup Parekh', email: 'anup@elegrow.com', role: 'Engineer' },
  { id: '60007088440', displayName: 'Karishma Prajapati', email: 'karishma@elegrow.com', role: 'Engineer' }
];
mockPortalUsers.push(...mockProjectUsers['27127000009850033']);

let currentAccessToken = config.zoho.accessToken || '';
let refreshInFlight = null;
const GENERIC_CLIENT_LABELS = new Set(['', '-', 'allexternal', 'external', 'client', 'all external']);
const GENERIC_LIFECYCLE_STAGE_LABELS = new Set([
  'active',
  'open',
  'opened',
  'archived',
  'archive',
  'closed',
  'complete',
  'completed',
  'in progress',
  'progress'
]);

function asText(value) {
  if (value === undefined || value === null) {
    return '';
  }
  if (typeof value === 'string' || typeof value === 'number') {
    return String(value);
  }
  if (typeof value === 'object') {
    const keys = ['id', 'id_string', 'project_id', 'value', 'text', 'name'];
    for (const key of keys) {
      const nested = asText(value[key]);
      if (nested) {
        return nested;
      }
    }
  }
  return '';
}

function sanitizeClientName(value) {
  const text = asText(value).trim();
  if (!text) {
    return '';
  }
  const normalized = text.toLowerCase();
  if (GENERIC_CLIENT_LABELS.has(normalized)) {
    return '';
  }
  return text;
}

function extractProjectId(project) {
  const candidates = [];
  const directKeys = ['id_string', 'id', 'project_id', 'projectId', 'projectid'];
  for (const key of directKeys) {
    const rawValue = project?.[key];
    if (rawValue === undefined || rawValue === null) {
      continue;
    }

    if (typeof rawValue === 'string' || typeof rawValue === 'number') {
      const value = String(rawValue).trim();
      if (value) {
        candidates.push(value);
      }
      continue;
    }

    if (typeof rawValue === 'object') {
      const nestedKeys = ['id', 'id_string', 'project_id', 'value', 'text'];
      for (const nestedKey of nestedKeys) {
        const nestedValue = asText(rawValue[nestedKey]).trim();
        if (nestedValue) {
          candidates.push(nestedValue);
        }
      }
    }
  }

  const selfUrl =
    asText(project?.link?.self?.url) ||
    asText(project?.links?.self?.href) ||
    asText(project?.self?.url) ||
    '';
  if (selfUrl) {
    const match = selfUrl.match(/\/projects\/([^/?#]+)/i);
    if (match && match[1]) {
      candidates.push(match[1]);
    }
  }

  if (!candidates.length) {
    return '';
  }

  const uniqueCandidates = [...new Set(candidates)];
  const numericCandidate = uniqueCandidates.find((id) => /^\d{8,}$/.test(id));
  if (numericCandidate) {
    return numericCandidate;
  }

  const idLikeCandidate = uniqueCandidates.find((id) => /^[a-z0-9-]{6,}$/i.test(id));
  return idLikeCandidate || '';
}

function readCustomValue(project, aliases) {
  if (!project || typeof project !== 'object') {
    return '';
  }

  const normalizedAliases = aliases.map((name) => String(name).toLowerCase());

  if (Array.isArray(project.custom_fields)) {
    for (const field of project.custom_fields) {
      const key = String(field?.label || field?.name || '').toLowerCase();
      if (normalizedAliases.some((alias) => key.includes(alias))) {
        return field?.value || field?.text || '';
      }
    }
  }

  for (const [key, value] of Object.entries(project)) {
    const normalizedKey = key.toLowerCase();
    if (normalizedAliases.some((alias) => normalizedKey.includes(alias))) {
      if (typeof value === 'string' || typeof value === 'number') {
        return String(value);
      }
    }
  }

  return '';
}

// Read a custom field value by label/name alias from project.custom_fields ONLY.
// Unlike readCustomValue, this does NOT scan arbitrary top-level project keys
// (which include noisy fields such as bug_client_permission="allexternal" that
// falsely match the "client" alias), so it is safe for ambiguous aliases.
function readCustomFieldValue(project, aliases) {
  if (!project || !Array.isArray(project.custom_fields)) {
    return '';
  }
  const normalizedAliases = aliases.map((name) => String(name).toLowerCase());
  for (const field of project.custom_fields) {
    const key = String(field?.label || field?.name || '').toLowerCase();
    if (normalizedAliases.some((alias) => key.includes(alias))) {
      return field?.value || field?.text || '';
    }
  }
  return '';
}

function readNestedText(source, keys) {
  if (!source || typeof source !== 'object') {
    return '';
  }
  for (const key of keys) {
    const value = source[key];
    if (value === undefined || value === null) {
      continue;
    }
    if (typeof value === 'string' || typeof value === 'number') {
      const text = String(value).trim();
      if (text) {
        return text;
      }
      continue;
    }
    if (typeof value === 'object') {
      const nested = readNestedText(value, [
        'name',
        'display_name',
        'full_name',
        'value',
        'text',
        'label',
        'display_value',
        'status',
        'status_name',
        'stage',
        'stage_name'
      ]);
      if (nested) {
        return nested;
      }
    }
  }
  return '';
}

function normalizeStageText(value) {
  const text = asText(value).trim();
  if (!text || /^\d+$/.test(text)) {
    return '';
  }

  const compact = text.replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!compact) {
    return '';
  }

  if (compact === compact.toLowerCase()) {
    return compact
      .split(' ')
      .map((word) => word.charAt(0).toUpperCase() + word.slice(1))
      .join(' ');
  }

  return compact;
}

function isGenericLifecycleStage(value) {
  const normalized = normalizeStageText(value).toLowerCase();
  return GENERIC_LIFECYCLE_STAGE_LABELS.has(normalized);
}

function readCustomFieldCandidates(project, aliases) {
  if (!project || typeof project !== 'object' || !Array.isArray(project.custom_fields)) {
    return [];
  }

  const normalizedAliases = aliases.map((name) => String(name).toLowerCase());
  const values = [];
  for (const field of project.custom_fields) {
    const key = String(field?.label || field?.name || '').toLowerCase();
    if (!normalizedAliases.some((alias) => key.includes(alias))) {
      continue;
    }
    values.push(field?.display_value, field?.value, field?.text, field?.name);
  }

  return values.map(normalizeStageText).filter(Boolean);
}

function extractOwnerName(project) {
  const direct =
    asText(project.owner_name) ||
    asText(project.owner_full_name) ||
    asText(project.owner_display_name) ||
    asText(project.owner_user_name) ||
    asText(project.project_owner_name) ||
    '';
  if (direct) {
    return direct;
  }

  const nestedOwner =
    readNestedText(project, ['owner', 'project_owner', 'owner_details', 'owner_data']) ||
    readCustomValue(project, ['owner', 'project owner', 'responsible']) ||
    '';
  return nestedOwner;
}

function extractStage(project) {
  const nestedCandidates = [
    readNestedText(project, [
      'project_stage',
      'stage_details',
      'project_status',
      'status_details',
      'status_data',
      'status_info',
      'status_obj',
      'custom_status',
      'custom_stage',
      'stage'
    ]),
    readNestedText(project, ['status']),
    readNestedText(project, ['status_name', 'stage_name']),
    asText(project.project_stage_name),
    asText(project.custom_status_name),
    asText(project.custom_stage_name),
    asText(project.stage_name),
    asText(project.stage),
    asText(project.current_status),
    asText(project.status_text),
    asText(project.status_label),
    readCustomValue(project, ['project stage', 'stage']),
    ...readCustomFieldCandidates(project, ['project stage', 'stage', 'status'])
  ]
    .map(normalizeStageText)
    .filter(Boolean);

  const directCandidates = [
    asText(project.project_status),
    asText(project.status_name),
    asText(project.status_type),
    asText(project.status)
  ]
    .map(normalizeStageText)
    .filter(Boolean);

  const allCandidates = [...nestedCandidates, ...directCandidates];
  if (!allCandidates.length) {
    return '';
  }

  const nonGeneric = allCandidates.find((value) => !isGenericLifecycleStage(value));
  return nonGeneric || allCandidates[0];
}

// Parse client/customer info out of the Zoho project NAME string. Used ONLY as a
// fallback when the authoritative custom fields are empty. Two real-world formats
// exist on the live portal:
//   A) dash format      SEQ-DEPT-FY-CUST           e.g. 025-D02_PSS-26-SPPL  (DEPT may contain "_")
//   B) underscore fmt   SEQ_MM_FY_DEPT_CUST[_...]  e.g. 561_03_26_TS_AIL_EA_NEO
// Returns null for names matching neither, so callers keep the raw name and leave
// client blank (never guess, never show garbage).
function parseProjectNameParts(rawName) {
  const name = asText(rawName).trim();
  if (!name) {
    return null;
  }

  // Format A: exactly SEQ-DEPT-FY-CUST when split on "-".
  const dashParts = name.split('-').map((part) => part.trim());
  if (dashParts.length === 4) {
    const [seq, dept, fy, customer] = dashParts;
    if (/^\d{2,4}$/.test(seq) && /^\d{2,4}$/.test(fy) && customer) {
      return { format: 'dash', seq, dept, fy, customer };
    }
  }

  // Format B: SEQ_MM_FY_DEPT_CUST with optional trailing scope tokens (customer = 5th token).
  const underscoreParts = name.split('_').map((part) => part.trim());
  if (underscoreParts.length >= 5) {
    const [seq, mm, fy, dept, customer] = underscoreParts;
    if (/^\d{2,4}$/.test(seq) && /^\d{1,2}$/.test(mm) && /^\d{2,4}$/.test(fy) && customer) {
      return { format: 'underscore', seq, mm, fy, dept, customer };
    }
  }

  return null;
}

function normalizeProject(project) {
  const id = extractProjectId(project);
  const name = asText(project.name || project.project_name);
  const parsedName = parseProjectNameParts(name);
  const projectNumber =
    readCustomValue(project, ['project number', 'project no']) ||
    asText(project.project_no) ||
    asText(project.project_number) ||
    name ||
    '';
  const workOrderNo =
    readCustomValue(project, ['work order', 'wo number', 'work order no']) ||
    asText(project.work_order_no) ||
    '';
  const clientName =
    sanitizeClientName(readCustomFieldValue(project, ['client'])) ||
    sanitizeClientName(project.client_name) ||
    sanitizeClientName(project.client?.name) ||
    sanitizeClientName(project.customer_name) ||
    sanitizeClientName(project.customer?.name) ||
    (parsedName ? sanitizeClientName(parsedName.customer) : '') ||
    '';
  const updatedAt =
    asText(project.last_updated_time) ||
    asText(project.last_updated_time_long) ||
    asText(project.updated_time) ||
    asText(project.updated_time_long) ||
    asText(project.last_update_time) ||
    asText(project.modified_time) ||
    asText(project.modified_time_long) ||
    asText(project.updated_at) ||
    '';
  const createdAt =
    asText(project.created_time) ||
    asText(project.created_time_long) ||
    asText(project.created_at) ||
    asText(project.created_date_long) ||
    asText(project.create_time) ||
    '';

  const ownerName = extractOwnerName(project);
  const stage = extractStage(project);
  const completion =
    asText(project.completed_percent) ||
    asText(project.percent_complete) ||
    asText(project.completion_percentage) ||
    '';
  const isClosed = String(project.is_closed || '').toLowerCase() === 'true';
  const derivedStage =
    stage ||
    (isClosed ? 'Closed' : '') ||
    (completion === '100' ? 'Completed' : '') ||
    (completion ? `In Progress (${completion}%)` : '');

  return {
    id: asText(id),
    name: asText(name),
    projectNumber: asText(projectNumber),
    clientName: sanitizeClientName(clientName),
    workOrderNo: asText(workOrderNo),
    ownerName: asText(ownerName),
    ownerEmail: asText(project.owner_email || project.owner?.email),
    ownerId: asText(project.owner_id),
    stage: asText(derivedStage),
    updatedAt: asText(updatedAt),
    createdAt: asText(createdAt)
  };
}

function collectProjectList(payload) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  if (Array.isArray(payload)) {
    return payload;
  }

  const possibleKeys = [
    'projects',
    'project',
    'data',
    'result',
    'response',
    'items',
    'records'
  ];

  for (const key of possibleKeys) {
    if (Array.isArray(payload[key])) {
      return payload[key];
    }
    if (payload[key] && typeof payload[key] === 'object') {
      const nested = collectProjectList(payload[key]);
      if (nested.length > 0) {
        return nested;
      }
    }
  }

  return [];
}

function collectProjectDetail(payload) {
  if (!payload || typeof payload !== 'object') {
    return null;
  }

  if (Array.isArray(payload)) {
    return payload[0] || null;
  }

  if (payload.project && typeof payload.project === 'object') {
    return payload.project;
  }

  const list = collectProjectList(payload);
  if (list.length) {
    return list[0];
  }

  if (payload.id || payload.id_string || payload.project_id) {
    return payload;
  }

  return null;
}

function extractTaskId(task) {
  const candidates = [];
  const keys = ['id_string', 'id', 'task_id', 'taskid', 'taskId'];
  for (const key of keys) {
    const value = asText(task?.[key]).trim();
    if (value) {
      candidates.push(value);
    }
  }

  const selfUrl =
    asText(task?.link?.self?.url) ||
    asText(task?.links?.self?.href) ||
    asText(task?.self?.url) ||
    '';
  if (selfUrl) {
    const match = selfUrl.match(/\/tasks\/([^/?#]+)/i);
    if (match && match[1]) {
      candidates.push(match[1]);
    }
  }

  if (!candidates.length) {
    return '';
  }

  const unique = [...new Set(candidates)];
  const numeric = unique.find((id) => /^\d{8,}$/.test(id));
  if (numeric) {
    return numeric;
  }

  return unique[0];
}

function collectTaskList(payload) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  if (Array.isArray(payload)) {
    return payload;
  }

  const keys = ['tasks', 'task', 'data', 'result', 'response', 'records', 'items'];
  for (const key of keys) {
    if (Array.isArray(payload[key])) {
      return payload[key];
    }
    if (payload[key] && typeof payload[key] === 'object') {
      const nested = collectTaskList(payload[key]);
      if (nested.length > 0) {
        return nested;
      }
    }
  }

  return [];
}

function normalizeTask(task) {
  const id = extractTaskId(task);
  const name =
    asText(task?.name) ||
    asText(task?.task_name) ||
    asText(task?.title) ||
    asText(task?.task_title) ||
    '';
  const status =
    normalizeStageText(
      readNestedText(task, ['status', 'status_name', 'task_status', 'state', 'status_details']) ||
        asText(task?.status) ||
        asText(task?.task_status) ||
        asText(task?.state) ||
        ''
    ) || '';
  const ownerName =
    asText(task?.owner_name) ||
    asText(task?.person_responsible_name) ||
    readNestedText(task, ['owner', 'owner_details', 'assigned_to', 'created_by']) ||
    '';
  const dueDate =
    asText(task?.due_date) ||
    asText(task?.end_date) ||
    asText(task?.due_date_format) ||
    asText(task?.end_date_format) ||
    asText(task?.due_date_long) ||
    asText(task?.end_date_long) ||
    '';
  const startDate =
    asText(task?.start_date) ||
    asText(task?.start_date_format) ||
    asText(task?.start_date_long) ||
    '';
  const percentComplete =
    asText(task?.completed_percent) ||
    asText(task?.percent_complete) ||
    asText(task?.completion_percentage) ||
    asText(task?.progress) ||
    '';
  const priority = asText(task?.priority) || asText(task?.priority_name) || '';
  const taskListName = readNestedText(task, ['tasklist', 'task_list', 'tasklist_name']) || '';
  const milestoneName = readNestedText(task, ['milestone', 'milestone_name']) || '';
  const closedFlag =
    String(task?.is_closed || task?.isclosed || task?.closed || '')
      .trim()
      .toLowerCase() === 'true';
  const effectiveStatus = status || (closedFlag ? 'Closed' : '');

  // Owners with their ids and e-mail, so a task owner can be matched to a Zoho user, and the custom
  // fields as the payload carries them (label, API name, value). Empty custom fields are omitted by
  // Zoho, so a field that is not here has no value on this task.
  const ownerList = (Array.isArray(task?.details?.owners) ? task.details.owners : Array.isArray(task?.owners) ? task.owners : [])
    .map((owner) => ({ id: asText(owner?.id || owner?.zuid), zpuid: asText(owner?.zpuid), name: asText(owner?.name || owner?.full_name), email: asText(owner?.email).toLowerCase() }))
    .filter((owner) => owner.id || owner.name || owner.email);
  const customFields = (Array.isArray(task?.custom_fields) ? task.custom_fields : [])
    .map((field) => ({ label: asText(field?.label_name), apiName: asText(field?.column_name || field?.field_name || field?.api_name), value: asText(field?.value) }))
    .filter((field) => field.label || field.apiName);

  return {
    id: asText(id),
    name: asText(name).replace(/&amp;/g, '&'),
    taskListId: asText(task?.tasklist?.id_string || task?.tasklist?.id),
    owners: ownerList,
    customFields,
    status: asText(effectiveStatus),
    ownerName: asText(ownerName || (ownerList[0] && ownerList[0].name)),
    dueDate: asText(dueDate),
    startDate: asText(startDate),
    percentComplete: asText(percentComplete),
    priority: asText(priority),
    taskListName: asText(taskListName).replace(/&amp;/g, '&'),
    milestoneName: asText(milestoneName)
  };
}

function hasRefreshCredentials() {
  return (
    Boolean(config.zoho.refreshToken) &&
    Boolean(config.zoho.clientId) &&
    Boolean(config.zoho.clientSecret)
  );
}

function isAxiosStatus(error, status) {
  return axios.isAxiosError(error) && error.response?.status === status;
}

async function refreshZohoAccessToken() {
  if (!hasRefreshCredentials()) {
    throw new Error(
      'Missing Zoho refresh credentials. Set ZOHO_REFRESH_TOKEN, ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET.'
    );
  }

  if (refreshInFlight) {
    return refreshInFlight;
  }

  refreshInFlight = (async () => {
    const base = String(config.zoho.accountsBaseUrl || '').replace(/\/+$/, '');
    const url = `${base}/oauth/v2/token`;
    const body = new URLSearchParams({
      refresh_token: config.zoho.refreshToken,
      client_id: config.zoho.clientId,
      client_secret: config.zoho.clientSecret,
      grant_type: 'refresh_token'
    });

    const response = await axios.post(url, body.toString(), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      timeout: 15000
    });

    const token = response?.data?.access_token;
    if (!token) {
      const payload = response?.data;
      const details =
        payload && typeof payload === 'object'
          ? payload.error ||
            payload.error_description ||
            payload.message ||
            payload.code ||
            JSON.stringify(payload)
          : '';
      throw new Error(
        `Zoho refresh did not return access_token.${details ? ` Response: ${details}` : ''}`
      );
    }

    currentAccessToken = token;
    return token;
  })();

  try {
    return await refreshInFlight;
  } finally {
    refreshInFlight = null;
  }
}

function toZohoErrorMessage(error) {
  if (!axios.isAxiosError(error)) {
    return error?.message || 'Unknown Zoho API error';
  }

  const status = error.response?.status;
  const data = error.response?.data;
  const apiMessage =
    data?.message ||
    data?.error?.message ||
    data?.error?.code ||
    data?.code ||
    '';

  if (status === 401) {
    const refreshHint = hasRefreshCredentials()
      ? ' Auto-refresh is configured; verify refresh token/client credentials and accounts domain.'
      : ' Configure refresh credentials to auto-recover token expiry.';
    return (
      'Zoho authentication failed (401). Access token is invalid/expired, ' +
      'or token data-center does not match API domain (zoho.in vs zoho.com).' +
      refreshHint
    );
  }

  if (status === 403) {
    return 'Zoho access denied (403). OAuth scopes or portal permissions are insufficient.';
  }

  if (status) {
    return `Zoho API request failed (${status})${apiMessage ? `: ${apiMessage}` : ''}`;
  }

  return error.message || 'Zoho API request failed';
}

function ensureZohoCredentials() {
  if (!config.zoho.portalId) {
    throw new Error('Missing Zoho credentials. Configure ZOHO_PORTAL_ID in .env');
  }

  if (!currentAccessToken) {
    if (hasRefreshCredentials()) {
      return refreshZohoAccessToken();
    }
    throw new Error(
      'Missing Zoho access token. Configure ZOHO_ACCESS_TOKEN or refresh credentials in .env'
    );
  }

  return Promise.resolve(currentAccessToken);
}

function buildZohoProjectsBasePath() {
  const customEndpoint = String(config.zoho.projectsEndpoint || '').trim();
  if (customEndpoint) {
    if (customEndpoint.includes('/portal/') && customEndpoint.includes('/projects')) {
      return customEndpoint.replace(/\/+$/, '').replace(/\/projects\/?$/, '/projects');
    }

    if (customEndpoint.includes('/portal/')) {
      return `${customEndpoint.replace(/\/+$/, '')}/projects`;
    }

    return `${customEndpoint.replace(/\/+$/, '')}/portal/${config.zoho.portalId}/projects`;
  }

  return `${config.zoho.baseUrl.replace(/\/+$/, '')}/portal/${config.zoho.portalId}/projects`;
}

function buildZohoPortalUsersBasePath() {
  return `${config.zoho.baseUrl.replace(/\/+$/, '')}/portal/${config.zoho.portalId}/users`;
}

async function requestZohoGet(url, params = {}, allowRetry = true) {
  try {
    return await axios.get(url, {
      headers: {
        Authorization: `Zoho-oauthtoken ${currentAccessToken}`
      },
      params,
      timeout: 15000
    });
  } catch (error) {
    if (allowRetry && isAxiosStatus(error, 401) && hasRefreshCredentials()) {
      await refreshZohoAccessToken();
      return requestZohoGet(url, params, false);
    }
    throw error;
  }
}

async function fetchProjectDetailFromZoho(projectId) {
  await ensureZohoCredentials();
  const projectBase = buildZohoProjectsBasePath();
  const endpoints = [`${projectBase}/${projectId}/`, `${projectBase}/${projectId}`];

  let lastError = null;
  for (const endpoint of endpoints) {
    try {
      const response = await requestZohoGet(endpoint, {}, true);
      const detail = collectProjectDetail(response.data);
      if (detail && typeof detail === 'object') {
        return detail;
      }
    } catch (error) {
      lastError = error;
      if (!isAxiosStatus(error, 404) && !isAxiosStatus(error, 400)) {
        throw error;
      }
    }
  }

  if (lastError) {
    return null;
  }
  return null;
}

async function enrichProjectsWithDetailStatus(projects) {
  if (!config.zoho.enrichProjectStage || !Array.isArray(projects) || projects.length === 0) {
    return projects;
  }

  const enriched = projects.map((project) => ({ ...project }));
  const candidates = enriched.filter(
    (project) =>
      (project.id && /^\d{8,}$/.test(String(project.id))) &&
      (isGenericLifecycleStage(project.stage) || !String(project.stage || '').trim())
  );

  if (!candidates.length) {
    return enriched;
  }

  const concurrency = 4;
  let index = 0;
  const workers = new Array(Math.min(concurrency, candidates.length)).fill(null).map(async () => {
    while (index < candidates.length) {
      const currentIndex = index;
      index += 1;

      const candidate = candidates[currentIndex];
      try {
        const detailRaw = await fetchProjectDetailFromZoho(candidate.id);
        if (!detailRaw) {
          continue;
        }

        const detail = normalizeProject(detailRaw);
        if (detail.stage && (!candidate.stage || isGenericLifecycleStage(candidate.stage))) {
          candidate.stage = detail.stage;
        }
        if (!candidate.ownerName && detail.ownerName) {
          candidate.ownerName = detail.ownerName;
        }
        if (!candidate.updatedAt && detail.updatedAt) {
          candidate.updatedAt = detail.updatedAt;
        }
      } catch (_error) {
        // Ignore detail-level failures and keep list-level values.
      }
    }
  });

  await Promise.all(workers);
  return enriched;
}

function normalizeUser(user) {
  const id = asText(user.id || user.user_id || user.zpuid || user.uid);
  const firstName = asText(user.first_name || user.firstname);
  const lastName = asText(user.last_name || user.lastname);
  const fullName =
    asText(user.display_name) ||
    asText(user.name) ||
    asText(user.full_name) ||
    `${firstName} ${lastName}`.trim() ||
    asText(user.email) ||
    '';
  const email = asText(user.email || user.email_id || user.mail || user.primary_email);
  const role = asText(user.role || user.role_name || user.user_role);
  const companyName = asText(
    user.company_name ||
      user.company ||
      user.organization_name ||
      user.account_name ||
      user.client_name ||
      user.customer_name
  );
  const userType = asText(user.user_type || user.type || user.member_type || user.category);
  const isClientFlag = String(user.is_client || user.client_user || '').toLowerCase() === 'true';
  const isClientByType = /client/i.test(userType);
  const isClientByRole = /client/i.test(role);
  const isClient = isClientFlag || isClientByType || isClientByRole;

  return {
    id: asText(id),
    displayName: asText(fullName),
    email: asText(email),
    role: asText(role),
    companyName: asText(companyName),
    userType: asText(userType),
    isClient
  };
}

function collectUserList(payload) {
  if (!payload || typeof payload !== 'object') {
    return [];
  }

  if (Array.isArray(payload)) {
    return payload;
  }

  const keys = ['users', 'user', 'client_users', 'clientUsers', 'data', 'result', 'response', 'records', 'items'];
  for (const key of keys) {
    if (Array.isArray(payload[key])) {
      return payload[key];
    }
    if (payload[key] && typeof payload[key] === 'object') {
      const nested = collectUserList(payload[key]);
      if (nested.length > 0) {
        return nested;
      }
    }
  }

  const nestedObjects = Object.values(payload).filter(
    (value) => value && typeof value === 'object' && !Array.isArray(value)
  );
  if (nestedObjects.length) {
    const likelyUsers = nestedObjects.filter((value) => {
      return Boolean(
        value.user_id ||
          value.zpuid ||
          value.uid ||
          value.email ||
          value.email_id ||
          value.display_name ||
          value.name
      );
    });
    if (likelyUsers.length) {
      return likelyUsers;
    }
  }

  return [];
}

function filterUsersToOrgDomain(users) {
  const configured = String(config.zoho.organizationUserEmailDomain || '')
    .trim()
    .toLowerCase()
    .replace(/^@/, '');

  if (!configured) {
    return users;
  }

  return users.filter((user) => {
    const email = String(user?.email || '').trim().toLowerCase();
    return email.endsWith(`@${configured}`);
  });
}

function isOrgEmail(email) {
  const configured = String(config.zoho.organizationUserEmailDomain || '')
    .trim()
    .toLowerCase()
    .replace(/^@/, '');
  if (!configured) {
    return false;
  }
  return String(email || '')
    .trim()
    .toLowerCase()
    .endsWith(`@${configured}`);
}

function filterUsersToClientUsers(users) {
  const unique = new Map();
  for (const user of users) {
    const byFlag = Boolean(user.isClient);
    const byType = /client/i.test(String(user.userType || ''));
    const byRole = /client/i.test(String(user.role || ''));
    const byEmail = user.email ? !isOrgEmail(user.email) : false;
    if (!(byFlag || byType || byRole || byEmail)) {
      continue;
    }
    const key = `${user.id}::${String(user.email || '').toLowerCase()}::${String(user.displayName || '').toLowerCase()}`;
    if (!unique.has(key)) {
      unique.set(key, user);
    }
  }
  return Array.from(unique.values());
}

async function fetchFromZoho(query) {
  await ensureZohoCredentials();
  const endpoint = `${buildZohoProjectsBasePath()}/`;

  const range = 200;
  let index = 1;
  const maxPages = 50;
  const uniqueMap = new Map();

  try {
    for (let page = 0; page < maxPages; page += 1) {
      const params = {
        index,
        range
      };
      if (query) {
        params.search_term = query;
      }

      const response = await requestZohoGet(endpoint, params, true);
      const pageProjects = collectProjectList(response.data).map(normalizeProject);
      if (!pageProjects.length) {
        break;
      }

      let addedInThisPage = 0;
      for (const project of pageProjects) {
        const key = `${project.id}::${project.projectNumber}::${project.name}`;
        if (!uniqueMap.has(key)) {
          uniqueMap.set(key, project);
          addedInThisPage += 1;
        }
      }

      // Stop if server repeats the same page or returns fewer records than requested range.
      if (addedInThisPage === 0 || pageProjects.length < range) {
        break;
      }

      index += 1;
    }
  } catch (error) {
    throw new Error(toZohoErrorMessage(error));
  }

  const projects = Array.from(uniqueMap.values());
  const stageEnrichedProjects = await enrichProjectsWithDetailStatus(projects);

  if (!query) {
    return stageEnrichedProjects;
  }

  const needle = query.toLowerCase();
  return stageEnrichedProjects.filter((project) => {
    return [
      project.id,
      project.name,
      project.projectNumber,
      project.clientName,
      project.workOrderNo
    ]
      .join(' ')
      .toLowerCase()
      .includes(needle);
  });
}

async function fetchProjectUsersFromZoho(projectId) {
  await ensureZohoCredentials();
  const projectBase = buildZohoProjectsBasePath();
  const endpoints = [
    `${projectBase}/${projectId}/users/`,
    `${projectBase}/${projectId}/users`
  ];

  let lastError = null;
  for (const endpoint of endpoints) {
    try {
      const response = await requestZohoGet(endpoint, {}, true);
      const users = collectUserList(response.data).map(normalizeUser);
      const unique = new Map();
      for (const user of users) {
        if (!user.displayName) {
          continue;
        }
        const key = `${user.id}::${user.email}::${user.displayName}`;
        if (!unique.has(key)) {
          unique.set(key, user);
        }
      }
      return Array.from(unique.values());
    } catch (error) {
      lastError = error;
      // try fallback endpoint for 404 only
      if (!isAxiosStatus(error, 404)) {
        throw new Error(toZohoErrorMessage(error));
      }
    }
  }

  throw new Error(toZohoErrorMessage(lastError));
}

async function fetchPortalUsersFromZoho() {
  await ensureZohoCredentials();
  const usersBase = buildZohoPortalUsersBasePath();
  const endpoints = [`${usersBase}/`, usersBase];
  const unique = new Map();
  let endpointWorked = false;
  let lastError = null;

  for (const endpoint of endpoints) {
    let index = 1;
    try {
      for (let page = 0; page < 25; page += 1) {
        const response = await requestZohoGet(endpoint, { index, range: 200 }, true);
        endpointWorked = true;
        const users = collectUserList(response.data).map(normalizeUser);

        if (!users.length) {
          break;
        }

        for (const user of users) {
          if (!user.displayName) {
            continue;
          }
          const key = `${user.id}::${user.email}::${user.displayName}`;
          if (!unique.has(key)) {
            unique.set(key, user);
          }
        }

        if (users.length < 200) {
          break;
        }
        index += 1;
      }

      if (endpointWorked) {
        break;
      }
    } catch (error) {
      lastError = error;
      if (!isAxiosStatus(error, 404) && !isAxiosStatus(error, 400)) {
        throw new Error(toZohoErrorMessage(error));
      }
    }
  }

  if (!endpointWorked && lastError) {
    throw new Error(toZohoErrorMessage(lastError));
  }

  return filterUsersToOrgDomain(Array.from(unique.values()));
}

async function requestClientUsersFromEndpoints(projectId) {
  const projectBase = buildZohoProjectsBasePath();
  const endpoints = [
    { url: `${projectBase}/${projectId}/users/`, params: { user_type: 'client' } },
    { url: `${projectBase}/${projectId}/users`, params: { user_type: 'client' } },
    { url: `${projectBase}/${projectId}/users/`, params: { usertype: 'client' } },
    { url: `${projectBase}/${projectId}/users`, params: { usertype: 'client' } },
    { url: `${projectBase}/${projectId}/clientusers/`, params: {} },
    { url: `${projectBase}/${projectId}/clientusers`, params: {} },
    { url: `${projectBase}/${projectId}/clients/`, params: {} },
    { url: `${projectBase}/${projectId}/clients`, params: {} }
  ];

  let lastError = null;
  for (const endpoint of endpoints) {
    try {
      const response = await requestZohoGet(endpoint.url, endpoint.params, true);
      const users = collectUserList(response.data).map(normalizeUser);
      const clientUsers = filterUsersToClientUsers(users);
      if (clientUsers.length) {
        return clientUsers;
      }
    } catch (error) {
      lastError = error;
      if (!isAxiosStatus(error, 404) && !isAxiosStatus(error, 400)) {
        throw new Error(toZohoErrorMessage(error));
      }
    }
  }

  if (lastError) {
    return [];
  }
  return [];
}

async function fetchProjectClientUsersFromZoho(projectId) {
  await ensureZohoCredentials();
  const fromClientEndpoints = await requestClientUsersFromEndpoints(projectId);
  if (fromClientEndpoints.length) {
    return fromClientEndpoints;
  }

  // Fallback: derive client users from complete users list.
  const allUsers = await fetchProjectUsersFromZoho(projectId);
  return filterUsersToClientUsers(allUsers);
}

async function fetchProjectTasksFromZoho(projectId, options = {}) {
  await ensureZohoCredentials();
  const projectBase = buildZohoProjectsBasePath();
  const endpointCandidates = [`${projectBase}/${projectId}/tasks/`, `${projectBase}/${projectId}/tasks`];

  const query = String(options.query || '').trim().toLowerCase();
  const statusFilter = normalizeStageText(options.status || '').toLowerCase();
  const startIndex = Math.max(1, Number.parseInt(options.index || '1', 10) || 1);
  const range = Math.min(200, Math.max(1, Number.parseInt(options.range || '200', 10) || 200));
  const maxPages = 50;
  const uniqueMap = new Map();

  let endpointWorked = false;
  let lastError = null;

  for (const endpoint of endpointCandidates) {
    let index = startIndex;
    try {
      for (let page = 0; page < maxPages; page += 1) {
        // Some Zoho portals reject status query param on /tasks with 400 ("Given URL is wrong").
        // Fetch all tasks and apply status/query filtering locally for stable behavior.
        const response = await requestZohoGet(endpoint, { index, range }, true);
        endpointWorked = true;
        const pageTasks = collectTaskList(response.data)
          .map(normalizeTask)
          .filter((task) => task.id || task.name);

        if (!pageTasks.length) {
          break;
        }

        let addedInPage = 0;
        for (const task of pageTasks) {
          const key = `${task.id}::${task.name}`;
          if (!uniqueMap.has(key)) {
            uniqueMap.set(key, task);
            addedInPage += 1;
          }
        }

        if (addedInPage === 0 || pageTasks.length < range) {
          break;
        }

        index += 1;
      }

      // Stop after first working endpoint to avoid duplicate calls.
      break;
    } catch (error) {
      lastError = error;
      // fallback endpoint for variants returning 400/404
      if (!isAxiosStatus(error, 400) && !isAxiosStatus(error, 404)) {
        throw new Error(toZohoErrorMessage(error));
      }
    }
  }

  if (!endpointWorked && lastError) {
    throw new Error(toZohoErrorMessage(lastError));
  }

  let tasks = Array.from(uniqueMap.values());
  if (statusFilter) {
    tasks = tasks.filter((task) => String(task.status || '').toLowerCase().includes(statusFilter));
  }
  if (query) {
    tasks = tasks.filter((task) =>
      [task.id, task.name, task.status, task.ownerName, task.taskListName, task.milestoneName]
        .join(' ')
        .toLowerCase()
        .includes(query)
    );
  }
  return tasks;
}

async function resolveProjectReferenceToId(projectRef) {
  const raw = String(projectRef || '').trim();
  if (!raw) {
    return '';
  }

  if (/^\d{8,}$/.test(raw)) {
    return raw;
  }

  let matches = await fetchFromZoho(raw);
  const needle = raw.toLowerCase();
  const exactByRef = (projects) =>
    projects.find((project) => {
      return [project.id, project.projectNumber, project.name]
        .map((value) => String(value || '').toLowerCase())
        .includes(needle);
    });

  let exact = exactByRef(matches);

  if (!matches.length || !exact) {
    const allProjects = await fetchFromZoho('');
    const contains = allProjects.filter((project) => {
      return [project.id, project.projectNumber, project.name]
        .join(' ')
        .toLowerCase()
        .includes(needle);
    });
    matches = contains.length ? contains : allProjects;
    exact = exactByRef(matches);
  }

  const selected = exact || matches[0];
  const resolved = String(selected?.id || '').trim();
  return resolved || raw;
}

async function getProjects(query = '') {
  if (config.zoho.useMock) {
    if (!query) {
      return mockProjects;
    }

    const needle = query.toLowerCase();
    return mockProjects.filter((project) => {
      return [
        project.id,
        project.name,
        project.projectNumber,
        project.clientName,
        project.workOrderNo
      ]
        .join(' ')
        .toLowerCase()
        .includes(needle);
    });
  }

  return fetchFromZoho(query);
}

async function getProjectUsers(projectId) {
  const id = String(projectId || '').trim();
  if (!id) {
    throw new Error('Project ID is required to fetch Zoho project users.');
  }

  if (config.zoho.useMock) {
    return filterUsersToOrgDomain(mockProjectUsers[id] || []);
  }

  const resolvedId = await resolveProjectReferenceToId(id);
  const users = await fetchProjectUsersFromZoho(resolvedId);
  return filterUsersToOrgDomain(users);
}

async function getPortalUsers() {
  if (config.zoho.useMock) {
    return filterUsersToOrgDomain(mockPortalUsers);
  }

  return fetchPortalUsersFromZoho();
}

async function getProjectClientUsers(projectId) {
  const id = String(projectId || '').trim();
  if (!id) {
    throw new Error('Project ID is required to fetch Zoho client users.');
  }

  if (config.zoho.useMock) {
    return mockProjectClientUsers[id] || [];
  }

  const resolvedId = await resolveProjectReferenceToId(id);
  return fetchProjectClientUsersFromZoho(resolvedId);
}

// One project (name, stage, owner), for deciding whether it is a D03 project.
async function getProject(projectRef) {
  const id = String(projectRef || '').trim();
  if (!id) {
    return null;
  }
  if (config.zoho.useMock) {
    return mockProjects.find((project) => project.id === id || project.name === id) || null;
  }
  const resolvedId = await resolveProjectReferenceToId(id);
  const detail = await fetchProjectDetailFromZoho(resolvedId);
  return detail ? normalizeProject(detail) : null;
}

async function getProjectTasks(projectRef, options = {}) {
  const id = String(projectRef || '').trim();
  if (!id) {
    throw new Error('Project ID is required to fetch Zoho project tasks.');
  }

  if (config.zoho.useMock) {
    const query = String(options.query || '').trim().toLowerCase();
    const status = String(options.status || '').trim().toLowerCase();
    let tasks = mockRawProjectTasks[id] ? mockRawProjectTasks[id].map(normalizeTask) : mockProjectTasks[id] || [];
    if (status) {
      tasks = tasks.filter((task) => String(task.status || '').toLowerCase().includes(status));
    }
    if (query) {
      tasks = tasks.filter((task) =>
        [task.id, task.name, task.status, task.ownerName, task.taskListName]
          .join(' ')
          .toLowerCase()
          .includes(query)
      );
    }
    return tasks;
  }

  const resolvedId = await resolveProjectReferenceToId(id);
  return fetchProjectTasksFromZoho(resolvedId, options);
}

// Projects created before this date still sync and stay resolvable (for manual
// mapping of stragglers), but are hidden from the picker dropdown list. Filtering
// uses the normalized createdAt, sourced from Zoho's created_date_long (epoch-ms).
const PICKER_MIN_CREATED_MS = Date.parse('2026-01-01T00:00:00Z');

function projectCreatedMs(project) {
  const raw = asText(project && project.createdAt).trim();
  if (!raw) {
    return null;
  }
  if (/^\d{10,}$/.test(raw)) {
    return Number(raw);
  }
  const parsed = Date.parse(raw);
  return Number.isFinite(parsed) ? parsed : null;
}

function filterProjectsForPicker(projects) {
  return (Array.isArray(projects) ? projects : []).filter((project) => {
    const ms = projectCreatedMs(project);
    // Fail-open: keep projects with a missing/unparseable created date rather than
    // risk hiding a valid recent project.
    if (ms === null) {
      return true;
    }
    return ms >= PICKER_MIN_CREATED_MS;
  });
}

module.exports = {
  normalizeTask,
  getProject,
  getProjects,
  getPortalUsers,
  getProjectUsers,
  getProjectClientUsers,
  getProjectTasks,
  filterProjectsForPicker
};
