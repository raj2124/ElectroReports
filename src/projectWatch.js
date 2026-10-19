// New D03 project watch (Session 6D). Zoho is polled for D03 projects that have no workspace yet; the
// owners of the project's ESHA Onsite tasks (the project owner when there are none) get an in-app
// notice with two questions: phones/tablets allowed on site, and mobile data available. The answers
// are stored on the project and set the default data-collection mode. A notice is never recreated
// once it exists, whether it was answered, closed by a colleague's answer or dismissed.

const defaultConfig = require('./config');
const { isD03Name, buildTaskPlan, nameOwners } = require('./zohoTasks');
const { ensureProject, findProject } = require('./workspace');

const MIN_CREATED_MS = Date.parse('2026-01-01T00:00:00Z');
const text = (value) => String(value === undefined || value === null ? '' : value).trim();

function createdMs(project) {
  const raw = text(project && project.createdAt);
  if (/^\d{10,}$/.test(raw)) return Number(raw);
  const parsed = Date.parse(raw);
  return Number.isFinite(parsed) ? parsed : null;
}

// Phones/tablets allowed AND mobile data available -> enter on site; otherwise printed sheets.
function dataCollectionModeFor({ phonesAllowed, mobileData }) {
  return phonesAllowed === true && mobileData === true ? 'On-site entry' : 'Printed sheets';
}

const personKey = (email, name) => text(email).toLowerCase() || text(name).toLowerCase();

async function recipientsFor(project, zoho) {
  const tasks = await zoho.getProjectTasks(project.id);
  const plan = buildTaskPlan(tasks, { d03: true });
  let users = [];
  try {
    users = await zoho.getProjectUsers(project.id);
  } catch (_error) {
    users = [];
  }
  nameOwners(plan, users);
  const onsite = plan.tasks.filter((task) => task.esha && task.onsite);
  const byKey = new Map();
  onsite.forEach((task) => task.owners.forEach((owner) => {
    const key = personKey(owner.email, owner.fullName);
    if (key && !byKey.has(key)) byKey.set(key, { key, name: owner.fullName || owner.name, reason: 'onsite_owner' });
  }));
  if (!byKey.size) {
    const key = personKey(project.ownerEmail, project.ownerName);
    if (key) byKey.set(key, { key, name: project.ownerName, reason: 'project_owner' });
  }
  return { recipients: [...byKey.values()], onsiteTasks: onsite.map((task) => task.name) };
}

// One poll. Returns what it looked at and which notices it created.
async function pollNewProjects({ store, zoho, config = defaultConfig }) {
  const source = zoho || require('./zohoClient');
  const statuses = config.zoho.notifyStatuses.map((status) => status.toLowerCase());
  const projects = await source.getProjects();
  const candidates = projects.filter((project) => {
    const created = createdMs(project);
    return isD03Name(project.name) && created !== null && created >= MIN_CREATED_MS && statuses.includes(text(project.stage).toLowerCase()) && !findProject(store, { zohoProjectId: project.id });
  });
  const created = [];
  for (const project of candidates) {
    const { recipients, onsiteTasks } = await recipientsFor(project, source);
    recipients.forEach((recipient) => {
      const notice = store.addNotice({
        zohoProjectId: project.id,
        recipientKey: recipient.key,
        recipientName: recipient.name,
        payload: { projectName: project.name, projectNo: project.projectNumber || project.name, clientName: project.clientName || '', zohoProjectId: project.id, ownerName: project.ownerName || '', stage: project.stage || '', reason: recipient.reason, onsiteTasks }
      });
      if (notice) created.push({ noticeId: notice.id, project: project.name, recipient: recipient.name, reason: recipient.reason });
    });
  }
  return { checked: projects.length, candidates: candidates.map((project) => project.name), created };
}

function answerError(message, code) {
  return Object.assign(new Error(message), { code });
}

// Records the two answers: creates the project's workspace if it has none, stores the answers and the
// default mode on it, and closes the other owners' open notices for the project.
function answerNotice(store, id, { phonesAllowed, mobileData, by = '' }) {
  const notice = store.getNotice(id);
  if (!notice) throw answerError('Notice not found.', 'NOT_FOUND');
  if (notice.status !== 'open') throw answerError('This notice has already been answered or dismissed.', 'NOT_OPEN');
  if (typeof phonesAllowed !== 'boolean' || typeof mobileData !== 'boolean') {
    throw answerError('Answer both questions (yes or no).', 'ANSWERS_REQUIRED');
  }
  const info = notice.payload;
  const { project } = ensureProject(store, {
    projectNo: info.projectNo,
    clientName: info.clientName || info.projectNo,
    zohoProjectId: info.zohoProjectId,
    zohoProjectName: info.projectName,
    zohoProjectOwner: info.ownerName,
    zohoProjectStage: info.stage
  });
  const mode = dataCollectionModeFor({ phonesAllowed, mobileData });
  const answeredAt = new Date().toISOString();
  store.updateProject(project.id, { siteAccess: { phonesAllowed, mobileData, answeredBy: by || notice.recipientName, answeredAt }, dataCollectionMode: mode });
  store.resolveNotice(id, { status: 'answered', answers: { phonesAllowed, mobileData }, by: by || notice.recipientName });
  store.closeOpenNotices(info.zohoProjectId, { exceptId: id, by: by || notice.recipientName });
  return { projectId: project.id, dataCollectionMode: mode };
}

function dismissNotice(store, id, by = '') {
  const notice = store.getNotice(id);
  if (!notice) throw answerError('Notice not found.', 'NOT_FOUND');
  if (notice.status !== 'open') throw answerError('This notice has already been answered or dismissed.', 'NOT_OPEN');
  return store.resolveNotice(id, { status: 'dismissed', by });
}

// Timer. `intervalMinutes` 0 turns it off. Errors are logged without detail that could hold a token.
function startProjectWatch({ store, config = defaultConfig, log = console }) {
  const minutes = config.zoho.pollMinutes;
  if (!(minutes > 0)) return { stop() {} };
  const run = () => pollNewProjects({ store, config }).catch((error) => log.error(`Zoho project watch failed: ${error && error.message ? String(error.message).slice(0, 160) : 'unknown error'}`));
  const first = setTimeout(run, 10000);
  const timer = setInterval(run, minutes * 60 * 1000);
  first.unref && first.unref();
  timer.unref && timer.unref();
  return { stop() { clearTimeout(first); clearInterval(timer); } };
}

module.exports = { MIN_CREATED_MS, dataCollectionModeFor, pollNewProjects, answerNotice, dismissNotice, startProjectWatch };
