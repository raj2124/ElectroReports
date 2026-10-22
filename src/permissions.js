// Who may do what (Session 6E). Enforced on the server: the browser only hides buttons.
//
// Roles are stored per user (several per person are fine); admin implies every role. Two layers:
//   1. ROUTE_RULES: which role a route needs. Anything not listed that changes data is admin-only
//      (default deny), so a new route cannot be left open by accident.
//   2. Report rules: ownership (a daily is edited by the person who created it, or an admin; a report
//      with no creator is admin-only) and the review gate: the person who PREPARED a report can never
//      check or approve it. An admin can override that with a typed reason, recorded on the transition.
//   A checker who is also the approver is allowed; the store reports it as a warning.

const { getReviewStatus } = require('./reportCollectionModel');
const { currentRound, lastOf, transitionAction } = require('./review');

const ROLES = ['engineer', 'reviewer', 'approver', 'admin'];
const ANY = 'any';
const OVERRIDE_MIN_LENGTH = 10;

const hasRole = (user, ...roles) => Boolean(user) && (user.roles.includes('admin') || roles.some((role) => user.roles.includes(role)));

// [method, path pattern, roles]. First match wins. ANY = any signed-in user.
const ROUTE_RULES = [
  ['*', /^\/api\/admin(\/|$)/, ['admin']],
  ['POST', /^\/api\/zoho\/poll$/, ['admin']],
  ['POST', /^\/api\/schedule\/refresh$/, [ANY]],

  ['POST', /^\/api\/reports$/, ['engineer']],
  ['PUT', /^\/api\/reports\/[^/]+$/, ['engineer']],
  ['DELETE', /^\/api\/reports\/[^/]+$/, ['engineer']],
  ['POST', /^\/api\/reports\/[^/]+\/ai\/generate$/, ['engineer']],
  ['POST', /^\/api\/reports\/[^/]+\/pdf$/, [ANY]],
  ['PUT', /^\/api\/reports\/[^/]+\/review\/ticks$/, ['reviewer']],
  ['POST', /^\/api\/reports\/[^/]+\/review\/submit$/, ['engineer']],
  ['POST', /^\/api\/reports\/[^/]+\/review\/check$/, ['reviewer']],
  ['POST', /^\/api\/reports\/[^/]+\/review\/approve$/, ['approver']],
  ['POST', /^\/api\/reports\/[^/]+\/review\/return$/, ['reviewer', 'approver']],
  ['POST', /^\/api\/reports\/[^/]+\/review\/[^/]+$/, ['admin']],
  ['POST', /^\/api\/reports\/[^/]+\/reissue$/, ['approver']],
  ['PUT', /^\/api\/reports\/[^/]+\/references$/, ['reviewer']],
  ['PUT', /^\/api\/reports\/[^/]+\/client-logo$/, ['approver']],
  ['PUT', /^\/api\/client-logo$/, ['approver']],
  ['DELETE', /^\/api\/client-logo$/, ['approver']],

  ['POST', /^\/api\/projects$/, ['engineer']],
  ['PUT', /^\/api\/projects\/[^/]+$/, ['engineer']],
  ['POST', /^\/api\/projects\/[^/]+\/(days|final\/assemble|zoho-refresh)$/, ['engineer']],

  ['POST', /^\/api\/sheets\/print$/, ['engineer']],
  ['POST', /^\/api\/ocr\/[^/]+\/preview$/, ['engineer']],
  ['POST', /^\/api\/assessments$/, [ANY]],
  ['POST', /^\/api\/notices\/[^/]+\/(answer|dismiss)$/, [ANY]],
  ['PUT', /^\/api\/profile$/, [ANY]]
];

// The roles a request needs, or null when no rule covers it (writes: default deny to admin).
function ruleFor(method, pathname) {
  const verb = String(method).toUpperCase();
  const rule = ROUTE_RULES.find(([ruleMethod, pattern]) => (ruleMethod === '*' || ruleMethod === verb) && pattern.test(pathname));
  if (rule) {
    return rule[2];
  }
  return verb === 'GET' || verb === 'HEAD' ? [ANY] : ['admin'];
}

function mayCall(user, method, pathname) {
  const roles = ruleFor(method, pathname);
  return roles.includes(ANY) || hasRole(user, ...roles);
}

// ---- Reports ------------------------------------------------------------------------------------
const lower = (value) => String(value || '').trim().toLowerCase();

// The people who prepared a report: its creator, and whoever submitted it for review.
function isPreparer(store, user, report) {
  if (report.createdBy && report.createdBy === user.id) {
    return true;
  }
  const submits = (store.listTransitions(report.id) || []).map((row) => ({ ...row, action: transitionAction(row) })).filter((row) => row.action === 'submit');
  return submits.some((row) => (row.userId ? row.userId === user.id : lower(row.actor) === lower(user.name)));
}

// Edit / delete of a report: its creator, or an admin. A report with no creator is admin-only.
function mayEditReport(user, report) {
  return hasRole(user, 'admin') || Boolean(report.createdBy && report.createdBy === user.id);
}

function mayOverride(user, reason) {
  return hasRole(user, 'admin') && String(reason || '').trim().length >= OVERRIDE_MIN_LENGTH;
}

// The review step `action` on `report` by `user`, after the role rule passed. Returns
// { allowed, code, message, canOverride, overridden }.
function reviewGate(store, user, action, report, { overrideReason = '' } = {}) {
  if (action === 'submit') {
    return mayEditReport(user, report)
      ? { allowed: true }
      : { allowed: false, code: 'NOT_OWNER', message: report.createdBy ? 'Only the engineer who created this report (or an admin) can submit it.' : 'This report has no creator: only an admin can submit it.', canOverride: false };
  }
  if (action === 'check' || action === 'approve' || action === 'ticks') {
    if (isPreparer(store, user, report)) {
      const verb = action === 'approve' ? 'approve' : 'check';
      if (mayOverride(user, overrideReason)) {
        return { allowed: true, overridden: true, reason: String(overrideReason).trim() };
      }
      return {
        allowed: false,
        code: 'SELF_REVIEW_BLOCKED',
        message: `You prepared this report, so you cannot ${verb} it.${hasRole(user, 'admin') ? ` An admin can override with a reason of at least ${OVERRIDE_MIN_LENGTH} characters.` : ''}`,
        canOverride: hasRole(user, 'admin')
      };
    }
  }
  return { allowed: true };
}

// What the review screen may offer this user for this report.
function reviewAccess(store, user, report) {
  const status = getReviewStatus(report);
  const preparer = isPreparer(store, user, report);
  const admin = hasRole(user, 'admin');
  const round = currentRound((store.listTransitions(report.id) || []).map((row) => ({ ...row, action: transitionAction(row) })));
  return {
    submit: status === 'draft' && hasRole(user, 'engineer') && mayEditReport(user, report),
    check: status === 'in_review' && !lastOf(round, 'check') && hasRole(user, 'reviewer') && (!preparer || admin),
    approve: status === 'in_review' && Boolean(lastOf(round, 'check')) && hasRole(user, 'approver') && (!preparer || admin),
    return: status === 'in_review' && hasRole(user, 'reviewer', 'approver'),
    reissue: status === 'approved' && hasRole(user, 'approver'),
    tick: status !== 'approved' && hasRole(user, 'reviewer') && (!preparer || admin),
    isPreparer: preparer,
    needsOverride: preparer && admin,
    edit: mayEditReport(user, report)
  };
}

module.exports = { ROLES, ANY, OVERRIDE_MIN_LENGTH, ROUTE_RULES, hasRole, ruleFor, mayCall, isPreparer, mayEditReport, mayOverride, reviewGate, reviewAccess };
