// Standards Acceptance List loader (ET-D03-REG-002).
//
// Every limit, formula and classification used in an ESHA report comes from the
// project's resolved profile. A profile is ETPL-DEFAULT, overridden only by the
// criteria a client profile lists. Each resolved criterion carries its value or
// rule, basis, source and clause so the report can print exactly what was used.

const fs = require('fs');
const path = require('path');

const REGISTER_PATH = path.join(__dirname, '..', 'config', 'ET-D03-REG-002_Standards_Acceptance_List.json');
const DEFAULT_PROFILE_ID = 'ETPL-DEFAULT';
// Copy-me scaffolding in the register; its values are placeholders, never a real profile.
const TEMPLATE_PROFILE_IDS = new Set(['CLIENT-TEMPLATE']);

const BASIS_LABELS = Object.freeze({
  standard: 'Standard',
  etpl_acceptance: 'ETPL acceptance criterion',
  client_spec: 'Client specification'
});

const registerCache = new Map();

function loadRegister(filePath = REGISTER_PATH) {
  if (registerCache.has(filePath)) {
    return registerCache.get(filePath);
  }
  const register = JSON.parse(fs.readFileSync(filePath, 'utf8'));
  const base = register?.profiles?.[DEFAULT_PROFILE_ID]?.criteria;
  if (!Array.isArray(base) || !base.length) {
    throw new Error(`Acceptance register ${filePath} has no ${DEFAULT_PROFILE_ID} criteria.`);
  }
  const ids = new Set();
  base.forEach((criterion) => {
    if (!criterion.id || ids.has(criterion.id)) {
      throw new Error(`Acceptance register has a missing or duplicate criterion id: ${criterion.id}`);
    }
    ids.add(criterion.id);
  });
  Object.entries(register.profiles).forEach(([profileId, profile]) => {
    if (profileId === DEFAULT_PROFILE_ID) {
      return;
    }
    (profile.criteria || []).forEach((override) => {
      if (!ids.has(override.id)) {
        throw new Error(`Acceptance profile ${profileId} overrides unknown criterion ${override.id}.`);
      }
    });
  });
  registerCache.set(filePath, register);
  return register;
}

function listProfiles(options = {}) {
  const register = loadRegister(options.registerPath);
  return Object.entries(register.profiles)
    .filter(([profileId]) => !TEMPLATE_PROFILE_IDS.has(profileId))
    .map(([profileId, profile]) => ({
      id: profileId,
      description: profile.description || '',
      clientSpecDocument: profile.client_spec_document || ''
    }));
}

// Resolve a profile id to its full criteria set. An empty id means the default.
// An unknown or template id throws: silently falling back would print a source
// that does not match the project's contract.
function resolveProfile(profileId, options = {}) {
  const register = loadRegister(options.registerPath);
  const requested = String(profileId === undefined || profileId === null ? '' : profileId).trim() || DEFAULT_PROFILE_ID;
  if (TEMPLATE_PROFILE_IDS.has(requested)) {
    throw new Error(`"${requested}" is a template, not a selectable acceptance profile.`);
  }
  const profile = register.profiles[requested];
  if (!profile) {
    throw new Error(`Unknown acceptance profile "${requested}".`);
  }

  const criteria = {};
  register.profiles[DEFAULT_PROFILE_ID].criteria.forEach((criterion) => {
    criteria[criterion.id] = { ...criterion, profileId: DEFAULT_PROFILE_ID, overridden: false };
  });
  if (requested !== DEFAULT_PROFILE_ID) {
    (profile.criteria || []).forEach((override) => {
      criteria[override.id] = { ...criteria[override.id], ...override, profileId: requested, overridden: true };
    });
  }

  return {
    profileId: requested,
    description: profile.description || '',
    clientSpecDocument: profile.client_spec_document || '',
    register: { document: register.document, revision: register.revision },
    criteria,
    get(id) {
      const criterion = criteria[id];
      if (!criterion) {
        throw new Error(`Acceptance criterion ${id} is not in profile ${requested}.`);
      }
      return criterion;
    },
    // Numeric value of a limit-type criterion (CONT-LIMIT, LOOP-LIMIT, ...).
    value(id) {
      const criterion = this.get(id);
      if (!Number.isFinite(criterion.value)) {
        throw new Error(`Acceptance criterion ${id} in profile ${requested} has no numeric value.`);
      }
      return criterion.value;
    }
  };
}

// Row for the report's Basis of Assessment table.
function describeCriterion(criterion) {
  const value = Number.isFinite(criterion.value) ? ` (value: ${criterion.value} ${criterion.unit || ''})`.replace(/ \)$/, ')') : '';
  return {
    id: criterion.id,
    parameter: criterion.parameter,
    ruleUsed: `${criterion.rule}${value}`,
    basis: criterion.basis,
    basisLabel: BASIS_LABELS[criterion.basis] || criterion.basis,
    source: criterion.source,
    clause: criterion.clause,
    profileId: criterion.profileId,
    overridden: criterion.overridden
  };
}

module.exports = {
  DEFAULT_PROFILE_ID,
  BASIS_LABELS,
  loadRegister,
  listProfiles,
  resolveProfile,
  describeCriterion
};
