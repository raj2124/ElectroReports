// Zoho task plan for step 1 (Session 6D): which tasks of a D03 project are measurement tasks, which
// ESHA test each one is, and the planned quantity per test. Pure functions over the tasks as
// zohoClient.normalizeTask returns them, so the rules are testable without Zoho.
//
// Rules (agreed):
//   - Only [ESHA] tasks whose Category (UI label "Work Location") is Onsite are measurement tasks;
//     [IRT] tasks are not ours. A measurement task also needs its ESHA Test set.
//   - Planned Quantity is read from Onsite tasks only, and only when it is ONE clear number
//     ("170", "170 pits"). Anything else is left blank, shown as "Zoho says: ..." and warned about.
//   - A number in the task name is a hint ("Task name says 170 Pits"); it never fills the value.
//   - Two Onsite tasks with the same ESHA Test: equal quantities count once, different ones leave
//     the value blank with a warning. Nothing is ever summed.
//   - Custom fields are found in the task payload, by API name first and label second.

const config = require('./config');

// Our nine tests, as they are named in the ESHA Test pick list.
const ESHA_TESTS = [
  { id: 'soilResistivity', name: 'Soil Resistivity' },
  { id: 'electrodeResistance', name: 'Earth Electrode Resistance' },
  { id: 'continuityTest', name: 'Earth Continuity & Bonding' },
  { id: 'loopImpedanceTest', name: 'Loop Impedance' },
  { id: 'prospectiveFaultCurrent', name: 'Prospective Fault Current' },
  { id: 'riserIntegrityTest', name: 'Riser / Grid Integrity' },
  { id: 'groundResistance', name: 'Ground (Grid) Resistance' },
  { id: 'stepTouchVoltage', name: 'Step & Touch Voltage' },
  { id: 'towerFootingResistance', name: 'Tower Footing Resistance' }
];

// Labels (second choice) of the fields; the API names come from config / the payload.
const FIELD_LABELS = {
  category: ['Category', 'Work Location'],
  eshaTest: ['ESHA Test'],
  plannedQuantity: ['Planned Quantity']
};

const D03_NAME = /^\d{3}-D03_TS-\d{2}-[A-Z0-9]+$/i;
const UNITS = 'pits?|locations?|panels?|points?|towers?';
const text = (value) => String(value === undefined || value === null ? '' : value).trim();
const squash = (value) => text(value).toLowerCase().replace(/&amp;/g, '&').replace(/\band\b/g, '&').replace(/[^a-z0-9&]+/g, ' ').trim();

const isD03Name = (name) => D03_NAME.test(text(name));

// "170" and "170 pits" give 170. Everything else gives no number and says why; nothing is guessed.
function parseQuantity(raw) {
  const original = text(raw);
  if (!original) {
    return { status: 'blank', value: null, raw: '' };
  }
  const match = new RegExp(`^(\\d+)\\s*(?:${UNITS})?$`, 'i').exec(original);
  if (match && Number(match[1]) > 0) {
    return { status: 'ok', value: Number(match[1]), raw: original };
  }
  return { status: 'unclear', value: null, raw: original };
}

// "Earth Pit Resistance Testing - 170 Pits (...)" -> { value: 170, text: '170 Pits' }. Only when the
// name carries exactly one such number.
function nameHint(name) {
  const clean = text(name).replace(/&amp;/g, '&');
  const found = [...clean.matchAll(new RegExp(`(?<![\\d.\\-+/])(\\d+)\\s*(${UNITS})\\b`, 'gi'))];
  const distinct = [...new Set(found.map((match) => match[1]))];
  if (distinct.length !== 1 || Number(distinct[0]) <= 0) {
    return null;
  }
  return { value: Number(distinct[0]), text: `${found[0][1]} ${found[0][2]}` };
}

// The custom field of a task: API name first, label second. `override` is the configured API name.
function fieldOf(task, key, taskFields = config.zoho.taskFields) {
  const fields = Array.isArray(task && task.customFields) ? task.customFields : [];
  const apiName = text(taskFields && taskFields[key]).toLowerCase();
  const labels = FIELD_LABELS[key].map((label) => label.toLowerCase());
  return (apiName && fields.find((field) => text(field.apiName).toLowerCase() === apiName)) || fields.find((field) => labels.includes(text(field.label).toLowerCase())) || null;
}

function matchTest(value) {
  const wanted = squash(value);
  return wanted ? ESHA_TESTS.find((test) => squash(test.name) === wanted) || null : null;
}

function viewOf(task, taskFields) {
  const category = fieldOf(task, 'category', taskFields);
  const test = fieldOf(task, 'eshaTest', taskFields);
  const planned = fieldOf(task, 'plannedQuantity', taskFields);
  const esha = /^\s*\[ESHA\]/i.test(task.name);
  const irt = /^\s*\[IRT\]/i.test(task.name);
  const onsite = squash(category && category.value) === 'onsite';
  const matched = matchTest(test && test.value);
  return {
    esha,
    irt,
    onsite,
    category: text(category && category.value),
    eshaTestRaw: text(test && test.value),
    testId: matched ? matched.id : '',
    plannedRaw: text(planned && planned.value),
    apiNames: { category: category && category.apiName, eshaTest: test && test.apiName, plannedQuantity: planned && planned.apiName }
  };
}

// The plan for one project. `d03` is false for any other project: tasks are listed, but no
// measurement task, test or quantity is derived (the custom fields are for D03 projects only).
function buildTaskPlan(tasks, { d03 = true, taskFields = config.zoho.taskFields } = {}) {
  const list = Array.isArray(tasks) ? tasks : [];
  const warnings = [];
  const rows = list.map((task) => {
    const view = viewOf(task, taskFields);
    const row = {
      id: task.id,
      name: text(task.name),
      taskListId: text(task.taskListId),
      taskList: text(task.taskListName) || 'Tasks',
      owners: Array.isArray(task.owners) ? task.owners : [],
      ownerName: text(task.ownerName),
      startDate: text(task.startDate),
      dueDate: text(task.dueDate),
      status: text(task.status),
      closed: task.closed === true,
      duration: text(task.duration),
      category: view.category,
      esha: view.esha,
      irt: view.irt,
      onsite: view.onsite,
      eshaTest: view.eshaTestRaw,
      testId: view.testId,
      plannedRaw: view.plannedRaw,
      measurement: false,
      selected: false,
      warnings: []
    };
    if (d03 && view.esha && view.onsite) {
      if (!view.eshaTestRaw) {
        row.warnings.push('Onsite task without an ESHA Test: not selected.');
      } else if (!view.testId) {
        row.warnings.push(`Unknown ESHA Test value "${view.eshaTestRaw}": not selected.`);
      } else {
        row.measurement = true;
        row.selected = true;
      }
    }
    row.warnings.forEach((message) => warnings.push(`${row.name}: ${message}`));
    return { row, view };
  });

  const tests = {};
  if (d03) {
    ESHA_TESTS.forEach((test) => {
      const group = rows.filter(({ row }) => row.measurement && row.testId === test.id);
      if (!group.length) {
        return;
      }
      const parsed = group.map(({ row }) => ({ row, quantity: parseQuantity(row.plannedRaw) }));
      const unclear = parsed.filter((entry) => entry.quantity.status === 'unclear');
      const clear = [...new Set(parsed.filter((entry) => entry.quantity.status === 'ok').map((entry) => entry.quantity.value))];
      const result = { testId: test.id, name: test.name, taskIds: group.map(({ row }) => row.id), planned: null, source: '', zohoSays: '', hint: '', warnings: [] };
      if (unclear.length) {
        result.zohoSays = [...new Set(unclear.map((entry) => entry.quantity.raw))].join(' / ');
        result.warnings.push(`Planned Quantity "${result.zohoSays}" on ${unclear.map((entry) => entry.row.name).join('; ')} is not one clear number: left blank.`);
      } else if (clear.length > 1) {
        result.warnings.push(`${test.name}: Onsite tasks give different planned quantities (${clear.join(', ')}): left blank, not summed.`);
      } else if (clear.length === 1) {
        result.planned = clear[0];
        result.source = 'zoho';
      }
      if (result.planned === null && !unclear.length) {
        const hints = parsed.filter((entry) => entry.quantity.status === 'blank').map((entry) => nameHint(entry.row.name)).filter(Boolean);
        const distinct = [...new Map(hints.map((hint) => [hint.text.toLowerCase(), hint])).values()];
        if (distinct.length) {
          result.hint = `Task name says ${distinct.map((hint) => hint.text).join(' / ')}`;
        }
      }
      result.warnings.forEach((message) => warnings.push(message));
      tests[test.id] = result;
    });
  }

  const taskLists = [];
  rows.forEach(({ row }) => {
    let entry = taskLists.find((candidate) => candidate.name === row.taskList && candidate.id === row.taskListId);
    if (!entry) {
      entry = { id: row.taskListId, name: row.taskList, taskIds: [] };
      taskLists.push(entry);
    }
    entry.taskIds.push(row.id);
  });
  const discovered = { category: '', eshaTest: '', plannedQuantity: '' };
  rows.forEach(({ view }) => Object.keys(discovered).forEach((key) => {
    discovered[key] = discovered[key] || text(view.apiNames[key]) || '';
  }));
  Object.keys(discovered).forEach((key) => {
    discovered[key] = discovered[key] || text(taskFields && taskFields[key]);
  });

  return { d03, tasks: rows.map(({ row }) => row), taskLists, tests, warnings, fields: discovered };
}

// Full names for the owners (the task payload carries a short name such as "anup"), from the project
// users, matched by id and then e-mail.
function nameOwners(plan, users) {
  const list = Array.isArray(users) ? users : [];
  plan.tasks.forEach((task) => {
    task.owners = task.owners.map((owner) => {
      const user = list.find((entry) => (owner.id && text(entry.id) === owner.id) || (owner.email && text(entry.email).toLowerCase() === owner.email));
      return { ...owner, fullName: (user && text(user.displayName)) || owner.name };
    });
    task.ownerName = task.owners.map((owner) => owner.fullName).filter(Boolean).join(', ') || task.ownerName;
  });
  return plan;
}

module.exports = { ESHA_TESTS, FIELD_LABELS, isD03Name, parseQuantity, nameHint, fieldOf, matchTest, buildTaskPlan, nameOwners };
