// Zoho task plan for step 1 (Session 6D): quantity parsing, name hints, duplicate Onsite tasks,
// field discovery by API name then label, D03-only, and the mock project built from the real
// portal shapes (056-D03_TS-26-CLL). No network: ZOHO_USE_MOCK.
process.env.ZOHO_USE_MOCK = 'true';
const assert = require('assert');
const { parseQuantity, nameHint, buildTaskPlan, nameOwners, isD03Name, matchTest, ESHA_TESTS, fieldOf } = require('../src/zohoTasks');
const zoho = require('../src/zohoClient');

let passed = 0;
async function check(name, fn) {
  await fn();
  passed += 1;
  console.log(`  ok - ${name}`);
}

const field = (label, value, apiName = `UDF_${label.replace(/\W/g, '').toUpperCase()}`) => ({ label, apiName, value });
// A task as zohoClient.normalizeTask returns it.
const task = (id, name, { onsite = true, test = '', qty, owners = [{ id: '1', name: 'anup', email: 'anup@elegrow.com' }], extra = [] } = {}) => ({
  id,
  name,
  taskListId: 'tl1',
  taskListName: onsite ? 'Onsite' : 'Offsite',
  owners,
  ownerName: owners[0] && owners[0].name,
  status: 'Planned',
  startDate: '10-19-2026',
  dueDate: '10-23-2026',
  customFields: [field('Work Location', onsite ? 'Onsite' : 'Offsite', 'UDF_CHAR81'), ...(test ? [field('ESHA Test', test)] : []), ...(qty !== undefined ? [field('Planned Quantity', qty)] : []), ...extra]
});
const planOf = (tasks) => buildTaskPlan(tasks, { d03: true, taskFields: { category: 'category', eshaTest: '', plannedQuantity: '' } });
const earth = (qty, id = 'e1', name = '[ESHA] Earth Pit Resistance Testing - 170 Pits (Fall-of-Potential / 3-Pole)') => task(id, name, { test: 'Earth Electrode Resistance', qty });

(async () => {
  await check('quantity: one clear number is read, with or without a unit', () => {
    assert.deepStrictEqual(parseQuantity('170'), { status: 'ok', value: 170, raw: '170' });
    assert.deepStrictEqual(parseQuantity('170 pits'), { status: 'ok', value: 170, raw: '170 pits' });
    assert.strictEqual(parseQuantity(' 4 Locations ').value, 4);
    assert.strictEqual(parseQuantity('1 tower').value, 1);
  });

  await check('quantity: "approx 30", "30-35", "25 + 5", blank and the unclear rest are never guessed', () => {
    assert.deepStrictEqual(parseQuantity('approx 30'), { status: 'unclear', value: null, raw: 'approx 30' });
    ['30-35', '25 + 5', '30 to 35', '170 nos', '17.5', '1,200', '0', '10 pits 5 panels', 'about 30 pits'].forEach((raw) => {
      assert.strictEqual(parseQuantity(raw).value, null, raw);
      assert.strictEqual(parseQuantity(raw).status, 'unclear', raw);
    });
    assert.deepStrictEqual(parseQuantity(''), { status: 'blank', value: null, raw: '' });
    assert.strictEqual(parseQuantity(undefined).status, 'blank');
    assert.strictEqual(parseQuantity('   ').status, 'blank');
  });

  await check('plan: a clear quantity fills the test, marked as from Zoho', () => {
    const plan = planOf([earth('170 pits')]);
    assert.deepStrictEqual([plan.tests.electrodeResistance.planned, plan.tests.electrodeResistance.source, plan.tests.electrodeResistance.warnings], [170, 'zoho', []]);
  });

  await check('plan: an unclear quantity stays blank, shows "Zoho says" and warns', () => {
    ['approx 30', '30-35', '25 + 5'].forEach((raw) => {
      const plan = planOf([earth(raw)]);
      const entry = plan.tests.electrodeResistance;
      assert.strictEqual(entry.planned, null, raw);
      assert.strictEqual(entry.zohoSays, raw);
      assert.strictEqual(entry.hint, '', 'no hint when Zoho has a (unclear) value');
      assert.ok(entry.warnings.length === 1 && entry.warnings[0].includes(raw));
      assert.ok(plan.warnings.some((warning) => warning.includes(raw)));
    });
  });

  await check('plan: a blank quantity with a number in the task name is a hint only, never filled', () => {
    const plan = planOf([earth('')]);
    const entry = plan.tests.electrodeResistance;
    assert.strictEqual(entry.planned, null);
    assert.strictEqual(entry.hint, 'Task name says 170 Pits');
    assert.strictEqual(planOf([earth(undefined)]).tests.electrodeResistance.hint, 'Task name says 170 Pits');
    // no hint when the name has no, or two different, quantities
    assert.strictEqual(planOf([earth('', 'e1', '[ESHA] Earth Pit Resistance Testing')]).tests.electrodeResistance.hint, '');
    assert.strictEqual(planOf([earth('', 'e1', '[ESHA] Testing - 10 Pits and 4 Locations')]).tests.electrodeResistance.hint, '');
    assert.strictEqual(planOf([earth('', 'e1', '[ESHA] Testing - 30-35 Pits')]).tests.electrodeResistance.hint, '');
  });

  await check('plan: two Onsite tasks, same ESHA Test: equal quantities count once; different leave it blank with a warning (never summed)', () => {
    const equal = planOf([earth('170', 'e1'), earth('170 pits', 'e2', '[ESHA] Earth Pit Resistance Testing - second visit')]);
    assert.strictEqual(equal.tests.electrodeResistance.planned, 170);
    assert.deepStrictEqual(equal.tests.electrodeResistance.taskIds, ['e1', 'e2']);
    const different = planOf([earth('100', 'e1'), earth('70', 'e2', '[ESHA] Earth Pit Resistance Testing - second visit')]);
    assert.strictEqual(different.tests.electrodeResistance.planned, null);
    assert.ok(/different planned quantities \(100, 70\)/.test(different.tests.electrodeResistance.warnings[0]));
    assert.ok(different.warnings.length === 1);
    // a clear value next to an unclear one is not resolved by guessing either
    const mixed = planOf([earth('100', 'e1'), earth('approx 70', 'e2', '[ESHA] Earth Pit Resistance Testing - second visit')]);
    assert.strictEqual(mixed.tests.electrodeResistance.planned, null);
    assert.strictEqual(mixed.tests.electrodeResistance.zohoSays, 'approx 70');
  });

  await check('plan: only Onsite [ESHA] tasks with an ESHA Test are measurement tasks; [IRT] and Offsite tasks are not', () => {
    const plan = planOf([
      earth('170'),
      task('o1', '[ESHA] Pit Resistance Data Tabulation', { onsite: false, test: 'Earth Electrode Resistance', qty: '999' }),
      task('i1', '[IRT] Thermography Survey', { test: 'Loop Impedance', qty: '5' }),
      task('n1', '[ESHA] Onsite work with no ESHA Test'),
      task('u1', '[ESHA] Onsite work, odd test', { test: 'Thermal Imaging' })
    ]);
    assert.deepStrictEqual(plan.tasks.filter((entry) => entry.measurement).map((entry) => entry.id), ['e1']);
    assert.deepStrictEqual(plan.tasks.filter((entry) => entry.selected).map((entry) => entry.id), ['e1']);
    assert.strictEqual(plan.tests.electrodeResistance.planned, 170, 'a quantity on an Offsite task is not read');
    assert.strictEqual(plan.tests.loopImpedanceTest, undefined);
    assert.ok(plan.warnings.some((warning) => /no ESHA Test|without an ESHA Test/.test(warning)));
    assert.ok(plan.warnings.some((warning) => warning.includes('Unknown ESHA Test value "Thermal Imaging"')), 'an unknown value is a warning, not an error');
  });

  await check('ESHA Test names: all nine of ours match; case and "and"/"&" do not matter', () => {
    assert.strictEqual(ESHA_TESTS.length, 9);
    ESHA_TESTS.forEach((test) => assert.strictEqual(matchTest(test.name).id, test.id));
    assert.strictEqual(matchTest('earth continuity and bonding').id, 'continuityTest');
    assert.strictEqual(matchTest('Step and Touch Voltage').id, 'stepTouchVoltage');
    assert.strictEqual(matchTest('Something else'), null);
  });

  await check('fields are found by API name first, label second; the env override names a field', () => {
    const t = { customFields: [{ label: 'Work Location', apiName: 'UDF_CHAR81', value: 'Onsite' }, { label: 'Category', apiName: 'category', value: 'Offsite' }] };
    assert.strictEqual(fieldOf(t, 'category', { category: 'category' }).value, 'Offsite', 'API name wins over the label');
    assert.strictEqual(fieldOf(t, 'category', { category: 'UDF_CHAR81' }).value, 'Onsite');
    assert.strictEqual(fieldOf({ customFields: [{ label: 'Work Location', apiName: 'UDF_CHAR81', value: 'Onsite' }] }, 'category', { category: 'category' }).value, 'Onsite', 'falls back to the label (v1 payload)');
    assert.strictEqual(fieldOf({ customFields: [{ label: 'x', apiName: 'task_cf_0006', value: '9 pits' }] }, 'plannedQuantity', { plannedQuantity: 'task_cf_0006' }).value, '9 pits');
    assert.strictEqual(fieldOf({ customFields: [] }, 'eshaTest', {}), null);
  });

  await check('only D03 projects use the custom fields; the name pattern is NNN-D03_TS-YY-CLIENT', () => {
    ['056-D03_TS-26-CLL', '024-D03_TS-26-GFL', '004-D03_TS-26-APL'].forEach((name) => assert.ok(isD03Name(name), name));
    ['Corporate Website Revamp', '025-D02_PSS-26-SPPL', '561_03_26_TS_AIL_EA_NEO', '56-D03_TS-26-CLL'].forEach((name) => assert.ok(!isD03Name(name), name));
    const plan = buildTaskPlan([earth('170')], { d03: false });
    assert.deepStrictEqual([plan.d03, Object.keys(plan.tests), plan.tasks.filter((entry) => entry.selected).length], [false, [], 0]);
    assert.strictEqual(plan.tasks.length, 1, 'the tasks are still listed');
  });

  await check('mock project 056-D03_TS-26-CLL (real shapes): tasklists, measurement tasks pre-selected, quantities and engineer', async () => {
    const tasks = await zoho.getProjectTasks('27127000009850033');
    assert.strictEqual(tasks.length, 6);
    const plan = nameOwners(buildTaskPlan(tasks, { d03: true }), await zoho.getProjectUsers('27127000009850033'));
    assert.deepStrictEqual(plan.taskLists.map((list) => list.name), ['Kick-off Meeting & Input Data', 'Onsite Survey / Measurement Activities', 'Data Analysis', 'Draft Report Preparation']);
    const selected = plan.tasks.filter((entry) => entry.selected);
    assert.deepStrictEqual(selected.map((entry) => entry.testId), ['soilResistivity', 'electrodeResistance']);
    assert.deepStrictEqual([plan.tests.soilResistivity.planned, plan.tests.electrodeResistance.planned], [4, 170]);
    assert.ok(selected.every((entry) => entry.ownerName === 'Anup Parekh'), 'the engineer is the task owner, by full name');
    assert.ok(plan.tasks.find((entry) => entry.name.startsWith('[IRT]')).irt);
    assert.ok(!plan.tasks.find((entry) => entry.name.startsWith('[IRT]')).selected);
    assert.ok(!plan.tasks.find((entry) => /Data Tabulation/.test(entry.name)).measurement, 'an Offsite task with an ESHA Test is not a measurement task');
    assert.strictEqual(plan.tasks[0].name.includes('&amp;'), false, 'names are decoded');
    assert.deepStrictEqual(plan.warnings, []);
  });

  await check('task normalisation keeps owners (with ids), the tasklist and the custom fields as the payload gives them', async () => {
    const raw = {
      id_string: '1', name: '[ESHA] X', status: { name: 'Planned' }, start_date: '10-19-2026', end_date: '10-23-2026',
      tasklist: { id_string: 'tl9', name: 'Onsite' }, details: { owners: [{ id: '60007088431', name: 'anup', zpuid: '27', email: 'Anup@Elegrow.com' }] },
      custom_fields: [{ column_name: 'UDF_CHAR81', label_name: 'Work Location', value: 'Onsite' }]
    };
    const normalized = zoho.normalizeTask(raw);
    assert.deepStrictEqual(normalized.owners, [{ id: '60007088431', zpuid: '27', name: 'anup', email: 'anup@elegrow.com' }]);
    assert.deepStrictEqual(normalized.customFields, [{ label: 'Work Location', apiName: 'UDF_CHAR81', value: 'Onsite' }]);
    assert.deepStrictEqual([normalized.taskListId, normalized.taskListName, normalized.ownerName], ['tl9', 'Onsite', 'anup']);
    assert.deepStrictEqual(zoho.normalizeTask({ id: '2', name: 'No fields' }).customFields, [], 'empty custom fields are omitted by Zoho');
    assert.strictEqual(zoho.normalizeTask({ id: '3', name: 'A &amp; B', tasklist: { id_string: 't', name: 'Kick-off Meeting &amp; Input Data' } }).taskListName, 'Kick-off Meeting & Input Data', 'HTML entities in tasklist names are decoded');
  });

  console.log(`\n${passed} zoho-task checks passed.`);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
