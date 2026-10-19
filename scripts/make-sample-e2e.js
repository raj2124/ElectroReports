// End-to-end sample: a project workspace with two site days created through the HTTP routes, the
// final report assembled from them, the review steps done, and the final PDF exported through the
// app. Writes reference/screenshots/sample-final-e2e.pdf (git-ignored) so the whole path can be read.
// Usage: node scripts/make-sample-e2e.js
const fs = require('fs');
const os = require('os');
const path = require('path');
const net = require('net');
const { spawn } = require('child_process');
const { buildAllTestsPayload } = require('./fixtures/golden-all-tests');

const ROOT = path.join(__dirname, '..');
const freePort = () => new Promise((resolve) => {
  const server = net.createServer().listen(0, '127.0.0.1', () => {
    const { port } = server.address();
    server.close(() => resolve(port));
  });
});

// Day 1: soil, the first six pits, continuity, loop and fault current. Day 2: the other four pits,
// riser, ground resistance, step & touch and tower footing.
function splitDays(all) {
  const pick = (ids, extra) => {
    const draft = { project: { ...all.project }, tests: Object.fromEntries(ids.map((id) => [id, true])), sheetHeaders: Object.fromEntries(Object.entries(all.sheetHeaders).filter(([id]) => ids.includes(id))), ...extra };
    ids.forEach((id) => {
      if (!(id in draft)) draft[id] = all[id];
    });
    return draft;
  };
  return [
    pick(['soilResistivity', 'electrodeResistance', 'continuityTest', 'loopImpedanceTest', 'prospectiveFaultCurrent'], { electrodeResistance: all.electrodeResistance.slice(0, 6) }),
    pick(['electrodeResistance', 'riserIntegrityTest', 'groundResistance', 'stepTouchVoltage', 'towerFootingResistance'], { electrodeResistance: all.electrodeResistance.slice(6) })
  ];
}

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'sample-e2e-'));
  const port = await freePort();
  const server = spawn(process.execPath, [path.join(ROOT, 'src', 'server.js')], {
    env: { ...process.env, ELECTROREPORTS_PORT: String(port), ELECTROREPORTS_DATA_FILE: path.join(dir, 'reports.json'), ELECTROREPORTS_GENERATED_DIR: path.join(dir, 'pdfs'), ANTHROPIC_API_KEY: '', OPENAI_API_KEY: '', GEMINI_API_KEY: '' },
    stdio: ['ignore', 'pipe', 'inherit']
  });
  await new Promise((resolve, reject) => {
    server.stdout.on('data', (chunk) => String(chunk).includes('running') && resolve());
    server.on('exit', (code) => reject(new Error(`server exited early (${code})`)));
    setTimeout(() => reject(new Error('server did not start')), 15000);
  });
  const call = async (method, url, body) => {
    const response = await fetch(`http://127.0.0.1:${port}${url}`, { method, headers: { 'Content-Type': 'application/json' }, body: body === undefined ? undefined : JSON.stringify(body) });
    const data = await response.json().catch(() => null);
    if (!response.ok) throw new Error(`${method} ${url}: ${response.status} ${JSON.stringify(data)}`);
    return data;
  };
  try {
    const all = buildAllTestsPayload();
    const { projectId } = await call('POST', '/api/projects', { project: { ...all.project, zohoProjectId: 'Z-1001', zohoProjectName: 'GFL ESHA' } });
    await call('PUT', `/api/projects/${projectId}`, { project: { plannedQuantities: all.project.plannedQuantities } });
    const [dayOne, dayTwo] = splitDays(all);
    const first = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-14', report: { ...dayOne, project: { ...dayOne.project, nextDayPlan: 'Remaining pits, riser, grid, step & touch and tower footing.' } } });
    const second = await call('POST', `/api/projects/${projectId}/days`, { workDate: '2026-09-15', report: { ...dayTwo, project: { ...dayTwo.project, nextDayPlan: 'Report assembly.' } } });
    const step = (id, name, body) => call('POST', `/api/reports/${id}/review/${name}`, body);
    const tickAll = async (id) => {
      const ticks = Object.fromEntries((await call('GET', `/api/reports/${id}/review`)).checklist.map((item) => [item.id, true]));
      await call('PUT', `/api/reports/${id}/review/ticks`, { ticks });
    };
    const people = { prepared: { actor: 'H. S. Patil', designation: 'Site Engineer' }, checked: { actor: 'A. S. Mahto', designation: 'Senior Engineer' }, approved: { actor: 'J. C. Shah', designation: 'Head - Earthing Services' } };
    for (const id of [first.id, second.id]) {
      await step(id, 'submit', people.prepared);
      await tickAll(id);
      await step(id, 'check', people.checked);
      await step(id, 'approve', people.approved);
    }
    const assembled = await call('POST', `/api/projects/${projectId}/final/assemble`, { actor: people.prepared.actor });
    const finalId = assembled.report.id;
    await step(finalId, 'submit', people.prepared);
    await tickAll(finalId);
    await step(finalId, 'check', people.checked);
    const state = await call('GET', `/api/reports/${finalId}/review`);
    const acknowledgements = Object.fromEntries(state.acknowledgements.map((item) => [item.key, 'Acknowledged for the sample: Annexure B photographs follow; the narrative was read against the tables.']));
    await step(finalId, 'approve', { ...people.approved, acknowledgements });
    const exported = await call('POST', `/api/reports/${finalId}/pdf`);
    const target = path.join(ROOT, 'reference', 'screenshots', 'sample-final-e2e.pdf');
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(path.join(dir, 'pdfs', path.basename(exported.pdfUrl)), target);
    console.log(`days ${assembled.includedDates.join(', ')}; unapproved: ${assembled.flaggedDates.join(', ') || 'none'}; ${assembled.supersededReadings} superseded reading(s)`);
    console.log(`written ${target}`);
  } finally {
    server.kill();
  }
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
