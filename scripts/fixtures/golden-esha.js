// Sample data from reference/generators/make_report.py (standard v0.4). Shared by the golden
// checks and the report-PDF tests so both assert the same numbers.
// soil[location][direction] = [spacing a (m), R (ohm)] — the sheet records R, rho = 2*pi*a*R.
const SOIL = {
  'Utility Yard': {
    d1: [[0.5, 9.8], [1.0, 4.6], [1.5, 2.9], [2.0, 2.0], [2.5, 1.5], [3.0, 1.2]],
    d2: [[0.5, 10.4], [1.0, 4.9], [1.5, 3.1], [2.0, 2.2], [2.5, 1.6], [3.0, 1.3]]
  },
  'Reactor Bay': {
    d1: [[0.5, 8.1], [1.0, 3.9], [1.5, 2.6], [2.0, 1.8], [2.5, 1.4], [3.0, 1.1]],
    d2: [[0.5, 8.6], [1.0, 4.1], [1.5, 2.7], [2.0, 1.9], [2.5, 1.5], [3.0, 1.2]]
  }
};

// tag, location, type, l (m), d (mm), R without grid (null = not measured), codes, remark
const PITS = [
  ['EP-01', 'Transformer TR-1', 'Rod', 3.0, 40, 0.82, [], ''],
  ['EP-02', 'Transformer TR-1', 'Rod', 3.0, 40, 1.10, [], ''],
  ['EP-03', 'MCC Room', 'Rod', 3.0, 40, 1.45, ['C1'], 'Rod heavily corroded at clamp'],
  ['EP-04', 'MCC Room', 'Rod', 3.0, 40, 0.95, [], ''],
  ['EP-05', 'Reactor Bay R-201', 'Rod', 3.0, 40, 1.30, ['C8'], 'No pit tag'],
  ['EP-06', 'Reactor Bay R-202', 'Rod', 3.0, 40, 2.05, [], ''],
  ['EP-07', 'Cooling Tower', 'Rod', 3.0, 40, 7.12, ['C2'], 'Strip disconnected from rod'],
  ['EP-08', 'Pump House', 'Rod', 3.0, 40, null, ['C4'], 'Chamber buried under paving'],
  ['EP-09', 'Pipe Rack PR-3', 'Rod', 3.0, 40, 6.40, [], ''],
  ['EP-10', 'DG Area', 'Strip', 6.0, 50, 2.40, [], '25x5 strip, width 50 mm entered']
];

// main location, point, towards equipment, towards grid, codes
const RISERS = [
  ['Reactor Bay', 'R-201 (R1)', 0.21, 0.12, []],
  ['Reactor Bay', 'R-201 (R2)', 0.34, 0.20, []],
  ['Reactor Bay', 'R-202 (R1)', 1.42, 0.31, ['C6']],
  ['Pump House', 'P-101 A', 0.18, 0.09, []],
  ['Pump House', 'P-101 B', 0.27, 0.19, []],
  ['Cooling Tower', 'CT Fan-1', 0.44, 0.38, []]
];

// panel, feeder, device, A, kA (null = not recorded), Zs R/Y/B (ohm), PFC R/Y/B (kA)
const PANELS = [
  ['PMCC-1', 'Incomer', 'ACB', 1600, 50, [0.62, 0.65, 0.61], [6.9, 6.7, 7.1]],
  ['MCC-2', '7F2 Feeder', 'MCCB', 250, 36, [1.21, 1.35, 1.92], [2.1, 2.0, 1.4]],
  ['MLDB-1', 'Incomer', 'MCCB', 125, 25, [1.05, 1.10, 1.02], [2.6, 2.5, 2.7]],
  ['APDB-2', 'Incomer', 'MCB', 63, null, [1.40, 1.38, 1.44], [1.9, 1.9, 1.8]]
];

const PHASES = ['R-E', 'Y-E', 'B-E'];
const panelRows = (extra) =>
  PANELS.flatMap((panel) =>
    PHASES.map((phase, i) => ({
      location: panel[0],
      feederTag: panel[1],
      deviceType: panel[2],
      deviceRating: String(panel[3]),
      breakingCapacity: panel[4] === null ? '' : String(panel[4]),
      measuredPoints: phase,
      ...extra(panel, i)
    }))
  );

function buildPayload() {
  const soilRows = (rows) => rows.map(([spacing, resistance]) => ({ spacing: String(spacing), resistivity: String(resistance) }));
  return {
    project: {
      projectNo: '024-D03_TS-26-GFL',
      clientName: 'Gujarat Fluorochemicals Ltd',
      siteLocation: 'Dahej, Bharuch',
      workOrder: '1002-1261201769',
      reportDate: '2026-09-14',
      engineerName: 'H. S. Patil'
    },
    tests: {
      soilResistivity: true,
      electrodeResistance: true,
      loopImpedanceTest: true,
      prospectiveFaultCurrent: true,
      riserIntegrityTest: true
    },
    soilResistivity: {
      locations: Object.entries(SOIL).map(([name, dirs]) => ({
        name,
        inputUnit: 'ohm',
        direction1: soilRows(dirs.d1),
        direction2: soilRows(dirs.d2)
      }))
    },
    electrodeResistance: PITS.map(([tag, location, type, l, d, r, codes, remark]) => ({
      tag,
      location,
      electrodeType: type,
      length: String(l),
      diameter: String(d),
      resistanceWithoutGrid: r === null ? '' : String(r),
      conditionCodes: codes,
      observation: remark
    })),
    loopImpedanceTest: panelRows((panel, i) => ({ loopImpedance: String(panel[5][i]) })),
    prospectiveFaultCurrent: panelRows((panel, i) => ({ prospectiveFaultCurrent: String(panel[6][i]) })),
    riserIntegrityTest: RISERS.map(([mainLocation, measurementPoint, equipment, grid, codes]) => ({
      mainLocation,
      measurementPoint,
      resistanceTowardsEquipment: String(equipment),
      resistanceTowardsGrid: String(grid),
      conditionCodes: codes
    }))
  };
}

module.exports = { buildPayload, SOIL, PITS, RISERS, PANELS };
