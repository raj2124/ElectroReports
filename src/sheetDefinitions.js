// Measurement sheet definitions: the data-entry screens, as data.
//
// One entry per approved sheet MS-01..MS-09 (ET-D03-FRM-001..009), with the header
// fields and the row columns in the order they are printed
// (reference/generators/make_sheets_v3.py). The server validates against these
// definitions and serves them (GET /api/sheets); the browser renders every entry
// screen from them, so the app and the printed sheets cannot drift apart. A test
// compares this file with make_sheets_v3.py.
//
// Field kinds:  text · number · date · tick (one checkbox) · ticks (choose one of
//               `options`; `multi` = choose several, used for condition codes) ·
//               serial (row number) · fixed (printed, not entered) · group (several
//               small fields sharing one printed header cell).
// A field with `extra: true` is not on the printed sheet (the app needs it, e.g. plate
// size); every other field is a printed column or header cell, in print order.
// `legacy: true` fields are kept so older reports still load but are not shown.

const { resolveProfile, loadRegister } = require('./acceptance');

const CONDITION_CODE_LEGEND = Object.freeze([
  { code: 'C1', text: 'Corroded' },
  { code: 'C2', text: 'Strip / test link disconnected' },
  { code: 'C3', text: 'Single earth connection (optional)' },
  { code: 'C4', text: 'Buried / chamber not accessible or damaged' },
  { code: 'C5', text: 'Electrode damaged' },
  { code: 'C6', text: 'Loose connection' },
  { code: 'C7', text: 'Joint not cleaned / tinned / protected' },
  { code: 'C8', text: 'No tag / label' },
  { code: 'OTH', text: 'Other – write remark' }
]);

const INSTRUMENTS = Object.freeze([
  { id: 'mi3290', label: 'Metrel MI 3290 GX Earth Analyser', registerName: 'MI 3290 GX' },
  { id: 'mi3152', label: 'Metrel MI 3152 EurotestXC', registerName: 'MI 3152' },
  { id: 'kyoritsu4105a', label: 'Kyoritsu 4105A Digital Earth Tester', registerName: 'Kyoritsu 4105A' },
  { id: 'kyoritsu4118a', label: 'Kyoritsu 4118A Loop / PSC Tester', registerName: 'Kyoritsu 4118A' },
  { id: 'mi3295', label: 'Metrel MI 3295 Step & Contact Voltage Meter (optional – confirm availability)', registerName: 'MI 3295' }
]);

// Presets for the soil spacing set (a per-project setting).
const SOIL_SPACING_PRESETS = Object.freeze({
  'etpl-default': {
    label: 'ETPL default',
    spacings: ['0.5', '1.0', '1.5', '2.0', '2.5', '3.0', '3.5', '4.0', '4.5', '5.0'],
    directions: 4
  },
  'is3043-substation': {
    label: 'IS 3043 substation',
    spacings: ['2', '5', '10', '15', '25', '50'],
    directions: 8
  }
});
const DEFAULT_SOIL_PRESET = 'etpl-default';

const PHASES = ['R-E', 'Y-E', 'B-E'];
const DEVICE_TYPES = ['ACB', 'MCCB', 'MCB', 'SFU', 'Fuse'];
const GROUND_CONDITIONS = ['Dry', 'Damp', 'Wet'];
const ALL_CODES = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8'];

// ---- field builders --------------------------------------------------------------
const serial = (label = 'Sr.') => ({ key: 'srNo', label, kind: 'serial' });
const text = (key, label, o = {}) => ({ key, label, kind: 'text', ...o });
const number = (key, label, o = {}) => ({ key, label, kind: 'number', inputMode: 'decimal', min: 0, ...o });
const ticks = (key, label, options, o = {}) => ({ key, label, kind: 'ticks', options, ...o });
const codes = (allowed, label = 'Condition codes') => ({ key: 'conditionCodes', label, kind: 'ticks', multi: true, options: allowed, codeLegend: true });
const fixed = (key, label, values, display) => ({ key, label, kind: 'fixed', values, display: display || values });
const photoRef = (label = 'Photo ref.') => text('photoRef', label);

// ---- header fields common to every sheet (the printed instrument row and footer) ----
const SHEET_COMMON_HEADER = Object.freeze([
  { key: 'instrument', label: 'Instrument used', kind: 'ticks', optionsFromSheet: 'instruments' },
  text('serialNo', 'Instrument Sr. No.'),
  { key: 'calibrationDue', label: 'Calibration due', kind: 'date' },
  text('timeStart', 'Time start', { inputMode: 'text' }),
  text('timeFinish', 'Time finish', { inputMode: 'text' })
]);

// Job-level header cells (printed on every sheet; entered once per job/day).
const JOB_HEADER = Object.freeze([
  { key: 'reportDate', label: 'Work date', kind: 'date', required: true },
  { key: 'engineerName', label: 'ETPL engineer(s)', kind: 'text', required: true },
  { key: 'clientRepresentative', label: 'Client representative', kind: 'text' },
  { key: 'groundCondition', label: 'Ground condition', kind: 'ticks', options: GROUND_CONDITIONS }
]);

const SHEETS = [
  {
    id: 'soilResistivity',
    code: 'MS-01',
    docNo: 'ET-D03-FRM-001',
    title: 'Soil Resistivity',
    reference: 'Wenner 4-probe · IS 3043:2018 Cl. 40 · IEEE 81-2012 Cl. 7',
    registerKey: 'soil_resistivity',
    instruments: ['mi3290', 'kyoritsu4105a'],
    layout: 'soilMatrix',
    headerScope: 'location',
    headerFields: [
      text('name', 'Location name'),
      ticks('inputUnit', 'Reading unit', [
        { value: 'ohm', label: 'R (Ω)' },
        { value: 'ohm-m', label: 'ρ (Ω·m)' }
      ], { required: true }),
      number('probeDepthCm', 'Probe depth (cm)', { unit: 'cm', min: 0, warnOutside: [10, 15], warnCode: 'PROBE_DEPTH', warnText: 'IS 3043:2018 Cl. 40.4.1 gives a probe depth of 10–15 cm' }),
      {
        key: 'bearingsA', label: 'Direction 1 / 2 bearing', kind: 'group',
        fields: [number('direction1', 'D1', { unit: '°', max: 360, store: 'bearings' }), number('direction2', 'D2', { unit: '°', max: 360, store: 'bearings' })]
      },
      {
        key: 'bearingsB', label: 'Direction 3 / 4 bearing', kind: 'group',
        fields: [number('direction3', 'D3', { unit: '°', max: 360, store: 'bearings' }), number('direction4', 'D4', { unit: '°', max: 360, store: 'bearings' })]
      },
      ticks('weather', 'Weather', ['Sunny', 'Cloudy', 'After rain'])
    ],
    rowFields: [
      serial(),
      { key: 'spacing', label: 'Probe spacing a (m)', kind: 'fixed', unit: 'm', fromPreset: true },
      number('direction1', 'Direction 1 reading', { matrix: true, min: 0, exclusiveMin: true }),
      number('direction2', 'Direction 2 reading', { matrix: true, min: 0, exclusiveMin: true }),
      number('direction3', 'Direction 3 reading', { matrix: true, min: 0, exclusiveMin: true }),
      number('direction4', 'Direction 4 reading', { matrix: true, min: 0, exclusiveMin: true }),
      text('remark', 'Remark')
    ],
    notes: [
      'Write the value shown on the instrument and tick its unit in the header: R (Ω) or ρ (Ω·m). The app calculates ρ = 2πaR, mean per spacing, average and IS 3043 Table 4 corrosivity.',
      'Four probes in a straight line at equal spacing a, driven 10–15 cm.',
      'Substations / generating stations: at least 8 directions from the centre (Cl. 40.2.1) – use the IS 3043 substation spacing set.'
    ]
  },
  {
    id: 'electrodeResistance',
    code: 'MS-02',
    docNo: 'ET-D03-FRM-002',
    title: 'Earth Electrode (Pit) Resistance',
    reference: 'Fall of potential 61.8% · IS 3043:2018 Cl. 14.2, 41 · IEEE 81-2012 Cl. 8',
    registerKey: 'earth_electrode',
    instruments: ['mi3290', 'kyoritsu4105a'],
    layout: 'rows',
    headerFields: [
      number('currentElectrodeDistanceM', 'Current electrode distance C (m)', { unit: 'm' }),
      { key: 'auxRodsOk', label: 'Auxiliary rods', kind: 'tick', tickLabel: '12.5 mm MS, ≤ 1 m deep' },
      ticks('check90180', '90°/180° check done', ['Yes', 'No'])
    ],
    rowFields: [
      serial(),
      text('tag', 'Pit TAG No.'),
      text('location', 'Location / name of pit'),
      ticks('electrodeType', 'Electrode type', ['Plate', 'Pipe', 'Rod', 'Strip'], { default: 'Rod' }),
      ticks('materialType', 'Material', ['Cu', 'Cu-bonded', 'GI', 'MS', 'CI', 'SS'], { default: 'Cu' }),
      number('length', 'Length l (m)', { unit: 'm', exclusiveMin: true }),
      number('diameter', 'Dia d / strip width (mm)', { unit: 'mm', exclusiveMin: true }),
      number('plateLengthM', 'Plate length (m)', { unit: 'm', exclusiveMin: true, extra: true, showIf: { key: 'electrodeType', equals: 'Plate' } }),
      number('plateWidthM', 'Plate width (m)', { unit: 'm', exclusiveMin: true, extra: true, showIf: { key: 'electrodeType', equals: 'Plate' } }),
      number('resistanceWithoutGrid', 'R without grid at 61.8% (Ω)', { unit: 'Ω' }),
      number('checkR50', 'Check R at 50% (Ω)', { unit: 'Ω' }),
      number('checkR70', 'Check R at 70% (Ω)', { unit: 'Ω' }),
      number('resistanceWithGrid', 'R with grid (Ω)', { unit: 'Ω' }),
      ticks('method', 'Method', ['61.8%', 'Clamp', 'Slope']),
      codes(ALL_CODES),
      photoRef(),
      text('observation', 'Remark')
    ],
    notes: [
      "Open the test link for 'R without grid'; restore it after. Read at 61.8%, then move P to 50% and 70%: if the three readings agree the 61.8% value is valid (Cl. 41.1.6).",
      'Current electrode C at least 30 m from a rod, pipe or plate (Cl. 41.1.5); auxiliary rods 12.5 mm MS, up to 1 m deep (Cl. 41.1.4). Record C distance in the header.',
      'Plate: enter the plate size (m × m); the app uses the area of both sides. Pre-printed tags come from the asset register; new pits on a continuation sheet.'
    ]
  },
  {
    id: 'continuityTest',
    code: 'MS-03',
    docNo: 'ET-D03-FRM-003',
    title: 'Earth Continuity & Bonding',
    reference: 'Ground integrity · IEEE 81-2012 Cl. 10 · IS 3043:2018 Cl. 17.2.3',
    registerKey: 'earth_continuity',
    instruments: ['mi3152', 'mi3290'],
    layout: 'rows',
    headerFields: [
      number('testLeadResistance', 'Test-lead resistance (Ω)', { unit: 'Ω' }),
      { key: 'leadNullDone', label: 'Lead null done', kind: 'tick', tickLabel: 'Yes' },
      ticks('method', 'Method', [{ value: 'Low-ohm', label: 'Low-Ω' }, { value: 'High-current', label: 'High-current' }])
    ],
    rowFields: [
      serial(),
      text('tag', 'TAG No.'),
      text('mainLocation', 'Main location / area'),
      text('referencePoint', 'Reference point (earth bar / riser)'),
      text('measurementPoint', 'Tested point (equipment / riser)'),
      number('distanceM', 'Distance (m)', { unit: 'm' }),
      number('testCurrentA', 'Test current (A)', { unit: 'A' }),
      number('resistance', 'Measured R (Ω)', { unit: 'Ω' }),
      codes(['C1', 'C2', 'C6', 'C7']),
      photoRef(),
      text('comment', 'Remark'),
      { key: 'impedance', label: 'Impedance (Ω)', kind: 'number', legacy: true }
    ],
    notes: [
      'Short the leads and null the tester; write lead R in the header – it is subtracted from every reading (IEEE 81-2012 Cl. 10.2).',
      'Keep one reference point for the area so results can be compared point to point.',
      'High-current integrity test: write test current and use Remark for the voltage drop (V).'
    ]
  },
  {
    id: 'loopImpedanceTest',
    code: 'MS-04',
    docNo: 'ET-D03-FRM-004',
    title: 'Earth Fault Loop Impedance',
    reference: 'Zs, L-PE · IS 3043:2018 Cl. 42 · IS/IEC 60364-6',
    registerKey: 'loop_impedance',
    instruments: ['kyoritsu4118a', 'mi3152'],
    layout: 'phaseGroups',
    groupSize: 3,
    headerFields: [],
    rowFields: [
      { ...serial(), merge: true },
      text('location', 'Location of panel / equipment', { merge: true }),
      text('feederTag', 'Feeder name & tag no.', { merge: true }),
      ticks('deviceType', 'Device type', DEVICE_TYPES, { merge: true, default: 'MCB' }),
      number('deviceRating', 'Device rating In (A)', { unit: 'A', merge: true, exclusiveMin: true }),
      ticks('tripCurve', 'Trip curve', ['B', 'C', 'D', 'Adj.'], { merge: true }),
      fixed('measuredPoints', 'Measured point', PHASES),
      number('loopImpedance', 'Loop impedance Zs (Ω)', { unit: 'Ω' }),
      number('voltage', 'Voltage L-PE (V)', { unit: 'V', exclusiveMin: true, default: '230' }),
      { ...photoRef(), merge: true },
      text('remarks', 'Remark', { merge: true })
    ],
    notes: [
      'Live working: risk assessment, PPE and permit mandatory. Measure at the outgoing side of the device, one reading per phase to PE.',
      "Copy device type, rating In and trip curve from the nameplate – the app uses them for the device's disconnection limit.",
      "A panel is healthy only if all three phases meet the project's Zs criterion (Standards Acceptance List)."
    ]
  },
  {
    id: 'prospectiveFaultCurrent',
    code: 'MS-05',
    docNo: 'ET-D03-FRM-005',
    title: 'Prospective Fault & Short-circuit Current',
    reference: 'PFC (L-PE) & PSC (L-N) · IS/IEC 60364-4-43 Reg. 434.5.1',
    registerKey: 'pfc',
    instruments: ['kyoritsu4118a', 'mi3152'],
    layout: 'phaseGroups',
    groupSize: 3,
    headerFields: [],
    rowFields: [
      { ...serial(), merge: true },
      text('location', 'Location of panel / equipment', { merge: true }),
      text('feederTag', 'Feeder name & tag no.', { merge: true }),
      ticks('deviceType', 'Device type', DEVICE_TYPES, { merge: true, default: 'MCB' }),
      number('deviceRating', 'Rating In (A)', { unit: 'A', merge: true, exclusiveMin: true }),
      number('breakingCapacity', 'Breaking capacity Icu / Icn (kA)', { unit: 'kA', merge: true, exclusiveMin: true }),
      fixed('measuredPoints', 'Phase', PHASES, ['R', 'Y', 'B']),
      number('prospectiveFaultCurrent', 'PFC L-PE (kA)', { unit: 'kA' }),
      number('shortCircuitCurrentLN', 'PSC L-N (kA)', { unit: 'kA' }),
      number('voltage', 'Voltage L-N (V)', { unit: 'V', exclusiveMin: true, default: '230' }),
      { ...photoRef(), merge: true },
      text('comment', 'Remark', { merge: true }),
      { key: 'loopImpedance', label: 'Zs (Ω)', kind: 'number', legacy: true }
    ],
    notes: [
      "Breaking capacity (kA) from the nameplate is REQUIRED – without it the result is 'Not Assessable'.",
      'Record both PFC (L-PE) and PSC (L-N) per phase; the device must interrupt the higher of the two.',
      'Write values in kA as shown on the 4118A (convert A to kA if the display shows A).'
    ]
  },
  {
    id: 'riserIntegrityTest',
    code: 'MS-06',
    docNo: 'ET-D03-FRM-006',
    title: 'Riser Integrity',
    reference: 'Two-direction continuity · IEEE 81-2012 Cl. 10',
    registerKey: 'riser_integrity',
    instruments: ['mi3290', 'mi3152'],
    layout: 'rows',
    headerFields: [
      number('testLeadResistance', 'Test-lead resistance (Ω)', { unit: 'Ω' }),
      text('referenceRiser', 'Reference riser / point'),
      ticks('testCurrent', 'Test current', [{ value: '200 mA', label: '200 mA' }, { value: 'Other', label: 'Other' }]),
      number('testCurrentOther', 'Other test current (mA)', { unit: 'mA', extra: true, showIf: { key: 'testCurrent', equals: 'Other' } })
    ],
    rowFields: [
      serial(),
      text('mainLocation', 'Main location'),
      text('measurementPoint', 'Equipment / riser / measurement point'),
      number('resistanceTowardsEquipment', 'R towards equipment (Ω)', { unit: 'Ω' }),
      number('resistanceTowardsGrid', 'R towards grid (Ω)', { unit: 'Ω' }),
      codes(['C1', 'C2', 'C6', 'C7']),
      photoRef(),
      text('comment', 'Remark')
    ],
    notes: [
      'Null the leads first; write lead R in the header (IEEE 81-2012 Cl. 10.2).',
      'Reading 1 from the measurement point to the equipment body; reading 2 from the same point to the grid / reference riser.',
      'Use one reference riser for the whole area so risers can be compared with each other.'
    ]
  },
  {
    id: 'groundResistance',
    code: 'MS-07',
    docNo: 'ET-D03-FRM-007',
    title: 'Ground (Grid) Resistance',
    reference: '61.8% / slope / clamp · IS 3043:2018 Cl. 41, Table 17 · IEEE 81-2012 Cl. 8',
    registerKey: 'ground_resistance',
    instruments: ['mi3290', 'kyoritsu4105a'],
    layout: 'rows',
    headerFields: [
      number('maxGridDimensionL', 'Max. grid dimension L (m)', { unit: 'm', exclusiveMin: true }),
      number('table17DistanceC', 'Table 17 distance C used (m)', { unit: 'm', exclusiveMin: true }),
      ticks('check90180', '90°/180° check', ['Yes', 'No'])
    ],
    rowFields: [
      serial(),
      text('tag', 'Point / pit TAG'),
      text('location', 'Location of measured ground'),
      ticks('method', 'Method', ['61.8%', 'Slope', 'Clamp']),
      number('currentProbeDistanceM', 'Current probe distance C (m)', { unit: 'm', dashAllowed: true }),
      number('r50', 'R at 50% (Ω)', { unit: 'Ω', dashAllowed: true }),
      number('r618', 'R at 61.8% (Ω)', { unit: 'Ω' }),
      number('r70', 'R at 70% (Ω)', { unit: 'Ω', dashAllowed: true }),
      codes(['C1', 'C2', 'C6', 'C7']),
      text('remark', 'Remark')
    ],
    notes: [
      "Grid stays connected (in-service value). Choose C from Table 17 by the grid's maximum dimension L.",
      'The 61.8% value is valid when the 50% and 70% readings agree; if not, move C further away or use the slope method (Cl. 41.1.8).',
      "Clamp method: no auxiliary probes – write '—' in the distance and 50/70% columns."
    ]
  },
  {
    id: 'stepTouchVoltage',
    code: 'MS-08',
    docNo: 'ET-D03-FRM-008',
    title: 'Step & Touch Voltage',
    reference: 'Current injection · IEEE 81-2012 Cl. 9 · IS 3043:2018 Cl. 10.3 · IEEE 80-2013',
    registerKey: 'step_touch',
    instruments: ['mi3295', 'mi3290'],
    layout: 'rows',
    headerFields: [
      { key: 'faultCurrent', label: 'Earth fault current If (kA) & source', kind: 'group', fields: [number('faultCurrentKA', 'If', { unit: 'kA', exclusiveMin: true }), text('faultCurrentSource', 'Source')] },
      { key: 'clearingTime', label: 'Fault clearing time ts (s) & source', kind: 'group', fields: [number('faultClearingTimeS', 'ts', { unit: 's', exclusiveMin: true }), text('faultClearingSource', 'Source')] },
      number('remoteElectrodeDistanceM', 'Remote current electrode distance (m)', { unit: 'm' }),
      { key: 'surfaceLayer', label: 'Surface layer ρs (Ω·m) & thickness hs (m)', kind: 'group', fields: [number('surfaceResistivity', 'ρs', { unit: 'Ω·m', exclusiveMin: true }), number('surfaceThicknessM', 'hs', { unit: 'm', exclusiveMin: true })] },
      ticks('bodyWeightBasis', 'Body weight basis', [{ value: '50', label: '50 kg' }, { value: '70', label: '70 kg' }, { value: 'both', label: 'Both' }]),
      text('injectionPoint', 'Injection point'),
      { key: 'csOverrideGroup', label: 'Cs override (client design value) & source', kind: 'group', extra: true, fields: [number('csOverride', 'Cs', { exclusiveMin: true, max: 1 }), text('csSource', 'Source')] }
    ],
    rowFields: [
      serial(),
      text('location', 'Location / structure / measurement point'),
      ticks('measurementType', 'Measurement type', [
        { value: 'Step', label: 'Step' },
        { value: 'Touch-structure', label: 'Touch-structure' },
        { value: 'Touch-mesh', label: 'Touch-mesh' },
        { value: 'Touch-fence', label: 'Touch-fence' },
        { value: 'Touch-gate', label: 'Touch-gate' }
      ]),
      ticks('surface', 'Surface', ['Soil', 'Gravel', 'Concrete']),
      number('distanceM', 'Distance from structure (m)', { unit: 'm' }),
      number('testCurrentA', 'Test current It (A)', { unit: 'A', exclusiveMin: true }),
      number('v1', 'V1 (V)', { unit: 'V' }),
      number('v2', 'V2 (V)', { unit: 'V' }),
      number('v3', 'V3 (V)', { unit: 'V' }),
      number('vmax', 'Vmax (V)', { unit: 'V' }),
      photoRef(),
      text('remark', 'Remark')
    ],
    notes: [
      'Step: two probes 1 m apart. Touch: probe within 1 m of the structure (arm reach). Probes ~150 mm into moist soil (IEEE 81-2012 Cl. 9.2–9.3).',
      'Write the test current It for every row – the app scales each voltage by If / It before comparing with the tolerable limit.',
      'MI 3295 is not on the current instrument list – confirm availability.'
    ]
  },
  {
    id: 'towerFootingResistance',
    code: 'MS-09',
    docNo: 'ET-D03-FRM-009',
    title: 'Tower Footing Impedance',
    reference: 'Clamp method · IS 3043:2018 Cl. 41 · IEEE 81-2012 Cl. 8',
    registerKey: 'tower_footing',
    instruments: ['mi3290'],
    layout: 'towerGroups',
    groupSize: 4,
    headerFields: [
      text('lineName', 'Line / circuit name'),
      text('lineVoltageTowerType', 'Line voltage (kV) / tower type'),
      { key: 'zsat', label: 'Zsat (Ω) from acceptance list', kind: 'fixed', fromProfile: 'TOWER-ZSAT', readOnly: true }
    ],
    rowFields: [
      { ...serial(), merge: true },
      text('mainLocationTower', 'Tower no. / location', { merge: true }),
      fixed('measurementPointLocation', 'Foot', ['Foot-1', 'Foot-2', 'Foot-3', 'Foot-4'], ['F-1', 'F-2', 'F-3', 'F-4']),
      ticks('footToEarthingConnectionStatus', 'Foot-to-earth connection', ['Given', 'Not given'], { default: 'Given' }),
      number('measuredCurrentMa', 'Measured current I (mA)', { unit: 'mA' }),
      number('measuredImpedance', 'Measured impedance Z (Ω)', { unit: 'Ω' }),
      codes(['C1', 'C2', 'C5', 'C6']),
      photoRef(),
      text('remark', 'Remark')
    ],
    notes: [
      "Four feet per tower are one unit. Clamp each foot's earth connection in turn; the overhead earth wire stays connected.",
      "Mark 'Not given' where a foot has no earth connection – that foot is reported as a defect.",
      "The app calculates Zt from the four feet and compares it with Zsat from the project's acceptance profile."
    ]
  },
  {
    // Retired for new reports (folded into MS-03); kept so older reports still load.
    id: 'earthContinuityTest',
    code: 'legacy',
    docNo: '',
    title: 'Earth Continuity (retired – use MS-03)',
    reference: 'Replaced by MS-03 Earth Continuity & Bonding',
    registerKey: 'earth_continuity',
    instruments: ['mi3152', 'mi3290'],
    layout: 'rows',
    retired: true,
    headerFields: [],
    rowFields: [
      serial(),
      text('tag', 'Tag'),
      text('locationBuildingName', 'Location / Building'),
      number('distance', 'Distance'),
      number('measuredValue', 'Measured Value'),
      text('remark', 'Remark')
    ],
    notes: []
  }
];

// Printed column width fractions (the `w` of each col(...) in make_sheets_v3.py), one per
// printed row column in print order. The entry tables use the same proportions so a sheet
// looks like its printed form; a test compares these numbers with the generator.
const PRINTED_WIDTHS = Object.freeze({
  soilResistivity: [4, 10, 14, 14, 14, 14, 18],
  electrodeResistance: [2.3, 7, 9, 8, 12.5, 5, 6, 7, 6, 6, 6, 8.5, 11.5, 3.6, 6],
  continuityTest: [3.5, 8, 12, 13, 13, 7, 7.5, 8.5, 10.5, 5, 8],
  loopImpedanceTest: [2.6, 14, 13, 11, 7, 7, 5, 9, 7, 4.5, 11],
  prospectiveFaultCurrent: [2.6, 14, 13, 11, 6, 8, 4.2, 7.5, 7.5, 6, 4.5, 9],
  riserIntegrityTest: [4, 15, 20, 12.5, 12.5, 12, 6, 11],
  groundResistance: [3.6, 10.5, 16, 11, 9, 9.5, 9.5, 9.5, 9.5, 11.5],
  stepTouchVoltage: [2.6, 16, 15, 9, 6.5, 6.5, 6, 6, 6, 6, 4, 9],
  towerFootingResistance: [3.6, 14, 5.5, 11, 11, 11, 10.5, 5.5, 11]
});

SHEETS.forEach((sheet) => {
  const widths = PRINTED_WIDTHS[sheet.id];
  if (!widths) {
    return;
  }
  sheet.rowFields.filter((field) => !field.extra && !field.legacy).forEach((field, position) => {
    field.width = widths[position];
  });
});

const SHEET_BY_ID = Object.fromEntries(SHEETS.map((sheet) => [sheet.id, sheet]));

function getSheet(id) {
  return SHEET_BY_ID[id] || null;
}

function optionValue(option) {
  return option && typeof option === 'object' ? option.value : option;
}

function optionValues(field) {
  return (field.options || []).map(optionValue);
}

// Instruments allowed for a sheet, from the register (ET-D03-REG-002 `instruments`), so
// a change to the register changes what the sheets accept.
function instrumentsForSheet(sheetId) {
  const sheet = getSheet(sheetId);
  if (!sheet) {
    return [];
  }
  const registerList = loadRegister().instruments[sheet.registerKey] || [];
  return INSTRUMENTS.filter((instrument) => registerList.some((name) => name.startsWith(instrument.registerName)));
}

// Every field in a sheet's row layout, flattened (group fields expanded).
function flattenFields(fields) {
  return fields.flatMap((field) => (field.kind === 'group' ? field.fields : [field]));
}

// JSON-safe definitions for GET /api/sheets.
function publicDefinitions() {
  const profile = resolveProfile();
  return {
    sheets: SHEETS.map((sheet) => ({
      ...sheet,
      instrumentOptions: instrumentsForSheet(sheet.id).map(({ id, label }) => ({ value: id, label })),
      commonHeaderFields: SHEET_COMMON_HEADER.map((field) => (field.optionsFromSheet ? { ...field, options: instrumentsForSheet(sheet.id).map(({ id, label }) => ({ value: id, label })) } : field)),
      zsat: sheet.id === 'towerFootingResistance' ? profile.value('TOWER-ZSAT') : undefined
    })),
    jobHeader: JOB_HEADER,
    instruments: INSTRUMENTS,
    conditionCodes: CONDITION_CODE_LEGEND,
    soilSpacingPresets: SOIL_SPACING_PRESETS,
    defaultSoilSpacingPreset: DEFAULT_SOIL_PRESET,
    groundConditions: GROUND_CONDITIONS
  };
}

module.exports = {
  SHEETS,
  SHEET_COMMON_HEADER,
  JOB_HEADER,
  INSTRUMENTS,
  CONDITION_CODE_LEGEND,
  SOIL_SPACING_PRESETS,
  DEFAULT_SOIL_PRESET,
  GROUND_CONDITIONS,
  PHASES,
  getSheet,
  optionValue,
  optionValues,
  instrumentsForSheet,
  flattenFields,
  publicDefinitions
};
