// Second golden fixture: all nine measurement sheets (MS-01 to MS-09), with instruments, photo
// references on the P1 pits, soil test conditions and the 50/70% checks. It extends the first
// golden payload, so the soil, pit, loop, PFC and riser numbers are the same.
const { buildPayload } = require('./golden-esha');

function buildAllTestsPayload() {
  const base = buildPayload();
  const calibrationDue = '2027-03-31';
  return {
    ...base,
    project: {
      ...base.project,
      unit: 'FI Plant',
      unitCode: 'FI',
      siteLocation: '12/A GIDC Industrial Estate, Dahej, Bharuch, Gujarat 392130',
      clientRepresentative: 'R. K. Shah',
      groundCondition: 'Dry',
      plannedQuantities: { soilResistivity: 2, electrodeResistance: 10, continuityTest: 4, loopImpedanceTest: 4, prospectiveFaultCurrent: 4, riserIntegrityTest: 6, groundResistance: 2, stepTouchVoltage: 4, towerFootingResistance: 1 }
    },
    tests: {
      soilResistivity: true,
      electrodeResistance: true,
      continuityTest: true,
      loopImpedanceTest: true,
      prospectiveFaultCurrent: true,
      riserIntegrityTest: true,
      groundResistance: true,
      stepTouchVoltage: true,
      towerFootingResistance: true
    },
    sheetHeaders: {
      ...base.sheetHeaders,
      continuityTest: { instrument: 'mi3152', serialNo: 'MI3152-2208', calibrationDue, testLeadResistance: '0.10', leadNullDone: true },
      groundResistance: { instrument: 'mi3290', serialNo: 'MI3290-0412', calibrationDue, maxGridDimensionL: '40', table17DistanceC: '200' },
      stepTouchVoltage: { instrument: 'mi3295', serialNo: 'MI3295-0093', calibrationDue, faultCurrentKA: '20', faultCurrentSource: 'Client protection study', faultClearingTimeS: '0.5', faultClearingSource: 'Client protection study', bodyWeightBasis: '70', surfaceResistivity: '3000', surfaceThicknessM: '0.1', concreteResistivity: '5000', concreteThicknessM: '0.15' },
      towerFootingResistance: { instrument: 'mi3290', serialNo: 'MI3290-0412', calibrationDue }
    },
    soilResistivity: {
      locations: base.soilResistivity.locations.map((location, index) => ({ ...location, weather: index === 0 ? 'Sunny' : 'Cloudy', probeDepthCm: '12' }))
    },
    electrodeResistance: base.electrodeResistance.map((row) => ({
      ...row,
      // EP-03 and EP-07 are P1: EP-03 has its photo reference, EP-07 does not (the ack path); EP-08 (C4) has none either.
      photoRef: row.tag === 'EP-03' ? 'IMG_0312' : '',
      checkR50: row.resistanceWithoutGrid ? String((Number(row.resistanceWithoutGrid) * 1.02).toFixed(2)) : '',
      checkR70: row.resistanceWithoutGrid ? String((Number(row.resistanceWithoutGrid) * (row.tag === 'EP-06' ? 1.3 : 0.98)).toFixed(2)) : '',
      method: row.resistanceWithoutGrid ? '61.8%' : ''
    })),
    continuityTest: [
      { tag: 'CT-01', mainLocation: 'MCC Room', referencePoint: 'Main earth bar', measurementPoint: 'MCC-1 body', resistance: '0.32' },
      { tag: 'CT-02', mainLocation: 'MCC Room', referencePoint: 'Main earth bar', measurementPoint: 'MCC-2 body', resistance: '0.48' },
      { tag: 'CT-03', mainLocation: 'Reactor Bay', referencePoint: 'Bay earth bar', measurementPoint: 'R-201 vessel', resistance: '1.45', conditionCodes: ['C6'] },
      { tag: 'CT-04', mainLocation: 'Pump House', referencePoint: 'Bay earth bar', measurementPoint: 'P-101 motor', resistance: '0.61' }
    ],
    // Tags of the pits whose limit applies to the in-service grid reading
    groundResistance: [
      { tag: 'EP-02', location: 'Plant grid, north', method: '61.8%', currentProbeDistanceM: '200', r50: '0.52', r618: '0.50', r70: '0.49' },
      { tag: 'EP-04', location: 'Plant grid, south', method: '61.8%', currentProbeDistanceM: '200', r50: '0.56', r618: '0.55', r70: '0.54' }
    ],
    prospectiveFaultCurrent: base.prospectiveFaultCurrent.map((row) => ({ ...row, shortCircuitCurrentLN: String((Number(row.prospectiveFaultCurrent) * 0.8).toFixed(2)) })),
    stepTouchVoltage: [
      { location: 'Transformer TR-1 fence', measurementType: 'Touch-fence', surface: 'Gravel', distanceM: '1', testCurrentA: '5', v1: '0.06', v2: '0.08', v3: '0.05', vmax: '0.08' },
      { location: 'Transformer TR-1 yard', measurementType: 'Step', surface: 'Gravel', distanceM: '1', testCurrentA: '5', v1: '0.12', v2: '0.15', v3: '0.11', vmax: '0.15' },
      { location: 'Cooling tower structure', measurementType: 'Touch-structure', surface: 'Concrete', distanceM: '1', testCurrentA: '5', v1: '0.10', v2: '0.12', v3: '0.09', vmax: '0.12' },
      { location: 'Pump House walkway', measurementType: 'Step', surface: 'Soil', distanceM: '1', testCurrentA: '5', v1: '0.03', v2: '0.04', v3: '0.03', vmax: '0.04' }
    ],
    towerFootingResistance: [
      { mainLocationTower: 'Tower T1', readings: [1, 2, 3, 4].map((n) => ({ footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '10', measuredImpedance: String(2 + n * 0.2), conditionCodes: [], photoRef: '' })) }
    ]
  };
}

module.exports = { buildAllTestsPayload };
