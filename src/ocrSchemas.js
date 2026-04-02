const SOIL_RESISTIVITY_OCR_SCHEMA = {
  id: 'soilResistivity',
  label: 'Soil Resistivity Test OCR',
  supportedMimeTypes: new Set(['application/pdf', 'image/png', 'image/jpeg', 'image/jpg', 'image/webp']),
  template: {
    sheetType: 'soilResistivity',
    locations: [
      {
        name: 'Location name from the uploaded sheet',
        drivenElectrodeDiameter: '',
        drivenElectrodeLength: '',
        notes: '',
        direction1: [{ spacing: '0.5', resistivity: '58.7' }],
        direction2: [{ spacing: '0.5', resistivity: '30.2' }]
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'locations[0].direction1[2].resistivity',
        reason: 'Handwriting unclear or value partly hidden'
      }
    ]
  }
};

const ELECTRODE_RESISTANCE_OCR_SCHEMA = {
  id: 'electrodeResistance',
  label: 'Earth Electrode Resistance Test OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'electrodeResistance',
    rows: [
      {
        tag: 'P-1',
        location: 'Switchyard Near GT-1',
        electrodeType: 'Rod',
        materialType: 'Copper',
        length: '3',
        diameter: '25',
        resistanceWithoutGrid: '3.8',
        resistanceWithGrid: '2.9'
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'rows[0].resistanceWithGrid',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

const CONTINUITY_TEST_OCR_SCHEMA = {
  id: 'continuityTest',
  label: 'Continuity Test OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'continuityTest',
    rows: [
      {
        srNo: '1',
        mainLocation: 'Panel A',
        measurementPoint: 'Busbar',
        resistance: '0.3',
        impedance: '0.2'
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'rows[0].measurementPoint',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

const LOOP_IMPEDANCE_OCR_SCHEMA = {
  id: 'loopImpedanceTest',
  label: 'Loop Impedance Test OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'loopImpedanceTest',
    groups: [
      {
        location: 'STG-1 MCC',
        feederTag: 'O/G to Room AC unit',
        deviceType: 'MCB',
        deviceRating: '125',
        breakingCapacity: '25',
        rows: [
          { measuredPoints: 'R-E', loopImpedance: '0.8', voltage: '230' },
          { measuredPoints: 'Y-E', loopImpedance: '1.3', voltage: '239' },
          { measuredPoints: 'B-E', loopImpedance: '1.8', voltage: '237' }
        ]
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'groups[0].rows[1].loopImpedance',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

const PROSPECTIVE_FAULT_CURRENT_OCR_SCHEMA = {
  id: 'prospectiveFaultCurrent',
  label: 'Prospective Fault Current OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'prospectiveFaultCurrent',
    groups: [
      {
        location: 'STG-1 MCC',
        feederTag: 'O/G to Room AC unit',
        deviceType: 'MCB',
        deviceRating: '125',
        breakingCapacity: '25',
        rows: [
          { measuredPoints: 'R-E', loopImpedance: '0.8', prospectiveFaultCurrent: '4.2', voltage: '230' },
          { measuredPoints: 'Y-E', loopImpedance: '1.3', prospectiveFaultCurrent: '3.9', voltage: '239' },
          { measuredPoints: 'B-E', loopImpedance: '1.8', prospectiveFaultCurrent: '3.4', voltage: '237' }
        ]
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'groups[0].rows[0].prospectiveFaultCurrent',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

const RISER_INTEGRITY_OCR_SCHEMA = {
  id: 'riserIntegrityTest',
  label: 'Riser / Grid Integrity Test OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'riserIntegrityTest',
    rows: [
      {
        srNo: '1',
        mainLocation: 'MCC Room',
        measurementPoint: 'Riser Bond',
        resistanceTowardsEquipment: '0.4',
        resistanceTowardsGrid: '0.3'
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'rows[0].measurementPoint',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

const EARTH_CONTINUITY_OCR_SCHEMA = {
  id: 'earthContinuityTest',
  label: 'Earth Continuity Test OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'earthContinuityTest',
    rows: [
      {
        srNo: '1',
        tag: 'EQ-01',
        locationBuildingName: 'Switch Yard',
        distance: '18',
        measuredValue: '0.4'
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'rows[0].measuredValue',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

const TOWER_FOOTING_OCR_SCHEMA = {
  id: 'towerFootingResistance',
  label: 'Tower Footing Resistance OCR',
  supportedMimeTypes: SOIL_RESISTIVITY_OCR_SCHEMA.supportedMimeTypes,
  template: {
    sheetType: 'towerFootingResistance',
    groups: [
      {
        srNo: '1',
        mainLocationTower: 'Between ICL-2 & ICL-1 Line Near C.T.',
        readings: [
          { measurementPointLocation: 'Foot-1', footToEarthingConnectionStatus: 'Not Given', measuredCurrentMa: '10.8', measuredImpedance: '3.0' },
          { measurementPointLocation: 'Foot-2', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '10.7', measuredImpedance: '0.0' },
          { measurementPointLocation: 'Foot-3', footToEarthingConnectionStatus: 'Not Given', measuredCurrentMa: '10.8', measuredImpedance: '5.1' },
          { measurementPointLocation: 'Foot-4', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '10.9', measuredImpedance: '0.02' }
        ]
      }
    ],
    warnings: ['Any sheet-level warnings or assumptions'],
    uncertainFields: [
      {
        path: 'groups[0].readings[2].measuredImpedance',
        reason: 'Handwriting unclear or partly hidden'
      }
    ]
  }
};

module.exports = {
  SOIL_RESISTIVITY_OCR_SCHEMA,
  ELECTRODE_RESISTANCE_OCR_SCHEMA,
  CONTINUITY_TEST_OCR_SCHEMA,
  LOOP_IMPEDANCE_OCR_SCHEMA,
  PROSPECTIVE_FAULT_CURRENT_OCR_SCHEMA,
  RISER_INTEGRITY_OCR_SCHEMA,
  EARTH_CONTINUITY_OCR_SCHEMA,
  TOWER_FOOTING_OCR_SCHEMA
};
