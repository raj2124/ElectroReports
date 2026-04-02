const fs = require('fs');
const os = require('os');
const path = require('path');

const {
  normalizeReportInput,
  calculateSoilSummary,
  deriveElectrodeAssessment,
  deriveContinuityAssessment,
  deriveLoopImpedanceAssessment,
  deriveFaultCurrentAssessment,
  deriveRiserAssessment,
  deriveEarthContinuityAssessment,
  deriveTowerFootingAssessment,
  summarizeTowerGroups,
  buildTowerGroupKey,
  buildExecutiveSnapshot
} = require('../src/reportModel');
const { generateElectroReportPdf } = require('../src/pdfService');
const { createReportStore } = require('../src/reportStore');

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

function buildPayload() {
  return {
    project: {
      projectNo: 'ER-VERIFY-20260318',
      clientName: 'Verification Client',
      siteLocation: 'Surat Test Yard',
      workOrder: 'WO-VERIFY-001',
      reportDate: '2026-03-18',
      engineerName: 'Codex Verification'
    },
    tests: {
      soilResistivity: true,
      electrodeResistance: true,
      continuityTest: true,
      loopImpedanceTest: true,
      prospectiveFaultCurrent: true,
      riserIntegrityTest: true,
      earthContinuityTest: true,
      towerFootingResistance: true
    },
    soilResistivity: {
      locations: [
        {
          name: 'VAV Link 220kV Switch Yard',
          direction1: [
            { spacing: '0.5', resistivity: '60', rowObservation: 'Dry top layer observed.' },
            { spacing: '1.0', resistivity: '70' },
            { spacing: '1.5', resistivity: '80' },
            { spacing: '2.0', resistivity: '90' },
            { spacing: '2.5', resistivity: '100' },
            { spacing: '3.0', resistivity: '110' }
          ],
          direction2: [
            { spacing: '0.5', resistivity: '70' },
            { spacing: '1.0', resistivity: '80' },
            { spacing: '1.5', resistivity: '90' },
            { spacing: '2.0', resistivity: '100' },
            { spacing: '2.5', resistivity: '110' },
            { spacing: '3.0', resistivity: '120' }
          ],
          drivenElectrodeDiameter: '40',
          drivenElectrodeLength: '3',
          notes: 'Verification dataset for soil averaging.'
        }
      ]
    },
    electrodeResistance: [
      {
        tag: 'E-01',
        location: 'North Grid',
        electrodeType: 'Rod',
        materialType: 'Copper',
        length: '3',
        diameter: '25',
        resistanceWithoutGrid: '5.8',
        resistanceWithGrid: '4.2',
        rowObservation: 'Clamp cleaned before final reading.'
      },
      {
        tag: 'E-02',
        location: 'South Grid',
        electrodeType: 'Rod',
        materialType: 'Copper',
        length: '3',
        diameter: '25',
        resistanceWithoutGrid: '4.8',
        resistanceWithGrid: '5.2'
      }
    ],
    continuityTest: [
      { srNo: '1', mainLocation: 'Panel A', measurementPoint: 'Busbar', resistance: '0.3', impedance: '0.2' },
      { srNo: '2', mainLocation: 'Panel B', measurementPoint: 'Tray Joint', resistance: '0.8', impedance: '0.6' },
      { srNo: '3', mainLocation: 'Panel C', measurementPoint: 'Bond Lug', resistance: '1.4', impedance: '0.9' }
    ],
    loopImpedanceTest: [
      { srNo: '1', location: 'STG-1 MCC', feederTag: 'O/G to Room AC unit', deviceType: 'MCCB', deviceRating: '125', breakingCapacity: '25', measuredPoints: 'R-E', loopImpedance: '0.8', voltage: '230' },
      { srNo: '2', location: 'STG-1 MCC', feederTag: 'O/G to Room AC unit', deviceType: 'MCCB', deviceRating: '125', breakingCapacity: '25', measuredPoints: 'Y-E', loopImpedance: '1.3', voltage: '239' },
      { srNo: '3', location: 'STG-1 MCC', feederTag: 'O/G to Room AC unit', deviceType: 'MCCB', deviceRating: '125', breakingCapacity: '25', measuredPoints: 'B-E', loopImpedance: '1.8', voltage: '237' }
    ],
    prospectiveFaultCurrent: [
      {
        srNo: '1',
        location: 'Phase-1 MCC Room',
        feederTag: '2R4_O/G to PDB 23',
        deviceType: 'MCCB',
        deviceRating: '250',
        breakingCapacity: '20',
        measuredPoints: 'R-E',
        loopImpedance: '0.2',
        prospectiveFaultCurrent: '12',
        voltage: '238'
      },
      {
        srNo: '2',
        location: 'Phase-1 MCC Room',
        feederTag: '2R4_O/G to PDB 23',
        deviceType: 'MCCB',
        deviceRating: '250',
        breakingCapacity: '20',
        measuredPoints: 'Y-E',
        loopImpedance: '0.3',
        prospectiveFaultCurrent: '19',
        voltage: '238'
      },
      {
        srNo: '3',
        location: 'Phase-1 MCC Room',
        feederTag: '2R4_O/G to PDB 23',
        deviceType: 'MCCB',
        deviceRating: '250',
        breakingCapacity: '20',
        measuredPoints: 'B-E',
        loopImpedance: '0.4',
        prospectiveFaultCurrent: '24',
        voltage: '238'
      }
    ],
    riserIntegrityTest: [
      { srNo: '1', mainLocation: 'Riser A', measurementPoint: 'Joint 1', resistanceTowardsEquipment: '0.3', resistanceTowardsGrid: '0.4' },
      { srNo: '2', mainLocation: 'Riser B', measurementPoint: 'Joint 2', resistanceTowardsEquipment: '0.7', resistanceTowardsGrid: '0.4' },
      { srNo: '3', mainLocation: 'Riser C', measurementPoint: 'Joint 3', resistanceTowardsEquipment: '1.2', resistanceTowardsGrid: '0.6' }
    ],
    earthContinuityTest: [
      { srNo: '1', tag: 'EC-1', locationBuildingName: 'Warehouse', distance: '12', measuredValue: '0.3' },
      { srNo: '2', tag: 'EC-2', locationBuildingName: 'Substation', distance: '24', measuredValue: '0.7' },
      { srNo: '3', tag: 'EC-3', locationBuildingName: 'Control Room', distance: '30', measuredValue: '1.3' }
    ],
    towerFootingResistance: [
      {
        srNo: '1',
        mainLocationTower: 'Between ICL-2 & ICL-1 Line Near C.T.',
        readings: [
          { measurementPointLocation: 'Foot-1', footToEarthingConnectionStatus: 'Not Given', measuredCurrentMa: '10.8', measuredImpedance: '3.0' },
          { measurementPointLocation: 'Foot-2', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '10.7', measuredImpedance: '0.0' },
          { measurementPointLocation: 'Foot-3', footToEarthingConnectionStatus: 'Not Given', measuredCurrentMa: '10.8', measuredImpedance: '5.1', rowObservation: 'Reading stabilized after cable repositioning.' },
          { measurementPointLocation: 'Foot-4', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '10.9', measuredImpedance: '0.02' }
        ]
      },
      {
        srNo: '2',
        mainLocationTower: 'Tower T-102',
        readings: [
          { measurementPointLocation: 'Foot-1', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '12.0', measuredImpedance: '3.0' },
          { measurementPointLocation: 'Foot-2', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '12.0', measuredImpedance: '3.0' },
          { measurementPointLocation: 'Foot-3', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '12.0', measuredImpedance: '3.0' },
          { measurementPointLocation: 'Foot-4', footToEarthingConnectionStatus: 'Given', measuredCurrentMa: '12.0', measuredImpedance: '3.0' }
        ]
      }
    ]
  };
}

async function main() {
  const payload = buildPayload();
  const report = normalizeReportInput(payload);

  assert(report.soilResistivity.locations.length === 1, 'Soil should normalize to one location.');
  assert(report.soilResistivity.locations[0].direction1.length === 6, 'Soil direction 1 should normalize to 6 rows.');
  assert(report.electrodeResistance.length === 2, 'Electrode sheet should keep 2 rows.');
  assert(report.continuityTest.length === 3, 'Continuity sheet should keep 3 rows.');
  assert(report.loopImpedanceTest.length === 3, 'Loop impedance sheet should keep 3 rows.');
  assert(report.prospectiveFaultCurrent.length === 3, 'PFC sheet should keep 3 rows.');
  assert(report.riserIntegrityTest.length === 3, 'Riser sheet should keep 3 rows.');
  assert(report.earthContinuityTest.length === 3, 'Earth continuity sheet should keep 3 rows.');
  assert(report.towerFootingResistance.length === 2, 'Tower footing should keep 2 tower groups.');

  const soil = calculateSoilSummary(report);
  assert(soil.direction1Average === 85, `Expected soil direction 1 average 85, got ${soil.direction1Average}`);
  assert(soil.direction2Average === 95, `Expected soil direction 2 average 95, got ${soil.direction2Average}`);
  assert(soil.overallAverage === 90, `Expected soil overall average 90, got ${soil.overallAverage}`);
  assert(soil.category.label === 'Low', `Expected soil category Low, got ${soil.category.label}`);

  assert(deriveElectrodeAssessment(report.electrodeResistance[0]).status.label === 'Healthy', 'Electrode row 1 should be healthy.');
  assert(deriveElectrodeAssessment(report.electrodeResistance[1]).status.tone === 'critical', 'Electrode row 2 should be critical.');

  assert(deriveContinuityAssessment(report.continuityTest[0]).status.tone === 'healthy', 'Continuity row 1 should be healthy.');
  assert(deriveContinuityAssessment(report.continuityTest[1]).status.tone === 'warning', 'Continuity row 2 should be warning.');
  assert(deriveContinuityAssessment(report.continuityTest[2]).status.tone === 'critical', 'Continuity row 3 should be critical.');

  assert(deriveLoopImpedanceAssessment(report.loopImpedanceTest[0]).status.tone === 'healthy', 'Loop row 1 should be healthy.');
  assert(deriveLoopImpedanceAssessment(report.loopImpedanceTest[1]).status.tone === 'warning', 'Loop row 2 should be warning.');
  assert(deriveLoopImpedanceAssessment(report.loopImpedanceTest[2]).status.tone === 'critical', 'Loop row 3 should be critical.');

  assert(deriveFaultCurrentAssessment(report.prospectiveFaultCurrent[0]).status.tone === 'healthy', 'PFC row 1 should be healthy.');
  assert(deriveFaultCurrentAssessment(report.prospectiveFaultCurrent[1]).status.tone === 'warning', 'PFC row 2 should be warning.');
  assert(deriveFaultCurrentAssessment(report.prospectiveFaultCurrent[2]).status.tone === 'critical', 'PFC row 3 should be critical.');

  assert(deriveRiserAssessment(report.riserIntegrityTest[0]).status.tone === 'healthy', 'Riser row 1 should be healthy.');
  assert(deriveRiserAssessment(report.riserIntegrityTest[1]).status.tone === 'warning', 'Riser row 2 should be warning.');
  assert(deriveRiserAssessment(report.riserIntegrityTest[2]).status.tone === 'critical', 'Riser row 3 should be critical.');

  assert(deriveEarthContinuityAssessment(report.earthContinuityTest[0]).status.tone === 'healthy', 'Earth continuity row 1 should be healthy.');
  assert(deriveEarthContinuityAssessment(report.earthContinuityTest[1]).status.tone === 'warning', 'Earth continuity row 2 should be warning.');
  assert(deriveEarthContinuityAssessment(report.earthContinuityTest[2]).status.tone === 'critical', 'Earth continuity row 3 should be critical.');

  const towerSummaries = summarizeTowerGroups(report.towerFootingResistance);
  const towerA = report.towerFootingResistance[0];
  const towerAAssessment = deriveTowerFootingAssessment(towerA, towerSummaries.get(buildTowerGroupKey(towerA)));
  assert(towerAAssessment.totalImpedanceZt === 8.12, `Expected tower A Zt 8.12, got ${towerAAssessment.totalImpedanceZt}`);
  assert(towerAAssessment.totalCurrentItotal === 43.2, `Expected tower A current total 43.2, got ${towerAAssessment.totalCurrentItotal}`);
  assert(towerAAssessment.comment === 'Healthy', `Expected tower A remark Healthy, got ${towerAAssessment.comment}`);
  assert(report.towerFootingResistance[0].standardTolerableImpedanceZsat === '10', 'Tower footing Zsat should stay fixed at 10.');

  const towerB = report.towerFootingResistance[1];
  const towerBAssessment = deriveTowerFootingAssessment(towerB, towerSummaries.get(buildTowerGroupKey(towerB)));
  assert(towerBAssessment.totalImpedanceZt === 12, `Expected tower B Zt 12, got ${towerBAssessment.totalImpedanceZt}`);
  assert(towerBAssessment.totalCurrentItotal === 48, `Expected tower B current total 48, got ${towerBAssessment.totalCurrentItotal}`);
  assert(towerBAssessment.comment === 'Marginal', `Expected tower B remark Marginal, got ${towerBAssessment.comment}`);

  const snapshot = buildExecutiveSnapshot(report);
  assert(snapshot.selectedTests.length === 8, `Expected 8 selected tests, got ${snapshot.selectedTests.length}`);

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'electroreports-verify-'));
  const pdf = await generateElectroReportPdf(report, {
    generatedDir: tempDir,
    publicPathPrefix: '/generated-pdfs'
  });
  assert(fs.existsSync(pdf.filePath), 'Generated PDF file should exist.');
  assert(fs.statSync(pdf.filePath).size > 1024, 'Generated PDF should not be empty.');

  const storeFile = path.join(tempDir, 'reports.json');
  const store = createReportStore({ filePath: storeFile });
  const saved = store.addReport(report);
  assert(saved.id && saved.createdAt, 'Saved report should contain store metadata.');
  assert(store.listReports().length === 1, 'Store list should return the saved report.');
  assert(store.getReport(saved.id)?.id === saved.id, 'Store getReport should return the saved report.');
  assert(store.deleteReport(saved.id).deleted === true, 'Store delete should remove the saved report.');
  assert(store.listReports().length === 0, 'Store should be empty after delete.');

  fs.rmSync(tempDir, { recursive: true, force: true });

  console.log('[electroreports] Verification PASS');
  console.log(
    `[electroreports] Soil=${soil.overallAverage} ohm-m | ElectrodeCritical=1 | ContinuityWarnings=1 | ContinuityCritical=1 | TowerHealthy=${towerAAssessment.comment} | TowerMarginal=${towerBAssessment.comment}`
  );
}

module.exports = {
  buildPayload
};

if (require.main === module) {
  main().catch((error) => {
    console.error(`[electroreports] Verification FAIL: ${error.message}`);
    process.exit(1);
  });
}
