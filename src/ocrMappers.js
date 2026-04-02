const {
  createDefaultSoilLocation,
  createDefaultSoilRow,
  createDefaultElectrodeRow,
  createDefaultContinuityRow,
  createDefaultLoopImpedanceRow,
  createDefaultProspectiveFaultRow,
  createDefaultRiserIntegrityRow,
  createDefaultEarthContinuityRow,
  createDefaultTowerFootingGroup
} = require('./reportModel');

const PHASE_MEASURED_POINTS = ['R-E', 'Y-E', 'B-E'];

function mapSoilOcrPreviewToDraft(preview = {}) {
  const locations = (Array.isArray(preview.locations) ? preview.locations : []).map((location, index) => {
    const draftLocation = createDefaultSoilLocation(location.name || `Location ${index + 1}`);
    draftLocation.name = location.name || `Location ${index + 1}`;
    draftLocation.drivenElectrodeDiameter = String(location.drivenElectrodeDiameter || '').trim();
    draftLocation.drivenElectrodeLength = String(location.drivenElectrodeLength || '').trim();
    draftLocation.notes = String(location.notes || '').trim();
    draftLocation.direction1 = (Array.isArray(location.direction1) ? location.direction1 : []).map((row) => {
      const draftRow = createDefaultSoilRow(String(row.spacing || '').trim());
      draftRow.resistivity = String(row.resistivity || '').trim();
      return draftRow;
    });
    draftLocation.direction2 = (Array.isArray(location.direction2) ? location.direction2 : []).map((row) => {
      const draftRow = createDefaultSoilRow(String(row.spacing || '').trim());
      draftRow.resistivity = String(row.resistivity || '').trim();
      return draftRow;
    });
    return draftLocation;
  });

  return {
    locations
  };
}

module.exports = {
  mapSoilOcrPreviewToDraft,
  mapElectrodeOcrPreviewToDraft(preview = {}) {
    const rows = (Array.isArray(preview.rows) ? preview.rows : []).map((row) => {
      const draftRow = createDefaultElectrodeRow();
      draftRow.tag = String(row.tag || '').trim();
      draftRow.location = String(row.location || '').trim();
      draftRow.electrodeType = String(row.electrodeType || draftRow.electrodeType).trim() || draftRow.electrodeType;
      draftRow.materialType = String(row.materialType || draftRow.materialType).trim() || draftRow.materialType;
      draftRow.length = String(row.length || '').trim();
      draftRow.diameter = String(row.diameter || '').trim();
      draftRow.resistanceWithoutGrid = String(row.resistanceWithoutGrid || '').trim();
      draftRow.resistanceWithGrid = String(row.resistanceWithGrid || '').trim();
      return draftRow;
    });

    return { rows };
  },
  mapLoopImpedanceOcrPreviewToDraft(preview = {}) {
    const rows = [];
    (Array.isArray(preview.groups) ? preview.groups : []).forEach((group, groupIndex) => {
      (Array.isArray(group.rows) ? group.rows : []).forEach((row, rowIndex) => {
        const draftRow = createDefaultLoopImpedanceRow();
        draftRow.srNo = String(groupIndex * PHASE_MEASURED_POINTS.length + rowIndex + 1);
        draftRow.location = String(group.location || '').trim();
        draftRow.feederTag = String(group.feederTag || '').trim();
        draftRow.deviceType = String(group.deviceType || draftRow.deviceType).trim() || draftRow.deviceType;
        draftRow.deviceRating = String(group.deviceRating || '').trim();
        draftRow.breakingCapacity = String(group.breakingCapacity || '').trim();
        draftRow.measuredPoints = String(row.measuredPoints || PHASE_MEASURED_POINTS[rowIndex] || draftRow.measuredPoints).trim();
        draftRow.loopImpedance = String(row.loopImpedance || '').trim();
        draftRow.voltage = String(row.voltage || draftRow.voltage).trim() || draftRow.voltage;
        rows.push(draftRow);
      });
    });

    return { rows };
  },
  mapProspectiveFaultOcrPreviewToDraft(preview = {}) {
    const rows = [];
    (Array.isArray(preview.groups) ? preview.groups : []).forEach((group, groupIndex) => {
      (Array.isArray(group.rows) ? group.rows : []).forEach((row, rowIndex) => {
        const draftRow = createDefaultProspectiveFaultRow();
        draftRow.srNo = String(groupIndex * PHASE_MEASURED_POINTS.length + rowIndex + 1);
        draftRow.location = String(group.location || '').trim();
        draftRow.feederTag = String(group.feederTag || '').trim();
        draftRow.deviceType = String(group.deviceType || draftRow.deviceType).trim() || draftRow.deviceType;
        draftRow.deviceRating = String(group.deviceRating || '').trim();
        draftRow.breakingCapacity = String(group.breakingCapacity || '').trim();
        draftRow.measuredPoints = String(row.measuredPoints || PHASE_MEASURED_POINTS[rowIndex] || draftRow.measuredPoints).trim();
        draftRow.loopImpedance = String(row.loopImpedance || '').trim();
        draftRow.prospectiveFaultCurrent = String(row.prospectiveFaultCurrent || '').trim();
        draftRow.voltage = String(row.voltage || draftRow.voltage).trim() || draftRow.voltage;
        rows.push(draftRow);
      });
    });

    return { rows };
  },
  mapContinuityOcrPreviewToDraft(preview = {}) {
    const rows = (Array.isArray(preview.rows) ? preview.rows : []).map((row, index) => {
      const draftRow = createDefaultContinuityRow(String(row.srNo || index + 1).trim());
      draftRow.srNo = String(row.srNo || index + 1).trim();
      draftRow.mainLocation = String(row.mainLocation || '').trim();
      draftRow.measurementPoint = String(row.measurementPoint || '').trim();
      draftRow.resistance = String(row.resistance || '').trim();
      draftRow.impedance = String(row.impedance || '').trim();
      return draftRow;
    });

    return { rows };
  },
  mapRiserIntegrityOcrPreviewToDraft(preview = {}) {
    const rows = (Array.isArray(preview.rows) ? preview.rows : []).map((row, index) => {
      const draftRow = createDefaultRiserIntegrityRow();
      draftRow.srNo = String(row.srNo || index + 1).trim();
      draftRow.mainLocation = String(row.mainLocation || '').trim();
      draftRow.measurementPoint = String(row.measurementPoint || '').trim();
      draftRow.resistanceTowardsEquipment = String(row.resistanceTowardsEquipment || '').trim();
      draftRow.resistanceTowardsGrid = String(row.resistanceTowardsGrid || '').trim();
      return draftRow;
    });

    return { rows };
  },
  mapEarthContinuityOcrPreviewToDraft(preview = {}) {
    const rows = (Array.isArray(preview.rows) ? preview.rows : []).map((row, index) => {
      const draftRow = createDefaultEarthContinuityRow();
      draftRow.srNo = String(row.srNo || index + 1).trim();
      draftRow.tag = String(row.tag || '').trim();
      draftRow.locationBuildingName = String(row.locationBuildingName || '').trim();
      draftRow.distance = String(row.distance || '').trim();
      draftRow.measuredValue = String(row.measuredValue || '').trim();
      return draftRow;
    });

    return { rows };
  },
  mapTowerFootingOcrPreviewToDraft(preview = {}) {
    const rows = (Array.isArray(preview.groups) ? preview.groups : []).map((group, index) => {
      const draftGroup = createDefaultTowerFootingGroup();
      draftGroup.srNo = String(group.srNo || index + 1).trim();
      draftGroup.mainLocationTower = String(group.mainLocationTower || '').trim();
      draftGroup.readings = draftGroup.readings.map((reading, readingIndex) => ({
        ...reading,
        measurementPointLocation: String(group.readings?.[readingIndex]?.measurementPointLocation || reading.measurementPointLocation).trim(),
        footToEarthingConnectionStatus:
          String(group.readings?.[readingIndex]?.footToEarthingConnectionStatus || reading.footToEarthingConnectionStatus).trim() ||
          reading.footToEarthingConnectionStatus,
        measuredCurrentMa: String(group.readings?.[readingIndex]?.measuredCurrentMa || '').trim(),
        measuredImpedance: String(group.readings?.[readingIndex]?.measuredImpedance || '').trim()
      }));
      return draftGroup;
    });

    return { rows };
  }
};
