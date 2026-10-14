// Auto-generated remark text and how to tell it from an engineer's own words.
//
// Remarks are generated at render time from the live status (see assessmentService).
// Earlier versions saved the generated sentence into the row's remark field at
// normalize time, so existing reports hold text that can contradict the status now
// computed from the same row. Those sentences are recognised here and treated as
// "no remark typed"; anything else is the engineer's and is kept untouched.

const GENERATED_REMARK_PATTERNS = [
  // continuity / earth continuity
  /^Enter a (resistance reading|measured value) to assess (continuity|earth continuity) against the /,
  /^Continuity resistance is (within|(well )?above) the .+?;? (bond\/joint continuity is acceptable|inspect joints, terminations, and bonding integrity|urgent investigation of the continuity path is recommended)\./,
  /^Earth continuity is (within|(well )?above) the .+?; (the earth path is acceptable for this point|inspect the earth path, joints, and terminations|urgent investigation of the earth path is recommended)\./,
  // loop impedance
  /^Enter measured Zs (and verify it against the protective device's maximum permitted Zs under|to compare it with the) /,
  /^Measured Zs is (in the healthy range, but final compliance still depends on|within the .+? a panel is healthy only if all three phases are\.)/,
  /^Measured Zs is elevated; compare it carefully against/,
  /^Measured Zs is high and likely to challenge disconnection performance/,
  /^Measured Zs exceeds the .+?; verify the PE path and confirm disconnection performance/,
  // prospective fault current
  /^Enter (both prospective fault current and device breaking capacity|the prospective fault current) to verify that /,
  /^Prospective fault current (is comfortably within the device breaking capacity\.|is within the device breaking capacity, but the safety margin is small\.|exceeds the device breaking capacity; review protective device selection immediately\.)/,
  /^Device breaking capacity (was not recorded, so the fault current cannot be assessed|[\d.]+ kA (does not exceed|exceeds) the governing fault current of )/,
  // riser
  /^Enter both resistance readings to assess (continuity|the riser) against the /,
  /^Riser (continuity )?readings are (within the |high; investigate the riser path)/,
  /^One or both riser (continuity )?readings (are above|exceed) the /,
  // tower footing: the remark used to be the bare status word
  /^(Healthy|Marginal|Not Acceptable|-)$/
];

function isGeneratedRemark(text) {
  const value = String(text === undefined || text === null ? '' : text).trim();
  return Boolean(value) && GENERATED_REMARK_PATTERNS.some((pattern) => pattern.test(value));
}

// What the engineer actually typed: '' when the field is empty or holds generated text.
function engineerRemark(text) {
  const value = String(text === undefined || text === null ? '' : text).trim();
  return isGeneratedRemark(value) ? '' : value;
}

// Fields that held generated remark text, by section.
const REMARK_FIELDS = Object.freeze({
  continuityTest: 'comment',
  loopImpedanceTest: 'remarks',
  prospectiveFaultCurrent: 'comment',
  riserIntegrityTest: 'comment',
  earthContinuityTest: 'remark',
  towerFootingResistance: 'remarks',
  groundResistance: 'remark',
  stepTouchVoltage: 'remark'
});

// A copy of the report with saved generated remarks cleared. Used when loading a
// report, so no stored sentence can outlive the rules that wrote it.
function scrubGeneratedRemarks(report) {
  if (!report || typeof report !== 'object') {
    return report;
  }
  const scrubbed = { ...report };
  Object.entries(REMARK_FIELDS).forEach(([section, field]) => {
    if (Array.isArray(report[section])) {
      scrubbed[section] = report[section].map((row) => (row && typeof row === 'object' ? { ...row, [field]: engineerRemark(row[field]) } : row));
    }
  });
  return scrubbed;
}

module.exports = { isGeneratedRemark, engineerRemark, scrubGeneratedRemarks, REMARK_FIELDS };
