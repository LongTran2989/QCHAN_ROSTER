/**
 * Pure UTC -> Bangkok (UTC+7) shift classification for AC CHECK schedule WPs.
 * Source FROM/TO values from the AC CHECKS sheet are naive UTC datetimes.
 * No SpreadsheetApp dependency, so this can be exercised outside Apps Script (see test/shiftUtils.test.js).
 */

var SHIFT = { MORNING: "M", EVENING: "E" };
var BANGKOK_OFFSET_HOURS = 7;

/**
 * Shifts a naive-UTC Date by the Bangkok offset, returning a new Date whose
 * UTC-field getters (getUTCHours, getUTCDate, ...) read as Bangkok wall-clock values.
 * @param {Date} utcDate
 * @returns {Date}
 */
function toBangkokInstant(utcDate) {
  return new Date(utcDate.getTime() + BANGKOK_OFFSET_HOURS * 60 * 60 * 1000);
}

/**
 * Evening: local hour in [11:00, 23:00). Morning: local hour in [23:00, 11:00) (wraps midnight).
 * @param {Date} bangkokInstant - a Date already shifted via toBangkokInstant.
 * @returns {string} SHIFT.MORNING or SHIFT.EVENING
 */
function getShiftForBangkokInstant(bangkokInstant) {
  var hour = bangkokInstant.getUTCHours();
  return (hour >= 11 && hour < 23) ? SHIFT.EVENING : SHIFT.MORNING;
}

/**
 * Classifies a WP's start/end into Bangkok-local day + shift, and flags whether
 * it needs night-shift coverage (starts Evening, ends Morning on a later calendar day).
 * @param {Date} fromUTC
 * @param {Date} toUTC
 */
function classifyWP(fromUTC, toUTC) {
  var fromBangkok = toBangkokInstant(fromUTC);
  var toBangkok = toBangkokInstant(toUTC);

  var fromShift = getShiftForBangkokInstant(fromBangkok);
  var toShift = getShiftForBangkokInstant(toBangkok);

  var fromDayKey = Date.UTC(fromBangkok.getUTCFullYear(), fromBangkok.getUTCMonth(), fromBangkok.getUTCDate());
  var toDayKey = Date.UTC(toBangkok.getUTCFullYear(), toBangkok.getUTCMonth(), toBangkok.getUTCDate());

  var nightShiftRequired = (fromShift === SHIFT.EVENING) && (toShift === SHIFT.MORNING) && (toDayKey > fromDayKey);

  return {
    fromBangkok: fromBangkok,
    toBangkok: toBangkok,
    fromDay: fromBangkok.getUTCDate(),
    toDay: toBangkok.getUTCDate(),
    fromShift: fromShift,
    toShift: toShift,
    nightShiftRequired: nightShiftRequired
  };
}

function pad2(n) {
  return (n < 10 ? "0" : "") + n;
}

/**
 * Formats a Bangkok-shifted instant (as produced by toBangkokInstant) as "dd/mm HH:mm".
 * @param {Date} bangkokInstant
 * @returns {string}
 */
function formatBangkokTime(bangkokInstant) {
  return pad2(bangkokInstant.getUTCDate()) + "/" + pad2(bangkokInstant.getUTCMonth() + 1) +
    " " + pad2(bangkokInstant.getUTCHours()) + ":" + pad2(bangkokInstant.getUTCMinutes());
}

function shiftName(shift) {
  return shift === SHIFT.EVENING ? "Evening" : "Morning";
}

/**
 * Multi-line tooltip text for a WP's start/end: times and shift names only (the
 * timezone is not shown to the user).
 * @param {ReturnType<typeof classifyWP>} shiftInfo
 * @returns {string}
 */
function buildShiftNote(shiftInfo) {
  var lines = [
    "Start: " + formatBangkokTime(shiftInfo.fromBangkok) + " (" + shiftName(shiftInfo.fromShift) + " shift)",
    "End: " + formatBangkokTime(shiftInfo.toBangkok) + " (" + shiftName(shiftInfo.toShift) + " shift)"
  ];
  if (shiftInfo.nightShiftRequired) {
    lines.push("⚠ Night shift coverage required");
  }
  return lines.join("\n");
}

/**
 * Compact suffix appended to a check label: a moon when night-shift coverage is required, else
 * empty. Shift letters (E/M) live in the cell note, not the label.
 * @param {ReturnType<typeof classifyWP>|null|undefined} shiftInfo
 * @returns {string}
 */
function buildShiftLabelSuffix(shiftInfo) {
  return (shiftInfo && shiftInfo.nightShiftRequired) ? " \u{1F319}" : "";
}

/**
 * Single-line label: [star ][<reg> ]<check>[ moon][ middot <assignee>]. Newlines are collapsed to
 * spaces so the cell never needs to wrap (keeps short-TAT rows at the fixed row height). An
 * omitted/empty acReg adds nothing.
 * @param {{acCheck: string, isChk: boolean, shiftInfo: object|null, assigneeShort?: string, acReg?: string}} opts
 * @returns {string}
 */
function buildCheckLabel(opts) {
  var flat = function (s) { return String(s == null ? "" : s).replace(/\s*[\r\n]+\s*/g, " "); };
  var reg = flat(opts.acReg).trim();
  var label = (opts.isChk ? "★ " : "") + (reg ? reg + " " : "") + flat(opts.acCheck) + buildShiftLabelSuffix(opts.shiftInfo);
  if (opts.assigneeShort) label += " · " + flat(opts.assigneeShort);
  return label;
}

/**
 * Text for a reg cell: the reg, plus " NS" for 0.5-day TAT WPs.
 * @param {string} acReg
 * @param {number|string} tat
 * @returns {string}
 */
function buildRegLabel(acReg, tat) {
  return String(tat) === "0.5" ? acReg + " NS" : acReg;
}

/**
 * Rewrites the assignee part of an existing check label (the text after " · "), keeping the
 * rest. Legacy two-line labels are reduced to their first line. Empty name removes the assignee.
 * @param {string} label
 * @param {string} shortName
 * @returns {string} single-line label
 */
function setLabelAssignee(label, shortName) {
  var base = String(label == null ? "" : label).split(/\r?\n/)[0];
  var idx = base.indexOf(" · ");
  if (idx >= 0) base = base.slice(0, idx);
  return shortName ? base + " · " + shortName : base;
}

// Allow `node test/shiftUtils.test.js` to require() this file, while GAS still sees plain globals.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    SHIFT: SHIFT,
    BANGKOK_OFFSET_HOURS: BANGKOK_OFFSET_HOURS,
    toBangkokInstant: toBangkokInstant,
    getShiftForBangkokInstant: getShiftForBangkokInstant,
    classifyWP: classifyWP,
    formatBangkokTime: formatBangkokTime,
    buildShiftNote: buildShiftNote,
    buildShiftLabelSuffix: buildShiftLabelSuffix,
    buildCheckLabel: buildCheckLabel,
    buildRegLabel: buildRegLabel,
    setLabelAssignee: setLabelAssignee
  };
}
