/**
 * node test/shiftUtils.test.js
 * Plain-Node tests for ShiftUtils.js — no Google Apps Script APIs involved.
 */
var assert = require("assert");
var ShiftUtils = require("../ShiftUtils.js");
var SHIFT = ShiftUtils.SHIFT;
var classifyWP = ShiftUtils.classifyWP;
var getShiftForBangkokInstant = ShiftUtils.getShiftForBangkokInstant;
var toBangkokInstant = ShiftUtils.toBangkokInstant;
var formatBangkokTime = ShiftUtils.formatBangkokTime;
var buildShiftNote = ShiftUtils.buildShiftNote;
var buildShiftLabelSuffix = ShiftUtils.buildShiftLabelSuffix;
var buildCheckLabel = ShiftUtils.buildCheckLabel;
var setLabelAssignee = ShiftUtils.setLabelAssignee;

var passed = 0;
var failed = [];

function test(name, fn) {
  try {
    fn();
    passed++;
  } catch (err) {
    failed.push({ name: name, error: err });
  }
}

function utc(iso) {
  return new Date(iso);
}

// --- Boundary checks on the shift split itself ---

test("10:59 Bangkok is Morning (just before the Evening boundary)", function () {
  var d = toBangkokInstant(utc("2024-01-05T03:59:00Z")); // + 7h = 10:59
  assert.strictEqual(getShiftForBangkokInstant(d), SHIFT.MORNING);
});

test("11:00 Bangkok is Evening (start of Evening boundary)", function () {
  var d = toBangkokInstant(utc("2024-01-05T04:00:00Z")); // + 7h = 11:00
  assert.strictEqual(getShiftForBangkokInstant(d), SHIFT.EVENING);
});

test("22:59 Bangkok is Evening (just before Morning boundary)", function () {
  var d = toBangkokInstant(utc("2024-01-05T15:59:00Z")); // + 7h = 22:59
  assert.strictEqual(getShiftForBangkokInstant(d), SHIFT.EVENING);
});

test("23:00 Bangkok is Morning (start of Morning boundary, wraps midnight)", function () {
  var d = toBangkokInstant(utc("2024-01-05T16:00:00Z")); // + 7h = 23:00
  assert.strictEqual(getShiftForBangkokInstant(d), SHIFT.MORNING);
});

test("00:00 Bangkok is Morning", function () {
  var d = toBangkokInstant(utc("2024-01-05T17:00:00Z")); // + 7h = 00:00 next day
  assert.strictEqual(getShiftForBangkokInstant(d), SHIFT.MORNING);
});

// --- Full WP classification ---

test("same-day Evening job: no night shift needed", function () {
  var r = classifyWP(utc("2024-01-05T10:00:00Z"), utc("2024-01-05T12:00:00Z")); // 17:00 -> 19:00 Bangkok
  assert.strictEqual(r.fromShift, SHIFT.EVENING);
  assert.strictEqual(r.toShift, SHIFT.EVENING);
  assert.strictEqual(r.fromDay, 5);
  assert.strictEqual(r.toDay, 5);
  assert.strictEqual(r.nightShiftRequired, false);
});

test("Evening start, Morning finish next day: night shift required", function () {
  // 15:00Z -> 22:00 Bangkok (Evening, day 5); 20:00Z -> 03:00 Bangkok next day (Morning, day 6)
  var r = classifyWP(utc("2024-01-05T15:00:00Z"), utc("2024-01-05T20:00:00Z"));
  assert.strictEqual(r.fromShift, SHIFT.EVENING);
  assert.strictEqual(r.toShift, SHIFT.MORNING);
  assert.strictEqual(r.fromDay, 5);
  assert.strictEqual(r.toDay, 6);
  assert.strictEqual(r.nightShiftRequired, true);
});

test("Morning start and Morning finish same wrapped day: not a night-shift flag", function () {
  // 17:00Z Jan5 -> 00:00 Bangkok Jan6 (Morning); 18:00Z Jan5 -> 01:00 Bangkok Jan6 (Morning)
  var r = classifyWP(utc("2024-01-05T17:00:00Z"), utc("2024-01-05T18:00:00Z"));
  assert.strictEqual(r.fromShift, SHIFT.MORNING);
  assert.strictEqual(r.toShift, SHIFT.MORNING);
  assert.strictEqual(r.nightShiftRequired, false);
});

test("Evening start, Evening finish next day (spans a full day+): not flagged as night-shift-only", function () {
  // starts Evening day 5, ends Evening day 6 -- crosses a night but ALSO covers day-shift hours,
  // so it is not the narrow "starts evening, ends morning next day" case.
  var r = classifyWP(utc("2024-01-05T10:00:00Z"), utc("2024-01-06T12:00:00Z")); // 17:00 Jan5 -> 19:00 Jan6
  assert.strictEqual(r.fromShift, SHIFT.EVENING);
  assert.strictEqual(r.toShift, SHIFT.EVENING);
  assert.strictEqual(r.nightShiftRequired, false);
});

test("UTC day-bucketing bug regression: 23:00 UTC start lands on the NEXT Bangkok calendar day", function () {
  // Old code used the raw (unshifted) date's getDate(), which would have bucketed this on day 5.
  // Correctly shifted to Bangkok, 23:00Z Jan5 -> 06:00 Jan6.
  var r = classifyWP(utc("2024-01-05T23:00:00Z"), utc("2024-01-05T23:30:00Z"));
  assert.strictEqual(r.fromDay, 6);
  assert.strictEqual(r.toDay, 6);
  assert.strictEqual(r.fromShift, SHIFT.MORNING);
});

// --- formatting helpers ---

test("formatBangkokTime pads day/month/hour/minute to 2 digits", function () {
  var d = toBangkokInstant(utc("2024-01-05T02:05:00Z")); // + 7h = 09:05 Jan 5
  assert.strictEqual(formatBangkokTime(d), "05/01 09:05");
});

test("buildShiftLabelSuffix is empty for an Evening start without night shift", function () {
  var r = classifyWP(utc("2024-01-05T10:00:00Z"), utc("2024-01-05T12:00:00Z"));
  assert.strictEqual(buildShiftLabelSuffix(r), "");
});

test("buildShiftLabelSuffix is empty for a Morning start", function () {
  var r = classifyWP(utc("2024-01-05T00:00:00Z"), utc("2024-01-05T02:00:00Z"));
  assert.strictEqual(buildShiftLabelSuffix(r), "");
});

test("buildShiftLabelSuffix is a moon when night shift coverage is required", function () {
  var r = classifyWP(utc("2024-01-05T15:00:00Z"), utc("2024-01-05T20:00:00Z"));
  assert.strictEqual(buildShiftLabelSuffix(r), " 🌙");
});

test("buildShiftLabelSuffix is null-safe", function () {
  assert.strictEqual(buildShiftLabelSuffix(null), "");
  assert.strictEqual(buildShiftLabelSuffix(undefined), "");
});

test("buildCheckLabel combines star, check, moon and assignee on one line", function () {
  var label = buildCheckLabel({ acCheck: "4C", isChk: true, shiftInfo: { nightShiftRequired: true }, assigneeShort: "Nguyen T." });
  assert.strictEqual(label, "★ 4C 🌙 · Nguyen T.");
});

test("buildCheckLabel is just the check when plain", function () {
  assert.strictEqual(buildCheckLabel({ acCheck: "4C", isChk: false, shiftInfo: null }), "4C");
});

test("buildCheckLabel puts the reg after the star and before the check", function () {
  var label = buildCheckLabel({ acReg: "A350-1", acCheck: "4C", isChk: true, shiftInfo: { nightShiftRequired: true }, assigneeShort: "Nguyen T." });
  assert.strictEqual(label, "★ A350-1 4C 🌙 · Nguyen T.");
});

test("buildCheckLabel with a reg and no star", function () {
  assert.strictEqual(buildCheckLabel({ acReg: "A350-1", acCheck: "4C", isChk: false, shiftInfo: null }), "A350-1 4C");
});

test("buildCheckLabel never contains a newline, even in the reg", function () {
  var label = buildCheckLabel({ acReg: "A350\n-1", acCheck: "4C", isChk: false, shiftInfo: null });
  assert.ok(label.indexOf("\n") === -1, JSON.stringify(label));
});

test("buildCheckLabel never contains a newline", function () {
  var label = buildCheckLabel({ acCheck: "4C\nLONG", isChk: false, shiftInfo: null, assigneeShort: "A\nB" });
  assert.ok(label.indexOf("\n") === -1, JSON.stringify(label));
});

test("buildShiftNote includes both times and the night-shift warning when required", function () {
  var r = classifyWP(utc("2024-01-05T15:00:00Z"), utc("2024-01-05T20:00:00Z"));
  var note = buildShiftNote(r);
  assert.ok(note.indexOf("Start: 05/01 22:00 (Evening shift)") !== -1, note);
  assert.ok(note.indexOf("End: 06/01 03:00 (Morning shift)") !== -1, note);
  assert.ok(note.indexOf("Hanoi") === -1 && note.indexOf("Bangkok") === -1, note);
  assert.ok(note.indexOf("Night shift coverage required") !== -1, note);
});

test("buildShiftNote omits the warning when night shift isn't required", function () {
  var r = classifyWP(utc("2024-01-05T10:00:00Z"), utc("2024-01-05T12:00:00Z"));
  var note = buildShiftNote(r);
  assert.ok(note.indexOf("Night shift coverage required") === -1, note);
});

test("setLabelAssignee appends the assignee to an unassigned label", function () {
  assert.strictEqual(setLabelAssignee("★ 4C 🌙", "Tran B."), "★ 4C 🌙 · Tran B.");
});

test("setLabelAssignee keeps star, reg and moon on a new-format label", function () {
  assert.strictEqual(setLabelAssignee("★ A350-1 4C 🌙", "Tran B."), "★ A350-1 4C 🌙 · Tran B.");
  assert.strictEqual(setLabelAssignee("A350-1 4C · Nguyen T.", ""), "A350-1 4C");
});

test("setLabelAssignee replaces an existing assignee", function () {
  assert.strictEqual(setLabelAssignee("4C · Nguyen T.", "Tran B."), "4C · Tran B.");
});

test("setLabelAssignee with empty name removes the assignee", function () {
  assert.strictEqual(setLabelAssignee("4C · Nguyen T.", ""), "4C");
});

test("setLabelAssignee output never contains a newline, even for a legacy two-line label", function () {
  var out = setLabelAssignee("4C\nNguyen T.", "Tran B.");
  assert.ok(out.indexOf("\n") === -1, JSON.stringify(out));
  assert.strictEqual(out, "4C · Tran B.");
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
