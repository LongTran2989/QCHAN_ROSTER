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

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
