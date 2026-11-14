/**
 * node test/scheduleStyle.test.js
 * Plain-Node tests for ScheduleStyle.js and the schedule palette in Config.js.
 */
var assert = require("assert");
var COLORS = require("../Config.js").CONFIG.COLORS;
var Style = require("../ScheduleStyle.js");

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

var PASTEL_KEYS = ["A320", "A350", "B787", "DEFAULT", "EA_LAN", "EA_HAN"];
var FILL_KEYS = PASTEL_KEYS.concat(["NIGHT_SHIFT_FLAG"]);
var HEX = /^#[0-9A-Fa-f]{6}$/;

test("contrastRatio black on white is 21", function () {
  assert.ok(Math.abs(Style.contrastRatio("#000000", "#FFFFFF") - 21) < 0.01);
});

test("contrastRatio of a color with itself is 1", function () {
  assert.ok(Math.abs(Style.contrastRatio("#B71C1C", "#B71C1C") - 1) < 0.001);
});

test("pickFontColor returns white on the night-shift red", function () {
  assert.strictEqual(Style.pickFontColor("#B71C1C"), "#FFFFFF");
});

test("pickFontColor returns black on pastel yellow", function () {
  assert.strictEqual(Style.pickFontColor("#FFF59D"), "#000000");
});

FILL_KEYS.concat(["CHANGE_HIGHLIGHT", "UNASSIGNED_BORDER"]).forEach(function (key) {
  test("palette " + key + " is a 6-digit hex", function () {
    assert.ok(HEX.test(COLORS[key]), key + " = " + COLORS[key]);
  });
});

FILL_KEYS.forEach(function (key) {
  test("palette " + key + " has >= 4.5:1 contrast with its auto font color", function () {
    var fill = COLORS[key];
    var ratio = Style.contrastRatio(fill, Style.pickFontColor(fill));
    assert.ok(ratio >= 4.5, key + " ratio " + ratio);
  });
});

PASTEL_KEYS.forEach(function (key) {
  test("pastel " + key + " has >= 7:1 contrast with black", function () {
    var ratio = Style.contrastRatio(COLORS[key], "#000000");
    assert.ok(ratio >= 7, key + " ratio " + ratio);
  });
});

test("CONFIG.LAYOUT.ROW_HEIGHT is 21", function () {
  assert.strictEqual(require("../Config.js").CONFIG.LAYOUT.ROW_HEIGHT, 21);
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
