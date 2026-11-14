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

var PASTEL_KEYS = ["A320", "A350", "B787", "DEFAULT", "EA_LAN", "EA_HAN", "NIGHT_SHIFT_FLAG"];
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

function style(over) {
  var base = { acType: "A350", isChk: false, nightShiftRequired: false, assigned: true, changed: false };
  Object.keys(over || {}).forEach(function (k) { base[k] = over[k]; });
  return Style.resolveCheckStyle(base, COLORS);
}

test("resolveCheckStyle: CHK + night is black bold on light red", function () {
  var s = style({ isChk: true, nightShiftRequired: true });
  assert.strictEqual(s.labelBg, "#EF9A9A");
  assert.strictEqual(s.labelFont, "#000000");
  assert.strictEqual(s.bold, true);
});

test("resolveCheckStyle: plain A350 uses the blue bar with black, non-bold text", function () {
  var s = style();
  assert.strictEqual(s.barColor, "#81D4FA");
  assert.strictEqual(s.labelBg, "#81D4FA");
  assert.strictEqual(s.labelFont, "#000000");
  assert.strictEqual(s.bold, false);
});

test("resolveCheckStyle: A320 and A321 share the yellow bar", function () {
  assert.strictEqual(style({ acType: "A320" }).barColor, "#FFF59D");
  assert.strictEqual(style({ acType: "A321" }).barColor, "#FFF59D");
});

test("resolveCheckStyle: unknown AC type falls back to DEFAULT", function () {
  assert.strictEqual(style({ acType: "XYZ" }).barColor, COLORS.DEFAULT);
});

test("resolveCheckStyle: changed + unassigned shows the solid thick highlight", function () {
  var s = style({ changed: true, assigned: false });
  assert.deepStrictEqual(s.border, { color: "#6A1B9A", style: "SOLID_THICK" });
});

test("resolveCheckStyle: unassigned only shows the dashed border", function () {
  var s = style({ assigned: false });
  assert.deepStrictEqual(s.border, { color: "#424242", style: "DASHED" });
});

test("resolveCheckStyle: assigned and unchanged has no border", function () {
  assert.strictEqual(style().border, null);
});

["A320", "A321", "A350", "B787", "XYZ"].forEach(function (type) {
  [false, true].forEach(function (night) {
    test("resolveCheckStyle: " + type + (night ? " night" : "") + " label font has >= 4.5:1 contrast", function () {
      var s = style({ acType: type, nightShiftRequired: night });
      var ratio = Style.contrastRatio(s.labelBg, s.labelFont);
      assert.ok(ratio >= 4.5, "ratio " + ratio);
    });
  });
});

// Review finding: label text overflows past its own cell, so the font must also be legible on
// every bar color and on plain white, not just on the label cell's own fill.
["A320", "A350", "B787", "DEFAULT"].forEach(function (key) {
  test("night label font stays legible over spilled " + key + " bar cells", function () {
    var s = style({ nightShiftRequired: true });
    assert.ok(Style.contrastRatio(COLORS[key], s.labelFont) >= 4.5, key);
  });
});

test("night label font stays legible over plain white cells", function () {
  assert.ok(Style.contrastRatio("#FFFFFF", style({ nightShiftRequired: true }).labelFont) >= 4.5);
});

test("orderBorders puts SOLID_THICK borders after DASHED ones, preserving relative order", function () {
  var ops = [
    { row: 1, border: { style: "SOLID_THICK" } },
    { row: 2, border: { style: "DASHED" } },
    { row: 3, border: { style: "SOLID_THICK" } },
    { row: 4, border: { style: "DASHED" } }
  ];
  var ordered = Style.orderBorders(ops).map(function (o) { return o.row; });
  assert.deepStrictEqual(ordered, [2, 4, 1, 3]);
  assert.deepStrictEqual(ops.map(function (o) { return o.row; }), [1, 2, 3, 4], "input must not be mutated");
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
