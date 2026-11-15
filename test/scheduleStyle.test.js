/**
 * node test/scheduleStyle.test.js
 * Plain-Node tests for ScheduleStyle.js and the schedule palette in Config.js.
 */
var assert = require("assert");
var COLORS = require("../Config.js").CONFIG.COLORS;
var Style = require("../ScheduleStyle.js");

function srgbToLab(hex) {
  var lin = [1, 3, 5].map(function (i) {
    var c = parseInt(hex.substr(i, 2), 16) / 255;
    return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  // sRGB -> XYZ (D65), then normalised by the D65 white point.
  var x = (0.4124564 * lin[0] + 0.3575761 * lin[1] + 0.1804375 * lin[2]) / 0.95047;
  var y = (0.2126729 * lin[0] + 0.7151522 * lin[1] + 0.0721750 * lin[2]) / 1.0;
  var z = (0.0193339 * lin[0] + 0.1191920 * lin[1] + 0.9503041 * lin[2]) / 1.08883;
  var f = function (t) { return t > 216 / 24389 ? Math.cbrt(t) : (24389 / 27 * t + 16) / 116; };
  var fx = f(x), fy = f(y), fz = f(z);
  return [116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)];
}

// Test-only helper (kept out of the Apps Script runtime): perceptual distance between two hex
// colors, CIELAB delta E76 (D65, sRGB). Roughly: < 10 looks alike, >= 25 is clearly different.
function colorDistance(hexA, hexB) {
  var a = srgbToLab(hexA);
  var b = srgbToLab(hexB);
  return Math.sqrt(Math.pow(a[0] - b[0], 2) + Math.pow(a[1] - b[1], 2) + Math.pow(a[2] - b[2], 2));
}

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

// Pastel fills must hit >= 7:1 against black. EA_LAN (taupe) is only ~6.3:1, so it is a fill
// (>= 4.5:1 with its auto font) but not a "pastel".
var PASTEL_KEYS = ["A320", "A350", "B787", "DEFAULT", "EA_HAN", "NIGHT_SHIFT_FLAG"];
var FILL_KEYS = PASTEL_KEYS.concat(["EA_LAN"]);
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

FILL_KEYS.concat(["CHANGE_HIGHLIGHT", "UNASSIGNED_BORDER", "NIGHT_REG_FONT"]).forEach(function (key) {
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

// --- station colors, distance, phase style ---

test("station colors and night reg font have the specified values", function () {
  assert.strictEqual(COLORS.EA_HAN, "#66BB6A");
  assert.strictEqual(COLORS.EA_LAN, "#A1887F");
  assert.strictEqual(COLORS.NIGHT_REG_FONT, "#C62828");
});

test("night reg font has >= 4.5:1 contrast on white", function () {
  assert.ok(Style.contrastRatio(COLORS.NIGHT_REG_FONT, "#FFFFFF") >= 4.5);
});

test("colorDistance of a color with itself is 0", function () {
  assert.strictEqual(colorDistance("#000000", "#000000"), 0);
  assert.strictEqual(colorDistance("#66BB6A", "#66BB6A"), 0);
});

test("colorDistance black/white is about 100 (CIELAB L*)", function () {
  var d = colorDistance("#000000", "#FFFFFF");
  assert.ok(Math.abs(d - 100) < 0.5, "d " + d);
});

["EA_HAN", "EA_LAN"].forEach(function (key) {
  ["A320", "A350", "B787", "DEFAULT", "NIGHT_SHIFT_FLAG"].concat(key === "EA_HAN" ? ["EA_LAN"] : ["EA_HAN"]).forEach(function (other) {
    test("colorDistance " + key + " vs " + other + " is >= 25", function () {
      var d = colorDistance(COLORS[key], COLORS[other]);
      assert.ok(d >= 25, "d " + d);
    });
  });
});

test("stationFill maps L-HAN to EA_LAN and everything else to EA_HAN", function () {
  assert.strictEqual(Style.stationFill("L-HAN", COLORS), COLORS.EA_LAN);
  assert.strictEqual(Style.stationFill("HAN", COLORS), COLORS.EA_HAN);
});

["HAN", "L-HAN"].forEach(function (station) {
  var fill = station === "HAN" ? "#66BB6A" : "#A1887F";
  test("resolvePhaseStyle " + station + " night: station fill, black label, bold red reg", function () {
    var s = Style.resolvePhaseStyle({ station: station, nightShiftRequired: true }, COLORS);
    assert.strictEqual(s.labelBg, fill);
    assert.notStrictEqual(s.labelBg, COLORS.NIGHT_SHIFT_FLAG);
    assert.strictEqual(s.labelFont, "#000000");
    assert.strictEqual(s.bold, true);
    assert.strictEqual(s.regFont, "#C62828");
    assert.strictEqual(s.regBold, true);
  });
  test("resolvePhaseStyle " + station + " non-night: no reg font override", function () {
    var s = Style.resolvePhaseStyle({ station: station, nightShiftRequired: false }, COLORS);
    assert.strictEqual(s.labelBg, fill);
    assert.strictEqual(s.regFont, null);
    assert.strictEqual(s.regBold, false);
  });
  test("resolvePhaseStyle " + station + " label font has >= 4.5:1 contrast", function () {
    var s = Style.resolvePhaseStyle({ station: station, nightShiftRequired: false }, COLORS);
    assert.ok(Style.contrastRatio(s.labelBg, s.labelFont) >= 4.5);
  });
});

test("resolveCheckStyle with opts.fill uses it for the bar and label", function () {
  var s = style({ fill: COLORS.EA_LAN });
  assert.strictEqual(s.barColor, COLORS.EA_LAN);
  assert.strictEqual(s.labelBg, COLORS.EA_LAN);
  assert.ok(Style.contrastRatio(s.labelBg, s.labelFont) >= 4.5);
});

test("resolveCheckStyle with opts.fill still goes light red on the label when night", function () {
  var s = style({ fill: COLORS.EA_LAN, nightShiftRequired: true });
  assert.strictEqual(s.barColor, COLORS.EA_LAN);
  assert.strictEqual(s.labelBg, "#EF9A9A");
  assert.strictEqual(s.labelFont, "#000000");
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
