/**
 * node test/scheduleLayout.test.js
 * Plain-Node tests for ScheduleLayout.js (Phase slot packing + section row layout).
 */
var assert = require("assert");
var Layout = require("../ScheduleLayout.js");

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

// Deterministic PRNG so a failing property is reproducible.
function rng(seed) {
  var s = seed;
  return function () {
    s = (s * 1664525 + 1013904223) % 4294967296;
    return s / 4294967296;
  };
}

function layout(over) {
  var base = { phaseHeaderRow: 76, phaseDepth: 1, normalCount: 10, otherCount: 0, stoCount: 3 };
  Object.keys(over || {}).forEach(function (k) { base[k] = over[k]; });
  return Layout.computeScheduleLayout(base);
}

test("assignPhaseSlots stacks same-day entries and counts depth", function () {
  var r = Layout.assignPhaseSlots([5, 5, 5, 7, 5, 7]);
  assert.deepStrictEqual(r.slots, [0, 1, 2, 0, 3, 1]);
  assert.strictEqual(r.depth, 4);
});

test("assignPhaseSlots of an empty list", function () {
  assert.deepStrictEqual(Layout.assignPhaseSlots([]), { slots: [], depth: 0 });
});

test("assignPhaseSlots: no two entries share (col, slot) for random inputs", function () {
  var rand = rng(12345);
  for (var iter = 0; iter < 200; iter++) {
    var n = Math.floor(rand() * 60);
    var cols = [];
    for (var i = 0; i < n; i++) cols.push(3 + Math.floor(rand() * 31));
    var r = Layout.assignPhaseSlots(cols);
    assert.strictEqual(r.slots.length, cols.length);
    var seen = {};
    cols.forEach(function (c, idx) {
      var key = c + ":" + r.slots[idx];
      assert.ok(!seen[key], "duplicate " + key + " for " + JSON.stringify(cols));
      seen[key] = true;
      assert.ok(r.slots[idx] < r.depth);
    });
  }
});

test("computeScheduleLayout reproduces today's rows with no Other short group", function () {
  var l = layout();
  assert.strictEqual(l.phaseHeaderRow, 76);
  assert.strictEqual(l.phaseStartRow, 77);
  assert.strictEqual(l.normalHeaderRow, 79);
  assert.strictEqual(l.normalStartRow, 80);
  assert.strictEqual(l.otherHeaderRow, null);
  assert.strictEqual(l.otherStartRow, null);
  assert.strictEqual(l.endOfListRow, 92);
  assert.strictEqual(l.stoStartRow, 93);
});

test("lastRow is the last row the layout writes (END OF LIST when there is no STO)", function () {
  assert.strictEqual(layout().lastRow, 95);
  assert.strictEqual(layout({ stoCount: 0 }).lastRow, 92);
  assert.strictEqual(layout({ otherCount: 4, stoCount: 2 }).lastRow, 100);
});

test("computeScheduleLayout places the Other short group between Normal and END OF LIST", function () {
  var l = layout({ otherCount: 4 });
  assert.strictEqual(l.otherHeaderRow, 91);
  assert.strictEqual(l.otherStartRow, 92);
  assert.strictEqual(l.endOfListRow, 98);
  assert.strictEqual(l.stoStartRow, 99);
});

test("NORMAL CHECKS header sits on the row right after the last Phase line", function () {
  assert.strictEqual(layout({ phaseDepth: 0 }).phaseRowCount, 0);
  assert.strictEqual(layout({ phaseDepth: 0 }).normalHeaderRow, 77);
  assert.strictEqual(layout({ phaseDepth: 1 }).phaseRowCount, 2);
  var l = layout({ phaseDepth: 6 });
  assert.strictEqual(l.phaseRowCount, 12);
  assert.strictEqual(l.normalHeaderRow, l.phaseStartRow + 2 * 5 + 1 + 1);
});

test("computeScheduleLayout: sections never overlap (property)", function () {
  for (var depth = 0; depth <= 12; depth++) {
    for (var counts = 0; counts <= 40; counts += 1) {
      [0, counts].forEach(function (other) {
        var l = layout({ phaseDepth: depth, normalCount: counts, otherCount: other, stoCount: 5 });
        var lastPhaseRow = l.phaseStartRow + 2 * (depth - 1) + 1;
        if (depth > 0) assert.ok(lastPhaseRow < l.normalHeaderRow, "phase " + depth);
        assert.ok(l.phaseHeaderRow < l.phaseStartRow);
        assert.ok(l.normalHeaderRow < l.normalStartRow);
        var lastNormalRow = l.normalStartRow + counts - 1;
        if (other > 0) {
          assert.ok(lastNormalRow < l.otherHeaderRow, "normal before other");
          assert.ok(l.otherHeaderRow < l.otherStartRow);
          var lastOtherRow = l.otherStartRow + other - 1;
          assert.ok(lastOtherRow < l.endOfListRow, "other before END OF LIST");
        } else {
          assert.strictEqual(l.otherHeaderRow, null);
          assert.strictEqual(l.otherStartRow, null);
          assert.ok(lastNormalRow < l.endOfListRow, "normal before END OF LIST");
        }
        assert.strictEqual(l.stoStartRow, l.endOfListRow + 1);
        assert.ok(l.endOfListRow > l.normalStartRow);
      });
    }
  }
});

test("computeScheduleLayout with zero counts everywhere is still valid", function () {
  var l = layout({ phaseDepth: 0, normalCount: 0, otherCount: 0, stoCount: 0 });
  assert.ok(l.endOfListRow > l.normalStartRow);
  assert.strictEqual(l.otherHeaderRow, null);
  assert.strictEqual(l.stoStartRow, l.endOfListRow + 1);
});

test("compareScheduleRows: start day, then end day, then reg, mixing A/C types", function () {
  // [type, reg, check, from, to]
  var rows = [
    ["A350", "Z", "", 5, 9],
    ["A321", "B", "", 1, 8],
    ["A350", "A", "", 1, 3],
    ["A321", "C", "", 1, 8],
    ["A350", "Y", "", 5, 6]
  ];
  var order = rows.slice().sort(Layout.compareScheduleRows).map(function (r) { return r[1]; });
  assert.deepStrictEqual(order, ["A", "B", "C", "Y", "Z"]);
});

test("compareByTypeThenStair: A/C type groups first, staircase within each", function () {
  var rows = [
    ["A350", "Z", "", 1, 4],
    ["A321", "B", "", 6, 8],
    ["A350", "A", "", 2, 3],
    ["A321", "C", "", 1, 2]
  ];
  var order = rows.slice().sort(Layout.compareByTypeThenStair).map(function (r) { return r[1]; });
  assert.deepStrictEqual(order, ["C", "B", "Z", "A"]);
});

test("compareByTypeThenStair: A320 and A321 form one staircase", function () {
  var rows = [
    ["A321", "B", "", 5, 9],
    ["A320", "C", "", 1, 2],
    ["A350", "D", "", 1, 2],
    ["A321 ", "A", "", 3, 4]
  ];
  var order = rows.slice().sort(Layout.compareByTypeThenStair).map(function (r) { return r[1]; });
  assert.deepStrictEqual(order, ["C", "A", "B", "D"]);
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
