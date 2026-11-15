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

test("PHASE_MIN_ROWS is 9", function () {
  assert.strictEqual(Layout.PHASE_MIN_ROWS, 9);
});

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
  assert.strictEqual(l.normalHeaderRow, 86);
  assert.strictEqual(l.normalStartRow, 87);
  assert.strictEqual(l.otherHeaderRow, null);
  assert.strictEqual(l.otherStartRow, null);
  assert.strictEqual(l.endOfListRow, 99);
  assert.strictEqual(l.stoStartRow, 100);
});

test("computeScheduleLayout places the Other short group between Normal and END OF LIST", function () {
  var l = layout({ otherCount: 4 });
  assert.strictEqual(l.otherHeaderRow, 98);
  assert.strictEqual(l.otherStartRow, 99);
  assert.strictEqual(l.endOfListRow, 105);
  assert.strictEqual(l.stoStartRow, 106);
});

test("phaseRowCount is 9 for depth 0 and 1, and grows 2 rows per extra level", function () {
  assert.strictEqual(layout({ phaseDepth: 0 }).phaseRowCount, 9);
  assert.strictEqual(layout({ phaseDepth: 1 }).phaseRowCount, 9);
  var l = layout({ phaseDepth: 6 });
  assert.strictEqual(l.phaseRowCount, 13);
  assert.strictEqual(l.normalHeaderRow, 77 + 13);
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

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
