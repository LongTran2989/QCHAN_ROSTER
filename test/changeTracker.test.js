/**
 * node test/changeTracker.test.js
 * Plain-Node tests for ChangeTracker.js's pure diff logic — no Google Apps Script APIs.
 */
var assert = require("assert");
var ChangeTracker = require("../ChangeTracker.js");
var diffWPLists = ChangeTracker.diffWPLists;
var buildChangeLogRows = ChangeTracker.buildChangeLogRows;
var buildAssignmentChangeLogRow = ChangeTracker.buildAssignmentChangeLogRow;

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

function row(overrides) {
  return Object.assign({
    pjid: "PJ1",
    acReg: "VN-A123",
    acCheck: "C4000",
    from: 1000,
    to: 2000,
    station: "HAN",
    assignedPerson: ""
  }, overrides);
}

test("identical rows produce no diff entries", function () {
  var oldRows = [row()];
  var newRows = [row()];
  var result = diffWPLists(oldRows, newRows);
  assert.deepStrictEqual(result.added, []);
  assert.deepStrictEqual(result.removed, []);
  assert.deepStrictEqual(result.changed, []);
});

test("a PJID only in newRows is Added", function () {
  var result = diffWPLists([], [row()]);
  assert.strictEqual(result.added.length, 1);
  assert.strictEqual(result.added[0].pjid, "PJ1");
  assert.strictEqual(result.removed.length, 0);
  assert.strictEqual(result.changed.length, 0);
});

test("a PJID only in oldRows is Removed", function () {
  var result = diffWPLists([row()], []);
  assert.strictEqual(result.removed.length, 1);
  assert.strictEqual(result.removed[0].pjid, "PJ1");
  assert.strictEqual(result.added.length, 0);
  assert.strictEqual(result.changed.length, 0);
});

test("a single differing field produces one Changed entry with one changedField", function () {
  var oldRows = [row({ from: 1000 })];
  var newRows = [row({ from: 5000 })];
  var result = diffWPLists(oldRows, newRows);
  assert.strictEqual(result.changed.length, 1);
  assert.strictEqual(result.changed[0].pjid, "PJ1");
  assert.deepStrictEqual(result.changed[0].changedFields, [{ field: "from", oldValue: 1000, newValue: 5000 }]);
});

test("multiple differing fields are all captured on one Changed entry", function () {
  var oldRows = [row({ from: 1000, to: 2000, station: "HAN" })];
  var newRows = [row({ from: 5000, to: 6000, station: "L-HAN" })];
  var result = diffWPLists(oldRows, newRows);
  assert.strictEqual(result.changed.length, 1);
  var fields = result.changed[0].changedFields.map(function (f) { return f.field; });
  assert.deepStrictEqual(fields.sort(), ["from", "station", "to"]);
});

test("assignedPerson differences are NOT reported as a diff (tracked separately by the sidebar)", function () {
  var oldRows = [row({ assignedPerson: "Alice" })];
  var newRows = [row({ assignedPerson: "Bob" })];
  var result = diffWPLists(oldRows, newRows);
  assert.deepStrictEqual(result.changed, []);
});

test("added/removed/changed can all occur in the same diff", function () {
  var oldRows = [row({ pjid: "PJ1" }), row({ pjid: "PJ2", acCheck: "C1" })];
  var newRows = [row({ pjid: "PJ1" }), row({ pjid: "PJ2", acCheck: "C2" }), row({ pjid: "PJ3" })];
  var result = diffWPLists(oldRows, newRows);
  assert.strictEqual(result.added.length, 1);
  assert.strictEqual(result.added[0].pjid, "PJ3");
  assert.strictEqual(result.removed.length, 0);
  assert.strictEqual(result.changed.length, 1);
  assert.strictEqual(result.changed[0].pjid, "PJ2");
});

test("buildChangeLogRows emits one row per added/removed WP and one row per changed field", function () {
  var diffResult = diffWPLists(
    [row({ pjid: "PJ1", from: 1000, to: 2000 })],
    [row({ pjid: "PJ1", from: 1000, to: 9999 }), row({ pjid: "PJ2" })]
  );
  var rows = buildChangeLogRows(diffResult, "2024-01-01 00:00", "tester@example.com", "JAN24 R0");
  // 1 Added (PJ2) + 1 Changed field (to) for PJ1
  assert.strictEqual(rows.length, 2);
  var types = rows.map(function (r) { return r[6]; });
  assert.deepStrictEqual(types.sort(), ["Added", "Changed"]);
});

test("buildAssignmentChangeLogRow shapes a single Assigned entry", function () {
  var r = buildAssignmentChangeLogRow("2024-01-01 00:00", "tester@example.com", "JAN24 R0", "PJ1", "VN-A123", "C4000", "Alice", "Bob");
  assert.strictEqual(r[6], "Assigned");
  assert.strictEqual(r[7], "assignedPerson");
  assert.strictEqual(r[8], "Alice");
  assert.strictEqual(r[9], "Bob");
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
