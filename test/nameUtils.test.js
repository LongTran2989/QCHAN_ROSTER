/**
 * node test/nameUtils.test.js
 */
var assert = require("assert");
var NameUtils = require("../NameUtils.js");
var toInitialsWithFirstName = NameUtils.toInitialsWithFirstName;
var splitAssignedPeople = NameUtils.splitAssignedPeople;
var joinAssignedPeople = NameUtils.joinAssignedPeople;
var formatAssignedShortNames = NameUtils.formatAssignedShortNames;

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

test("Tran Thanh Long -> TTLong", function () {
  assert.strictEqual(toInitialsWithFirstName("Trần Thanh Long"), "TTLong");
});

test("Vu Hong Hai -> VHHải (given name keeps its diacritics)", function () {
  assert.strictEqual(toInitialsWithFirstName("Vũ Hồng Hải"), "VHHải");
});

test("two-word name uses one initial", function () {
  assert.strictEqual(toInitialsWithFirstName("Nguyen Long"), "NLong");
});

test("single-word name is returned unchanged", function () {
  assert.strictEqual(toInitialsWithFirstName("Long"), "Long");
});

test("extra whitespace between/around words is ignored", function () {
  assert.strictEqual(toInitialsWithFirstName("  Tran   Thanh  Long  "), "TTLong");
});

test("empty/blank input returns empty string", function () {
  assert.strictEqual(toInitialsWithFirstName(""), "");
  assert.strictEqual(toInitialsWithFirstName("   "), "");
});

test("splitAssignedPeople splits a comma-joined cell into trimmed names", function () {
  assert.deepStrictEqual(splitAssignedPeople("Tran Thanh Long, Vu Hong Hai"), ["Tran Thanh Long", "Vu Hong Hai"]);
});

test("splitAssignedPeople drops empty entries from stray commas/whitespace", function () {
  assert.deepStrictEqual(splitAssignedPeople("Tran Thanh Long,  , "), ["Tran Thanh Long"]);
});

test("splitAssignedPeople of empty/blank input is an empty array", function () {
  assert.deepStrictEqual(splitAssignedPeople(""), []);
  assert.deepStrictEqual(splitAssignedPeople(null), []);
});

test("splitAssignedPeople of a single legacy (pre-multi-assignee) name is a one-element array", function () {
  assert.deepStrictEqual(splitAssignedPeople("Tran Thanh Long"), ["Tran Thanh Long"]);
});

test("joinAssignedPeople is the inverse of splitAssignedPeople for clean input", function () {
  assert.strictEqual(joinAssignedPeople(["Tran Thanh Long", "Vu Hong Hai"]), "Tran Thanh Long, Vu Hong Hai");
});

test("joinAssignedPeople drops blank/whitespace-only names", function () {
  assert.strictEqual(joinAssignedPeople(["Tran Thanh Long", "  ", ""]), "Tran Thanh Long");
});

test("joinAssignedPeople of an empty list is an empty string", function () {
  assert.strictEqual(joinAssignedPeople([]), "");
});

test("formatAssignedShortNames compacts each name and joins them", function () {
  assert.strictEqual(formatAssignedShortNames(["Trần Thanh Long", "Vũ Hồng Hải"]), "TTLong, VHHải");
});

test("formatAssignedShortNames of an empty list is an empty string", function () {
  assert.strictEqual(formatAssignedShortNames([]), "");
});

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
