/**
 * node test/nameUtils.test.js
 */
var assert = require("assert");
var toInitialsWithFirstName = require("../NameUtils.js").toInitialsWithFirstName;

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

// --- report ---

console.log(passed + " passed, " + failed.length + " failed");
failed.forEach(function (f) {
  console.log("FAIL: " + f.name);
  console.log("  " + f.error.message);
});
process.exit(failed.length > 0 ? 1 : 0);
