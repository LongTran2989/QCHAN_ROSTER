/** node test/externalSource.test.js */
var assert = require("assert");
var m = require("../ExternalSource.js");

assert.strictEqual(m.buildDownloadUrl("https://x/a?e=1"), "https://x/a?e=1&download=1");
assert.strictEqual(m.buildDownloadUrl("https://x/a"), "https://x/a?download=1");
assert.strictEqual(m.buildDownloadUrl("https://x/a?download=1"), "https://x/a?download=1");
assert.strictEqual(m.looksLikeXlsx([0x50, 0x4b, 3, 4, 0]), true);
assert.strictEqual(m.looksLikeXlsx([-80, 0x4b, 3, 4, 0]), false);
assert.strictEqual(m.looksLikeXlsx([0x3c, 0x68, 0x74, 0x6d, 0x6c]), false); // "<html"
assert.strictEqual(m.looksLikeXlsx(null), false);
console.log("externalSource: all passed");
