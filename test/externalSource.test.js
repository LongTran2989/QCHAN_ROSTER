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
assert.strictEqual(m.mergeCookies("", ["FedAuth=abc; path=/; HttpOnly", "rtFa=x; secure"]), "FedAuth=abc; rtFa=x");
assert.strictEqual(m.mergeCookies("FedAuth=old", "FedAuth=new; path=/"), "FedAuth=new");
assert.strictEqual(m.mergeCookies("a=1", undefined), "a=1");
assert.strictEqual(m.resolveRedirect("https://h.com/x/y?z", "/p/q.xlsx?ga=1"), "https://h.com/p/q.xlsx?ga=1");
assert.strictEqual(m.resolveRedirect("https://h.com/x", "https://o.com/a"), "https://o.com/a");
console.log("externalSource: all passed");
