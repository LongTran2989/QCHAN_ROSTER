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
// convertRawToSchedule: wall-clock date/time cells -> naive-UTC FROM/TO, by header name
var pad = function (n) { return (n < 10 ? "0" : "") + n; };
var fmt = function (d, p) { // fake sheet-zone formatter: read the Date's UTC fields
  return p === "HH:mm" ? pad(d.getUTCHours()) + ":" + pad(d.getUTCMinutes())
    : d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
};
var hdr = ["AC Type", "AC Model", "AC Reg.", "WP No.", "WP Desc.", "WP Status Name", "Start Date", "Start Time", "End Date", "End Time", "Station", "Dept.", "Maint. Center", "AC Opr.", "TAT", "Name"];
var row = ["32S", "A321", "A356", "A356-051226-CHK-EQ03", "EQ03", "In Preparation", new Date(Date.UTC(2026, 11, 1)), new Date(Date.UTC(1899, 11, 30, 11, 0)), new Date(Date.UTC(2026, 11, 2)), new Date(Date.UTC(1899, 11, 30, 23, 30)), "HAN", "NTHN", "H", "HVN", 1, "EA-BASE-HAN"];
var conv = m.convertRawToSchedule([hdr, row, ["32S", "A320", "A1", "", "x", "", "", "", "", "", "HAN", "", "", "", 1, ""]], fmt);
assert.strictEqual(conv.length, 2); // blank WP No. row dropped
assert.strictEqual(conv[1][0], "A321");          // AC_TYPE = AC Model
assert.strictEqual(conv[1][1], "A356");          // AC_REG
assert.strictEqual(conv[1][2], "EQ03");          // AC_CHECK = WP Desc.
assert.strictEqual(conv[1][3].toISOString(), "2026-12-01T11:00:00.000Z");
assert.strictEqual(conv[1][4].toISOString(), "2026-12-02T23:30:00.000Z");
assert.strictEqual(conv[1][5], 1);               // TAT
assert.strictEqual(conv[1][6], "HAN");           // STATION
assert.strictEqual(conv[1][8], "A356-051226-CHK-EQ03"); // PJID
assert.throws(function () { m.convertRawToSchedule([["AC Type"]], fmt); }, /missing column/);
console.log("externalSource: all passed");
