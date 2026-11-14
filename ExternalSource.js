/**
 * Reads the schedule from a SharePoint/OneDrive Excel link: downloads the .xlsx, converts
 * it to a throw-away Google Sheet (Drive advanced service), returns the values, then
 * deletes the temp file.
 *
 * Note: SharePoint's anonymous link sets a session cookie on the first redirect and rejects the
 * file request without it (UrlFetchApp's auto-redirect drops cookies, which produced the HTML
 * "Something went wrong" page), so redirects are followed manually with cookies forwarded.
 * If it ever fails, set CONFIG.EXTERNAL_SOURCE.ENABLED = false to use the SCHEDULE tab.
 */

/** Adds the direct-download flag to a SharePoint sharing link. */
function buildDownloadUrl(url) {
  if (/[?&]download=1(&|$)/.test(url)) return url;
  return url + (url.indexOf("?") === -1 ? "?" : "&") + "download=1";
}

/** True if bytes start with the ZIP magic "PK" (.xlsx is a zip) — false for a login/HTML page. */
function looksLikeXlsx(bytes) {
  return !!bytes && bytes.length > 4 && (bytes[0] & 255) === 0x50 && (bytes[1] & 255) === 0x4b;
}

/**
 * Turns Set-Cookie header value(s) into a "name=value; name2=value2" Cookie header, merged
 * into any cookies already held. SharePoint's anonymous share link issues a session cookie on
 * the first hop and refuses the redirected file request without it; UrlFetchApp's automatic
 * redirect following does not carry cookies across hops.
 */
function mergeCookies(existing, setCookie) {
  var jar = {};
  (existing || "").split(";").forEach(function (c) {
    var i = c.indexOf("=");
    if (i > 0) jar[c.slice(0, i).trim()] = c.slice(i + 1).trim();
  });
  var list = setCookie == null ? [] : (Array.isArray(setCookie) ? setCookie : [setCookie]);
  list.forEach(function (sc) {
    var first = String(sc).split(";")[0];
    var i = first.indexOf("=");
    if (i > 0) jar[first.slice(0, i).trim()] = first.slice(i + 1).trim();
  });
  return Object.keys(jar).map(function (k) { return k + "=" + jar[k]; }).join("; ");
}

/** Resolves a Location header (absolute or root-relative) against the current URL. */
function resolveRedirect(base, location) {
  if (/^https?:\/\//i.test(location)) return location;
  var origin = base.match(/^https?:\/\/[^\/]+/i)[0];
  return location.charAt(0) === "/" ? origin + location : origin + "/" + location;
}

/** Downloads the Excel file as a Blob. Throws a clear error if the link needs sign-in. */
function fetchExternalXlsxBlob() {
  var url = buildDownloadUrl(CONFIG.EXTERNAL_SOURCE.URL);
  var cookies = "";
  var resp, code, bytes;
  // Follow redirects by hand (max 6 hops), forwarding the cookies the server itself issues.
  for (var hop = 0; hop < 6; hop++) {
    var options = { followRedirects: false, muteHttpExceptions: true };
    if (cookies) options.headers = { Cookie: cookies };
    resp = UrlFetchApp.fetch(url, options);
    code = resp.getResponseCode();
    var headers = resp.getAllHeaders();
    cookies = mergeCookies(cookies, headers["Set-Cookie"] || headers["set-cookie"]);
    var loc = headers["Location"] || headers["location"];
    if (code >= 300 && code < 400 && loc) {
      url = resolveRedirect(url, Array.isArray(loc) ? loc[0] : loc);
      continue;
    }
    break;
  }
  bytes = resp.getContent();
  if (code !== 200 || !looksLikeXlsx(bytes)) {
    // Show what came back instead of a generic message: a 200 with a non-xlsx body usually
    // means SharePoint returned an HTML sign-in or viewer page.
    var type = resp.getHeaders()["Content-Type"] || resp.getHeaders()["content-type"] || "unknown";
    var text = resp.getContentText();
    var titleMatch = text.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    var title = titleMatch ? titleMatch[1].replace(/\s+/g, " ").trim() : "(no title)";
    throw new Error("External source download failed (HTTP " + code + ", type " + type + "), " +
      "not an .xlsx file. Page title: " + title);
  }
  return resp.getBlob().setName("external_source.xlsx");
}

/** Header names of the raw AMOS export (tab "RAW") mapped to what the sync needs. */
var RAW_HEADERS = {
  AC_TYPE: "AC Model", AC_REG: "AC Reg.", AC_CHECK: "WP Desc.", PJID: "WP No.",
  STATION: "Station", TAT: "TAT",
  START_DATE: "Start Date", START_TIME: "Start Time", END_DATE: "End Date", END_TIME: "End Time"
};

/**
 * Converts the raw AMOS export into the 9-column schedule layout the sync reads
 * (see SCHEDULE_INDEX in ScheduleManager.js): [AC_TYPE, AC_REG, AC_CHECK, FROM, TO, TAT,
 * STATION, NOTE, PJID]. Columns are found by header name so reordering doesn't break it.
 * FROM/TO are naive UTC Dates: the export's date + time wall-clock (which is UTC) read as UTC.
 * @param {Array<Array>} values 2D values of the RAW tab (first row = headers).
 * @param {function(Date,string):string} fmt formats a Date's wall clock in the sheet's zone
 *   ("yyyy-MM-dd" / "HH:mm"); in Apps Script: function (d, p) { return Utilities.formatDate(d, tz, p); }
 */
function convertRawToSchedule(values, fmt) {
  if (!values || !values.length) throw new Error("External source tab is empty.");
  var header = values[0].map(function (h) { return String(h).trim(); });
  var col = {};
  Object.keys(RAW_HEADERS).forEach(function (k) {
    var idx = header.indexOf(RAW_HEADERS[k]);
    if (idx === -1) {
      throw new Error("External source is missing column '" + RAW_HEADERS[k] + "'. Found: " + header.join(", "));
    }
    col[k] = idx;
  });

  function wallClock(dateCell, timeCell) {
    var d = dateCell instanceof Date ? fmt(dateCell, "yyyy-MM-dd") : String(dateCell).trim();
    var t = timeCell instanceof Date ? fmt(timeCell, "HH:mm") : String(timeCell).trim();
    var dm = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
    var tm = t.match(/^(\d{1,2}):(\d{2})/);
    if (!dm || !tm) return "";
    return new Date(Date.UTC(+dm[1], +dm[2] - 1, +dm[3], +tm[1], +tm[2]));
  }

  var out = [["AC_TYPE", "AC_REG", "AC_CHECK", "FROM", "TO", "TAT", "STATION", "NOTE", "PJID"]];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var pjid = r[col.PJID];
    if (pjid === "" || pjid == null) continue;
    out.push([
      r[col.AC_TYPE], r[col.AC_REG], r[col.AC_CHECK],
      wallClock(r[col.START_DATE], r[col.START_TIME]),
      wallClock(r[col.END_DATE], r[col.END_TIME]),
      r[col.TAT], r[col.STATION], "", pjid
    ]);
  }
  return out;
}

/**
 * Downloads and reads the external workbook.
 * @param {string} [sheetName] Tab to read; empty = first tab.
 * @param {string} [timeZone] Time zone used to read date cells (pass the roster spreadsheet's
 *   zone so dates come out exactly as the old in-spreadsheet source did).
 * @returns {Array<Array>} the tab converted to the 9-column schedule layout
 */
function getExternalSourceValues(sheetName, timeZone) {
  var blob = fetchExternalXlsxBlob();
  var temp = Drive.Files.create(
    { name: "tmp_external_source", mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );
  try {
    var ss = SpreadsheetApp.openById(temp.id);
    if (timeZone) ss.setSpreadsheetTimeZone(timeZone);
    var sh = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    if (!sh) throw new Error("Tab '" + sheetName + "' not found in external source.");
    var tz = ss.getSpreadsheetTimeZone();
    return convertRawToSchedule(sh.getDataRange().getValues(), function (d, p) {
      return Utilities.formatDate(d, tz, p);
    });
  } finally {
    Drive.Files.remove(temp.id);
  }
}

if (typeof module !== "undefined") {
  module.exports = { buildDownloadUrl: buildDownloadUrl, looksLikeXlsx: looksLikeXlsx, convertRawToSchedule: convertRawToSchedule, mergeCookies: mergeCookies, resolveRedirect: resolveRedirect };
}
