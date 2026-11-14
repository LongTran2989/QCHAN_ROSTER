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

/**
 * Downloads and reads the external workbook.
 * @param {string} [sheetName] Tab to read; empty = first tab.
 * @param {string} [timeZone] Time zone used to read date cells (pass the roster spreadsheet's
 *   zone so dates come out exactly as the old in-spreadsheet source did).
 * @returns {Array<Array>} 2D values of the tab
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
    return sh.getDataRange().getValues();
  } finally {
    Drive.Files.remove(temp.id);
  }
}

if (typeof module !== "undefined") {
  module.exports = { buildDownloadUrl: buildDownloadUrl, looksLikeXlsx: looksLikeXlsx, mergeCookies: mergeCookies, resolveRedirect: resolveRedirect };
}
