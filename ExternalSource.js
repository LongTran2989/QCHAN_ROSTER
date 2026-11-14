/**
 * Reads the schedule source that moved from Google Sheets to a public, view-only
 * SharePoint/OneDrive Excel link. Downloads the .xlsx, converts it to a throw-away Google
 * Sheet (Drive advanced service), returns the values, then deletes the temp file.
 * Requires the link to be "Anyone with the link" (no sign-in).
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

/** Downloads the Excel file as a Blob. Throws a clear error if the link needs sign-in. */
function fetchExternalXlsxBlob() {
  var resp = UrlFetchApp.fetch(buildDownloadUrl(CONFIG.EXTERNAL_SOURCE.URL), {
    followRedirects: true,
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var bytes = resp.getContent();
  if (code !== 200 || !looksLikeXlsx(bytes)) {
    // Show what came back instead of a generic message: a 200 with a non-xlsx body usually
    // means SharePoint returned an HTML sign-in or viewer page.
    var type = resp.getHeaders()["Content-Type"] || resp.getHeaders()["content-type"] || "unknown";
    var snippet = resp.getContentText().replace(/\s+/g, " ").slice(0, 200);
    throw new Error("External source download failed (HTTP " + code + ", type " + type + "), " +
      "not an .xlsx file. Body starts with: " + snippet);
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
  module.exports = { buildDownloadUrl: buildDownloadUrl, looksLikeXlsx: looksLikeXlsx };
}
