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
    throw new Error("External source download failed (HTTP " + code + "). " +
      "The link may no longer be public ('Anyone with the link') or has expired.");
  }
  return resp.getBlob().setName("external_source.xlsx");
}

/**
 * Downloads and reads the external workbook.
 * @param {string} [sheetName] Only this tab; omit for all tabs.
 * @returns {Object<string, Array<Array>>} map of tab name -> 2D values
 */
function getExternalSourceValues(sheetName) {
  var blob = fetchExternalXlsxBlob();
  var temp = Drive.Files.create(
    { name: "tmp_external_source", mimeType: MimeType.GOOGLE_SHEETS },
    blob
  );
  try {
    var ss = SpreadsheetApp.openById(temp.id);
    var out = {};
    ss.getSheets().forEach(function (sh) {
      if (!sheetName || sh.getName() === sheetName) {
        out[sh.getName()] = sh.getDataRange().getValues();
      }
    });
    if (sheetName && !out[sheetName]) {
      throw new Error("Tab '" + sheetName + "' not found in external source.");
    }
    return out;
  } finally {
    Drive.Files.remove(temp.id);
  }
}

/** True if the downloaded content differs from the last run (so downstream sync can be skipped). */
function externalSourceChanged(values) {
  var digest = Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(values))
  );
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty("EXTERNAL_SOURCE_HASH") === digest) return false;
  props.setProperty("EXTERNAL_SOURCE_HASH", digest);
  return true;
}

/**
 * Trigger entry point. Pulls the data and, if changed, hands it to the Calendar sync.
 * TODO: point onExternalSourceUpdated at the existing Calendar-sync function.
 */
function syncFromExternalSource() {
  var values = getExternalSourceValues(CONFIG.EXTERNAL_SOURCE.SHEET_NAME);
  if (!externalSourceChanged(values)) return;
  onExternalSourceUpdated(values);
}

/** Replace the body with a call into the Calendar sync, which now receives data instead of reading Google Sheets. */
function onExternalSourceUpdated(values) {
  console.log("External source changed; tabs: " + Object.keys(values).join(", "));
}

/** Run once to schedule the sync (every 30 minutes). */
function installExternalSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "syncFromExternalSource") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("syncFromExternalSource").timeBased().everyMinutes(30).create();
}

if (typeof module !== "undefined") {
  module.exports = { buildDownloadUrl: buildDownloadUrl, looksLikeXlsx: looksLikeXlsx };
}
