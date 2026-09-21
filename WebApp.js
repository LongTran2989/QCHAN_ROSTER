/**
 * Mobile-friendly Web App front end for the QC HAN menu actions.
 * Google Sheets' mobile apps don't render custom menus, so this gives phone
 * users a browser page that drives the same underlying functions.
 */

/**
 * Serves the Web App page. Deploy via Deploy > Test deployments (always latest
 * code) or Deploy > New deployment > Web app (a stable, versioned URL).
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('WebApp')
    .evaluate()
    .setTitle('QC HAN Roster')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Sheet names that are internal/system tabs and shouldn't be offered as a target roster tab.
 */
function web_getSystemSheetNames_() {
  return [
    CONFIG.SHEET_IDS.ROSTER_TEMPLATE,
    CONFIG.SHEET_IDS.EMAILS_SHEET,
    "CC_TEMP",
    "HUONG DAN",
    "Personel info",
    "DATE"
  ];
}

/**
 * Lists selectable roster tabs plus whichever one is currently active, for the sheet picker.
 * @returns {{sheets: string[], active: string}}
 */
function web_listSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var systemNames = web_getSystemSheetNames_();

  var sheets = ss.getSheets()
    .filter(function (sh) { return !sh.isSheetHidden() && systemNames.indexOf(sh.getName()) === -1; })
    .map(function (sh) { return sh.getName(); });

  return { sheets: sheets, active: ss.getActiveSheet().getName() };
}

/**
 * Points getActiveSheet()/getActiveSpreadsheet().getActiveSheet() at the sheet the user picked
 * in the Web App before running a menu action against it.
 * @param {string} sheetName
 */
function web_setActiveSheet_(sheetName) {
  if (!sheetName) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Sheet '" + sheetName + "' not found.");
  ss.setActiveSheet(sh);
}

/**
 * Web App version of "Create new Roster". Takes month/year directly instead of UI prompts.
 * @param {number|string} month
 * @param {number|string} year
 * @returns {string} status message
 */
function web_createNewRoster(month, year) {
  month = parseInt(month, 10);
  year = parseInt(year, 10);

  if (isNaN(month) || isNaN(year) || month <= 0 || month > 12 || year < 2022 || year > 2050) {
    throw new Error("Invalid month or year!");
  }

  fillDateToSheet(year, month);
  setBackgroundColor();
  return "Created roster " + textMonth(month) + year;
}

/**
 * Web App version of "Update A/C Schedules".
 * @param {string} sheetName
 * @returns {string} status message
 */
function web_updateACSchedules(sheetName) {
  web_setActiveSheet_(sheetName);
  doUpdateACSchedules();
  return "A/C schedules updated on '" + SpreadsheetApp.getActiveSheet().getName() + "'.";
}

/**
 * Web App version of "Update Public Roster".
 * @param {string} sheetName
 * @returns {string} URL of the public roster
 */
function web_updatePublicRoster(sheetName) {
  web_setActiveSheet_(sheetName);
  doUpdatePublicRoster();
  return "https://docs.google.com/spreadsheets/d/" + CONFIG.SHEET_IDS.PUBLIC_ROSTER;
}

/**
 * Web App version of "Xuất bảng chấm công".
 * @param {string} sheetName
 * @returns {string} status message
 */
function web_ccExport(sheetName) {
  web_setActiveSheet_(sheetName);
  doCcExport();
  return "Timesheet exported from '" + SpreadsheetApp.getActiveSheet().getName() + "'.";
}

/**
 * Web App version of "Publish as new revision".
 * @param {string} sheetName
 * @returns {{newCleanName: string, emailed: boolean}}
 */
function web_publishAsNewRevision(sheetName) {
  web_setActiveSheet_(sheetName);
  return doPublishAsNewRevision();
}
