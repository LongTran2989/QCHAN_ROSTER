/**
 * Common utilities for array/date manipulations and error-proof accesses.
 */

/**
 * Move an item in an array from one index to another inline.
 * @param {Array} arr - The array to modify.
 * @param {number} fromIndex - Original index.
 * @param {number} toIndex - Target index.
 */
function arraymove(arr, fromIndex, toIndex) {
  var element = arr[fromIndex];
  arr.splice(fromIndex, 1);
  arr.splice(toIndex, 0, element);
}

/**
 * Comparator function for sorting arrays primarily by the FROM date (index 3).
 * @param {Array} a - First element
 * @param {Array} b - Second element
 * @returns {number} Sorting order
 */
function comparator(a, b) {
  if (a[3] < b[3]) return -1;
  if (a[3] > b[3]) return 1;
  return 0;
}

/**
 * Calculate the number of days in a given month and year.
 * @param {number} year - The full year.
 * @param {number} month - The month (1-12).
 * @returns {number} The number of days.
 */
function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

/**
 * Convert Date integer day of the week to string abbreviation.
 * @param {number} day - Day index (0-6) where 0 is Sunday.
 * @returns {string} The text representation of the day.
 */
function textDay(day) {
  const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return days[day];
}

/**
 * Convert month index to capitalized abbreviation.
 * @param {number} month - Month index (1-12).
 * @returns {string} The text abbreviation of the month.
 */
function textMonth(month) {
  const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
  return months[month - 1];
}

/**
 * Creates an empty 2D array representing a grid, filled with a default value.
 * @param {number} rows 
 * @param {number} cols 
 * @param {any} defaultVal 
 * @returns {Array<Array<any>>}
 */
function createGrid(rows, cols, defaultVal) {
  var grid = [];
  for (var i = 0; i < rows; i++) {
    var row = [];
    for (var j = 0; j < cols; j++) {
      row.push(defaultVal);
    }
    grid.push(row);
  }
  return grid;
}

/**
 * Safe wrapper for opening spreadsheets to catch permission/not-found errors.
 * @param {string} id - The Spreadsheet ID.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function openSpreadsheetSafe(id) {
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    throw new Error(`Failed to open Spreadsheet (ID: ${id}): ${err.message}`);
  }
}

/**
 * Retrieves the email recipients list from the configured email sheet.
 * If the sheet doesn't exist, it creates one.
 * @returns {string} Comma-separated list of email addresses.
 */
function getEmailRecipients() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = CONFIG.SHEET_IDS.EMAILS_SHEET;
  var emailSheet = ss.getSheetByName(sheetName);

  if (!emailSheet) {
    emailSheet = ss.insertSheet(sheetName);
    emailSheet.hideSheet();
    emailSheet.getRange("A1").setValue("Email Recipients").setFontWeight("bold");
    return ""; // Return empty string as it's a newly created sheet
  }

  var lastRow = emailSheet.getLastRow();
  if (lastRow <= 1) return "";

  var emails = emailSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var recipientList = [];

  for (var i = 0; i < emails.length; i++) {
    var email = emails[i][0].toString().trim();
    if (email !== "") {
      recipientList.push(email);
    }
  }

  return recipientList.join(", ");
}

/**
 * Records the last update time and user to PropertiesService and to a designated cell.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to write the timestamp to.
 */
function recordUpdateMetadata(sheet) {
  try {
    var email = Session.getActiveUser().getEmail();
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var updateText = `Last updated: ${timestamp}`;

    var cell = sheet.getRange(CONFIG.ROSTER.UPDATE_INFO_CELL || "B75");
    cell.setValue(updateText);

    var props = PropertiesService.getDocumentProperties();
    props.setProperty("LAST_UPDATE_INFO", `${email}|${timestamp}`);
  } catch (err) {
    console.error("Failed to record metadata: " + err.message);
  }
}
