/**
 * Rebuilds the local "AC CHECKS" cache sheet from the external ops WP-tracking
 * spreadsheet (CONFIG.EXTERNAL_WP_SOURCE), replacing the old manual/unknown copy
 * process. Combines Start Date + Start Time (and End Date + End Time) into real
 * datetimes so downstream shift classification (see ShiftUtils.js) has an actual
 * time-of-day to work with, instead of the date-only values the old process left behind.
 */

// Column layout of the external "RAW" tab, resolved by header name (not fixed
// index) so a reordered/inserted column upstream doesn't silently corrupt data.
var EXTERNAL_WP_HEADERS = {
  AC_TYPE: "AC Type",
  AC_REG: "AC Reg.",
  WP_NO: "WP No.",
  WP_DESC: "WP Desc.",
  START_DATE: "Start Date",
  START_TIME: "Start Time",
  END_DATE: "End Date",
  END_TIME: "End Time",
  STATION: "Station",
  AC_OPR: "AC Opr.",
  TAT: "TAT"
};

/**
 * @param {Array<string>} headerRow
 * @param {string} headerName
 * @returns {number} column index, or -1 if not found
 */
function findExternalWPColumn(headerRow, headerName) {
  for (var i = 0; i < headerRow.length; i++) {
    if ((headerRow[i] + "").trim() === headerName) return i;
  }
  return -1;
}

/**
 * Combines a date-only cell and a time-only cell into one "naive UTC" Date, i.e. a
 * Date whose UTC-field getters read as the digits actually typed into the two cells
 * (per the external sheet's own timezone) -- matching the convention classifyWP()
 * in ShiftUtils.js already expects (it applies the +7 Bangkok/Hanoi shift on top).
 * @param {Date} dateVal
 * @param {Date} timeVal
 * @param {string} sourceTimeZone - the external spreadsheet's own timezone, so the
 *   extracted digits reflect what's actually displayed in the cell.
 * @returns {Date|null}
 */
function combineExternalDateAndTime(dateVal, timeVal, sourceTimeZone) {
  if (!(dateVal instanceof Date) || isNaN(dateVal.getTime())) return null;

  var datePart = Utilities.formatDate(dateVal, sourceTimeZone, "yyyy-MM-dd");
  var timePart = (timeVal instanceof Date && !isNaN(timeVal.getTime()))
    ? Utilities.formatDate(timeVal, sourceTimeZone, "HH:mm:ss")
    : "00:00:00";

  return new Date(datePart + "T" + timePart + "Z");
}

/**
 * Fetches the current WP list from the external ops sheet, rebuilds the local
 * "AC CHECKS" sheet from it, and returns the same shape doUpdateACSchedules()
 * already expects from sh_schedule.getDataRange().getValues() (header row + rows
 * matching SCHEDULE_INDEX), so the rest of the pipeline is unchanged.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh_schedule - the local "AC CHECKS" sheet to rebuild.
 * @returns {Array<Array<*>>}
 */
function fetchAndCacheExternalWPRows(sh_schedule) {
  var externalSp = openSpreadsheetSafe(CONFIG.EXTERNAL_WP_SOURCE.SPREADSHEET_ID);
  var externalSheet = externalSp.getSheetByName(CONFIG.EXTERNAL_WP_SOURCE.TAB_NAME);
  if (!externalSheet) {
    throw new Error("External WP sheet tab '" + CONFIG.EXTERNAL_WP_SOURCE.TAB_NAME + "' not found.");
  }
  var sourceTimeZone = externalSp.getSpreadsheetTimeZone();

  var raw = externalSheet.getDataRange().getValues();
  if (raw.length < 2) return [buildAcChecksHeaderRow()];

  var headerRow = raw[0];
  var col = {};
  for (var key in EXTERNAL_WP_HEADERS) {
    col[key] = findExternalWPColumn(headerRow, EXTERNAL_WP_HEADERS[key]);
    if (col[key] === -1) {
      throw new Error("External WP sheet is missing expected column '" + EXTERNAL_WP_HEADERS[key] + "'.");
    }
  }

  var outRows = [buildAcChecksHeaderRow()];
  for (var i = 1; i < raw.length; i++) {
    var row = raw[i];
    var wpNo = (row[col.WP_NO] + "").trim();
    if (wpNo === "") continue;

    var from = combineExternalDateAndTime(row[col.START_DATE], row[col.START_TIME], sourceTimeZone);
    var to = combineExternalDateAndTime(row[col.END_DATE], row[col.END_TIME], sourceTimeZone);
    var tat = row[col.TAT];
    if (!from || !to || tat === "") continue;

    var stationCode = (row[col.STATION] + "").trim();
    var station = CONFIG.STATION_CODE_MAP[stationCode] || stationCode;

    outRows.push([
      row[col.AC_TYPE],   // AC_TYPE
      row[col.AC_REG],    // AC_REG
      row[col.WP_DESC],   // AC_CHECK
      from,                // FROM
      to,                  // TO
      tat,                 // TAT
      station,             // STATION
      "",                  // NOTE (not sourced from the external sheet)
      wpNo,                // PJID
      row[col.AC_OPR]      // CUS
    ]);
  }

  sh_schedule.clearContents();
  sh_schedule.getRange(1, 1, outRows.length, outRows[0].length).setValues(outRows);

  return outRows;
}

function buildAcChecksHeaderRow() {
  return ["AC type", "AC Reg", "AC check", "From", "To", "TAT", "Station", "Note", "PJID", "CUS"];
}
