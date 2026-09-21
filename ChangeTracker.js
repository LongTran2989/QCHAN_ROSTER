/**
 * Tracks what changed in the AC CHECKS WP listing between runs of updateACSchedules,
 * and persists/reads the per-sheet snapshot + Change Log sheet used for that.
 *
 * diffWPLists() has no SpreadsheetApp dependency, so it's exercised outside Apps
 * Script (see test/changeTracker.test.js). Everything below it talks to Sheets.
 */

var DIFF_FIELDS = ["acReg", "acCheck", "from", "to", "station"];

/**
 * @typedef {Object} WPSnapshotRow
 * @property {string} pjid
 * @property {string} acReg
 * @property {string} acCheck
 * @property {number} from - epoch ms (Bangkok-corrected instant, pre month-clamping)
 * @property {number} to - epoch ms
 * @property {string} station
 * @property {string} assignedPerson
 * @property {number} fromDayCol - clamped day-of-month used as the label cell's column
 *   offset this run (CONFIG.ROSTER.LEFT_COL - 1 + fromDayCol). Rendering detail only,
 *   never compared by diffWPLists.
 */

/**
 * Compares the previous run's WP rows against this run's, keyed by PJID.
 * @param {WPSnapshotRow[]} oldRows
 * @param {WPSnapshotRow[]} newRows
 * @returns {{added: WPSnapshotRow[], removed: WPSnapshotRow[], changed: Array<{pjid: string, newRow: WPSnapshotRow, changedFields: Array<{field: string, oldValue: *, newValue: *}>}>}}
 */
function diffWPLists(oldRows, newRows) {
  var oldByPjid = {};
  for (var i = 0; i < oldRows.length; i++) {
    oldByPjid[oldRows[i].pjid] = oldRows[i];
  }

  var newByPjid = {};
  for (var j = 0; j < newRows.length; j++) {
    newByPjid[newRows[j].pjid] = newRows[j];
  }

  var added = [];
  var changed = [];

  for (var n = 0; n < newRows.length; n++) {
    var newRow = newRows[n];
    var oldRow = oldByPjid[newRow.pjid];

    if (!oldRow) {
      added.push(newRow);
      continue;
    }

    var changedFields = [];
    for (var f = 0; f < DIFF_FIELDS.length; f++) {
      var field = DIFF_FIELDS[f];
      if (oldRow[field] !== newRow[field]) {
        changedFields.push({ field: field, oldValue: oldRow[field], newValue: newRow[field] });
      }
    }

    if (changedFields.length > 0) {
      changed.push({ pjid: newRow.pjid, newRow: newRow, changedFields: changedFields });
    }
  }

  var removed = [];
  for (var o = 0; o < oldRows.length; o++) {
    if (!newByPjid[oldRows[o].pjid]) {
      removed.push(oldRows[o]);
    }
  }

  return { added: added, removed: removed, changed: changed };
}

/**
 * Flattens a diff result into Change Log row arrays, one per added/removed WP and one
 * per changed field, newest-first order is the caller's job (insert at the sheet's row 2).
 * @param {ReturnType<typeof diffWPLists>} diffResult
 * @param {string} timestamp
 * @param {string} user
 * @param {string} sheetName
 * @returns {Array<Array<string>>}
 */
function buildChangeLogRows(diffResult, timestamp, user, sheetName) {
  var rows = [];

  diffResult.added.forEach(function (row) {
    rows.push([timestamp, user, sheetName, row.pjid, row.acReg, row.acCheck, "Added", "", "", ""]);
  });

  diffResult.removed.forEach(function (row) {
    rows.push([timestamp, user, sheetName, row.pjid, row.acReg, row.acCheck, "Removed", "", "", ""]);
  });

  diffResult.changed.forEach(function (entry) {
    entry.changedFields.forEach(function (fc) {
      rows.push([
        timestamp, user, sheetName, entry.pjid, entry.newRow.acReg, entry.newRow.acCheck,
        "Changed", fc.field, fc.oldValue, fc.newValue
      ]);
    });
  });

  return rows;
}

/**
 * A single "Assigned" change-log row for the sidebar's applyAssignment() flow.
 */
function buildAssignmentChangeLogRow(timestamp, user, sheetName, pjid, acReg, acCheck, oldPerson, newPerson) {
  return [timestamp, user, sheetName, pjid, acReg, acCheck, "Assigned", "assignedPerson", oldPerson || "", newPerson || ""];
}

var CHANGE_LOG_HEADER = ["Timestamp", "User", "Sheet", "PJID", "AC Reg", "AC Check", "Change Type", "Field", "Old Value", "New Value"];
var SNAPSHOT_HEADER = ["PJID", "AC Reg", "AC Check", "From", "To", "Station", "AssignedPerson", "FromDayCol"];

// --- Sheets-facing I/O below this point ---

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} rosterSheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureSnapshotSheet(spreadsheet, rosterSheetName) {
  var name = CONFIG.SNAPSHOT_SHEET_PREFIX + rosterSheetName;
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.hideSheet();
    sheet.getRange(1, 1, 1, SNAPSHOT_HEADER.length).setValues([SNAPSHOT_HEADER]);
  }
  return sheet;
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} snapshotSheet
 * @returns {WPSnapshotRow[]}
 */
function readSnapshot(snapshotSheet) {
  var lastRow = snapshotSheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = snapshotSheet.getRange(2, 1, lastRow - 1, SNAPSHOT_HEADER.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if (v[0] === "") continue;
    rows.push({
      pjid: v[0] + "",
      acReg: v[1] + "",
      acCheck: v[2] + "",
      from: Number(v[3]),
      to: Number(v[4]),
      station: v[5] + "",
      assignedPerson: v[6] + "",
      fromDayCol: Number(v[7])
    });
  }
  return rows;
}

/**
 * Overwrites the snapshot sheet with the given rows.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} snapshotSheet
 * @param {WPSnapshotRow[]} rows
 */
function writeSnapshot(snapshotSheet, rows) {
  var lastRow = snapshotSheet.getLastRow();
  if (lastRow > 1) {
    snapshotSheet.getRange(2, 1, lastRow - 1, SNAPSHOT_HEADER.length).clearContent();
  }
  if (rows.length === 0) return;

  var values = rows.map(function (r) {
    return [r.pjid, r.acReg, r.acCheck, r.from, r.to, r.station, r.assignedPerson || "", r.fromDayCol];
  });
  snapshotSheet.getRange(2, 1, values.length, SNAPSHOT_HEADER.length).setValues(values);
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureChangeLogSheet(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_IDS.CHANGE_LOG_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_IDS.CHANGE_LOG_SHEET);
    sheet.getRange(1, 1, 1, CHANGE_LOG_HEADER.length).setValues([CHANGE_LOG_HEADER]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Inserts rows right after the header so the Change Log always reads newest-first.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} changeLogSheet
 * @param {Array<Array<string>>} rows
 */
function appendChangeLogEntries(changeLogSheet, rows) {
  if (rows.length === 0) return;
  changeLogSheet.insertRowsBefore(2, rows.length);
  changeLogSheet.getRange(2, 1, rows.length, CHANGE_LOG_HEADER.length).setValues(rows);
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    DIFF_FIELDS: DIFF_FIELDS,
    diffWPLists: diffWPLists,
    buildChangeLogRows: buildChangeLogRows,
    buildAssignmentChangeLogRow: buildAssignmentChangeLogRow,
    CHANGE_LOG_HEADER: CHANGE_LOG_HEADER,
    SNAPSHOT_HEADER: SNAPSHOT_HEADER
  };
}
