/**
 * Backend for the "Assign Personnel" sidebar: looks up WP rows by PJID (no bookmark,
 * so it stays correct regardless of gaps between the Normal/STO blocks), and applies
 * assignments directly to the sheet + snapshot + change log.
 */

/**
 * Scans column 38 (PJID) over the WP-listing area and maps PJID -> current row number.
 * Only Normal/STO rows carry a PJID marker there (Phase Checks don't), so this
 * naturally scopes assignment to the rows that have an Assigned Person column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object<string, number>}
 */
function findWPRowsByPJID(sheet) {
  var startRow = CONFIG.ROSTER.LOWER_ROW + 3;
  var numRows = 250;
  var values = sheet.getRange(startRow, 38, numRows, 1).getValues();

  var map = {};
  for (var i = 0; i < values.length; i++) {
    var pjid = (values[i][0] + "").trim();
    if (pjid !== "") {
      map[pjid] = startRow + i;
    }
  }
  return map;
}

/**
 * @returns {string[]} Non-empty names from the Personel info sheet, column B, row 2 down.
 */
function getPersonnelNames() {
  var sp = SpreadsheetApp.getActive();
  var sheet = sp.getSheetByName(CONFIG.PERSONNEL_SHEET.NAME);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.PERSONNEL_SHEET.START_ROW) return [];

  var count = lastRow - CONFIG.PERSONNEL_SHEET.START_ROW + 1;
  var values = sheet.getRange(CONFIG.PERSONNEL_SHEET.START_ROW, CONFIG.PERSONNEL_SHEET.NAME_COL, count, 1).getValues();

  var names = [];
  for (var i = 0; i < values.length; i++) {
    var name = (values[i][0] + "").trim();
    if (name !== "") names.push(name);
  }
  return names;
}

/**
 * Data for the sidebar: this sheet's assignable WPs (from the snapshot, cross-referenced
 * against which PJIDs currently have a row) plus the personnel dropdown list.
 */
function getAssignmentSidebarData() {
  var sp = SpreadsheetApp.getActive();
  var currentSheet = sp.getActiveSheet();

  var pjidToRow = findWPRowsByPJID(currentSheet);
  var snapshotSheet = ensureSnapshotSheet(sp, currentSheet.getName());
  var snapshotByPjid = {};
  readSnapshot(snapshotSheet).forEach(function (r) { snapshotByPjid[r.pjid] = r; });

  var wps = Object.keys(pjidToRow).map(function (pjid) {
    var snap = snapshotByPjid[pjid];
    return {
      pjid: pjid,
      acReg: snap ? snap.acReg : "",
      acCheck: snap ? snap.acCheck : "",
      fromDisplay: snap ? formatBangkokTime(new Date(snap.from)) : "",
      sortKey: snap ? snap.from : 0,
      assignedPerson: snap ? snap.assignedPerson : ""
    };
  });

  wps.sort(function (a, b) { return a.sortKey - b.sortKey; });

  return {
    sheetName: currentSheet.getName(),
    wps: wps,
    personnel: getPersonnelNames()
  };
}

/**
 * Writes personName as the Assigned Person for the given WP: the live grid cell, the
 * snapshot (so it survives the next redraw), and a Change Log entry if it actually changed.
 * @param {string} pjid
 * @param {string} personName - "" to unassign.
 */
function applyAssignment(pjid, personName) {
  var sp = SpreadsheetApp.getActive();
  var currentSheet = sp.getActiveSheet();

  var pjidToRow = findWPRowsByPJID(currentSheet);
  var row = pjidToRow[pjid];
  if (!row) {
    throw new Error("Could not find a row for PJID " + pjid + " on '" + currentSheet.getName() + "'. Try running Update A/C Schedules again.");
  }

  var snapshotSheet = ensureSnapshotSheet(sp, currentSheet.getName());
  var snapshotRows = readSnapshot(snapshotSheet);

  var targetRow = null;
  for (var i = 0; i < snapshotRows.length; i++) {
    if (snapshotRows[i].pjid === pjid) {
      targetRow = snapshotRows[i];
      break;
    }
  }
  if (!targetRow) {
    throw new Error("No snapshot entry for PJID " + pjid + ". Try running Update A/C Schedules again.");
  }

  var oldPerson = targetRow.assignedPerson || "";
  var newPerson = personName || "";
  targetRow.assignedPerson = newPerson;
  writeSnapshot(snapshotSheet, snapshotRows);

  currentSheet.getRange(row, CONFIG.ROSTER.LEFT_COL - 2).setValue(newPerson);

  if (oldPerson !== newPerson) {
    var changeLogSheet = ensureChangeLogSheet(sp);
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var user = Session.getActiveUser().getEmail();
    appendChangeLogEntries(changeLogSheet, [
      buildAssignmentChangeLogRow(timestamp, user, currentSheet.getName(), pjid, targetRow.acReg, targetRow.acCheck, oldPerson, newPerson)
    ]);
  }

  return { pjid: pjid, assignedPerson: newPerson };
}

/**
 * Opens the Assign Personnel sidebar. Wired to the QC HAN menu in Main.js.
 */
function showAssignmentSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("AssignmentSidebar").setTitle("Assign Personnel");
  SpreadsheetApp.getUi().showSidebar(html);
}
