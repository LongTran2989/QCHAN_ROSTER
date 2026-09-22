/**
 * Personnel assignment: editing a WP's check-label cell directly (double-click to edit,
 * type a name, Enter) is intercepted by onEdit() below, which reverts the raw text and
 * opens a prompt instead -- then applyAssignment() persists it to the sheet + snapshot +
 * change log, so it survives the next "Update A/C Schedules" redraw. Rows are looked up
 * by PJID (no bookmark), so this stays correct regardless of gaps between blocks.
 */

/**
 * Scans column 38 (PJID) over the WP-listing area and maps PJID -> current row number.
 * Only Normal/STO rows carry a PJID marker there (Phase Checks don't), so this
 * naturally scopes assignment to the rows that can carry an assignee.
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

  // The check label lives at "AC_CHECK (shift)[\nShortName]" -- rewrite only the second
  // line, so this stays correct regardless of what built the first line. Overflow (not
  // wrap) to match how ScheduleManager.js renders these labels -- wrapping here would
  // blow the row back out to full height on every reassignment.
  var labelCol = CONFIG.ROSTER.LEFT_COL - 1 + targetRow.fromDayCol;
  var labelCell = currentSheet.getRange(row, labelCol);
  var firstLine = (labelCell.getValue() + "").split("\n")[0];
  var shortName = newPerson ? toInitialsWithFirstName(newPerson) : "";
  labelCell.setValue(shortName ? firstLine + "\n" + shortName : firstLine).setWrap(false);

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
 * Simple onEdit trigger (auto-recognized by Apps Script, no manual trigger setup
 * needed). Editing a WP's check-label cell directly -- double-click, type, Enter --
 * reverts that raw text (the label is derived, not meant to be hand-typed) and opens
 * a prompt to set the assignee instead, going through applyAssignment() above so it's
 * persisted the same way a menu-driven assignment always was.
 */
function onEdit(e) {
  var range = e.range;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return; // ignore paste/fill

  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  if (web_getSystemSheetNames_().indexOf(sheetName) !== -1 ||
      sheetName.indexOf(CONFIG.SNAPSHOT_SHEET_PREFIX) === 0) {
    return;
  }

  var row = range.getRow();
  var col = range.getColumn();
  if (row < CONFIG.ROSTER.LOWER_ROW + 3 || col < CONFIG.ROSTER.LEFT_COL || col > CONFIG.ROSTER.RIGHT_COL) return;

  var pjid = (sheet.getRange(row, 38).getValue() + "").trim();
  if (pjid === "") return; // not an assignable WP row (e.g. a Phase Check, or blank space)

  // Restore the pre-edit text -- applyAssignment() reads the label's first line back
  // out of the cell, so it needs to see the real "AC_CHECK (shift)" text, not whatever
  // was just typed.
  range.setValue(e.oldValue !== undefined ? e.oldValue : range.getValue());

  var acCheck = "";
  try {
    var snapshotRows = readSnapshot(ensureSnapshotSheet(sheet.getParent(), sheetName));
    for (var i = 0; i < snapshotRows.length; i++) {
      if (snapshotRows[i].pjid === pjid) { acCheck = snapshotRows[i].acCheck; break; }
    }
  } catch (err) {
    // Non-fatal -- the prompt just won't show the check name.
  }

  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "Assign Personnel",
    "PJID " + pjid + (acCheck ? " (" + acCheck + ")" : "") + "\nEnter personnel name (leave blank to unassign):",
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  try {
    applyAssignment(pjid, result.getResponseText().trim());
  } catch (err) {
    ui.alert("Couldn't save assignment: " + err.message);
  }
}
