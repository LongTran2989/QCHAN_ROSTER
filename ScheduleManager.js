/**
 * Handles fetching, managing data parsing and drawing Aircraft schedules onto the Roster Grid.
 */

const SCHEDULE_INDEX = {
  AC_TYPE: 0,
  AC_REG: 1,
  AC_CHECK: 2,
  FROM: 3,
  TO: 4,
  TAT: 5,
  STATION: 6,
  NOTE: 7,
  PJID: 8
};

/**
 * Menu entry point: runs the schedule update and surfaces errors via the Sheets UI.
 */
function updateACSchedules() {
  try {
    doUpdateACSchedules();
  } catch(err) {
    console.error("Error in updateACSchedules: " + err.message);
    SpreadsheetApp.getUi().alert("Error during schedule fetch: " + err.message);
  }
}

/**
 * Fetches external schedule, filters arrays, prepares grid structures and batches them for visual layout rendering.
 * Contains no UI calls so it can also be driven from the Web App front end.
 */
function doUpdateACSchedules() {
    // One run at a time: overlapping runs (double click, menu + Web App) race on the
    // snapshot sheet, the roster area and the change log.
    var lock = LockService.getDocumentLock();
    if (!lock.tryLock(30000)) {
      throw new Error("Another schedule update is already running. Wait for it to finish, then try again.");
    }
    try {
      return runUpdateACSchedules();
    } finally {
      lock.releaseLock();
    }
}

function runUpdateACSchedules() {
    var sp = SpreadsheetApp.getActive();
    
    var value_schedule;
    if (CONFIG.EXTERNAL_SOURCE.ENABLED) {
      // Source moved to a public SharePoint Excel file (see ExternalSource.js)
      value_schedule = getExternalSourceValues(CONFIG.EXTERNAL_SOURCE.SHEET_NAME, sp.getSpreadsheetTimeZone());
    } else {
      var sh_schedule = sp.getSheetById(CONFIG.SHEET_IDS.SCHEDULE);
      if (!sh_schedule) throw new Error("Schedule source sheet not found");
      value_schedule = sh_schedule.getDataRange().getValues();
    }
    
    var currentSheet = sp.getActiveSheet();

    // Previous run's WP state (assignments, and a baseline to diff against) lives in a
    // hidden per-sheet snapshot rather than the old AJ1/AK1 row bookmark.
    var snapshotSheet = ensureSnapshotSheet(sp, currentSheet.getName());
    var previousSnapshotRows = readSnapshot(snapshotSheet);
    var previousByPjid = {};
    previousSnapshotRows.forEach(function (r) { previousByPjid[r.pjid] = r; });

    // Previous sortHAN() macro relied on AC CHECKS. We now sort in memory instead.

    var currentMonth_FirstDay = currentSheet.getRange("B1").getValue();
    if (!(currentMonth_FirstDay instanceof Date) || isNaN(currentMonth_FirstDay.getTime())) {
      throw new Error(
        "Cell B1 on '" + currentSheet.getName() + "' doesn't hold a valid date (got: " +
        JSON.stringify(currentMonth_FirstDay) + "). Check that B1's format wasn't changed " +
        "to plain text or cleared."
      );
    }
    currentMonth_FirstDay.setHours(0, 0, 0, 0);
    var currentMonth_LastDay = new Date(currentMonth_FirstDay.getFullYear(), currentMonth_FirstDay.getMonth() + 1, 0);

    var rawData = [];
    
    for (var i = 1; i < value_schedule.length; i++) {
      if (value_schedule[i][SCHEDULE_INDEX.FROM] === "" || value_schedule[i][SCHEDULE_INDEX.TO] === "" || value_schedule[i][SCHEDULE_INDEX.TAT] === "") continue;
      if (value_schedule[i][SCHEDULE_INDEX.STATION] === "HAN" || value_schedule[i][SCHEDULE_INDEX.STATION] === "L-HAN") {
        rawData.push(value_schedule[i]);
      }
    }
    
    // Safety: never wipe the roster for a source that yielded no HAN rows (wrong tab/columns).
    if (rawData.length === 0) {
      throw new Error("No usable HAN rows found in the schedule source (" + (value_schedule.length - 1) +
        " data rows read). The source layout probably doesn't match SCHEDULE_INDEX. Roster was not changed.");
    }

    // Sort rawData to mimic the old macro: Primary=AC_TYPE, Secondary=FROM
    rawData.sort(function(a, b) {
      if (a[0] < b[0]) return -1;
      if (a[0] > b[0]) return 1;
      
      var dateA = new Date(a[3]).getTime();
      var dateB = new Date(b[3]).getTime();
      if (dateA < dateB) return -1;
      if (dateA > dateB) return 1;
      
      return 0;
    });

    var filteredData = [];
    var filteredData_EA = [];
    // Source FROM/TO are naive UTC; classifyWP() corrects each to a real Bangkok (UTC+7)
    // instant and tells us which shift each end falls in. Keyed by row identity since the
    // row arrays survive the sort/filter below unchanged. See ShiftUtils.js.
    var shiftInfoByRow = new Map();

    for (var i = 0; i < rawData.length; i++) {
      var rawFromUTC = new Date(rawData[i][SCHEDULE_INDEX.FROM]);
      var rawToUTC = new Date(rawData[i][SCHEDULE_INDEX.TO]);

      var shiftInfo = classifyWP(rawFromUTC, rawToUTC);
      shiftInfoByRow.set(rawData[i], shiftInfo);

      // Bangkok calendar day of each end (as local midnight). Don't setHours() the shifted instant:
      // the script TZ is also Bangkok, so that would add the +7h offset twice (evening start -> next day).
      rawData[i][SCHEDULE_INDEX.FROM] = bangkokDayStart(shiftInfo.fromBangkok);
      rawData[i][SCHEDULE_INDEX.TO] = bangkokDayStart(shiftInfo.toBangkok);

      var validData = false;

      if (rawData[i][SCHEDULE_INDEX.FROM] >= currentMonth_FirstDay && rawData[i][SCHEDULE_INDEX.FROM] <= currentMonth_LastDay) {
        rawData[i][SCHEDULE_INDEX.FROM] = rawData[i][SCHEDULE_INDEX.FROM].getDate();
        rawData[i][SCHEDULE_INDEX.TO] = (rawData[i][SCHEDULE_INDEX.TO] > currentMonth_LastDay) ? currentMonth_LastDay.getDate() : rawData[i][SCHEDULE_INDEX.TO].getDate();
        validData = true;
      } else if (rawData[i][SCHEDULE_INDEX.TO] >= currentMonth_FirstDay && rawData[i][SCHEDULE_INDEX.TO] <= currentMonth_LastDay) {
        rawData[i][SCHEDULE_INDEX.FROM] = currentMonth_FirstDay.getDate();
        rawData[i][SCHEDULE_INDEX.TO] = rawData[i][SCHEDULE_INDEX.TO].getDate();
        validData = true;
      } else if (rawData[i][SCHEDULE_INDEX.FROM] <= currentMonth_FirstDay && rawData[i][SCHEDULE_INDEX.TO] >= currentMonth_LastDay) {
        rawData[i][SCHEDULE_INDEX.FROM] = currentMonth_FirstDay.getDate();
        rawData[i][SCHEDULE_INDEX.TO] = currentMonth_LastDay.getDate();
        validData = true;
      }

      if (validData) {
        if (rawData[i][SCHEDULE_INDEX.TAT] == "0.5" || rawData[i][SCHEDULE_INDEX.TAT] == "1") {
          filteredData_EA.push(rawData[i]);
        } else {
          filteredData.push(rawData[i]);
        }
      }
    }

    // Build this run's snapshot rows (Normal + STO + Phase checks) for diffing/persistence.
    // FROM/TO here are the full Bangkok-corrected instants (pre month-clamping), so a
    // shift-only change is still detected even when it doesn't move the display day.
    var newSnapshotRows = filteredData.concat(filteredData_EA).map(function (row) {
      var info = shiftInfoByRow.get(row);
      var pjid = row[SCHEDULE_INDEX.PJID] + "";
      var previous = previousByPjid[pjid];
      return {
        pjid: pjid,
        acReg: row[SCHEDULE_INDEX.AC_REG] + "",
        acCheck: row[SCHEDULE_INDEX.AC_CHECK] + "",
        from: info.fromBangkok.getTime(),
        to: info.toBangkok.getTime(),
        station: row[SCHEDULE_INDEX.STATION] + "",
        assignedPerson: previous ? previous.assignedPerson : "",
        fromDayCol: row[SCHEDULE_INDEX.FROM]
      };
    });

    var assignedPersonByPjid = {};
    newSnapshotRows.forEach(function (r) { assignedPersonByPjid[r.pjid] = r.assignedPerson; });

    var diffResult = diffWPLists(previousSnapshotRows, newSnapshotRows);
    var highlightedPjids = new Set(
      diffResult.added.map(function (r) { return r.pjid; })
        .concat(diffResult.changed.map(function (e) { return e.pjid; }))
    );

    // --- SPLIT THE LISTS ---
    // Short-TAT (0.5 / 1 day) WPs: only CHK ones are Phase checks (staff attend in person); the
    // rest render in the "OTHER CHECKS" group under NORMAL CHECKS (which also takes long non-CHK WPs).
    filteredData_EA = filteredData_EA.sort(comparator);
    var phaseChk = [];
    var otherShort = [];
    var filteredDataNormal = [];
    var filteredDataSTO = [];

    // A PJID starting with "D-A" always belongs in NORMAL CHECKS, as an A350 (sorted and
    // coloured so). Checked first, so it wins over the Phase / STO / TOW / Other rules.
    function isDAWP(row) {
      // Prefix match on the PJID (WP No.), e.g. "D-AIXQ_H-26"; "D-A" is the shared part so other
      // similar PJIDs are caught too.
      if ((row[SCHEDULE_INDEX.PJID] + "").trim().toUpperCase().indexOf("D-A") !== 0) return false;
      row[SCHEDULE_INDEX.AC_TYPE] = "A350";
      return true;
    }

    for (var i = 0; i < filteredData_EA.length; i++) {
      var eaPjid = filteredData_EA[i][SCHEDULE_INDEX.PJID] + "";
      if (isDAWP(filteredData_EA[i])) {
        filteredDataNormal.push(filteredData_EA[i]);
      } else if (eaPjid.indexOf("TOW") === -1 && eaPjid.indexOf("CHK") !== -1) {
        phaseChk.push(filteredData_EA[i]);
      } else {
        otherShort.push(filteredData_EA[i]);
      }
    }

    for (var i = 0; i < filteredData.length; i++) {
      var longPjid = filteredData[i][SCHEDULE_INDEX.PJID] + "";
      var longIsSTO = longPjid.indexOf("STO") !== -1;
      var longIsCHK = longPjid.indexOf("CHK") !== -1;
      if (isDAWP(filteredData[i])) {
        filteredDataNormal.push(filteredData[i]);
      } else if (longIsSTO && !longIsCHK) {
        filteredDataSTO.push(filteredData[i]);
      } else if (longIsCHK && (longIsSTO || longPjid.indexOf("TOW") === -1)) {
        // CHK WPs are Normal (a STO that is also a CHK included; TOW CHK stays in OTHER CHECKS).
        filteredDataNormal.push(filteredData[i]);
      } else {
        // Long-TAT non-CHK WPs (incl. TOW) join the other short checks in OTHER CHECKS.
        otherShort.push(filteredData[i]);
      }
    }

    // NORMAL: one staircase per A/C type (A321 first, then A350, ...). OTHER / STO: a single staircase.
    filteredDataNormal.sort(compareByTypeThenStair);
    otherShort.sort(compareScheduleRows);
    filteredDataSTO.sort(compareScheduleRows);

    // One row cursor decides every section's position (see ScheduleLayout.js), so blocks can't overlap.
    var phaseCols = phaseChk.map(function (row) { return CONFIG.ROSTER.LEFT_COL - 1 + row[SCHEDULE_INDEX.FROM]; });
    var phasePacking = assignPhaseSlots(phaseCols);
    var layout = computeScheduleLayout({
      phaseHeaderRow: CONFIG.ROSTER.LOWER_ROW + 3,
      phaseDepth: phasePacking.depth,
      normalCount: filteredDataNormal.length,
      otherCount: otherShort.length,
      stoCount: filteredDataSTO.length
    });

    // USER FEEDBACK: Assure clean slate format over drawing area. Sized from the layout so a tall
    // layout never leaves stale content/format behind; refuse (before touching the sheet) if the
    // layout would outgrow the rows the sidebar scans for PJIDs (findWPRowsByPJID).
    var drawStartRow = CONFIG.ROSTER.LOWER_ROW + 3;
    var rowsNeeded = layout.lastRow - drawStartRow + 1;
    if (rowsNeeded > CONFIG.LAYOUT.MAX_DRAW_ROWS) {
      throw new Error("The schedule needs " + rowsNeeded + " rows below the roster but at most " +
        CONFIG.LAYOUT.MAX_DRAW_ROWS + " are supported. Roster was not changed.");
    }
    var areaToClear = currentSheet.getRange(drawStartRow, CONFIG.ROSTER.LEFT_COL - 1, Math.max(200, rowsNeeded), 33 + 7);
    areaToClear.clearContent();
    areaToClear.clearFormat();
    areaToClear.clearNote();

    // Batch Output Definitions
    var renderPayloads = [];

    // --- PHASE CHECKS LOGIC ---
    renderPayloads.push({range: [layout.phaseHeaderRow, CONFIG.ROSTER.LEFT_COL - 1], val: "PHASE CHECKS", bg: null, color: null, bold: false});

    for (var i = 0; i < phaseChk.length; i++) {
        var phaseRow = phaseChk[i];
        var paintRow = layout.phaseStartRow + 2 * phasePacking.slots[i];
        var paintCol = phaseCols[i];

        var phaseShiftInfo = shiftInfoByRow.get(phaseRow);
        var phaseNight = !!(phaseShiftInfo && phaseShiftInfo.nightShiftRequired);
        var phaseStyle = resolvePhaseStyle({station: phaseRow[SCHEDULE_INDEX.STATION], nightShiftRequired: phaseNight}, CONFIG.COLORS);

        // Phase labels: no star (every Phase WP is a CHK) and no moon; night shows on the reg cell below.
        var phaseLabel = buildCheckLabel({acCheck: phaseRow[SCHEDULE_INDEX.AC_CHECK], isChk: false, shiftInfo: null});
        var phaseNote = (phaseRow[SCHEDULE_INDEX.NOTE] ? phaseRow[SCHEDULE_INDEX.NOTE] + "\n\n" : "") + (phaseShiftInfo ? buildShiftNote(phaseShiftInfo) : "");

        // Font is derived from the station fill so the label is always legible (never red-on-colour).
        renderPayloads.push({range: [paintRow, paintCol], val: phaseLabel, bg: phaseStyle.labelBg, color: phaseStyle.labelFont, bold: phaseStyle.bold, note: phaseNote});

        var nsLabel = buildRegLabel(phaseRow[SCHEDULE_INDEX.AC_REG], phaseRow[SCHEDULE_INDEX.TAT], phaseNight);
        // The reg cell is clipped like the label above it, so carry the full text in a note too.
        // It is unfilled: night coverage is the bold red font here (red never sits on a colored fill).
        renderPayloads.push({range: [paintRow + 1, paintCol], val: nsLabel, bg: null, color: phaseStyle.regFont, bold: phaseStyle.regBold, note: nsLabel + ""});
    }

    // Phase labels sit in a single day cell: clip (full text is in the cell note) and keep rows fixed-height.
    if (layout.phaseRowCount > 0) { // Sheets rejects zero-row ranges (no Phase WPs this run)
      currentSheet.getRange(layout.phaseStartRow, CONFIG.ROSTER.LEFT_COL, layout.phaseRowCount, 33).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("top");
      currentSheet.setRowHeights(layout.phaseStartRow, layout.phaseRowCount, CONFIG.LAYOUT.ROW_HEIGHT);
    }

    // --- NORMAL CHECKS LOGIC ---
    renderPayloads.push({range: [layout.normalHeaderRow, CONFIG.ROSTER.LEFT_COL - 1], val: "NORMAL CHECKS", bg: null, color: null, bold: false});
    var borderOps = [];
    drawChecksBlock(filteredDataNormal, layout.normalStartRow, assignedPersonByPjid, highlightedPjids, renderPayloads, shiftInfoByRow, currentSheet, borderOps);

    // --- OTHER SHORT CHECKS (non-CHK short-TAT WPs; station-colored; exported with Normal by Utils.js) ---
    if (otherShort.length > 0) {
      renderPayloads.push({range: [layout.otherHeaderRow, CONFIG.ROSTER.LEFT_COL - 1], val: "OTHER CHECKS", bg: null, color: null, bold: false});
      drawChecksBlock(otherShort, layout.otherStartRow, assignedPersonByPjid, highlightedPjids, renderPayloads, shiftInfoByRow, currentSheet, borderOps, {useStationFill: true});
    }

    // Utils.js (sheet export) parses this exact marker in column B.
    renderPayloads.push({range: [layout.endOfListRow, 2], val: "END OF LIST", bg: null, color: null, bold: false});

    // --- STO CHECKS LOGIC ---
    drawChecksBlock(filteredDataSTO, layout.stoStartRow, assignedPersonByPjid, highlightedPjids, renderPayloads, shiftInfoByRow, currentSheet, borderOps);

    // Dashed (unassigned) first, solid thick (changed) last so a changed row's box is never overdrawn
    // by a neighbouring unassigned row's shared edge.
    orderBorders(borderOps).forEach(function (op) {
      currentSheet.getRange(op.row, op.col, 1, op.width)
        .setBorder(true, true, true, true, false, false, op.border.color, SpreadsheetApp.BorderStyle[op.border.style]);
    });

    // Apply entire batched memory map
    for (var payload of renderPayloads) {
        var cell = currentSheet.getRange(payload.range[0], payload.range[1]);
        if (payload.val) cell.setValue(payload.val);
        if (payload.bg) cell.setBackground(payload.bg);
        if (payload.color) cell.setFontColor(payload.color);
        if (payload.bold) cell.setFontWeight("bold");
        if (payload.note) cell.setNote(payload.note);
    }
    
    // Record who ran this update and when
    recordUpdateMetadata(currentSheet);

    // Persist this run's state for next time's diff/highlight/assignment-restore.
    writeSnapshot(snapshotSheet, newSnapshotRows);

    var changeLogRows = buildChangeLogRows(
      diffResult,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      Session.getActiveUser().getEmail(),
      currentSheet.getName()
    );
    if (changeLogRows.length > 0) {
      var changeLogSheet = ensureChangeLogSheet(sp);
      appendChangeLogEntries(changeLogSheet, changeLogRows);
    }
}

/**
 * Helper to build Grid payloads for block definitions.
 * @param {{useStationFill?: boolean}} [options] useStationFill: color bars by HAN / L-HAN instead of A/C type.
 */
function drawChecksBlock(dataBlock, startRow, assignedPersonByPjid, highlightedPjids, renderQueue, shiftInfoByRow, currentSheet, borderOps, options) {
  for (var i = 0; i < dataBlock.length; i++) {
    var pjid = dataBlock[i][SCHEDULE_INDEX.PJID] + "";
    var acReg = dataBlock[i][SCHEDULE_INDEX.AC_REG];
    var assignedPerson = assignedPersonByPjid[pjid] || "";
    var shortName = assignedPerson ? toInitialsWithFirstName(assignedPerson) : "";

    // AC Reg column: keeps the "NS" marker for 0.5-day TAT rows (Other short checks).
    renderQueue.push({range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 1], val: buildRegLabel(acReg, dataBlock[i][SCHEDULE_INDEX.TAT])});

    var isChk = (dataBlock[i][SCHEDULE_INDEX.PJID] + "").indexOf("CHK") !== -1;

    var shiftInfo = shiftInfoByRow.get(dataBlock[i]);
    var style = resolveCheckStyle({
      acType: dataBlock[i][SCHEDULE_INDEX.AC_TYPE],
      isChk: isChk,
      nightShiftRequired: !!(shiftInfo && shiftInfo.nightShiftRequired),
      assigned: !!assignedPerson,
      changed: highlightedPjids.has(pjid),
      // Other short checks use the HAN / L-HAN station fill instead of the A/C-type color.
      fill: (options && options.useStationFill) ? stationFill(dataBlock[i][SCHEDULE_INDEX.STATION], CONFIG.COLORS) : undefined
    }, CONFIG.COLORS);
    var barColor = style.barColor;

    // Single-line label (assignee inline) so short-TAT rows keep the fixed row height.
    var checkLabel = buildCheckLabel({acReg: acReg, acCheck: dataBlock[i][SCHEDULE_INDEX.AC_CHECK], isChk: isChk, shiftInfo: shiftInfo, assigneeShort: shortName});
    var checkNote = (dataBlock[i][SCHEDULE_INDEX.NOTE] ? dataBlock[i][SCHEDULE_INDEX.NOTE] + "\n\n" : "") + (shiftInfo ? buildShiftNote(shiftInfo) : "");

    // Font color comes from the label cell's final fill (night red or bar color), never red text.
    renderQueue.push({
      range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.FROM]],
      val: checkLabel,
      note: checkNote,
      color: style.labelFont,
      bold: style.bold
    });

    var fromCol = CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.FROM];
    var toCol = CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.TO];

    for (var j = fromCol; j <= toCol; j++) {
      renderQueue.push({range: [startRow + i, j], bg: barColor});
    }

    // Tint just the start/end day cells (keeping the AC-type color across the rest of the bar)
    // when this WP genuinely needs night-shift coverage.
    if (shiftInfo && shiftInfo.nightShiftRequired) {
      renderQueue.push({range: [startRow + i, fromCol], bg: CONFIG.COLORS.NIGHT_SHIFT_FLAG});
      renderQueue.push({range: [startRow + i, toCol], bg: CONFIG.COLORS.NIGHT_SHIFT_FLAG});
    }

    // Row-level border from the AC Reg column through the end of the bar: solid thick purple when
    // new/changed this run, dashed dark gray when unassigned (changed wins; see resolveCheckStyle).
    // Queued (not applied here): adjacent rows share a grid line and the last setBorder wins it, so
    // the caller applies all dashed borders first and the changed (solid) boxes last (orderBorders).
    // Cleared by next run's clearFormat().
    if (style.border) {
      var borderStartCol = CONFIG.ROSTER.LEFT_COL - 1;
      borderOps.push({row: startRow + i, col: borderStartCol, width: toCol - borderStartCol + 1, border: style.border});
    }

    renderQueue.push({range: [startRow + i, 38], val: dataBlock[i][SCHEDULE_INDEX.PJID]});
  }

  // Labels are single-line and overflow across the (empty) bar cells to the right; rows keep a
  // fixed height so a long label on a short-TAT check never makes the row tall.
  if (dataBlock.length > 0) {
    currentSheet.getRange(startRow, CONFIG.ROSTER.LEFT_COL, dataBlock.length, CONFIG.ROSTER.RIGHT_COL - CONFIG.ROSTER.LEFT_COL + 1)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
    currentSheet.setRowHeights(startRow, dataBlock.length, CONFIG.LAYOUT.ROW_HEIGHT);
  }
}

// Legacy assignment-DB helpers, ported from the live Apps Script project.
// Not called by any current flow (updateACSchedules uses the snapshot sheet).

/**
 * Gets or creates the Assignments Database sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet 
 */
function getOrCreateAssignmentsDB(spreadsheet) {
  var dbSheet = spreadsheet.getSheetByName(CONFIG.SHEET_IDS.ASSIGNMENTS_DB);
  if (!dbSheet) {
    dbSheet = spreadsheet.insertSheet(CONFIG.SHEET_IDS.ASSIGNMENTS_DB);
    dbSheet.appendRow(["WP No.", "Assignment", "Last Updated", "Start Date", "End Date"]);
    dbSheet.getRange("A1:E1").setFontWeight("bold");
    dbSheet.setFrozenRows(1);
  }
  return dbSheet;
}

/**
 * Harvests current assignments from the active sheet, syncs them to the DB, 
 * and returns the full DB as an optimized lookup map.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} currentSheet 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {Array<Array<any>>} value_schedule The raw schedule data
 * @returns {Object} map of WP No -> Assignment
 */
function syncAssignmentsToDB(currentSheet, spreadsheet, value_schedule) {
  var dbSheet = getOrCreateAssignmentsDB(spreadsheet);
  
  var pjidDates = {};
  if (value_schedule) {
    for (var i = 1; i < value_schedule.length; i++) {
      var p = value_schedule[i][SCHEDULE_INDEX.PJID];
      if (p) {
        pjidDates[p] = {
          start: value_schedule[i][SCHEDULE_INDEX.FROM],
          end: value_schedule[i][SCHEDULE_INDEX.TO]
        };
      }
    }
  }

  // 1. Harvest from current sheet
  var harvested = [];
  if (currentSheet.getRange("AJ1").getValue() !== "") {
    var startRow = currentSheet.getRange("AJ1").getValue();
    var endRow = currentSheet.getRange("AK1").getValue() + startRow;
    var name = currentSheet.getRange(startRow, 2, endRow - startRow, 1).getValues();
    var id = currentSheet.getRange(startRow, 38, endRow - startRow, 1).getValues();
    for (var i = 0; i < id.length; i++){
      if (id[i][0] && id[i][0] !== "") {
        harvested.push({ pjid: id[i][0], assign: name[i][0] });
      }
    }
  }

  // 2. Read DB
  var dbData = dbSheet.getDataRange().getValues();
  var dbMap = {}; // Maps PJID to { row: rowNum, assign: assignment }
  for (var i = 1; i < dbData.length; i++) {
    var pjid = dbData[i][0];
    if (pjid) {
      dbMap[pjid] = { row: i + 1, assign: dbData[i][1], timestamp: dbData[i][2], start: dbData[i][3], end: dbData[i][4] };
    }
  }

  // 3. Update DB Map with harvested data
  var timestamp = new Date();
  var rowsToAppend = [];
  var dbUpdates = [];
  
  for (var i = 0; i < harvested.length; i++) {
    var pjid = harvested[i].pjid;
    var assign = harvested[i].assign;
    
    var sDate = pjidDates[pjid] ? pjidDates[pjid].start : "";
    var eDate = pjidDates[pjid] ? pjidDates[pjid].end : "";
    
    if (dbMap[pjid]) {
      var changed = false;
      var newStart = dbMap[pjid].start;
      var newEnd = dbMap[pjid].end;
      var newTs = dbMap[pjid].timestamp;

      // Update existing if changed
      if (dbMap[pjid].assign !== assign) {
        changed = true;
        newTs = timestamp;
      }
      if ((sDate && !dbMap[pjid].start) || (eDate && !dbMap[pjid].end)) {
        changed = true;
        if (sDate && !dbMap[pjid].start) newStart = sDate;
        if (eDate && !dbMap[pjid].end) newEnd = eDate;
      }
      
      if (changed) {
        dbUpdates.push({row: dbMap[pjid].row, vals: [[assign, newTs, newStart, newEnd]]});
        dbMap[pjid].assign = assign;
      }
    } else {
      // Add new
      rowsToAppend.push([pjid, assign, timestamp, sDate, eDate]);
      dbMap[pjid] = { assign: assign }; // Row doesn't matter for the return map
    }
  }

  // Execute updates
  for (var i = 0; i < dbUpdates.length; i++) {
    dbSheet.getRange(dbUpdates[i].row, 2, 1, 4).setValues(dbUpdates[i].vals);
  }

  // 4. Bulk append new rows
  if (rowsToAppend.length > 0) {
    dbSheet.getRange(dbSheet.getLastRow() + 1, 1, rowsToAppend.length, 5).setValues(rowsToAppend);
  }

  // 5. Build and return simple key-value map for rendering
  var returnMap = {};
  for (var key in dbMap) {
    returnMap[key] = dbMap[key].assign;
  }
  return returnMap;
}
