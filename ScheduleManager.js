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
 * Fetches external schedule, filters arrays, prepares grid structures and batches them for visual layout rendering.
 */
function updateACSchedules() {
  try {
    var sp = SpreadsheetApp.getActive();
    
    var sh_schedule = sp.getSheetById(CONFIG.SHEET_IDS.SCHEDULE);
    if (!sh_schedule) throw new Error("Schedule source sheet not found");
    
    var value_schedule = sh_schedule.getDataRange().getValues();
    
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

      rawData[i][SCHEDULE_INDEX.FROM] = new Date(shiftInfo.fromBangkok.getTime());
      rawData[i][SCHEDULE_INDEX.TO] = new Date(shiftInfo.toBangkok.getTime());

      rawData[i][SCHEDULE_INDEX.FROM].setHours(0, 0, 0, 0);
      rawData[i][SCHEDULE_INDEX.TO].setHours(0, 0, 0, 0);

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
        assignedPerson: previous ? previous.assignedPerson : ""
      };
    });

    var assignedPersonByPjid = {};
    newSnapshotRows.forEach(function (r) { assignedPersonByPjid[r.pjid] = r.assignedPerson; });

    var diffResult = diffWPLists(previousSnapshotRows, newSnapshotRows);
    var highlightedPjids = new Set(
      diffResult.added.map(function (r) { return r.pjid; })
        .concat(diffResult.changed.map(function (e) { return e.pjid; }))
    );

    // USER FEEDBACK: Assure clean slate format over drawing area (column A onward, so the
    // Assigned Person column gets cleared too, not just the AC Reg column rightward).
    var areaToClear = currentSheet.getRange(CONFIG.ROSTER.LOWER_ROW + 3, CONFIG.ROSTER.LEFT_COL - 2, 200, 33 + 8);
    areaToClear.clearContent();
    areaToClear.clearFormat();
    
    // Batch Output Definitions
    var renderPayloads = [];

    // --- PHASE CHECKS LOGIC ---
    filteredData_EA = filteredData_EA.sort(comparator);
    renderPayloads.push({range: [CONFIG.ROSTER.LOWER_ROW + 3, CONFIG.ROSTER.LEFT_COL - 1], val: "PHASE CHECKS", bg: null, color: null, bold: false});

    var eaCheckBlockLength = 6;
    var eaStartRow = CONFIG.ROSTER.LOWER_ROW + 3 + 1;
    var paintRow = eaStartRow;

    // Build EA Checks iteratively
    for (var i = 0; i < filteredData_EA.length; i++) {
        var paintCol = CONFIG.ROSTER.LEFT_COL - 1 + filteredData_EA[i][SCHEDULE_INDEX.FROM];
        
        // Offset logic if there are collisions
        if (renderPayloads.some(p => p.range[0] === paintRow && p.range[1] === paintCol)) {
            paintRow += 2;
            eaCheckBlockLength = paintRow - CONFIG.ROSTER.LOWER_ROW + 1;
        } else {
            paintRow = CONFIG.ROSTER.LOWER_ROW + 3 + 1;
        }

        var isHank = filteredData_EA[i][SCHEDULE_INDEX.STATION] === "L-HAN";
        var bgColor = isHank ? CONFIG.COLORS.EA_LAN : CONFIG.COLORS.EA_HAN;
        
        var isChk = (filteredData_EA[i][SCHEDULE_INDEX.PJID] + "").indexOf("CHK") !== -1;
        var fontCol = isChk ? "red" : "black";

        var eaShiftInfo = shiftInfoByRow.get(filteredData_EA[i]);
        var eaLabel = filteredData_EA[i][SCHEDULE_INDEX.AC_CHECK] + (eaShiftInfo ? buildShiftLabelSuffix(eaShiftInfo) : "");
        var eaNote = (filteredData_EA[i][SCHEDULE_INDEX.NOTE] ? filteredData_EA[i][SCHEDULE_INDEX.NOTE] + "\n\n" : "") + (eaShiftInfo ? buildShiftNote(eaShiftInfo) : "");
        var eaBg = (eaShiftInfo && eaShiftInfo.nightShiftRequired) ? CONFIG.COLORS.NIGHT_SHIFT_FLAG : bgColor;

        renderPayloads.push({range: [paintRow, paintCol], val: eaLabel, bg: eaBg, color: fontCol, bold: isChk, note: eaNote});
        
        var nsLabel = filteredData_EA[i][SCHEDULE_INDEX.TAT] == "0.5" ? `${filteredData_EA[i][SCHEDULE_INDEX.AC_REG]} NS` : filteredData_EA[i][SCHEDULE_INDEX.AC_REG];
        renderPayloads.push({range: [paintRow + 1, paintCol], val: nsLabel, bg: null, color: null, bold: false});
    }

    eaCheckBlockLength += 5;
    currentSheet.getRange(eaStartRow, CONFIG.ROSTER.LEFT_COL, Math.max(1, eaCheckBlockLength - 2), 33).setWrap(true).setVerticalAlignment("top");

    // --- NORMAL CHECKS LOGIC ---
    var normalCheckStartRow = CONFIG.ROSTER.LOWER_ROW + 3 + eaCheckBlockLength;

    var filteredDataNormal = [];
    var filteredDataSTO = [];
    
    for (var i = 0; i < filteredData.length; i++) {
      if ((filteredData[i][SCHEDULE_INDEX.PJID] + "").indexOf("STO") !== -1) {
        filteredDataSTO.push(filteredData[i]);
      } else {
        filteredDataNormal.push(filteredData[i]);
      }
    }

    renderPayloads.push({range: [normalCheckStartRow - 1, CONFIG.ROSTER.LEFT_COL - 1], val: "NORMAL CHECKS", bg: null, color: null, bold: false});
    drawChecksBlock(filteredDataNormal, normalCheckStartRow, assignedPersonByPjid, highlightedPjids, renderPayloads, shiftInfoByRow, currentSheet);

    var listEndRow = normalCheckStartRow + filteredDataNormal.length + 2;
    renderPayloads.push({range: [listEndRow, 2], val: "END OF LIST", bg: null, color: null, bold: false});

    // --- STO CHECKS LOGIC ---
    var stoStartRow = listEndRow + 1;
    drawChecksBlock(filteredDataSTO, stoStartRow, assignedPersonByPjid, highlightedPjids, renderPayloads, shiftInfoByRow, currentSheet);

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

  } catch(err) {
    console.error("Error in updateACSchedules: " + err.message);
    SpreadsheetApp.getUi().alert("Error during schedule fetch: " + err.message);
  }
}

/**
 * Helper to build Grid payloads for block definitions
 */
function drawChecksBlock(dataBlock, startRow, assignedPersonByPjid, highlightedPjids, renderQueue, shiftInfoByRow, currentSheet) {
  for (var i = 0; i < dataBlock.length; i++) {
    var pjid = dataBlock[i][SCHEDULE_INDEX.PJID] + "";
    var acReg = dataBlock[i][SCHEDULE_INDEX.AC_REG];
    var assignedPerson = assignedPersonByPjid[pjid] || "";

    renderQueue.push({range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 2], val: assignedPerson});
    renderQueue.push({range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 1], val: acReg});

    var isChk = (dataBlock[i][SCHEDULE_INDEX.PJID] + "").indexOf("CHK") !== -1;
    var fCol = isChk ? "red" : "black";

    var shiftInfo = shiftInfoByRow.get(dataBlock[i]);
    var checkLabel = dataBlock[i][SCHEDULE_INDEX.AC_CHECK] + (shiftInfo ? buildShiftLabelSuffix(shiftInfo) : "");
    var checkNote = (dataBlock[i][SCHEDULE_INDEX.NOTE] ? dataBlock[i][SCHEDULE_INDEX.NOTE] + "\n\n" : "") + (shiftInfo ? buildShiftNote(shiftInfo) : "");

    renderQueue.push({
      range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.FROM]],
      val: checkLabel,
      note: checkNote,
      color: fCol,
      bold: isChk
    });

    var barColor;
    switch(dataBlock[i][SCHEDULE_INDEX.AC_TYPE]) {
      case "A320": case "A321": barColor = CONFIG.COLORS.A320; break;
      case "A350": barColor = CONFIG.COLORS.A350; break;
      case "B787": barColor = CONFIG.COLORS.B787; break;
      default: barColor = CONFIG.COLORS.DEFAULT; break;
    }

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

    // Flag rows that are new or changed this run with a border from the Assigned Person
    // column through the end of the bar. Applied directly (not batched) since it spans a
    // range rather than a single cell; cleared automatically by next run's clearFormat().
    if (highlightedPjids.has(pjid)) {
      var highlightStartCol = CONFIG.ROSTER.LEFT_COL - 2;
      currentSheet.getRange(startRow + i, highlightStartCol, 1, toCol - highlightStartCol + 1)
        .setBorder(true, true, true, true, false, false, CONFIG.COLORS.CHANGE_HIGHLIGHT, SpreadsheetApp.BorderStyle.SOLID_THICK);
    }

    renderQueue.push({range: [startRow + i, 38], val: dataBlock[i][SCHEDULE_INDEX.PJID]});
  }
}
