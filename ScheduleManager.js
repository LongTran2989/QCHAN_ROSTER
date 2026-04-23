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
 * Assign and backup current data to the Assignment config sheet.
 * @returns {Array<Array<string>>}
 */
function assignCRSC() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = spreadsheet.getActiveSheet();

  if (currentSheet.getRange("AJ1").getValue() === "") return [];

  var startRow = currentSheet.getRange("AJ1").getValue();
  var endRow = currentSheet.getRange("AK1").getValue() + startRow;

  var name = currentSheet.getRange(startRow, 2, endRow - startRow, 1).getValues();
  var id = currentSheet.getRange(startRow, 38, endRow - startRow, 1).getValues();
  
  var rawID = [];
  for (var i = 0; i < id.length; i++){
    rawID.push([id[i][0], name[i][0]]);
  }

  var assignCSheet = spreadsheet.getSheetByName(CONFIG.SHEET_IDS.ASSIGN_C);
  if (assignCSheet && rawID.length > 0) {
    assignCSheet.clearContents();
    assignCSheet.getRange(1, 1, rawID.length, 2).setValues(rawID);
  }

  return rawID;
}

/**
 * Fetches external schedule, filters arrays, prepares grid structures and batches them for visual layout rendering.
 */
function updateACSchedules() {
  try {
    var sp = SpreadsheetApp.getActive();
    
    var sh_schedule = sp.getSheetById(CONFIG.SHEET_IDS.SCHEDULE);
    if (!sh_schedule) throw new Error("Schedule source sheet not found");
    
    var value_schedule = sh_schedule.getDataRange().getValues();
    var sh_accheck = sp.getSheetByName(CONFIG.SHEET_IDS.AC_CHECKS_NAME);
    
    sh_accheck.getDataRange().clearContent();
    sh_accheck.getRange(1, 1, value_schedule.length, value_schedule[0].length).setValues(value_schedule);

    assignCRSC();
    var prevAssignSheet = sp.getSheetByName(CONFIG.SHEET_IDS.ASSIGN_C);
    var previousAssignC = prevAssignSheet ? prevAssignSheet.getDataRange().getValues() : [];

    var currentSheet = sp.getActiveSheet();
    if (currentSheet.getName() === CONFIG.SHEET_IDS.AC_CHECKS_NAME) {
      SpreadsheetApp.getUi().alert("1.QUAY LẠI TAB ROSTER CẦN UPDATE / 2.CHỌN QC HAN/UPDATE AC SCHEDULES");
      return;
    }
    
    if (typeof sortHAN === 'function') {
      sortHAN();
      currentSheet.activate();
    }

    var currentMonth_FirstDay = currentSheet.getRange("B1").getValue();
    currentMonth_FirstDay.setHours(0, 0, 0, 0);
    var currentMonth_LastDay = new Date(currentMonth_FirstDay.getFullYear(), currentMonth_FirstDay.getMonth() + 1, 0);

    var initialData = sh_accheck.getDataRange().getValues();
    var rawData = [];
    
    for (var i = 1; i < initialData.length; i++) {
      if (initialData[i][SCHEDULE_INDEX.FROM] === "" || initialData[i][SCHEDULE_INDEX.TO] === "" || initialData[i][SCHEDULE_INDEX.TAT] === "") continue;
      if (initialData[i][SCHEDULE_INDEX.STATION] === "HAN" || initialData[i][SCHEDULE_INDEX.STATION] === "L-HAN") {
        rawData.push(initialData[i]);
      }
    }

    var filteredData = [];
    var filteredData_EA = [];
    
    for (var i = 0; i < rawData.length; i++) {
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

    // USER FEEDBACK: Assure clean slate format over drawing area
    var areaToClear = currentSheet.getRange(CONFIG.ROSTER.LOWER_ROW + 3, CONFIG.ROSTER.LEFT_COL - 1, 200, 33 + 7);
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
        
        renderPayloads.push({range: [paintRow, paintCol], val: filteredData_EA[i][SCHEDULE_INDEX.AC_CHECK], bg: bgColor, color: fontCol, bold: isChk, note: filteredData_EA[i][SCHEDULE_INDEX.NOTE]});
        
        var nsLabel = filteredData_EA[i][SCHEDULE_INDEX.TAT] == "0.5" ? `${filteredData_EA[i][SCHEDULE_INDEX.AC_REG]} NS` : filteredData_EA[i][SCHEDULE_INDEX.AC_REG];
        renderPayloads.push({range: [paintRow + 1, paintCol], val: nsLabel, bg: null, color: null, bold: false});
    }

    eaCheckBlockLength += 5;
    currentSheet.getRange(eaStartRow, CONFIG.ROSTER.LEFT_COL, Math.max(1, eaCheckBlockLength - 2), 33).setWrap(true).setVerticalAlignment("top");

    // --- NORMAL CHECKS LOGIC ---
    var normalCheckStartRow = CONFIG.ROSTER.LOWER_ROW + 3 + eaCheckBlockLength;
    currentSheet.getRange("AJ1").setValue(normalCheckStartRow);
    currentSheet.getRange("AK1").setValue(filteredData.length);

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
    drawChecksBlock(filteredDataNormal, normalCheckStartRow, previousAssignC, renderPayloads);
    
    var listEndRow = normalCheckStartRow + filteredDataNormal.length + 2;
    renderPayloads.push({range: [listEndRow, 2], val: "END OF LIST", bg: null, color: null, bold: false});
    
    // --- STO CHECKS LOGIC ---
    var stoStartRow = listEndRow + 1;
    drawChecksBlock(filteredDataSTO, stoStartRow, previousAssignC, renderPayloads);

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
    
  } catch(err) {
    console.error("Error in updateACSchedules: " + err.message);
    SpreadsheetApp.getUi().alert("Error during schedule fetch: " + err.message);
  }
}

/**
 * Helper to build Grid payloads for block definitions
 */
function drawChecksBlock(dataBlock, startRow, previousAssignC, renderQueue) {
  for (var i = 0; i < dataBlock.length; i++) {
    var acReg = dataBlock[i][SCHEDULE_INDEX.AC_REG];
    
    for (var z = 0; z < previousAssignC.length; z++) {
      if (dataBlock[i][SCHEDULE_INDEX.PJID] == previousAssignC[z][0]) {
        acReg = previousAssignC[z][1]; break;
      }
    }
    renderQueue.push({range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 1], val: acReg});

    var isChk = (dataBlock[i][SCHEDULE_INDEX.PJID] + "").indexOf("CHK") !== -1;
    var fCol = isChk ? "red" : "black";
    
    renderQueue.push({
      range: [startRow + i, CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.FROM]],
      val: dataBlock[i][SCHEDULE_INDEX.AC_CHECK],
      note: dataBlock[i][SCHEDULE_INDEX.NOTE],
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

    for (var j = CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.FROM]; j <= CONFIG.ROSTER.LEFT_COL - 1 + dataBlock[i][SCHEDULE_INDEX.TO]; j++) {
      renderQueue.push({range: [startRow + i, j], bg: barColor});
    }

    renderQueue.push({range: [startRow + i, 38], val: dataBlock[i][SCHEDULE_INDEX.PJID]});
  }
}
