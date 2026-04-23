/**
 * Handles exporting schedule data to timesheet formatting (Bảng Chấm Công).
 */

/**
 * Extracts and processes raw roster data into Timesheet sheets.
 */
function ccExport() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var currentSheet = spreadsheet.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
    
    // Manage exception
    const SHEET_NAME = currentSheet.getName();
    if (SHEET_NAME === "HUONG DAN" || SHEET_NAME === "Personel info") {
      ui.alert("Chọn tab roster của tháng cần xuất file chấm công!");
      return;
    }

    var baseDate = currentSheet.getRange("B1").getValue();
    var month = baseDate.getMonth();  // index 0
    var year = baseDate.getFullYear();
    var maxDaysOfMonth = daysInMonth(year, month + 1);
    
    var d = new Date(year, month);
    var arrDate = [];
    var arrDay = [];
    
    for (var i = 0; i < maxDaysOfMonth; i++) {
      arrDate.push(d.getDate());
      arrDay.push(d.getDay());
      d.setDate(d.getDate() + 1);
    }
    
    var numRows = CONFIG.ROSTER.LOWER_ROW - CONFIG.ROSTER.UPPER_ROW - 1;
    var numCols = CONFIG.ROSTER.RIGHT_COL - CONFIG.ROSTER.LEFT_COL + 5;
    
    // Extract actual worked schedule
    var rosterRawData = currentSheet.getRange(CONFIG.ROSTER.UPPER_ROW + 2, CONFIG.ROSTER.LEFT_COL - 2, numRows, numCols).getValues();
    var workDays = [];
    
    for (var i = 0; i < rosterRawData.length; i++) {
      // Filter out empty ID lines
      if (rosterRawData[i][33] !== "") {
        workDays.push(rosterRawData[i]);
      }
    }
    
    // Transpose and restructure columns for the export layout
    for (var i = 0; i < workDays.length; i++) {
      arraymove(workDays[i], 33, 1);
      arraymove(workDays[i], 34, 3);
    }
    
    createFormCC(spreadsheet, workDays, arrDate, arrDay, year, month);
    
  } catch (err) {
    console.error("Error in ccExport: " + err.message);
    SpreadsheetApp.getUi().alert("Error running ccExport: " + err.message);
  }
}

/**
 * Clones the Template and applies the processed work entries using batched network arrays.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet 
 * @param {Array<Array<any>>} workDays 
 * @param {Array<number>} arrDate 
 * @param {Array<number>} arrDay 
 * @param {number} year 
 * @param {number} month 
 */
function createFormCC(spreadsheet, workDays, arrDate, arrDay, year, month) {
  var cc_tempSheet = spreadsheet.getSheetByName("CC_TEMP");
  if (!cc_tempSheet) {
    throw new Error("Missing 'CC_TEMP' sheet in this document");
  }
  
  cc_tempSheet.activate();
  var dispMonth = month + 1;
  var ccSheetToExport = spreadsheet.duplicateActiveSheet().setName(`Cham cong T${dispMonth}`);
  
  cc_tempSheet.showSheet();
  cc_tempSheet.hideSheet();
  ccSheetToExport.activate();
  
  const CC_UPPER_ROW = 7;
  const CC_LOWER_ROW = CC_UPPER_ROW + workDays.length;
  const CC_LEFT_COL = 5;
  const CC_RIGHT_COL = CC_LEFT_COL + arrDate.length;
  
  // Format Dynamic Window Size
  ccSheetToExport.insertColumns(CC_LEFT_COL, arrDate.length - 28);
  ccSheetToExport.setColumnWidths(CC_LEFT_COL, arrDate.length, 20);
  ccSheetToExport.insertRows(CC_UPPER_ROW + 1, workDays.length - 1);

  ccSheetToExport.getRange(CC_UPPER_ROW - 1, CC_LEFT_COL, 1, arrDay.length).setValues([arrDay]);
  ccSheetToExport.getRange(CC_UPPER_ROW, CC_LEFT_COL, 1, arrDate.length).setValues([arrDate]);

  var colorGrid = createGrid(CC_LOWER_ROW - 6, arrDate.length, "white");
  var weekendIndices = [];
  
  for (var i = 0; i < arrDay.length; i++) {
    if (arrDay[i] === 6 || arrDay[i] === 0) {
      weekendIndices.push(i);
    }
  }

  // Paint weekends in gray efficiently
  for (var c of weekendIndices) {
    for (var r = 0; r < CC_LOWER_ROW - 6; r++) {
       colorGrid[r][c] = "gray"; 
    }
  }
  
  ccSheetToExport.getRange(CC_UPPER_ROW, CC_LEFT_COL, CC_LOWER_ROW - 6, arrDate.length).setBackgrounds(colorGrid);

  // Group formatting header
  ccSheetToExport.getRange(CC_UPPER_ROW - 1, CC_LEFT_COL, 1, arrDay.length)
    .merge()
    .setHorizontalAlignment("center")
    .setValue("5. Ngày làm việc trong tháng")
    .setBorder(null, null, true, null, null, null);

  // Batch Payload Output
  ccSheetToExport.getRange(CC_UPPER_ROW + 1, 1, workDays.length, workDays[0].length) 
    .setValues(workDays)
    .setBorder(true, true, true, true, true, true)
    .setHorizontalAlignment("center")
    .setFontSize(8);

  ccSheetToExport.getRange(CC_UPPER_ROW + 1, 3, workDays.length, 1).setHorizontalAlignment("left");
  ccSheetToExport.getRange(CC_UPPER_ROW + 1, CC_RIGHT_COL, workDays.length, 1).setBorder(true, true, true, true, true, true);
  
  var titleStr = `BẢNG CHẤM CÔNG THÁNG ${dispMonth < 10 ? '0' : ''}${dispMonth} NĂM ${year}`;
  ccSheetToExport.getRange(3, 1).setValue(titleStr);
}
