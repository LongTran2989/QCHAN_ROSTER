/**
 * Handles Roster sheet generation, calendar date formatting, and Public roster syncing.
 */

/**
 * Triggered by UI menu. Prompts for Month/Year to build a new Roster sheet.
 */
function createNewRoster() {
  try {
    var ui = SpreadsheetApp.getUi();
    var responseMonth = ui.prompt("Enter month (1, 2..., 12)");
    var responseYear = ui.prompt("Enter year (2022, 2023...)");

    var respondedMonth = parseInt(responseMonth.getResponseText());
    var respondedYear = parseInt(responseYear.getResponseText());

    if(isNaN(respondedMonth) || isNaN(respondedYear) || respondedMonth <= 0 || respondedMonth > 12 || respondedYear < 2022 || respondedYear > 2050) {
      ui.alert("Invalid month or year!");
      return;
    }

    fillDateToSheet(respondedYear, respondedMonth);
    setBackgroundColor();
  } catch (err) {
    console.error("Error in createNewRoster: " + err.message);
  }
}

/**
 * Prepares a new Roster Sheet by duplicating a template and populating date/day headers dynamically.
 * @param {number} year 
 * @param {number} month 
 */
function fillDateToSheet(year, month) {
  var targetMonthIndex = month - 1; // JS counts month starting with 0 = JAN
  var d = new Date(year, targetMonthIndex);
  var maxDaysOfMonth = daysInMonth(year, month);

  var arrDate = [];
  var arrDay = [];

  for (var i = 0; i < maxDaysOfMonth; i++){
    arrDate.push(d.getDate());
    arrDay.push(textDay(d.getDay()));
    d.setDate(d.getDate() + 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName(CONFIG.SHEET_IDS.ROSTER_TEMPLATE);
  if (!templateSheet) {
    throw new Error(`Template sheet '${CONFIG.SHEET_IDS.ROSTER_TEMPLATE}' not found.`);
  }

  templateSheet.activate();

  var sheetName = `${textMonth(month)}${year-2000} R0`;
  var sheet = ss.duplicateActiveSheet().setName(sheetName).activate();

  templateSheet.showSheet();
  templateSheet.hideSheet();

  sheet.activate();
  
  // Fill new sheet with Calendar info
  const CALENDAR_ROW = 2;
  const CALENDAR_COL = 3;
  
  // Clear maximum possible 31 items
  sheet.getRange(CALENDAR_ROW, CALENDAR_COL, 2, 31).clearContent();
  
  // Fill dynamically based on length
  var dateRange = sheet.getRange(CALENDAR_ROW, CALENDAR_COL, 1, maxDaysOfMonth);
  var dayRange = sheet.getRange(CALENDAR_ROW + 1, CALENDAR_COL, 1, maxDaysOfMonth);
  
  dateRange.setValues([arrDate]);
  dayRange.setValues([arrDay]);

  ss.getRange("B1").setValue(`${textMonth(month)} ${year}`);
  ss.getRange("B2").setValue("R0");
  ss.getRange("AL1").setValue(month);
}

/**
 * Scans the active Roster schedule entirely and formats Weekend columns in one single batch assignment.
 */
function setBackgroundColor() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var maxCols = 31; // From original logic, up to 31 columns max
  var rowHeight = CONFIG.ROSTER.LOWER_ROW - CONFIG.ROSTER.UPPER_ROW + 1;
  
  var dayValues = sheet.getRange(CONFIG.ROSTER.UPPER_ROW + 1, CONFIG.ROSTER.LEFT_COL, 1, maxCols).getValues()[0];
  var currentBgArray = sheet.getRange(CONFIG.ROSTER.UPPER_ROW, CONFIG.ROSTER.LEFT_COL, rowHeight, maxCols).getBackgrounds();
  
  // Process the entire block in memory
  for (var c = 0; c < maxCols; c++) {
    var dayText = dayValues[c];
    var paintColor = null;
    
    if (dayText === "Sat" || dayText === "Sun") {
      paintColor = CONFIG.COLORS.BG_SAT_SUN;
    } else if (dayText === "") {
      paintColor = CONFIG.COLORS.BG_NULL;
    }
    
    // Assign that color top to bottom for this column
    if (paintColor) {
      for (var r = 0; r < rowHeight; r++) {
        currentBgArray[r][c] = paintColor;
      }
    }
  }

  // Single batched api push
  sheet.getRange(CONFIG.ROSTER.UPPER_ROW, CONFIG.ROSTER.LEFT_COL, rowHeight, maxCols).setBackgrounds(currentBgArray);
}

/**
 * Copies the active Roster grid to a public spreadsheet securely, omitting extraneous rows.
 */
function updatePublicRoster() {
  try {
    var public_roster_sp = openSpreadsheetSafe(CONFIG.SHEET_IDS.PUBLIC_ROSTER);
    var active_roster_sp = SpreadsheetApp.getActiveSpreadsheet();
    var active_roster_sh = active_roster_sp.getActiveSheet()
    var active_roster_sh_name = active_roster_sh.getName()
    
    // Copy exact sheet snapshot over
    var public_roster_sh = active_roster_sh.copyTo(public_roster_sp)
    
    // Ensure only the new published sheet and "DATE" sheets remain
    var sheets = public_roster_sp.getSheets();
    for (const sheet of sheets) {
      if (sheet.getName() !== public_roster_sh.getName() && sheet.getName() !== "DATE") {
        public_roster_sp.deleteSheet(sheet)
      }
    }
    
    var date = new Date()
    var d = date.getDate()
    
    public_roster_sh.setName(active_roster_sh_name)
    // Wipe extra details outside typical domain
    public_roster_sh.getRange(43, 1, 100, 300).clear()
    
    // Create view restriction window:
    var range = public_roster_sh.getRange("C4:AG42")
    var values = range.getValues()
    var values2 = createGrid(values.length, values[0].length, "");
    
    // Filter purely in memory
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[0].length; j++) {
        var strVal = values[i][j].toString();
        // Render window strictly bounded -3 to +3 days from current date per original business logic
        if ((strVal === "BA1-C" || strVal === "BA2") && j < d + 3 && j >= d - 3) {
          values2[i][j] = values[i][j]
          
          if (i - 1 >= 0) values2[i - 1][j] = values[i - 1][j]
          if (i + 1 < values.length) values2[i + 1][j] = values[i + 1][j]
        }
      }
    }
    
    // Batch updates
    public_roster_sh.getDataRange().clearDataValidations()
    range.setBackground(CONFIG.COLORS.WHITE)
    range.setValues(values2)
    public_roster_sh.getRange(2, d + 2, 2, 1).setBackground(CONFIG.COLORS.A320) // Current day highlight
    
    showDialog()
  } catch (err) {
    console.error("Error in updatePublicRoster: " + err.message);
    SpreadsheetApp.getUi().alert("Error syncing to Public Roster: " + err.message);
  }
}
