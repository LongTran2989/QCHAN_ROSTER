function test() {
  var namedRange = SpreadsheetApp.getActiveSpreadsheet().getRangeByName("ROSTER_TABLE");  //getNamed range is SpreadSheet level
  console.log(namedRange);
}

function getFileName() {
  var fileName = SpreadsheetApp.getActive().getName();
  console.log(fileName);
  today = new Date;
  console.log(today.getFullYear());
  fillDateToSheet(2023, 6);
}

function test1(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getSheetByName("test");
  currentSheet.getRange("B1").setNote("ABC");
}

function test2() {
  var sheet = spreadsheet.getSheetByName("MAR23 R14");
  var data = sheet.getRange("B19:B21").getValues();
  var data1 = sheet.getRange("A4:A6").getValues();
  //console.log(data1);
  //console.log(daysInMonth(2023, 3));
  var idRange = sheet.getRange(ROSTER_UPPER_ROW + 2, ROSTER_LEFT_COL - 2, ROSTER_LOWER_ROW - ROSTER_UPPER_ROW, 1);
  var idData = idRange.getValues();
  
}

function test3(){
  var ar = ["A", "B", "C", "D", "E", "F"];
  arraymove(ar, 5, 1);
  console.log(ar);
}
