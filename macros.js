function sortHAN() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('AC CHECKS'), true);
  spreadsheet.getRange('D1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(4, true);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(1, true);
};