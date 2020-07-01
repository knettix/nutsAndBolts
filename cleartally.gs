function clearTally() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("Data Patch").getRange(5, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to clear the tally?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['C10:Q16', 'C20:Q26', 'C30:Q33', 'C37:Q42', 'C46:Q52', 'C58:Q63', 'C67:Q73', 'C77:Q85', 'C89:Q95', 'C99:Q106', 'C112:Q122', 'C128:Q149', 'C155:Q162', 'C168:Q176']).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
    spreadsheet.getRangeList(['C5:Q5']).activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

  
  spreadsheet.getRange('C5').activate();
  spreadsheet.getCurrentCell().setValue('LOCATION 1');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C5:P5'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C5:P5').activate();
  spreadsheet.getRange('Q5').activate();
  spreadsheet.getCurrentCell().setValue('SPARE');
  spreadsheet.getRange('A7').activate();
};