//Vectorworks CSV Import From /Google Drive/ImportExport/Import Export.csv



function vwImport() {


// Clear Old Data

  SpreadsheetApp.flush()
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var sheet = ss.getSheetByName("VW Import");
  sheet.getRange("A1:O").clearContent();
  
// Import New Data

  var fid = DriveApp.getFoldersByName("ImportExport").next().getId();
  var folder = DriveApp.getFolderById(fid);
  var file = folder.getFilesByName("VW Export.csv").next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

Browser.msgBox('Fixtures Imported.')

}

function hoistImport() {


// Clear Old Data

  SpreadsheetApp.flush()
     var ss = SpreadsheetApp.getActiveSpreadsheet();
     var sheet = ss.getSheetByName("Hoist Import");
  sheet.getRange("A1:CV").clearContent();
  
// Import New Data

  var fid = DriveApp.getFoldersByName("ImportExport").next().getId();
  var folder = DriveApp.getFolderById(fid);
  var file = folder.getFilesByName("Hoist Export.csv").next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

Browser.msgBox('Hoists Imported.')

}