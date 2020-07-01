// Exports CSV to /Google Drive/ImportExport/Import Export.csv
function exportToVectorworks() {

  
 var fid = DriveApp.getFoldersByName("ImportExport").next().getId();

 
var folder = DriveApp.getFolderById(fid);
  
var file = folder.getFilesByName('VW Export.csv');
 while (file.hasNext()) {//If there is another element in the iterator
    var thisFile = file.next();
    var idToDLET = thisFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);

     DriveApp.getFileById(idToDLET).setTrashed(true);
  }
    
    
 
  // Check that the file name entered wasn't empty
 // if (fileName.length !== 0) {
    // Add the ".csv" extension to the file name
    var fileName = "VW Export.csv";
    // Convert the range data to CSV format
    var csvFile = convertRangeToCsvFile_(fileName);
    // Create a file in Drive with the given name, the CSV data and MimeType (file type)
 // var folderID = "; // Folder id to save in a folder.
var folder = DriveApp.getFolderById(fid);
var newFile = folder.createFile(fileName, csvFile, MimeType.CSV);
  
  

    Browser.msgBox('Fixtures exported to Vectorworks.')
  
  }
 

 
function convertRangeToCsvFile_(csvFileName) {
  // Get the selected range in the spreadsheet
  
   var Avals = SpreadsheetApp.getActiveSpreadsheet().getRange("VW Export!A1:A").getValues();
  var mm = Avals.filter(String).length;
  var ws = SpreadsheetApp.getActiveSpreadsheet().getRange("VW Export!A1:O" + mm);
  

  
  try {
    var data = ws.getValues();
    var csvFile = undefined;
 
    // Loop through the data in the range and build a string with the CSV data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
 
        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}


////HOIST EXPORT SCRIPT


function exportHoists() {

  
 var fid = DriveApp.getFoldersByName("ImportExport").next().getId();

 
var folder = DriveApp.getFolderById(fid);
  
var file = folder.getFilesByName('Hoist Export.csv');
 while (file.hasNext()) {//If there is another element in the iterator
    var thisFile = file.next();
    var idToDLET = thisFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);

     DriveApp.getFileById(idToDLET).setTrashed(true);
  }
    
    
 
  // Check that the file name entered wasn't empty
 // if (fileName.length !== 0) {
    // Add the ".csv" extension to the file name
    var fileName2 = "Hoist Export.csv";
    // Convert the range data to CSV format
    var csvFile = convert_(fileName2);
    // Create a file in Drive with the given name, the CSV data and MimeType (file type)
 // var folderID = "; // Folder id to save in a folder.
var folder = DriveApp.getFolderById(fid);
var newFile = folder.createFile(fileName2, csvFile, MimeType.CSV);
  
  

    Browser.msgBox('Hoists exported to Vectorworks.')
  
  }
 

 
function convert_(csvFileName2) {
  // Get the selected range in the spreadsheet
  
   var Avals = SpreadsheetApp.getActiveSpreadsheet().getRange("Hoist Export!A1:A").getValues();
  var mm = Avals.filter(String).length;
  var ws = SpreadsheetApp.getActiveSpreadsheet().getRange("Hoist Export!A1:R" + mm);
  

  
  try {
    var data = ws.getValues();
    var csvFile = undefined;
 
    // Loop through the data in the range and build a string with the CSV data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
 
        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
