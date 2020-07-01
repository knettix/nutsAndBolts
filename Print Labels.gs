// Exports CSV to /Google Drive/ImportExport/Import Export.csv
function printLabels() {

  
 var fid = DriveApp.getFoldersByName("ImportExport").next().getId();

 
var folder = DriveApp.getFolderById(fid);
  
var file = folder.getFilesByName('Fixture Labels.csv');
 while (file.hasNext()) {//If there is another element in the iterator
    var thisFile = file.next();
    var idToDLET = thisFile.getId();
    //Logger.log('idToDLET: ' + idToDLET);

     DriveApp.getFileById(idToDLET).setTrashed(true);
  }
    
    
 
  // Check that the file name entered wasn't empty
 // if (fileName.length !== 0) {
    // Add the ".csv" extension to the file name
    var fileName1 = "Fixture Labels.csv";
    // Convert the range data to CSV format
    var csvFile = convertRange_(fileName1);
    // Create a file in Drive with the given name, the CSV data and MimeType (file type)
 // var folderID = "; // Folder id to save in a folder.
var folder = DriveApp.getFolderById(fid);
var newFile = folder.createFile(fileName1, csvFile, MimeType.CSV);
  
  

    Browser.msgBox('Labels sent to printer.')
  
  }
 

 
function convertRange_(csvFileName1) {
  // Get the selected range in the spreadsheet
  
   var Avals = SpreadsheetApp.getActiveSpreadsheet().getRange("8.Fixture Labels!A1:A").getValues();
  var mm = Avals.filter(String).length;
  var ws = SpreadsheetApp.getActiveSpreadsheet().getRange("8.Fixture Labels!A1:G" + mm);
  

  
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
