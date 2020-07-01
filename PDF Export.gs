//////////////////// TEST SCRIPTS ////////////////////

function doStuffPatch() {
  createFolder();
  
  Utilities.sleep(2000);
  
  
  
  createProjectFolder();
  
   Utilities.sleep(2000);
  
  exportPatchPDF();
  
  
}
 

function test() {


var temp = getProjectFolder()



Browser.msgBox(temp)





}





//////////////////// GET PROJECT NAME //////////////////// 
  
function getProjectName() {
  
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var projectName = spreadsheet.getRange("Setup!C2").getValues() + " v" + spreadsheet.getRange("Setup!C4").getValues();
  //var parentFolder = DriveApp.getFolderById(DriveApp.getFoldersByName(folderName).next().getId());
  
  
  return projectName;
}
  
//////////////////// GET PROJECT FOLDER NAME ////////////////////

function getProjectFolder() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var projectName = spreadsheet.getRange("Setup!C2").getValues() + " v" + spreadsheet.getRange("Setup!C4").getValues();
  var projectFolderID = DriveApp.getFolderById(DriveApp.getFoldersByName(projectName).next().getId());
  
  
//  Browser.msgBox(projectFolderID)
  
   return projectFolderID;
}












//////////////////// CREATE PAPERWORK FOLDER //////////////////// 



function createFolder(folderID, folderName){

  var folderID = DriveApp.getFolderById(DriveApp.getRootFolder().getId());
  var folderName = "Paperwork";
  var parentFolder = DriveApp.getFolderById(DriveApp.getRootFolder().getId());
  var subFolders = parentFolder.getFolders();
  var doesntExists = true;

  
// Check if folder already exists.
  while(subFolders.hasNext()){
    var folder = subFolders.next();
    
    //If the name exists return the id of the folder
    if(folder.getName() === folderName){
      doesntExists = false;
      newFolder = folder;
      return newFolder.getId();
    };
  };
  //If the name doesn't exists, then create a new folder
  if(doesntExists = true){
    //If the file doesn't exists
    newFolder = parentFolder.createFolder(folderName);
    return newFolder.getId();
  };

}
 
 
 
 
 
 
 
 
 //////////////////// CREATE PROJECT NAMED FOLDER - NAME FROM GETPROJECTNAME SCRIPT //////////////////// 

 
 

 function createProjectFolder(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
   var folderName = getProjectName();
 // var folderName = spreadsheet.getRange("Setup!C2").getValues() + " " + spreadsheet.getRange("Setup!C4").getValues();
  var parentFolder = DriveApp.getFolderById(DriveApp.getFoldersByName("Paperwork").next().getId());

 
  var subFolders = parentFolder.getFolders();
  var doesntExists = true;


// Check if folder already exists.
  while(subFolders.hasNext()){
    var folder = subFolders.next();
    
    //If the name exists return the id of the folder
    if(folder.getName() === folderName){
      doesntExists = false;
      newFolder = folder;
      return newFolder.getId();
    };
  };
  //If the name doesn't exists, then create a new folder
  if(doesntExists = true){
    //If the file doesn't exists
    newFolder = parentFolder.createFolder(folderName);
    return newFolder.getId();
  };
}






//////////////////// PDF EXPORT SCRIPTS ////////////////////



//////////////////// EXPORT PART AS PDF - USED BY EXPORT PATCH SCRIPT ////////////////////


function exportPartAsPDF(predefinedRanges, pagename) {
  var ui = SpreadsheetApp.getUi()
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  
  var selectedRanges
  var fileSuffix
  if (predefinedRanges) {
    selectedRanges = predefinedRanges
    fileSuffix = '-predefined'
  } else {
    var activeRangeList = spreadsheet.getActiveRangeList()
    if (!activeRangeList) {
      ui.alert('Please select at least one range to export')
      return
    }
    selectedRanges = activeRangeList.getRanges()
    fileSuffix = '-selected'
  }
  
  if (selectedRanges.length === 1) {
    // special export with formatting
    var currentSheet = selectedRanges[0].getSheet()
    var blob = _getAsBlob(spreadsheet.getUrl(), currentSheet, selectedRanges[0])
    
    

   //var fileName = getProjectName() + " " + "PATCH";
   var fileName = getProjectName() + " - " + pagename;
    _exportBlob(blob, fileName)
    return
  }
  
  var tempSpreadsheet = SpreadsheetApp.create(spreadsheet.getName() + fileSuffix)
  var tempSheets = tempSpreadsheet.getSheets()
  var sheet1 = tempSheets.length > 0 ? tempSheets[0] : undefined
  SpreadsheetApp.setActiveSpreadsheet(tempSpreadsheet)
  
  for (var i = 0; i < selectedRanges.length; i++) {
    var selectedRange = selectedRanges[i]
    var originalSheet = selectedRange.getSheet()
    var originalSheetName = originalSheet.getName()
    
    var destSheet = tempSpreadsheet.getSheetByName(originalSheetName)
    if (!destSheet) {
      destSheet = tempSpreadsheet.insertSheet(originalSheetName)
    }
    
    Logger.log('a1notation=' + selectedRange.getA1Notation())
    var destRange = destSheet.getRange(selectedRange.getA1Notation())
    destRange.setValues(selectedRange.getValues())
    destRange.setTextStyles(selectedRange.getTextStyles())
    destRange.setBackgrounds(selectedRange.getBackgrounds())
    destRange.setFontColors(selectedRange.getFontColors())
    destRange.setFontFamilies(selectedRange.getFontFamilies())
    destRange.setFontLines(selectedRange.getFontLines())
    destRange.setFontStyles(selectedRange.getFontStyles())
    destRange.setFontWeights(selectedRange.getFontWeights())
    destRange.setHorizontalAlignments(selectedRange.getHorizontalAlignments())
    destRange.setNumberFormats(selectedRange.getNumberFormats())
    destRange.setTextDirections(selectedRange.getTextDirections())
    destRange.setTextRotations(selectedRange.getTextRotations())
    destRange.setVerticalAlignments(selectedRange.getVerticalAlignments())
    destRange.setWrapStrategies(selectedRange.getWrapStrategies())
  }
  
  // remove empty Sheet1
  if (sheet1) {
    Logger.log('lastcol = ' + sheet1.getLastColumn() + ',lastrow=' + sheet1.getLastRow())
    if (sheet1 && sheet1.getLastColumn() === 0 && sheet1.getLastRow() === 0) {
      tempSpreadsheet.deleteSheet(sheet1)
    }
  }
  
  
 }
 
 



 //////////////////// EXPORT BLOB ////////////////////



function _exportBlob(blob, fileName, projectFolderID) {
  blob = blob.setName(fileName)
  
  var projectFolder = getProjectFolder(projectFolderID);
  var fid = DriveApp.getFoldersByName(projectFolder).next().getId();
  var folder = DriveApp.getFolderById(fid)


   var file = folder.getFilesByName(fileName);
   while (file.hasNext()) {//If there is another element in the iterator
    var thisFile = file.next();
    var idToDLET = thisFile.getId();
    Logger.log('idToDLET: ' + idToDLET);

     DriveApp.getFileById(idToDLET).setTrashed(true);
  }



var pdfFile = folder.createFile(blob,)
  
  

  

  
  
  
  // Display a modal dialog box with custom HtmlService content.
  const htmlOutput = HtmlService
    .createHtmlOutput('<p>Click to open <a href="' + pdfFile.getUrl() + '" target="_blank">' + fileName + '</a></p>')
    .setWidth(300)
    .setHeight(80)
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export Successful')
}




 
 //////////////////// EXPORT PATCH PDF ////////////////////



function exportPatchPDF() {

createFolder();
  
  Utilities.sleep(1000);
  
  
  
  createProjectFolder();
  
   Utilities.sleep(1000)
   
  var pagename = "Patch";


 var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var Avals = ss.getRange("Patch!A1:A").getValues();
  var Alast = Avals.filter(String).length;
  var LastRow = ("A1:I"+Alast);
  
  var range = ss.getRange(LastRow);
  ss.setNamedRange('patch_export', range);
  var rangeCheck = ss.getRangeByName('patch_export');
  var rangeCheckName = rangeCheck.getA1Notation();

  


  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var allNamedRanges = spreadsheet.getNamedRanges()
  var toPrintNamedRanges = []
  for (var i = 0; i < allNamedRanges.length; i++) {
    var namedRange = allNamedRanges[i]
    if (/^patch_export/.test(namedRange.getName())) {
      Logger.log('found named range ' + namedRange.getName())
      toPrintNamedRanges.push(namedRange.getRange())
    }
  }
  if (toPrintNamedRanges.length === 0) {
    SpreadsheetApp.getUi().alert('No print areas found. Please add at least one \'print_area_1\' named range in the menu Data > Named ranges.')
    return
  } else {
    toPrintNamedRanges.sort(function (a, b) {
      return a.getSheet().getIndex() - b.getSheet().getIndex()
    })
    exportPartAsPDF(toPrintNamedRanges, pagename)
  }
}









///////////DAN

function exportSpreadsheet() {
 
  //All requests must include id in the path and a format parameter
  //https://docs.google.com/spreadsheets/d/{SpreadsheetId}/export
 
  //FORMATS WITH NO ADDITIONAL OPTIONS
  //format=xlsx       //excel
  //format=ods        //Open Document Spreadsheet
  //format=zip        //html zipped          
  
  //CSV,TSV OPTIONS***********
  //format=csv        // comma seperated values
  //             tsv        // tab seperated values
  //gid=sheetId             // the sheetID you want to export, The first sheet will be 0. others will have a uniqe ID
  
  // PDF OPTIONS****************
  //format=pdf     
  //size=0,1,2..10             paper size. 0=letter, 1=tabloid, 2=Legal, 3=statement, 4=executive, 5=folio, 6=A3, 7=A4, 8=A5, 9=B4, 10=B5  
  //fzr=true/false             repeat row headers
  //fzc=true/false             repeat column headers
  //portrait=true/false        false =  landscape
  //fitw=true/false            fit window or actual size
  //gridlines=true/false
  //printtitle=true/false
  //pagenum=CENTER/UNDEFINED    CENTER = show page numbers / UNDEFINED = do not show
  //attachment = true/false     dunno? Leave this as true
  //gid=sheetId                 Sheet Id if you want a specific sheet. The first sheet will be 0. others will have a uniqe ID. Leave this off for all sheets. 
  //printnotes=false            Set to false if you don't want to export the notes embedded in a sheet
  //top_margin=[number]         Margins - you need to put all four in order fir it to works, and they have to be to 
  //left_margin=[number]          2DP. So 0.00 for zero margin.
  //right_margin=[number]
  //bottom_margin=[number]
  //horizontal_alignment=CENTER Horizontal Alignment: LEFT/CENTER/RIGHT
  //vertical_alignment=TOP      Vertical Alignment: TOP/MIDDLE/BOTTOM
  //scale=1/2/3/4               1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
  //pageorder=1/2               1= Down, then over / 2= Over, then down
  //sheetnames=true/false
  //range=[NamedRange]          Named ranges supported - see below
 
  // EXPORT RANGE OPTIONS FOR PDF
  //need all the below to export a range
  //gid=sheetId                must be included. The first sheet will be 0. others will have a uniqe ID
  //ir=false                   seems to be always false
  //ic=false                   same as ir
  //r1=Start Row number - 1        row 1 would be 0 , row 15 wold be 14
  //c1=Start Column number - 1     column 1 would be 0, column 8 would be 7   
  //r2=End Row number
  //c2=End Column number
 var ss = SpreadsheetApp.getActiveSpreadsheet();

  var ssid = ss.getSheetByName("Setup").getRange(1, 14, 1,1).getValues().toString();
  var url = "https://docs.google.com/spreadsheets/d/"+ssid+"/export"+
                                                        "?format=pdf&"+
                                                        "size=7&"+
                                                        "gid=94192874&"+
                                                        "fzr=true&"+
                                                        "portrait=false&"+
                                                        "fitw=true&"+
                                                        "gridlines=false&"+
                                                        "printtitle=true&"+
                                                        "sheetnames=false&"+
                                                        "pagenum=CENTER&"+
                                                        "attachment=true";
                                                        
  var params = {method:"GET",headers:{"authorization":"Bearer "+ ScriptApp.getOAuthToken()}};
  
  var response = UrlFetchApp.fetch(url, params).getBlob();
  // save to drive
  DriveApp.createFile(response);
  
  //or send as email
  /*
  MailApp.sendEmail(email, subject, body, {
        attachments: [{
            fileName: "TPS REPORT" + ".pdf",
            content: response.getBytes(),
            mimeType: "application/pdf"
        }]
    });};
  */
  
}
function gfk() {
  return SpreadsheetApp.getActiveSpreadsheet().getId();
}
