function onOpen() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [];
  menuEntries.push({name: "Show All Rows", functionName: "showAllRows"});
  menuEntries.push(null);
  menuEntries.push({name: "Hide Empty Rows", functionName: "hideEmptyRows"});

  ss.addMenu("Row Visibility", menuEntries);
  
  var menuEntries1 = [];
  menuEntries1.push({name: "Insert Dividing Border", functionName: "insborder"});
  menuEntries1.push(null);
  menuEntries1.push({name: "Remove Selected Dividing Border", functionName: "delborder"});
  menuEntries1.push(null);
  menuEntries1.push({name: "Reset all Borders", functionName: "delallborder"});

  ss.addMenu("Borders", menuEntries1);
  
 var ui = SpreadsheetApp.getUi();
    ui.createMenu("Export PDF")
        .addItem("Current Sheet", "exportCurrentPDF")
        .addItem("Export All Sheets", "exportAllSheets")
        .addItem("Export From DropDown (J2)", "ExportSheetFromJ2")
        .addSeparator()
        .addSubMenu(ui
        .createMenu("Header")
        .addItem("Add Header", "addHeader")
        .addItem("Remove Header", "deleteHeader"))
        .addSubMenu(ui.createMenu("Email").addItem("Send Folder As Mail", "sendFolderAsMail"))
        .addToUi();

} 

function showAllRows(){
  
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getActiveSheet();
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
    sheet.showRows(1, sheet.getLastRow());
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
    sheet.showRows(1,29);
  
  
}


function showDataRows(){
  
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getActiveSheet();
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
   sheet.showRows(1, 29);
var range = sheet.getRange(63, 1, sheet.getLastRow(), sheet.getLastColumn());
    sheet.showRows(63,sheet.getLastRow());
  
}

function hideEmptyRows(){
  
     var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getActiveSheet();
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
        //get the values to those rows
    var values = range.getValues();

    //go through every row
    for (var i=0; i<values.length; i++){

        //if row value is equal to empty  
        if(values[i][0] === ""){

        //hide that row
        sheet.hideRows(i+1);
          
          
        }
    }
}

function hiderows(){
  
     var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getActiveSheet();
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
        //get the values to those rows
    var values = range.getValues();

    //go through every row
    for (var i=6; i<values.length; i++){

        //if row value is equal to empty  
        if(values[i][3] === "-"){

        //hide that row
        sheet.hideRows(i+1);
        }
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getRange(1, 1, 28, 10);
        //get the values to those rows
    var values = range.getValues();

    //go through every row i is starting row
    for (var i=6; i<values.length; i++){

        //if row value is equal to empty  
        if(values[i][4] === ""){

        //hide that row
        sheet.hideRows(i+1);
        }}
    SpreadsheetApp.getActiveSpreadsheet().toast('Rows Hidden');
}


function hidemultirows(){
  
     var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getActiveSheet();
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());

    var values = range.getValues();

    for (var i=3; i<values.length; i++){
      
        if(values[i][11] === "-"){
        sheet.hideRows(i+1,12);
        }
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Rows Hidden');
}

function hideWeightRows(){
  
     var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getActiveSheet();
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
        //get the values to those rows
    var values = range.getValues();

    //go through every row
    for (var i=0; i<values.length; i++){

        //if row value is equal to empty  
      if(values[i][12] === ""){
    //hide that row
        sheet.hideRows(i+1);
      
      }
        if(values[i][12] === "-"){
         sheet.hideRows(i+1);       
          
          
                }}
    
}


function insborder()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = ss.getActiveSheet();  
  
   var cell= sheet.getActiveCell().getA1Notation();
   var rw=sheet.getRange(cell).getRow();
  
  var rng = sheet.getRange("A" + rw + ":L" + rw);

  rng.setBorder(null, null, true, null, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  
 }

function delborder()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getActiveSheet();  
  
   var cell= sheet.getActiveCell().getA1Notation();
   var rw=sheet.getRange(cell).getRow();
  
  var rng = sheet.getRange("A" + rw + ":L" + rw);

  rng.setBorder(null, null, true, null, false, false, '#c7c7c7', SpreadsheetApp.BorderStyle.SOLID);
}

function delallborder()
{ 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getActiveSheet();  
  
   var cell= sheet.getActiveCell().getA1Notation();
   var rw=sheet.getRange(cell).getRow();
  
  var rng = sheet.getRange("A4:L1000");

  rng.setBorder(null, null, true, null, false, true, '#c7c7c7', SpreadsheetApp.BorderStyle.SOLID);}
