function CountDemo() {
   
   var ss = SpreadsheetApp.getActiveSpreadsheet();
//Get current active Spreadsheet
 var sheet = ss.getSheetByName("1.Count");
   
   var sht = ss.getSheetByName("1.Patch");
   
   var Avals = sht.getRange("A1:A").getValues();
  var mm = Avals.filter(String).length;
//Get all values from the spreadsheet's rows
   var data = sht.getRange("J3:J" + mm).getValues();
//Create an array for non-duplicates
 var newData = [];
//Iterate through a row's cells
 for (var i in data) {
   var row = data[i];
   var duplicate = false;
   for (var j in newData) {
    if (row.join() == newData[j].join()) {
     duplicate = true;
    }
  }
//If not a duplicate, put in newData array
 if (!duplicate) {
  newData.push(row);
 }
}
//Delete the old Sheet and insert the newData array
   sheet.getRange("A6:A").clearContent();
   sheet.getRange("A6:A").clear();

   
SpreadsheetApp.flush();
 sheet.getRange(6, 1, newData.length, newData[0].length).setValues(newData);
var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
    sheet.showRows(1,sheet.getLastRow());
   
   SORT_ORDER = [

{column: 1, ascending: true}, // 1 = column number, sort by ascending order 

];



  var ss = SpreadsheetApp.getActiveSpreadsheet();
 // var sheet = ss.getSheetByName(SHEET_NAME);
  var range = sheet.getRange("A6:A");
  range.sort(SORT_ORDER);
     
     sheet.getRange("A6:H34").clear({formatOnly: true});
     
     
  sheet.getRange('B6:H34').activate();
  sheet.getRange('B6:H6').copyTo(sheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange('A6').activate();
     
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C5:C').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('F5:F').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('B1:H4').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_THICK);
  spreadsheet.getRange('B5:H34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_THICK)
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle');
     spreadsheet.getRange('B35:H36').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_THICK)
  spreadsheet.getRange('B5:H34').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, null, true, '#b7b7b7', SpreadsheetApp.BorderStyle.SOLID);
  spreadsheet.getRange('B5:C34').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, true, null, '#b7b7b7', SpreadsheetApp.BorderStyle.DASHED);
  spreadsheet.getRange('D5:F34').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, true, null, '#b7b7b7', SpreadsheetApp.BorderStyle.DASHED);
  spreadsheet.getRange('G5:H34').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, true, null, '#b7b7b7', SpreadsheetApp.BorderStyle.DASHED);
;
  
 SpreadsheetApp.flush();
 
 
  ss.toast('Count Updated...');
}

function mergeData() {var start = 6; // Start row number for values.
  var c = {};
  var k = "";
  var m = "";
  var offset = 0;
  var ss = SpreadsheetApp.getActiveSheet();

  // Retrieve values of column B.
  var data = ss.getRange(start, 2, ss.getLastRow(), 1).getValues().filter(String);

  // Retrieve the number of duplication values.
  data.forEach(function(e){c[e[0]] = c[e[0]] ? c[e[0]] + 1 : 1;});

  // Merge cells.
  data.forEach(function(e){
    if (k != e[0]) {
      ss.getRange(start + offset, 2, c[e[0]], 1).merge();
      ss.getRange(start + offset, 3, c[e[0]], 1).merge();
      ss.getRange(start + offset, 7, c[e[0]], 1).merge();
      ss.getRange(start + offset, 8, c[e[0]], 1).merge();
      offset += c[e[0]];
};
    
    k = e[0];
 
    }

    
  );
}
function hideEmpty() {var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName("1.Count");;
var range = sheet.getRange(1, 1, 34, 9);
        //get the values to those rows
    var values = range.getValues();

    //go through every row
    for (var i=6; i<values.length; i++){

        //if row value is equal to empty  
        if(values[i][0] === ""){

        //hide that row
        sheet.hideRows(i+1);
        }
    }
}
function mergeHide() {
  mergeData();
  hideEmpty();
}
     
function runFullCount() {
  CountDemo();
  mergeHide();
  dataPatchCount();
  jumpToPatch();
}

function jumpToPatch(){
  var spreadsheet=  SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('1.Patch');
  spreadsheet.setActiveSheet(sheet);
Browser.msgBox('Count Generated.')
}
