function CopyMode() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("1.Patch");
  
  var Avals = sheet.getRange("A1:A").getValues();
  var mm = Avals.filter(String).length;
  
  
  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection();
var activeRange = selection.getActiveRange();
  
    var cell= activeRange.getA1Notation();
  
  var col=sheet.getRange(cell).getColumn();
  
  if (col!=6){
    
    Browser.msgBox("Please select Mode within range.");
    return;
  }
  
  var frow=sheet.getRange(cell).getRow();
  var lrow= activeRange.getLastRow();
  
  var ftype = sheet.getRange(frow,4).getValue();
  var mode = sheet.getRange(frow,6).getValue();
  
  
 for (i=frow;i<=lrow;i++){    
  if (sheet.getRange("D" + i).getValue()==ftype) {
  sheet.getRange("F" + i).setValue(mode);
   }
 }
  
  
 SpreadsheetApp.flush();
 
 
  ss.toast('Mode updated.');  
 
  }


function CopyModeAll() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("1.Patch");
  
  var Avals = sheet.getRange("A1:A").getValues();
  var mm = Avals.filter(String).length;
  
  
  var cell= sheet.getActiveCell().getA1Notation();
    var frow=sheet.getRange(cell).getRow(); 
 
  
  var ftype = sheet.getRange(frow,4).getValue();
  var mode = sheet.getRange(frow,6).getValue();
  
  
 for (i=3;i<=mm;i++){    
  if (sheet.getRange("D" + i).getValue()==ftype) {
  sheet.getRange("F" + i).setValue(mode);
   }
 }
  
  
 SpreadsheetApp.flush();
 
 
  ss.toast('Mode updated.');  
 
  }



  

  