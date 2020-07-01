// function OldpopulateSheet() {
//   showAllRows();
//   CopyHeadNo();
// }

// function CopyHeadNo() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var source_sheet = ss.getSheetByName("VW Import");
//   var target_sheet = ss.getSheetByName("1.Patch");
//   var source_range = source_sheet.getRange("A2:A");
//   var target_range = target_sheet.getRange("A3:A");
// source_range.copyTo(target_range, {contentsOnly:true});

// SHEET_NAME = "1.Patch";
// SORT_DATA_RANGE = "A3:A";
// SORT_ORDER = [

// {column: 1, ascending: true}, // 1 = column number, sort by ascending order 

// ];

//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getSheetByName(SHEET_NAME);
//   var range = sheet.getRange(SORT_DATA_RANGE);
//   range.sort(SORT_ORDER);
  
//  SpreadsheetApp.flush();
  

  
  
  
//   sheet.getRange("F3:F1000").clearContent();
//   var Avals = sheet.getRange("A1:A").getValues();
//   var mm = Avals.filter(String).length;
  
//   for (i=3;i<=mm;i++){
//    sheet.getRange("F" + i).setValue(sheet.getRange("V" + i).getValue());
//   }
  
//    SpreadsheetApp.flush();
//    Browser.msgBox('Import Complete')
 
//   }
  
  
//   function validation() {
  
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sht = ss.getSheetByName("Patch");
//  ss.getRange("F3:F1000").clearContent();
//  ss.getRange("F3:F1000").clearDataValidations();
//  for(i=770; i<1002; i++) {
  
//  var cells = sht.getRange("F" + i + ":F" + i);
//  var rules = sht.getRange("V" + i + ":Y" + i);
  
//  var cells = sheet.getRange("B2:C15");
//  var rules = dates.getRange("A:A");
//  var validation = SpreadsheetApp.newDataValidation().requireValueInRange(rules).build();
//  cells.setDataValidation(validation);
//  sht.getRange("F" + i).setValue(sht.getRange("V" + i).getValue());
//  }
  
//  }
  
//  function SortHeadNo() {

//  SHEET_NAME = "Patch";
//  SORT_DATA_RANGE = "A3:A";
//  SORT_ORDER = [

//  {column: 1, ascending: true}, // 1 = column number, sort by ascending order 

//  ];



//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getSheetByName(SHEET_NAME);
//  var range = sheet.getRange(SORT_DATA_RANGE);
//  range.sort(SORT_ORDER);
//  ss.toast('Sort Complete.');
//  }
  

  
  
  
  
  
//  Populate Patch Sheet
  
  
  
  function populateSheet() {


  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source_sheet = ss.getSheetByName("VW Import");
  var target_sheet = ss.getSheetByName("1.Patch");
  var source_range = source_sheet.getRange("A2:A");
  var target_range = target_sheet.getRange("A3:A");
  
  
  
  target_sheet.getRange("A3:A").clearContent(); //Clears Head Numbers
  target_sheet.getRange("F3:F").clearContent(); // Clears Modes
    target_sheet.getRange("L3:L").clearContent(); //Clear Clamps
  
  ss.toast('Previous Patch Removed.'); //Announcement Cleared
  
   SpreadsheetApp.flush();
  
  source_range.copyTo(target_range, {contentsOnly:true}); //Copies New Head Numbers

  SHEET_NAME = "1.Patch";
  SORT_DATA_RANGE = "A3:A";
  SORT_ORDER = [
  
  {column: 1, ascending: true}, // 1 = column number, sort by ascending order 

  ];






  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("1.Patch");
  var range = sheet.getRange(SORT_DATA_RANGE);
  range.sort(SORT_ORDER);
  
  
     ss.toast('Unit ID Sorted.');
  
  
  SpreadsheetApp.flush();
  

  
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source_sheet = ss.getSheetByName("1.Patch");
  var target_sheet = ss.getSheetByName("1.Patch");
  var source_range = source_sheet.getRange("V3:V");
  var target_range = target_sheet.getRange("F3:F");
  source_range.copyTo(target_range, {contentsOnly:true});
  
  
   ss.toast('Default Modes Assigned.'); 
    
  ss.getRange('L3').activate();
  ss.getCurrentCell().setFormula('=VLOOKUP(D3,LOOKUP,9,0)');
  ss.getActiveRange().autoFill(ss.getRange('L3:L1001'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  ss.getRange('L3:L188').activate();
      Browser.msgBox('Patch Sheet Populated.')
  
 
  }