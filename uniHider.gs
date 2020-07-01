function delAllUni(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(5, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch all universes?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['A65:A134', 'A139:A208', 'A213:A282', 'A287:A356', 'A361:A430', 'A435:A504', 'A509:A578', 'A583:A652', 'A657:A726', 'A731:A800', 'A805:A874', 'A879:A948', 'A953:A1022', 'A1027:A1096', 'A1101:A1170', 'A1175:A1244', 'A1249:A1318', 'A1323:A1392', 'A1397:A1466', 'A1471:A1540', 'A1545:A1614', 'A1619:A1688', 'A1693:A1762', 'A1767:A1836']).activate()
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 



function uniA(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(5, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A65:A134').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniB(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(6, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A139:A208').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniC(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(7, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A213:A282').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniD(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(8, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A287:A356').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniE(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(9, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A361:A430').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniF(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(10, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A435:A504').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 


function uniG(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(11, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A509:A578').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniH(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(12, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A583:A652').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniI(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(13, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A657:A726').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniJ(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(14, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A731:A800').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniK(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(15, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A805:A874').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniL(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(16 , 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A879:A948').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniM(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(17, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A953:A1022').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniN(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(18, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1027:A1096').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniO(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(19, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1101:A1170').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniP(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(20, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1175:A1244').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniQ(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(21, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1249:A1318').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniR(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(22, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1323:A1392').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniS(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(23, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1397:A1466').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniT(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(24, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1471:A1540').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniU(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(25, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1545:A1614').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniV(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(26, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1619:A1688').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniW(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(27, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1693:A1762').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 

function uniX(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataUni = ss.getSheetByName("2.Data Patch").getRange(28, 2, 1,1).getValues().toString();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to unpatch ' + dataUni + ' poppet?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
      var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1767:A1836').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});;
} 