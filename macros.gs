
function UntitledMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('6:6').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false);
  spreadsheet.getRange('E8').activate();
};

function PDFMacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A:I').activate();
};

function copyformulas() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2:D17').activate();
  spreadsheet.getRange('D2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('D2').activate();
};


function copyToData() {
 var ss = SpreadsheetApp.getActiveSpreadsheet()
 var sheet = ss.getSheetByName('1.Count'); //replace with source Sheet tab name
 var range = sheet.getRange('J6:Q34'); //assign the range you want to copy
 var data = range.getValues();
 var ts = ss.getSheetByName('2.Data Patch'); //replace with destination Sheet tab name

ts.getRange("A32:H60").setValues(data); //you will need to define the size of the copied data see getRange()
}

function mergeDataSheet() {
  var start = 32; // Start row number for values.
  var c = {};
  var k = "";
  var offset = 0;
  var ss = SpreadsheetApp.getActiveSheet();

  // Retrieve values of column A.
  var data = ss.getRange(start, 1, 29, 1).getValues().filter(String);

  // Retrieve the number of duplication values.
  data.forEach(function(e){c[e[0]] = c[e[0]] ? c[e[0]] + 1 : 1;});

  // Merge cells.
  data.forEach(function(e){
    if (k != e[0]) {
      ss.getRange(start + offset, 1, c[e[0]], 2).merge();
      offset += c[e[0]];
    }
    k = e[0];
  });
}

                          
function hideEmptyData() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('2.Data Patch');;
var range = sheet.getRange(1,1,59,sheet.getLastColumn());
        //get the values to those rows
    var values = range.getValues();

    //go through every row
    for (var i=32; i<values.length; i++){

        //if row value is equal to empty  
        if(values[i][2] === ""){

        //hide that row
        sheet.hideRows(i+1);
        }
    }
         
ss.toast('Modes Inserted on Data Patch Sheet.');}

function resetFixtureOverview(){
  
   var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('2.Data Patch');;
var range = sheet.getRange(1, 1, 59, sheet.getLastColumn());
    sheet.showRows(31,30);
  var spreadsheet = ss.getSheetByName('2.Data Patch');
  spreadsheet.getRange('A32:B60').activate()
  .breakApart()
  .mergeAcross();
  spreadsheet.getRange('A32:H60').activate();
  spreadsheet.getActiveRangeList().setBorder(null, null, null, null, null, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
}
  
function dataPatchCount() {
 resetFixtureOverview();
 copyToData();
 mergeDataSheet();
 hideEmptyData();
}

function unpatchedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.setActiveSheet(ss.getSheetByName('2.Data Patch'), true);
  var showCell = ss.getRangeByName("A2028");
  showCell.activate();
}

function unpatchedHide() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var showCell = ss.getRangeByName("A3");
  showCell.activate();
}

function unpatchMulti() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('E3:I374').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('L3:L374').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('L363'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var showCell = ss.getRangeByName("A3");
  showCell.activate();
}
  
function unpatchedMulti() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName('3.Multi Patch'), true);
  var showCell = ss.getRangeByName("A1026");
  showCell.activate();
}


function ResetColCode() {
  var spreadsheet = SpreadsheetApp.getActive();
    var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Are you sure you want to reset ALL colours?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES)
  spreadsheet.getRange('D16').activate();
  spreadsheet.getRange('Q16:R46').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C16:C46').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C16').activate();
  spreadsheet.getCurrentCell().setFormula('=unique(Patch!E3:E)');
  spreadsheet.getRange('C1').activate();
};

function clearHP1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRangeList(['C6:C77', 'H84:I85']).activate()
  .clear({contentsOnly: true, skipFilteredRows: true});
    spreadsheet.getRange('K84').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('0')
  .build());
      spreadsheet.getRange('K85').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('0')
  .build());
  spreadsheet.getRange('C6:C11').activate();
     spreadsheet.getRange('C6:C77').activate();
  spreadsheet.getActiveRangeList().setBackground('#b8d6fb');
};



function populateRack1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C6:C77').activate();
  spreadsheet.getRange('AA6:AA77').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

};



function rack1Make18Way() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:R4').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('18 Way #  |  LOCATION')
  .setTextStyle(0, 8, SpreadsheetApp.newTextStyle()
  .setFontSize(24)
  .build())
  .build());
  spreadsheet.getRange('C24:C77').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('I84').activate();
  spreadsheet.getCurrentCell().setValue('0');
  spreadsheet.getRangeList(['I84', '84:85', '24:77']).activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B84'));
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('86:86').activate();
    spreadsheet.getRange('C6:C11').activate();
      spreadsheet.getRange('K84').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('0')
  .build());
      spreadsheet.getRange('K85').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('0')
  .build());
    spreadsheet.getRange('84:85').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B84'));
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('86:86').activate();
  

  
};

function rack1Make36Way() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:R4').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('36 Way #  |  LOCATION')
  .setTextStyle(0, 8, SpreadsheetApp.newTextStyle()
  .setFontSize(24)
  .build())
  .build());
  spreadsheet.getRange('24:77').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B75'));
  spreadsheet.getActiveSheet().showRows(19, 56);
  spreadsheet.getRange('C42:C77').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('42:77').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B69'));
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('75:75').activate();
  spreadsheet.getRange('84:87').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('A84'));
  spreadsheet.getActiveSheet().showRows(79, 7);
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('84:84').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B84'));
  spreadsheet.getActiveSheet().hideRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  spreadsheet.getRange('85:85').activate();
      spreadsheet.getRange('C6:C11').activate();
  
 
};





function rack1Make72Way() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:R4').activate();
  spreadsheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
  .setText('72 Way #  |  LOCATION')
  .setTextStyle(0, 8, SpreadsheetApp.newTextStyle()
  .setFontSize(24)
  .build())
  .build());
  spreadsheet.getRange('6:87').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B3'));
  spreadsheet.getActiveSheet().showRows(2, 87);
  spreadsheet.getRange('C6:C11').activate();
};

function rack13ph() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L5').activate();
  spreadsheet.getRange('o5:q87').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C6:C11').activate();
};

function rack1hpsoca1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L5').activate();
  spreadsheet.getRange('r5:t87').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C6:C11').activate();
};



function plus1() { 
  SpreadsheetApp.getActiveSheet().getActiveRange().setValue(SpreadsheetApp.getActiveSheet().getActiveRange().getValue() + 1); }
function plus5() { 
  SpreadsheetApp.getActiveSheet().getActiveRange().setValue(SpreadsheetApp.getActiveSheet().getActiveRange().getValue() + 5); }
function plus10() { 
  SpreadsheetApp.getActiveSheet().getActiveRange().setValue(SpreadsheetApp.getActiveSheet().getActiveRange().getValue() + 10); }
function minus1() { 
  SpreadsheetApp.getActiveSheet().getActiveRange().setValue(SpreadsheetApp.getActiveSheet().getActiveRange().getValue() - 1); }
function minus5() { 
  SpreadsheetApp.getActiveSheet().getActiveRange().setValue(SpreadsheetApp.getActiveSheet().getActiveRange().getValue() - 5); }
function minus10() { 
  SpreadsheetApp.getActiveSheet().getActiveRange().setValue(SpreadsheetApp.getActiveSheet().getActiveRange().getValue() - 10); }


function clearTally() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C8:Q14').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C18:Q24').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C28:Q31').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C35:Q40').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C44:Q50').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C56:Q64').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C68:Q74').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C78:Q84').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C90:Q99').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C105:Q126').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C132:Q139').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C145:Q153').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C3').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C3:P3'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A5').activate();
};

function setlocation() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('C16').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C16:C46').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('C16').activate();
  spreadsheet.getCurrentCell().setFormula('=unique(1.Patch!E3:E)');
  spreadsheet.getRange('C1').activate();
};

function dimmerByMulti() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A6').activate();
  spreadsheet.getRange('X6:Y54').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('G7:H54').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('G54'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true})
  .setBackground('#b8d6fb');
  spreadsheet.getRange('G7').activate();
};




function newmarkout() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('6.Hoists'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.getActiveSheet().setName('6.Markout');
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('6.Hoists'), true);
  spreadsheet.getRange('A5:O200').activate();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('6.Markout'), true);
  spreadsheet.getRange('A5:O200').activate();
  spreadsheet.getRange('6.Hoists!A5:O200').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getSheetByName("6.Markout").sort(2);

  spreadsheet.getRange('Q:AV').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('A3').activate();
  spreadsheet.getActiveRangeList().setFontColor('#4a86e8');
  spreadsheet.getCurrentCell().setValue('SORTED BY Y-AXIS');
    spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setFormula('=CONCATENATE(Setup!C2, " - Markout Sheet v", Setup!C4)');
  spreadsheet.getRange('A4').activate();
};


function delsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var delSheet = ss.getSheetByName("6.Markout");
  if (delSheet) {
    ss.setActiveSheet(delSheet);
    ss.deleteActiveSheet();
  }
    ss.setActiveSheet(ss.getSheetByName('6.Hoists'), true);

}

function markout() {
  delsheet();
  newmarkout();
}


function populateClamps() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('L3').activate();
  spreadsheet.getCurrentCell().setFormula('=VLOOKUP(D3,LOOKUP,9,0)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('L3:L188'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('L3:L188').activate();
};

function pcByUnit() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A6').activate();
  spreadsheet.getRange('Y6:Z30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('E6').activate();
  spreadsheet.getRange('AA6:AB30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('A7').activate();
};

function pcByMulti() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A6').activate();
  spreadsheet.getRange('AC6:AD30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('E6').activate();
  spreadsheet.getRange('AE6:AF30').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('E7:E12').activate();
};

function dimmerBySoca() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A6').activate();
  spreadsheet.getRange('X6:Y54').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('G6').activate();
  spreadsheet.getRange('AD6:AE54').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('G:G').activate();
  spreadsheet.getActiveSheet().showColumns(4, 4);
  spreadsheet.getRange('G:H').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('H1'));
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('I:I').activate();
};

function newbyunit() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A6').activate();
  spreadsheet.getRange('X6:Y54').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('A7:B54').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B54'));
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true})
  .setBackground('#b8d6fb');
  spreadsheet.getRange('G6').activate();
  spreadsheet.getRange('Z6:AA54').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('G:H').activate();
  spreadsheet.getRange('I:I').activate();
  spreadsheet.getActiveSheet().showColumns(4, 5);
  spreadsheet.getRange('E:F').activate();
  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('G:G').activate();
  spreadsheet.getRange('A7').activate();
};


function resetBreaks() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('5:1900').activate();
  spreadsheet.getActiveSheet().autoResizeRows(5, 1900);
};
