function onEdit(e){
   var sheet = e.range.getSheet();
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if(sheet.getName() == 'Patch' && col==1){
 
  var xx=row;                            //last row with data in TCB
  sheet.getRange(xx,6).setValue(sheet.getRange(xx,22).getValue());
    
  }
  
}
