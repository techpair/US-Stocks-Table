
function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
 // var range = e.range;
  //range.setNote('Last modified: ' + new Date());

UpdateStocksTable();

}

function sortTable(){

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
  
  var sheet = spreadsheet.getActiveSheet(); //spreadsheet.getSheetByName('المستثمر الذكي');
  
  sheet.getRange(5, 1, sheet.getLastRow()-1, 9).activate();

  spreadsheet.getActiveRange().sort([{column: 9, ascending: false}]); 

}


function UpdateStocksTable() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  //var dataRange = sheet.getDataRange();
  //var values = dataRange.getValues();

  var headerRow = findCellRow("Stock Symbol");
  var headerCol = findCellColumn("Stock Symbol");

Logger.log(headerRow);

  var Direction=SpreadsheetApp.Direction;
  var lastRow =ss.getRange("A"+(ss.getLastRow()+1)).getNextDataCell(Direction.UP).getRow();
  var lastRow2 =ss.getRange("B"+(ss.getLastRow()+1)).getNextDataCell(Direction.UP).getRow();
  
  if(lastRow2<lastRow){

      for (var i = lastRow2; i < lastRow; i++) {
          
          //update current price =googlefinance(A5)

          var cell = sheet.getRange(i+1,2);
          //cell.setFormula("=GOOGLEFINANCE(A"+ (i+1) + ")");
          cell.setValue("<Enter Entry Price>");
          

          //update yesterday price =googlefinance(A5,"closeYest")
          var cell2 = sheet.getRange(i+1,3);
          //cell2.setFormula("=GOOGLEFINANCE(A" + (i+1) + "," + String.fromCharCode(34) + "closeYest" + String.fromCharCode(34) + ")" );

          cell2.setFormula("=GOOGLEFINANCE(A"+ (i+1) + ")");

          //update profit/loss =googlefinance(A5,"change")
          //var cell3 = sheet.getRange(i+1,4);
          //cell3.setFormula("=IFERROR(GOOGLEFINANCE(A" + (i+1) + "," + String.fromCharCode(34) + "change" + String.fromCharCode(34) + "),0.00)" );

          //cell3.setNumberFormat("#,##0.00");

          //update profit/loss percentage =googlefinance(A5,"changepct")
          //var cell4 = sheet.getRange(i+1,5);
          //cell4.setFormula("=IFERROR(GOOGLEFINANCE(A" + (i+1) + "," + String.fromCharCode(34) + "changepct" + String.fromCharCode(34) + ")/100,0.00%)" );

          //cell4.setNumberFormat("0.00%");


      }
  

  }

  for (var i = headerRow+1; i <= lastRow; i++) {
    
        //update profit/loss =googlefinance(A5,"change")
      var cell3 = sheet.getRange(i,4);
      cell3.setFormula("=IFERROR(C" + (i) + "-" + "B" + (i) + ",0.00)");
    Logger.log("=IFERROR(B" + (i) + "-" + "C" + (i) + ",0.00)");

      cell3.setNumberFormat("#,##0.00");

      //update profit/loss percentage =googlefinance(A5,"changepct")
      var cell4 = sheet.getRange(i,5);
      cell4.setFormula("=IFERROR(D" + (i) + "/" + "B" + (i) + ",0.00%)");

      cell4.setNumberFormat("0.00%");

      

      //if (cell4.getValue() < 0) {
      //if (sheet.getRange(i,2).getValue()>sheet.getRange(i,3).getValue()) {
      
      //cell4.setFontColor('red');
    //}else if (sheet.getRange(i,2).getValue()<sheet.getRange(i,3).getValue()){
      //cell4.setFontColor('#38761d');//dark green 2
    //}else{
      //  cell4.setFontColor('#000000');
    //}

    //var cell5 = sheet.getRange(i,8);
    //if (cell5.getValue()=='Buy') {
      //cell5.setFontColor('#38761d');
    //}else if (cell5.getValue()=='Sell'){
     // cell5.setFontColor('red');
    //}else if (cell5.getValue()=='Hold'){
      //cell5.setFontColor('orange');
    //}

  }
  //Logger.log(lastRow);
//Logger.log(lastRow2);



}

function findCellRow(strKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == strKeyword) {
        row = values[i][j+1];
        //Logger.log(row);
       //Logger.log(i+1); // This is your row number
       return i+1;
      }
    }    
  }  
}

function findCellColumn(strKeyword) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  for (var i = 0; i < values.length; i++) {
    var row = "";
    for (var j = 0; j < values[i].length; j++) {     
      if (values[i][j] == strKeyword) {
        row = values[i][j+1];
        //Logger.log(row);
       //Logger.log(j+1); // This is your row number
       return j+1;
      }
    }    
  }  
}