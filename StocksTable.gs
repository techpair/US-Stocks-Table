
function onEdit(e) {
  // Set a comment on the edited cell to indicate when it was changed.
 // var range = e.range;
  //range.setNote('Last modified: ' + new Date());

UpdateStocksTable();

}

function UpdateStocksTable() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  //var dataRange = sheet.getDataRange();
  //var values = dataRange.getValues();

  var headerRow = findCellRow("Stock Symbol");
  var headerCol = findCellColumn("Stock Symbol");
  var Direction=SpreadsheetApp.Direction;
  var lastRow =ss.getRange("A"+(ss.getLastRow()+1)).getNextDataCell(Direction.UP).getRow();
  var lastRow2 =ss.getRange("B"+(ss.getLastRow()+1)).getNextDataCell(Direction.UP).getRow();
  
  if(lastRow2<lastRow){

      for (var i = lastRow2; i < lastRow; i++) {
          
          //update current price =googlefinance(A5)

          var cell = sheet.getRange(i+1,2);
          cell.setFormula("=GOOGLEFINANCE(A"+ (i+1) + ")");

          //update yesterday price =googlefinance(A5,"closeYest")
          var cell2 = sheet.getRange(i+1,3);
          cell2.setFormula("=GOOGLEFINANCE(A" + (i+1) + "," + String.fromCharCode(34) + "closeYest" + String.fromCharCode(34) + ")" );

          //update profit/loss =googlefinance(A5,"change")
          var cell3 = sheet.getRange(i+1,4);
          cell3.setFormula("=GOOGLEFINANCE(A" + (i+1) + "," + String.fromCharCode(34) + "change" + String.fromCharCode(34) + ")" );

          cell3.setNumberFormat("#,##0.00");

          //update profit/loss percentage =googlefinance(A5,"changepct")
          var cell4 = sheet.getRange(i+1,5);
          cell4.setFormula("=GOOGLEFINANCE(A" + (i+1) + "," + String.fromCharCode(34) + "changepct" + String.fromCharCode(34) + ")/100" );

          cell4.setNumberFormat("0.00%");


      }
  

  }

  for (var i = headerRow+1; i <= lastRow; i++) {
    
        //update profit/loss =googlefinance(A5,"change")
      var cell3 = sheet.getRange(i,4);
      cell3.setFormula("=C" + (i) + "-" + "B" + (i));
    //Logger.log("=B" + (i) + "-" + "C" + (i));

      cell3.setNumberFormat("#,##0.00");

      //update profit/loss percentage =googlefinance(A5,"changepct")
      var cell4 = sheet.getRange(i,5);
      cell4.setFormula("=D" + (i) + "/" + "C" + (i));

      cell4.setNumberFormat("0.00%");
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