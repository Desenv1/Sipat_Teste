function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameArray = []; 
  var sheets = ss.getSheets(); //Vetor de planilhas.
   //Loop que forma um vetor com os nomes das salas.
  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }
  
  //Loop que move as planilhas para ordem alfabética.
  for( var j = 0; j < sheets.length; j++ ) {
    var nameSheet = sheetNameArray[j];
    var sheet = ss.getSheetByName(nameSheet);
    var rows = sheet.getMaxRows();
    var columns = sheet.getMaxColumns();
    if(rows > 200) {
      sheet.deleteRows(200,rows-200);
    }
    if (columns > 14) {
      sheet.deleteColumns(14,columns-14);
    }
  }
  Browser.msgBox("Linhas e Colunas removidas!")
}
