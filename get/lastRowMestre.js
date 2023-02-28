//Função para obter a última linha escrita da planilha mestre, com excessão da coluna de sincronização.
function lastRowMestre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  return lastRow(mestre);
}

//Função para obter a última linha escrita de uma planilha
function lastRow(sheet) {
  var plaquetasRange = sheet.getRange("A:A"); //Coluna de plaquetas.
  var plaquetas = plaquetasRange.getValues(); //Coluna de plaquetas.
  var sheetLR = 0;
  for(var i=1; i < plaquetas.length; i++){
    if (plaquetas[i][0] == 0) { // plaqueta vazia
      sheetLR = i;
      break;
    }
  }
  return sheetLR;
}
//Função para obter a última linha escrita da planilha mestre, com excessão da coluna de sincronização.
function lastRowOfColumn(column) {
  var sheetLR = column.length;
  for(var i=1; i < sheetLR; i++){
    if (column[i][0] == 0) {
      sheetLR = i;
      break;
    }
  }
  return sheetLR;
}