//Função para obter a última linha escrita da planilha mestre, com excessão da coluna de sincronização.
function lastRowMestre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var mestreLR = mestre.getLastRow(); //Última linha escrita da planilha mestre.
  var salas = mestre.getRange(1,8,mestreLR,1); //Coluna de plaquetas.
  let values = salas.getValues();
  for(var i=1; i < values.length; i++){
    if (values[i][0] == 0) {
      mestreLR = i;
      break;
    }
  }
  return mestreLR;
}

//Função para obter a última linha escrita de uma planilha
function lastRow(sheet) {
  var sheetLR = sheet.getLastRow(); //Última linha escrita da planilha .
  var plaquetas = sheet.getRange(1,1,sheetLR,1).getValues(); //Coluna de plaquetas.
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