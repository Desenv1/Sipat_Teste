//Com o nome de uma sala de argumento, a função retorna o setor da mesma.
function getSetorResponsavel(sala) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var salas = dados.getRange(1,colSala,400,1); //Coluna de salas.
  var procurarSala = salas.createTextFinder(sala); //Procurar pela sala na coluna.
  var cellSala = procurarSala.findNext(); //Célula onde a sala está localizada.
  //Se a sala não for encontrada, finaliza a função.
  if (cellSala == null) {
    return
  }
  var rowSala = cellSala.getRow(); //Linha da sala.
  var setor = dados.getRange(rowSala,colSetor).getValue(); //Obtém o setor da sala.
  return setor; //Retorna o setor da sala.
}
