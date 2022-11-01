//Com o nome de uma sala de argumento, a função retorna o nome do deu responsável.
function getNomeResponsavel(sala) {
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
  var nome = dados.getRange(rowSala,colNome).getValue(); //Obtém o nome do responsável.
  return nome; //Retorna o nome do responsável.
}
