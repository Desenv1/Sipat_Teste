//Função para obter uma sala.
function findSala(destino) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var salas = dados.getRange(1,colAlias,1000,1); //Coluna de apelidos.
  var procurar = salas.createTextFinder(destino); //Procurar pela sala na coluna.
  var cellSala = procurar.findNext(); //Célula onde a sala está localizada.
  if (cellSala == null) {
    return cellSala
  }
  var rowSala = cellSala.getRow(); //Pega a linha da célula da sala.
  var sala = dados.getRange(rowSala,colSala).getValue(); //Pega o nome da sala.
  return sala; //Retorna a sala.
}
