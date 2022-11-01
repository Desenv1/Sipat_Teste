//Função para obter o setor de uma sala.
function findSetor(nameSetor) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var setores = dados.getRange(1,colSetor,1000,1); //Coluna de setores.
  var procurar = setores.createTextFinder(nameSetor); //Procurar pela existência do setor na coluna.
  var cellSetor = procurar.findNext(); //Célula onde o setor está localizado.
  if (cellSetor == null) {
    return cellSetor
  }
  else return nameSetor; //Retorna o setor.
}
