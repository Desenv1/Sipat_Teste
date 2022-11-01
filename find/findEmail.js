//Função para obter um e-mail.
function findEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var emails = dados.getRange(1,colEmail,1000,1); //Coluna de setores.
  var procurar = emails.createTextFinder(email); //Procurar pela existência do setor na coluna.
  var cellEmail = procurar.findNext(); //Célula onde o setor está localizado.
  if (cellEmail == null) {
    return cellEmail
  }
  else return email; //Retorna o e-mail.
}
