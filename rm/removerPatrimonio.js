//Função para remover um patrimônio.
function removerPatrimonio() {
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var mestreLR = lastRowMestre();
  var plaquetas = mestre.getRange(1,1,mestreLR,1); //Coluna de plaquetas.
  var salaAtual = ss.getActiveSheet(); //Planilha ativa.
  var nomeSala = salaAtual.getName(); //Nome da planilha ativa.
  var emailResponsavel = getEmailResponsavel(nomeSala); //E-mail do responsável por essa planilha.
  //Se os e-mails não coincidirem, o usuário é avisado que ele não tem permissão para realizar essa ação. 
  if(emailUser != emailResponsavel) {
    Browser.msgBox('Você não tem autorização para realizar ações nessa sala.',Browser.Buttons.OK);
    return;
  }
  //O usuário é requisitado para digitar a linha do patrimônio que ele deseja remover.
  var linhaPatrimonio = Browser.inputBox('Remoção de Patrimônio', 'Digite a linha do patrimônio que deseja remover.', Browser.Buttons.OK_CANCEL);
  if (linhaPatrimonio == 'cancel') return;
  //Se for a linha do cabeçalho, o usuário é avisado.
  else if (linhaPatrimonio == 1) {
    SpreadsheetApp.getUi().alert('Linha 1 bloqueada! Por favor, repita o procedimento.');
    return
  }
  var indices = indexColums(salaAtual);
  var plaqueta = salaAtual.getRange(linhaPatrimonio,indices[0]).getValue(); //Plaqueta do patrimônio.
  //O usuário é perguntado se realmente deseja remover tal patrimônio.
  var check = Browser.msgBox('Remoção de Patrimônio', 'Você realmente deseja remover o patrimônio ' + plaqueta + ' presente na linha ' 
                                + linhaPatrimonio +'?', Browser.Buttons.YES_NO);
  //Se não, a função é cancelada.
  if (check == 'no') {
    Browser.msgBox('Operação cancelada.',Browser.Buttons.OK);
    return
  }
  var procurar = plaquetas.createTextFinder(plaqueta); // Cria uma busca da plaqueta.
  var cellPlaqueta = procurar.findNext();
  
  if (cellPlaqueta == null) { }
  else {
    var linhaPlaqueta = cellPlaqueta.getRow();
    mestre.deleteRow(linhaPlaqueta); //Deleta o patrimônio da planilha mestre.
  }
  salaAtual.deleteRow(linhaPatrimonio); //Deleta o patrimônio da sala atual.
  //E-mail de confirmação.
    MailApp.sendEmail(emailResponsavel,
                    "Remoção de Patrimônio - SiPat",
                    "Patrimônio: " + plaqueta + "\n" +
                    "Sala: " + nomeSala + "\n" +
                    "Responsável da Sala: " + emailResponsavel + "\n" +
                    "Situação: Removido."
  + "\n-------------------------------------------------------\n"
  + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    
  
  
  Browser.msgBox('Patrimônio removido com sucesso!',Browser.Buttons.OK); //Alerta de confirmação.
}
