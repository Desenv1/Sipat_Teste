//Função para remover um patrimônio.
function removerPatrimonio() {
  var ui = SpreadsheetApp.getUi();
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var salaAtual = ss.getActiveSheet(); //Planilha ativa.
  var nomeSala = salaAtual.getName(); //Nome da planilha ativa.
  var emailResponsavel = getEmailResponsavel(nomeSala); //E-mail do responsável por essa planilha.
  //Se os e-mails não coincidirem, o usuário é avisado que ele não tem permissão para realizar essa ação. 
  if(emailUser != emailResponsavel && emailUser != 'wagnercoltec@teiacoltec.org' && emailUser != 'desenv1@teiacoltec.org') {
    Browser.msgBox('Você não tem autorização para realizar ações nessa sala.',Browser.Buttons.OK);
    return;
  }
  //O usuário é requisitado para digitar a linha do patrimônio que ele deseja remover.
  var linhaPatrimonio = Browser.inputBox('Remoção de Patrimônio', 'Digite as linhas do patrimônio que deseja remover.', Browser.Buttons.OK_CANCEL);
  if (linhaPatrimonio == 'cancel') return;
  //Se for a linha do cabeçalho, o usuário é avisado.
  var lines = getLines(linhaPatrimonio);
  if (lines.indexOf(1) == 0) {
    SpreadsheetApp.getUi().alert('Linha 1 bloqueada! Confira se as linhas estão corretas');
    lines.shift();
  }
  var indices = indexColums(salaAtual);
  var patrimonios = getValuesByIndex(salaAtual,indices,lines); //Plaqueta do patrimônio.
  var patStr = String();
  for(let j of patrimonios){
    patStr += j[0] + ' - ' + j[2] + '\n';
  }
  //O usuário é perguntado se realmente deseja remover tal patrimônio.
  var check = ui.alert('Remoção de Patrimônio', 'Você realmente deseja remover os seguintes patrimônios?\n'
                               + patStr, ui.ButtonSet.YES_NO);
  //Se não, a função é cancelada.
  if (check == ui.Button.NO) {
    Browser.msgBox('Operação cancelada.',Browser.Buttons.OK);
    return
  }
  manageMestre(mestre,'remove',patrimonios);
  lines.reverse();
  for(let i of lines){
    salaAtual.deleteRows(i,1);
  }
  //E-mail de confirmação.
    MailApp.sendEmail(emailUser,
                    "Remoção de Patrimônio - SiPat",
                    "Sala: " + nomeSala + "\n" +
                    "Responsável da Sala: " + emailResponsavel + "\n" +
                    "Patrimônios:\n" + patStr + "\n" +
                    "Situação: Removido."
  + "\n-------------------------------------------------------\n"
  + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    
  
  
  ui.alert('Patrimônio removido com sucesso!',Browser.Buttons.OK); //Alerta de confirmação.
}
