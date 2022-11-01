//Função de adicionar usuário.
function adicionar(){
  var emailUser = Session.getActiveUser().getEmail(); //Pega o e-mail do usuário.
  if (emailUser != 'desenv1@teiacoltec.org' && emailUser != 'patrimonio@teiacoltec.org' && emailUser != 'wagnercoltec@teiacoltec.org') {
    SpreadsheetApp.getUi().alert('Você não está autorizado a realizar esse tipo de operação!');  //Se o usuário não tiver a permissão, ele não conseguirá adicionar usuário novo.
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName('Dados dos responsáveis');  //Planilha de dados dos responsáveis.
  var sheetLista = ss.getSheetByName('Lista de Transferências'); //Planilha da lista de transferências.
  var dadosLR = sheetDados.getLastRow(); //Última linha da planilha de dados.
  
  var nomeCadastro = Browser.inputBox('Digite o nome do usuário a ser adicionado.',Browser.Buttons.OK_CANCEL); //O administrador é solicitado para digitar o nome do novo usuário.
  if (nomeCadastro == 'cancel' || nomeCadastro == 0) return;
  var emailCadastro = Browser.inputBox('Digite o e-mail do usuário a ser adicionado.',Browser.Buttons.OK_CANCEL); //O administrador é solicitado para digitar o e-mail do novo usuário.
  if (emailCadastro == 'cancel' || emailCadastro == 0) return;
  var aliasCadastro = Browser.inputBox('Digite o código da sala que será destinada a responsabilidade ao usuário.',Browser.Buttons.OK_CANCEL); //O administrador é solicitado para digitar o código da nova sala.
  if (aliasCadastro == 'cancel') return;
  var salaCadastro = Browser.inputBox('Digite o nome da sala que será destinada a responsabilidade ao usuário. Caso a sala não exista, será automaticamente criada.',Browser.Buttons.OK_CANCEL); //O administrador é solicitado para digitar o nome da nova sala.
  if (salaCadastro == 'cancel') return;
  var setorCadastro = Browser.inputBox('Digite o nome do setor da sala que será destinada a responsabilidade ao usuário.',Browser.Buttons.OK_CANCEL); //O administrador é solicitado para digitar o setor da nova sala.
  if (setorCadastro == 'cancel') return;
  var orgaoCadastro = Browser.inputBox('Digite o nome do órgão da sala que será destinada a responsabilidade ao usuário.',Browser.Buttons.OK_CANCEL); //O administrador é solicitado para digitar o órgão da nova sala.
  if (orgaoCadastro == 'cancel') return;
  
  var procuraSala = ss.getSheetByName(salaCadastro);
  if (procuraSala == null){
    createSala(salaCadastro);
    procuraSala = ss.getSheetByName(salaCadastro);
    var protectionSala = procuraSala.protect();
  }
  else {
    var protectionSala = procuraSala.protect();
  }
  // Caso a sala não exista, ela é criada. Depois disso, protejida.
  
  var salas = sheetDados.getRange(1,colSala,1000,1); //Coluna de salas.
  var procurar = salas.createTextFinder(salaCadastro); //Procurar pela sala na coluna.
  var cellSala = procurar.findNext(); //Célula onde a sala está localizada.
  //Todos os dados da sala nova são adicionados.
  if (cellSala == null) {
    sheetDados.getRange(dadosLR+1,colOrgao).setValue(orgaoCadastro);
    sheetDados.getRange(dadosLR+1,colSetor).setValue(setorCadastro);
    sheetDados.getRange(dadosLR+1,colAlias).setValue(aliasCadastro);
    sheetDados.getRange(dadosLR+1,colSala).setValue(salaCadastro);
    sheetDados.getRange(dadosLR+1,colNome).setValue(nomeCadastro);
    sheetDados.getRange(dadosLR+1,colEmail).setValue(emailCadastro);
    protectionSala.addEditor(emailUser);
    protectionSala.removeEditors(protectionSala.getEditors());
    if (protectionSala.canDomainEdit()) { // Verifica se a sala pode ser editada por qualquer usuário do domínio. Se sim, remove essa possibilidade.
      protectionSala.setDomainEdit(false);
    }
    protectionSala.addEditor(emailCadastro); // O novo usuário é adicionado como editor.
    protectionSala.addEditor("wagnercoltec@teiacoltec.org"); // O Wagner é adicionado como editor.
    SpreadsheetApp.getUi().alert(salaCadastro + ' criada! Usuário '+ emailCadastro + ' cadastrado na(o) '+ salaCadastro + ' com sucesso!'); //Aviso confirmando o novo cadastro.
    MailApp.sendEmail(emailCadastro,
                      "SiPat - Responsabilidade de Salas",
                      "O patrimônio da sala " + salaCadastro + " foi atribuído à sua responsabilidade."
                      + "\n-------------------------------------------------------\n"
                      + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  
  } // E-mail de confirmação enviado.
  else {
    var linhaSala = cellSala.getRow();
    var emailSala = sheetDados.getRange(linhaSala, colEmail).getValue();
    if (emailSala == emailCadastro) {
      protectionSala.addEditor(emailCadastro);
      SpreadsheetApp.getUi().alert('Usuário '+ emailCadastro + ' já cadastrado para a(o) '+ salaCadastro + '!');
      return
    } // Caso já esteja cadastrado, o administrador é avisado.
    else {
      sheetDados.getRange(dadosLR+1,colOrgao).setValue(orgaoCadastro);
      sheetDados.getRange(dadosLR+1,colSetor).setValue(setorCadastro);
      sheetDados.getRange(dadosLR+1,colAlias).setValue(aliasCadastro);
      sheetDados.getRange(dadosLR+1,colSala).setValue(salaCadastro);
      sheetDados.getRange(dadosLR+1,colNome).setValue(nomeCadastro);
      sheetDados.getRange(dadosLR+1,colEmail).setValue(emailCadastro);
      protectionSala.addEditor(emailCadastro);
      SpreadsheetApp.getUi().alert('Usuário '+ emailCadastro + ' cadastrado na(o) '+ salaCadastro + ' com sucesso!');
      MailApp.sendEmail(emailCadastro,
                        "SiPat - Responsabilidade de Salas",
                        "O patrimônio da sala "+ salaCadastro + " foi atribuído à sua responsabilidade."
                        + "\n-------------------------------------------------------\n"
                        + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  
    }
    //Todos dados adicionados ou atualizados e e-mail de confirmação enviado
  }
  var rangeTotal = sheetDados.getRange("A2:F");
  rangeTotal.sort(SORT_ORDER2); // Salas colocadas em ordem alfabética.
  sortSheets(); //Reorganiza as salas em ordem alfabética.
}