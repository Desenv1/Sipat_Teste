//Função para remover usuário de uma sala.
function remover(){
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  //Se a pessoa não estiver autorizada, ela será avisada.
  if (emailUser != 'dener@teiacoltec.org' && emailUser != 'rafael@teiacoltec.org' && emailUser != 'patrimonio@teiacoltec.org') {
    SpreadsheetApp.getUi().alert('Você não está autorizado a realizar esse tipo de operação!');
    return
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName('Dados dos responsáveis'); //Planilha dos dados dos responsáveis.
  var dadosLR = sheetDados.getLastRow(); //Última linha escrita dos dados dos responsáveis.
  //O usuário é requisitado para digitar o e-mail do usuário que será removido.
  var emailCadastro = Browser.inputBox('Digite o e-mail do usuário a ser removido.',Browser.Buttons.OK_CANCEL);
  if (emailCadastro == 'cancel' || emailCadastro == 0) return;
  var emailCadastroCheck = findEmail(emailCadastro); //O e-mail é procurado na planilha de dados.
  //Se ele não estiver cadastrado, o usuário será avisado.
  if (emailCadastroCheck == null) {
    SpreadsheetApp.getUi().alert('O email '+ emailCadastro + ' não corresponde a nenhum usuário cadastrado! Por favor, tente novamente.');
    return
  }
  //O usuário é requisitado para digitar o código da sala que será removida da responsabilidade do usuário.
  var salaCadastro = Browser.inputBox('Digite o código da sala que será removida da responsabilidade do usuário.',Browser.Buttons.OK_CANCEL);
  if (salaCadastro == 'cancel' || salaCadastro == 0) return;
  //Se esse código não estiver cadastrado para nenhuma sala, o usuário será avisado.
  var salaCadastroCheck = findSala(salaCadastro);
  if (salaCadastroCheck == null) {
    SpreadsheetApp.getUi().alert('O código '+ salaCadastro + ' não está cadastrado para nenhuma sala!');
    return
  }
  var procuraSala = ss.getSheetByName(salaCadastroCheck); //Planilha da sala.
  //Se a sala não estiver cadastrada, o usuário será avisado.
  if (procuraSala == null) {
    SpreadsheetApp.getUi().alert('A(o) '+ salaCadastroCheck + ' não está cadastrada!');
    return
  }
  else {
    var protection = procuraSala.protect();
  }
  //O usuário é perguntado se deseja remover a sala do Sipat também.
  var removerSala = Browser.msgBox('Deseja também remover a sala ' + salaCadastroCheck + ' do Sipat?',Browser.Buttons.YES_NO_CANCEL);
  if (removerSala == 'cancel') return;
  //Se sim, ele recebe uma confirmação para ter certeza.
  if (removerSala == 'yes') {
    var removerSala = Browser.msgBox('Tem certeza que deseja remover a sala  ' + salaCadastroCheck + ' do Sipat?',Browser.Buttons.YES_NO_CANCEL);
  }
  if (removerSala == 'cancel') return;
  
  var salas = sheetDados.getRange(1,colSala,1000,1); //Coluna de salas.
  var procurar = salas.createTextFinder(salaCadastroCheck); //Procurar pela sala na coluna.
  var cellSala = procurar.findNext(); //Célula onde a sala está localizada.
  //Se a sala não for localizada, o usuário é avisado.
  if (cellSala == null) {
    SpreadsheetApp.getUi().alert('A(o) '+ salaCadastro + ' não está cadastrada para nenhum usuário!')
    return
  }
  else {
    while (cellSala != null) {
      var linhaSala = cellSala.getRow(); //Linha da sala.
      var emailSala = sheetDados.getRange(linhaSala,colEmail).getValue(); //E-mail do responsável pela sala.
      //Se os e-mails coincidirem, o usuário é removido da lista de editores e a linha da sala é deletada.
      if (emailSala == emailCadastro) {
        protection.removeEditor(emailCadastro);
        sheetDados.deleteRow(linhaSala);
        //Se o usuário desejou remover a sala, a planilha da mesma será deletada.
        if (removerSala == 'yes') {
          var sheetRemove = ss.getSheetByName(salaCadastroCheck);
          ss.deleteSheet(sheetRemove);
        }
        //Os dados são rearranjados em ordem alfabética.
        var rangeTotal = sheetDados.getRange("A2:F");
        rangeTotal.sort(SORT_ORDER2);
        //E-mail de confirmação.
        MailApp.sendEmail(emailCadastro,
                          "SiPat - Responsabilidade de Salas",
                          "O patrimônio da sala "+ salaCadastro + " foi removido da sua responsabilidade."
                          + "\n-------------------------------------------------------\n"
                          + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  
        SpreadsheetApp.getUi().alert('Usuário '+ emailCadastro + ' removido da(o) '+ salaCadastro + ' com sucesso!');
        return
      }
      //Se o e-mail não coincidir, procura a sala novamente.
      else {
        var cellSala = procurar.findNext();
      }
      SpreadsheetApp.getUi().alert('Usuário '+ emailCadastro + ' não está cadastrado para a(o) '+ salaCadastro + '.'); //Alerta de confirmação.
    }
  }
}
