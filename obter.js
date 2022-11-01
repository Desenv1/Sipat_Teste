// Enviar e-mails avisando das transferências e das negações: principalmente das negações para o proprietário pegar o patrimônio de volta.
//Função para obter patrimônio desfeito.
function obter(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var mestreLR = lastRowMestre(); //Última linha escrita da planilha mestre.
  var lista = ss.getActiveSheet(); //Planilha ativa.
  var nomeLista = lista.getSheetName(); //Nome da planilha ativa.
  //Se for a Lista de desfazimento, a função prossegue.

  //Caso o usuário não esteja na lista de desfazimento, ele será avisado.
  if(nomeLista != 'Lista de desfazimento'){
    SpreadsheetApp.getUi().alert('Página errada! Você deve realizar essa ação apenas na página "Lista de desfazimento".')
    return;
  }
  

  var history = ss.getSheetByName('Histórico de Transferência'); //Planilha da lista de transferências.
  var historyLR = lastRow(history); ////Última linha escrita da lista de transferências.
  var linha = Browser.inputBox('Obtenção de Patrimônio Desfeito - Patrimônio', 'Digite o número da linha do patrimônio.', Browser.Buttons.OK_CANCEL); //Usuário é solicitado para digitar a linha do patrimônio que deseja obter.
  if (linha == 'cancel') {
    return
  }
  //Se a linha do cabeçalho for escolhida, o usuário é alertado.
  else if (linha == 1) {
    SpreadsheetApp.getUi().alert('Linha 1 bloqueada! Por favor, repita o procedimento.');
    return
  }
  //Usuário é solicitado para digitar a sala para qual deseja transferir o patrimônio.
  var destino = Browser.inputBox('Obtenção de Patrimônio Desfeito - Destino', 'Digite o código da sala para qual você deseja transferir esse patrimõnio.', Browser.Buttons.OK_CANCEL);
  if (destino == 'cancel') {
    return
  }
  
  var origem = lista.getRange(linha,8).getValue(); //Sala de origem do patrimônio.
  var sheetOrigem = ss.getSheetByName(origem); //Planilha da sala de origem do patrimônio.
  var origemLR = sheetOrigem.getLastRow(); //Última linha escrita da sala de origem.
  var destino = findSala(destino); //Procura a existência da sala destino.
  //Se não existir, o usuário é avisado.
  if(destino == null){
    SpreadsheetApp.getUi().alert('Sala inexistente!  Por favor, repita o procedimento.');
    return
  }
  var sheetDestino = ss.getSheetByName(destino); //Planilha da sala destino
  var destinoLR = lastRow(sheetDestino); //Última linha escrita da sala destino.
  
  var valoresPat = lista.getRange(linha,1,1,8).getValues(); //Pega todos os valores do patrimônio a ser transferido.
  var plaqueta = valoresPat[0][0]; //Plaqueta do patrimônio
  
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  
  var emailRemetente = getEmailResponsavel(origem); //E-mail do responsável pela sala de origem.
  var emailDestinatario = getEmailResponsavel(destino); //E-mail do responsável pela sala destino.
  //Pergunta se o usuário autoriza a transferência.

  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  //Se o usuário não tiver permissão para realizar tal transferência, ele será avisado.
  if(emailUser != emailDestinatario){
    SpreadsheetApp.getUi().alert('Você não está autorizado a realizar essa transferência!');
    return;
  }

  var permissao = Browser.msgBox('Autorização de Transferência - Patrimônio', 'Você autoriza a transferência do patrimônio ' + plaqueta 
                                  + ' para a(o) ' + destino + '?', Browser.Buttons.YES_NO_CANCEL);
  
  var valoresPat = lista.getRange(linha,1,1,8).getValues(); //Pega os valores do patrimônio a ser transferido.
  var salaAnterior = valoresPat[0][8]; //Sala anterior do patrimônio.

  var plaquetas = mestre.getRange(1,1,lastRowMestre(),1); //Coluna de plaquetas na planilha mestre
  var plaqueta = valoresPat[0][0]; //Plaqueta do patrimônio transferido.
  var procurar = plaquetas.createTextFinder(plaqueta); //Procurar pela plaqueta na coluna.
  var cellPlaqueta = procurar.findNext(); //Célula onde a plaqueta está localizada.
  
  //Se não for encontrada, o patrimônio é adicionado.
  if (cellPlaqueta == null) {
    mestre.getRange(mestreLR+1,1,1,8).setValues(valoresPat);
    var rowPlaqueta = mestreLR+1;
  }
  else {
    var rowPlaqueta = cellPlaqueta.getRow(); //Linha do patrimônio na planilha mestre.
  }
  var indices = indexColums(sheetDestino);
  //Se a tranferência for permitida...
  if (permissao == 'yes'){
    var patHistory = valoresPat;
    patHistory[0][7] = now;
    patHistory[0].push(destino);
    patHistory[0].push('Obtido da lista de desfazimento.');
    patHistory[0].push('Transferência');
    //Dados da transferência inseridos na lista.
    history.getRange(historyLR+1,1,1,11).setValues(patHistory);
    mestre.getRange(rowPlaqueta,7,1,3).setValues([[now, destino, salaAnterior]]); //Muda a data da última transferência.
    //Novo patrimônio inserido na sala destino.
    setValueByIndex(sheetDestino,indices,destinoLR+1,valoresPat[0]);
    sheetDestino.getRange(destinoLR+1,indices[6]).setValue(now);
    sheetDestino.getRange(destinoLR+1,indices[7]).setValue(origem);
    sheetDestino.getRange(destinoLR+1,indices[8]).setValue('Transferência');
    lista.deleteRow(linha); //Deletado da lista de desfazimento.
    
    autoDoc(sheetOrigem, sheetDestino, patHistory); //Relatório gerado.
    //E-mail de confirmação.
    MailApp.sendEmail(emailRemetente, emailDestinatario,
                      "Transferência de Patrimônio - SiPat - Obtenção de patrimônio desfeito realizada!",
                      "Patrimônio: " + plaqueta + "\n" +
                      "Origem: " + origem + "\n" +
                      "Destino: " + destino + "\n" +
                      "Responsável do destino: " + emailDestinatario + "\n" +
                      "Situação: Patrimônio obtido pelo responsável do destino. "
                      + "\n-------------------------------------------------------\n"
                      + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    
  }
  //Caso o usuário não tenha permitido, a função não acontece.
  else if (permissao == 'no' || permissao == 'cancel'){
    return;
  }
}
