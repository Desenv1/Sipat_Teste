//Função para transferir um ou mais patrimônios para determinada sala.
function transferir(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slave = ss.getActiveSheet(); //Planilha ativa.
  var slaveName = slave.getSheetName(); //Nome da planilha ativa.
  var lista = ss.getSheetByName('Lista de transferências'); //Planilha da lista de transferências.
  var origem = slave.getName(); //Nome da planilha ativa.
  var emailResponsavel = getEmailResponsavel(slaveName); //E-mail do responsável pela sala.

  var emailUser = Session.getActiveUser(); //E-mail do usuário.
  //Se os e-mails não coincidirem, o usuário será alertado que ele não tem permissão para isso.
  if (emailResponsavel != emailUser) {
    Browser.msgBox('Você não está autorizado a realizar a transferência nessa planilha ativa. Por favor, tente em uma sala de sua responsabilidade.', Browser.Buttons.OK);
    return
  }
  var selection = Browser.msgBox('Deseja transferir os patrimônios selecionados na tabela? Ao selecionar YES os patrimônios a serem transferidos serão os que ja estão selecionados e se for selecionado NO terá que ser digitado as linhas onde estão os patrimônios.',Browser.Buttons.YES_NO_CANCEL);
  if(selection == 'no'){
  //Usuário é solicitado para digitar as linhas dos patrimônios que deseja transferir.
    var linha = Browser.inputBox('Transferência de Patrimônio - Patrimônio', 'Digite o(s) número(s) da(s) linha(s) do(s) patrimônio(s) a ser transferido(s). Caso seja mais de um, separe as linhas por vírgulas e defina intervalos por "-". Não coloque espaço. Exemplo: "5,6-8,13-16,17"', Browser.Buttons.OK_CANCEL);
  }
  else if(selection == 'yes'){
    linha = bySelection();
  }
  else{
    return;
  }
  if (linha == 'cancel' || linha == 0) {
    return
  }
  var linhas = getLines(linha);
  
  var numLinhas = linhas.length; //Tamanho do vetor.
  if(linhas.length == 0){
    SpreadsheetApp.getUi().alert('O número de patrimônios a transferir é igual a 0. Nenhuma alteração será feita.');
    return;
  }
  //Usuário é solicitado para digitar a sala destino dos patrimônios que deseja transferir.
  var destino = Browser.inputBox('Transferência de Patrimônio - Destino', 'Digite o código da sala de destino do patrimônio.', Browser.Buttons.OK_CANCEL)
  if (destino == 'cancel' || linha == 0) {
    SpreadsheetApp.getUi().alert('Processo encerrado!');
    return 
  }
  var destino = findSala(destino); //Procura pela sala digitada.
  //Se ela não existir, o usuário é avisado.
  if(destino == null){
    SpreadsheetApp.getUi().alert('Sala inexistente!  Por favor, repita o procedimento.');
    return
  }
  
  var pass = '0';
  while (pass != '1') {
    //Usuário é solicitado para digitar qual opção de transferência ele deseja.
    var opcao = Browser.inputBox('Transferência de Patrimônio - Finalidade', 'Digite a letra para o tipo de operação.\n'+
                                 'T - Transferência\n E - Empréstimo\n M - Manutenção', Browser.Buttons.OK_CANCEL)
    //Caso seja transferência, não há data de devolução.
    if (opcao == 'T') {
      opcao = 'Transferência';
      pass = '1';
      var dataDevolucao = '-';
    }
    //Caso seja empréstimo, há data de devolução.
    else if (opcao == 'E') {
      opcao = 'Empréstimo';
      pass = '1';
      //Usuário é solicitado para digitar qual a previsão da data de devolução.
      var dataDevolucao = Browser.inputBox('Transferência de Patrimônio - Data de Devolução', 'Qual a data prevista para devolução?', Browser.Buttons.OK_CANCEL);
    }
    //Caso seja manutenção, há data de devolução.
    else if (opcao == 'M') {
      opcao = 'Manutenção';
      pass = '1';
      //Usuário é solicitado para digitar qual a previsão da data de devolução.
      var dataDevolucao = Browser.inputBox('Transferência de Patrimônio - Data de Devolução', 'Qual a data prevista para devolução?', Browser.Buttons.OK_CANCEL);
    }
    else if(opcao == 'cancel') {
      return
    }
    //Caso o usuário digite uma opção inexistente, ele é avisado.
    else {
      Browser.msgBox('A opção digitada é inválida. Por favor, repita o procedimento.',Browser.Buttons.OK);
      pass = '0';
    }
  }
  //atualizarSala();

  var indices = indexColums(slave);
  //Loop que insere os patrimônios na lista.
  var patrimonios = getValuesByIndex(slave,indices,linhas);
  patrimonios.filter(function (value){ // tirando todos os patrimonios sem plaqueta
    return value[0] == 0;
  })
  for(let i of patrimonios){
    i.splice(indices.length - 4, 4);
    i.splice(i.length,0,origem,destino,opcao,dataDevolucao);
  }

  manageTrans(lista,'add patris', patrimonios,0);

  //lista.getRange('A2:L1000').setFontSize(10);
  var emailRemetente = getEmailResponsavel(origem); //E-mail do responsável pela sala de origem.
  var emailDestinatario = getEmailResponsavel(destino); //E-mail do responsável pela sala destino.

  
  //E-mails de confirmação enviados de acordo com o número de patrimônios transferidos.
  if(numLinhas == 1){
    var plaqueta = patrimonios[0][0];
    MailApp.sendEmail(emailRemetente,
                      "Transferência de Patrimônio - SiPat - Notificação de Envio",
                      "Patrimônio: " + plaqueta + "\n" +
                      "Origem: " + origem + "\n" +
                      "Destino: " + destino + "\n" +
                      "Responsável do destino: " + emailDestinatario + "\n" +
                      "Situação: Aguardando autorização na Lista de Transferências. "
                      + "\n-------------------------------------------------------\n"
                      + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    
    
    MailApp.sendEmail(emailDestinatario,
                      "Transferência de Patrimônio - SiPat - Notificação de recebimento",
                      "Patrimônio: " + plaqueta + "\n" +
                      "Origem: " + origem + "\n" +
                      "Destino: " + destino + "\n" +
                      "Responsável do destino: " + emailDestinatario + "\n" +
                      "Situação: Aguardando autorização na Lista de Transferências. \nAcesse: " + ss.getUrl() + " para autorizar ou não a transferência."
    + "\n-------------------------------------------------------\n"
    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    
  }
  
  else{
    MailApp.sendEmail(emailRemetente,
                      "Transferência de Patrimônios - SiPat - Notificação de Envio",
                      "Origem: " + origem + "\n" +
                      "Destino: " + destino + "\n" +
                      "Responsável do destino: " + emailDestinatario + "\n" +
                      "Situação: Aguardando autorização na Lista de Transferências. "
                      + "\n-------------------------------------------------------\n"
                      + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  
    
    MailApp.sendEmail(emailDestinatario,
                      "Transferência de Patrimônios - SiPat - Notificação de recebimento",
                      "Origem: " + origem + "\n" +
                      "Destino: " + destino + "\n" +
                      "Responsável do destino: " + emailDestinatario + "\n" +
                      "Situação: Aguardando autorização na Lista de Transferências. \nAcesse: " + ss.getUrl() + " para autorizar ou não a transferência."
    + "\n-------------------------------------------------------\n"
    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    
  }
  
  var aux = -1;
  //Loop para deletar os patrimônios da planilha da sala de origem.
  for(var i = 0; i<numLinhas; i++){
    aux++;
    if(i>0)slave.deleteRow(linhas[i]-aux);
    else slave.deleteRow(linhas[i]);
  }
}