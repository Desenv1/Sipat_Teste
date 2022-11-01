// Enviar e-mails avisando das transferências e das negações: principalmente das negações para o proprietário pegar o patrimônio de volta.
//Função para permissão ou não de uma transferência.
function permissoes(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var lista = ss.getActiveSheet(); //Planilha ativa.
  var listaNome = lista.getSheetName(); //Nome da planilha ativa.
  var ui = SpreadsheetApp.getUi(); 

  var history = ss.getSheetByName("Histórico de Transferência");

  //Se o usuário estiver na página errada ele é avisado.
  
  if(listaNome != 'Lista de transferências'){
    SpreadsheetApp.getUi().alert('Página errada! Você deve realizar essa ação apenas na página "Lista de transferências".');
    return;
  }
  var selection = Browser.msgBox('Deseja transferir os patrimônios selecionados na tabela? Ao selecionar SIM, a permissão prosseguirá com os patrimônios selecionados e se for escolhido NÃO, aparecerá uma próxima janela no qual deverá ser digitado as linhas.',Browser.Buttons.YES_NO_CANCEL);
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
  if(numLinhas == 0){
    SpreadsheetApp.getUi().alert('O número de patrimônios a transferir é igual a 0. Nenhuma alteração será feita.');
  return;
  }

  var patrimonios = manageTrans(lista, 'get data', 0, linhas);

  var deniedPatris = manageTrans(lista,'get background', 0, linhas);

  const allEqual = arr => arr.every( v => v === arr[0] ); // verifica se os elem. do array são iguais
  var salasOring = patrimonios.map(function(value) { return value[7]; });
  var salasDest =  patrimonios.map(function(value) { return value[8]; });
  var finalidades = patrimonios.map(function(value) { return value[10]; });
  var consistencia = allEqual(salasOring) && allEqual(finalidades) && allEqual(salasDest);
  if(!consistencia){
    SpreadsheetApp.getUi().alert('A sala de origem, a finalidade ou a sala destino de algum patrimônio difere de outro. Lembre-se que eles têm que ter esses dados iguais para a permissão múltipla de patrimônios.');
      return;
  }
  
  var origem = patrimonios[0][7]; //Sala de origem dos patrimônios. 
  var sheetOrigem = ss.getSheetByName(origem); //Planilha da sala de origem.
  var destino = patrimonios[0][8]; //Sala destino dos patrimônios.
  var sheetDestino = ss.getSheetByName(destino); //Planilha da sala destino.
  
  // Data
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');

  // pegando os emails
  var emailRemetente = getEmailResponsavel(origem); //E-mail do responsável da sala de origem.
  var emailDestinatario = getEmailResponsavel(destino); //E-mail do responsável da sala destino.
  
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  if(emailDestinatario != emailUser && 
     emailUser != 'desenv1@teiacoltec.org' &&
     emailUser != 'wagnercoltec@teiacoltec.org'){ //Se o usuário não estiver autorizado, ele será avisado.
    SpreadsheetApp.getUi().alert('Você não está autorizado a realizar essa transferência!');
    return;
  }
  
  var patStr = String();

  for(let j of patrimonios){
	  patStr += j[0] + '   :   ' + j[2] + '\n';
  }
  var msg = 'Você autoriza a transferência dos patrimônios abaixo para a(o) ' + destino + '?\n\n' + patStr;

  var permissao = ui.alert('Autorização de Transferência', msg, ui.ButtonSet.YES_NO_CANCEL);
  
  var indices = indexColums(sheetDestino);

  

  patrimonios = transpose(patrimonios);
  var valoresPat = patrimonios.slice(0,12);
  valoresPat.splice(8,2);
  valoresPat[6].fill(now);
  valoresPat = transpose(valoresPat);
  if(permissao == ui.Button.YES){
    patrimonios.splice(9,0,Array(numLinhas).fill("Autorizada")); // Insere Autorizada com a qntd de itens
    patrimonios[6].fill(now);
    patrimonios = transpose(patrimonios);
    let valMestre = [now, destino, origem];
	// inserindo os patrimônios
    insertPatrisByIndex(sheetDestino,indices,valoresPat);
	// inserindo no histórico
    manageHistory(history, patrimonios);
	// modificando a planilha mestre
  var patTrans = transpose(patrimonios);
  var salaDest = patTrans[8];
  patTrans.splice(8,4);
  patTrans.splice(7,0,salaDest);
  patrimonios = transpose(patTrans);
	manageMestre(mestre, 'update', patrimonios);
	manageMestre(mestre, 'sort', 0);
  
	// removendo da lista de transferência
	manageTrans(lista, 'remove lines', patrimonios, 0);

    MailApp.sendEmail(emailRemetente,
                          "Transferência de Patrimônio - SiPat - Transferência Aceita",
                          "Patrimônios: " + patStr + "\n" +
                          "Origem: " + origem + "\n" +
                          "Destino: " + destino + "\n" +
                          "Responsável do destino: " + emailDestinatario + "\n" +
                          "Situação: Aceito. "
                          + "\n-------------------------------------------------------\n"
                          + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");
  }


  else if(permissao == ui.Button.NO){
    patrimonios = transpose(patrimonios);
    manageTrans(lista, 'deny trans', patrimonios,0);

    MailApp.sendEmail(emailRemetente,
                          "Transferência de Patrimônio - SiPat - Transferência Negada",
                          "Patrimônios: " + patStr + "\n" +
                          "Origem: " + origem + "\n" +
                          "Destino: " + destino + "\n" +
                          "Responsável do destino: " + emailDestinatario + "\n" +
                          "Situação: Negado.  Recupere seu patrimônio na Lista de Transferências em: " + ss.getUrl()
        + "\n-------------------------------------------------------\n"
        + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");

  }

  else if (permissao == 'cancel' || permissao == 0) {
    return
  }


  if(linhas.length>0 && permissao == "yes"){
    autoDoc(sheetOrigem, sheetDestino, patrimonios); //Relatório automático da transferência.
  }
}

