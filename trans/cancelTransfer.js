// Função de cancelar transferência.
function cancelTransfer(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var lista = ss.getActiveSheet(); //Planilha ativa.
  var listaNome = lista.getSheetName(); //Nome da página ativa.
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  var history = ss.getSheetByName('Histórico de Transferência'); //Planilha mestre.
  var historyLR = lastRow(history);
  
  var emailUser = Session.getActiveUser(); //E-mail do usuário.

  if(listaNome != 'Lista de transferências' && listaNome != 'Lista de desfazimento'){
    //Caso esteja na página errada, essa ação não é permitida.
    SpreadsheetApp.getUi().alert('Página errada! Você deve realizar essa ação apenas na página "Lista de transferências".'); 
    return;
  }

  var linha = Browser.inputBox('Cancelamento de Transferência - Patrimônio', 'Digite o(s) número(s) da(s) linha(s) do(s) patrimônio(s) que terá ' +
  'sua transferência cancelada. Caso seja mais de um, separe-os por vírgulas em ordem CRESCENTE.', Browser.Buttons.OK_CANCEL); //Usuário é solicitado para digitar as linhas dos patrimônios que terão sua transferência cancelada.
  if (linha == 'cancel') return
  var linhas = getLines(linha); //Vetor de linhas.
  var notRemoved = Array();

  var patris = manageTrans(lista,'get data', 0, linhas);


  for(var j = 0;j<patris.length;j++){
    var salaOrigem = patris[j][7]; //Sala de origem.
    var sheetOrigem = ss.getSheetByName(salaOrigem); //Planilha da sala de origem.
    var origemLR = lastRow(sheetOrigem); //Última linha escrita da sala de origem.
    var emailOrigem = getEmailResponsavel(salaOrigem); //E-mail do responsável da sala de origem.
    var plaqueta = patris[j][0]; //Plaqueta do patrimônio
    
    //Se a pessoa for autorizada a fazer essa função e o patrimônio estiver aguardando autorização.
    if (emailOrigem != emailUser) { 
      notRemoved.push(j);
      linhas.splice(j+1,1);
      Browser.msgBox('Não é possível cancelar a transferência do patrimônio ' + plaqueta + ' pois você não está autorizado. ' +
                      'Aperte OK para continuar ou CANCEL para interromper os próximos cancelamentos de transferências.', Browser.Buttons.OK_CANCEL);
      continue;
    }
    var opcao = patris[j][10]; //Opção de transferência.
    var destino = patris[j][8]; //Sala destino.
    var sheetDestino = ss.getSheetByName(destino); //Planilha da sala destino.

    var indices = indexColums(sheetOrigem);
    indices.splice(indices.length - 3, 3);
    var emailDestinatario = getEmailResponsavel(destino); //E-mail do responsável da sala destino.
    var valoresPat = patris[j].slice(0,8); //Dados do patrimônio.
	  valoresPat[6] = now;
    
    //Os dados dos patrimônios são devolvidos à planilha da sala origem.
    patris[j][6] = now; //A data é atualizada.
    patris[j][10] = '-';
    patris[j][11] = '-';
    setValueByIndex(sheetOrigem,indices,origemLR+1, valoresPat);
    
	patris[j].splice(9,0,'Transferência cancelada'); //A situação de transferência é atualizada para cancelada.
    
  }
  patris = patris.filter(function(value, index){
    return notRemoved.indexOf(index) == -1;
  });
  if (patris.length == 0){
    return
  }
  manageHistory(history,patris);
  plaquetas = (transpose(patris))[0];
  
  manageTrans(lista, 'remove lines', patris, 0);

  plaquetas = concatPlaqueta(plaquetas);
  MailApp.sendEmail(emailOrigem, emailDestinatario,
                      "Transferência de Patrimônio - SiPat - Transferência Cancelada",
                      "Patrimônio: " + plaquetas + "\n" +
                      "Origem: " + salaOrigem + "\n" +
                      "Destino: " + destino + "\n" +
                      "Responsável do destino: " + emailDestinatario + "\n" +
                      "Situação: Transferência cancelada."
      + "\n-------------------------------------------------------\n"
      + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");
    //Um E-mail é enviado confirmando a transferência.
}

function concatPlaqueta(plaquetas){
  var plaqString = String();
  for(let i = 0; i < plaquetas.length-1; i++){
    plaqString += plaquetas[i] + ", ";
  }
  plaqString += plaquetas.slice(-1);
}
