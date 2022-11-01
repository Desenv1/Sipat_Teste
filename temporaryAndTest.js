

//Função de adicionar usuário.
function adicionarVariasSalas(){
  var emailUser = Session.getActiveUser().getEmail(); //Pega o e-mail do usuário.
  if (emailUser != 'desenv1@teiacoltec.org' && emailUser != 'patrimonio@teiacoltec.org' && emailUser != 'wagnercoltec@teiacoltec.org') {
    SpreadsheetApp.getUi().alert('Você não está autorizado a realizar esse tipo de operação!');  //Se o usuário não tiver a permissão, ele não conseguirá adicionar usuário novo.
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDados = ss.getSheetByName('Dados dos responsáveis');  //Planilha de dados dos responsáveis.
  var dadosLR = sheetDados.getLastRow(); //Última linha da planilha de dados.
  
  var nss = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1yZ2eGMOIwLkTt0iXQVfFntzmEJ_6JTHNGOv48vztrok/edit#gid=0');
  var nss_news = nss.getSheetByName("CRIAR");
  var vals = nss_news.getRange(3,2,28,7).getValues();


  for(var i = 0; i < vals.length; i++, dadosLR++){
    var nomeCadastro = vals[i][5];
    var emailCadastro = vals[i][6];
    var aliasCadastro = vals[i][3];
    var salaCadastro = vals[i][4];
    var setorCadastro = vals[i][0];
    var orgaoCadastro = vals[i][1];
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
      SpreadsheetApp.getUi().alert(salaCadastro + ' criada! Usuário '+ emailCadastro + ' cadastrado na(o) '+ salaCadastro + ' com sucesso!'); //Aviso confirmando o novo cadastro.
      /*MailApp.sendEmail(emailCadastro,
                        "SiPat - Responsabilidade de Salas",
                        "O patrimônio da sala " + salaCadastro + " foi atribuído à sua responsabilidade."
                        + "\n-------------------------------------------------------\n"
                        + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  */
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
        SpreadsheetApp.getUi().alert('Sala já existente');
        /*sheetDados.getRange(dadosLR+1,colOrgao).setValue(orgaoCadastro);
        sheetDados.getRange(dadosLR+1,colSetor).setValue(setorCadastro);
        sheetDados.getRange(dadosLR+1,colAlias).setValue(aliasCadastro);
        sheetDados.getRange(dadosLR+1,colSala).setValue(salaCadastro);
        sheetDados.getRange(dadosLR+1,colNome).setValue(nomeCadastro);
        sheetDados.getRange(dadosLR+1,colEmail).setValue(emailCadastro);
        protectionSala.addEditor(emailCadastro);
        SpreadsheetApp.getUi().alert('Usuário '+ emailCadastro + ' cadastrado na(o) '+ salaCadastro + ' com sucesso!');
        /*MailApp.sendEmail(emailCadastro,
                          "SiPat - Responsabilidade de Salas",
                          "O patrimônio da sala "+ salaCadastro + " foi atribuído à sua responsabilidade."
                          + "\n-------------------------------------------------------\n"
                          + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  */
      }
      //Todos dados adicionados ou atualizados e e-mail de confirmação enviado
    }
}
  sortSheets();
  var rangeTotal = sheetDados.getRange("A2:F");
  rangeTotal.sort(SORT_ORDER2); // Salas colocadas em ordem alfabética.
}



function myFunction() {
	// Display a modeless dialog box with custom HtmlService content.
	var htmlOutput = HtmlService.createHtmlOutputFromFile("interface.html");
	SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'My add-on');
  }

  function createSala(salaCadastro) {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var lista = ss.getSheetByName("Lista de transferências"); //Vetor de todas as páginas.
	var titleRange = lista.getRange('C15');
	var code = titleRange.getBackground();
	var finish = "finished";
  
  }

function moverHistoricoTransferencia() {

	var sipatSheet = SpreadsheetApp.getActiveSpreadsheet();
	var lista = sipatSheet.getSheetByName("Lista de transferências");
	var listaLR = lista.getLastRow();
	var plaquetas = lista.getRange(2,1,listaLR-1,1).getValues();
	for(let i = 0; i < listaLR-1; i++){
	  if(plaquetas[i][0] == 0){
		listaLR = i+1;
		break;
	  }
	}
  
	var historySheet = sipatSheet.getSheetByName("Histórico de Transferência");
	var historyLR= lastRow(historySheet);
  
	var dadosTransferencia = lista.getRange(2,1,listaLR-1,13).getValues();
	var transferidos = Array();
	for(let i = 0, aux = 0; i < dadosTransferencia.length; i++){
	  if (dadosTransferencia[i][9] == "Autorizada" || dadosTransferencia[i][9] == "Transferência cancelada" || dadosTransferencia[i][9] == "Retomado"){
		transferidos.push(dadosTransferencia[i]);
		lista.deleteRow(i+2-aux);
		aux++;
	  }
	}
  
	var aux4 = historySheet.getRange(historyLR+1,1,transferidos.length,13).setValues(transferidos);
  }