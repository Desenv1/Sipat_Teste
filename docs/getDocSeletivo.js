//Documento de transferências filtradas por data.
function getDocSeletivo(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lista = ss.getSheetByName('Lista de transferências'); //Planilha da lista de transferências.
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  var dataInicio = Browser.inputBox('Gerar Nota de Movimentação', 'Digite a data inicial para as transferências.', Browser.Buttons.OK_CANCEL); //Usuário é solicitado para digitar a data inicial para o filtro.
  if (dataInicio == 'cancel' || dataInicio == 0) return;
  var dataInicio = new Date(dataInicio);
  var dataInicioCont = dataInicio.getTime();
  var dataFim = Browser.inputBox('Gerar Nota de Movimentação', 'Digite a data final para as transferências.', Browser.Buttons.OK_CANCEL); //Usuário é solicitado para digitar a data final para o filtro.
  if (dataFim == 'cancel' || dataFim == 0) return;
  var dataFim = new Date(dataFim);
  var dataFimCont = dataFim.getTime();
  //Aviso caso ocorra erro de lógica entre as datas.
  if (dataFimCont - dataInicioCont < 0) {
    Browser.msgBox('Data de início mais recente que a data final. Por favor, tente novamente.', Browser.Buttons.OK);
    return
  }
  //Escolha para o usuário de apenas uma sala ou todas.
  var salaSelecionada = Browser.inputBox('Gerar Nota de Movimentação','Digite o código da sala que deseja verificar as transferências. '+
                                         'Se deseja ver as movimentações de um período para todas as salas, digite Todas', Browser.Buttons.OK_CANCEL);
  if (salaSelecionada == 'cancel') return;
  //Caso seja uma sala específica.
  else if (salaSelecionada != 'Todas') {
    salaSelecionada = findSala(salaSelecionada); //Procura a sala.
    //Se ela não existir o usuário é avisado.
    if (salaSelecionada == null) {
      Browser.msgBox('Sala inexistente. Por favor, tente novamente.', Browser.Buttons.OK);
      return
    }
  }
  
  var ultimaLinha = lista.getLastRow(); //Última linha da lista de transferências.
  
    // Definir estrutura do arquivo.  
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1oSuV2X5QifI7vyJaQdyH47sPuG-aq0xz4WpIsyADhCw/edit');
  var body = doc.getBody();
  var tabelas = body.getTables();
  var cabecalho = tabelas[0];
  var move = tabelas[1];
  var numLinhasMove = 120;

  // Garante a limpeza do documento antes de colocar os dados

  for (var j = 2; j < numLinhasMove; j++) {
    move.getCell(j,0).clear();
    move.getCell(j,1).clear(); 
    move.getCell(j,2).clear(); 
    move.getCell(j,3).clear(); 
    move.getCell(j,4).clear(); 
  }

  var dataInicioDoc = new Date(dataInicio).toLocaleString('pt-BR'); //Data para o documento.
  var dataFimDoc = new Date(dataFim).toLocaleString('pt-BR');
  //Dados inseridos no cabeçalho do documento.
  cabecalho.getCell(0,1).setText(now).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
  cabecalho.getCell(1,1).setText(dataInicioDoc).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
  cabecalho.getCell(2,1).setText(dataFimDoc).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
  cabecalho.getCell(3,1).setText(salaSelecionada).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
 
  var k = 0;
  //Loop que varre todas as transferências.
  for (var i = 2; i <= ultimaLinha; i++) {
    var dataPatrimonio = lista.getRange(i,7).getValue(); //Obtém a data.
    var dataPatrimonioCont = new Date(dataPatrimonio).getTime();
    //Se estiver fora do período desejado, passa para a próxima.
    if (dataPatrimonioCont - dataFimCont > 0) {
      break;
    }
    else {    
      var salaPatrimonioOrigem = lista.getRange(i,8).getValue(); //Sala de origem do patrimônio.
      var salaPatrimonioDestino = lista.getRange(i,9).getValue(); //Sala destino do patrimônio.
      var nomeResponsavelDestino = getNomeResponsavel(salaPatrimonioDestino); //Nome do responsável pela sala destino.
      //Dados do patrimônio.
      var plaqueta = lista.getRange(i,1).getValue();
      var patrimonio = lista.getRange(i,2).getValue();
      var descricaoPat = lista.getRange(i,3).getValue();
      //Caso o usuário queira uma sala específica.
      if (salaSelecionada != 'Todas') {
        //Condição para confirmar que a sala participou da transfrência. Se sim, os dados são adicionados ao documento.
        if (salaSelecionada == salaPatrimonioOrigem || salaSelecionada == salaPatrimonioDestino) {
          move.getCell(k+2,0).setText(plaqueta).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
          move.getCell(k+2,1).setText(patrimonio).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
          move.getCell(k+2,2).setText(descricaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
          move.getCell(k+2,3).setText(salaPatrimonioOrigem).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
          move.getCell(k+2,4).setText(salaPatrimonioDestino).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
          k = k+1;
        }
        else continue;
      }
      //Caso o usuário não deseje uma sala específica, todas as transferências nesse período serão adicionadas ao documento.
      else {
        cabecalho.getCell(0,1).setText(now);
        move.getCell(k+2,0).setText(plaqueta).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,1).setText(patrimonio).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,2).setText(descricaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,3).setText(salaPatrimonioOrigem).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,4).setText(salaPatrimonioDestino).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        k= k+1;
      }
    }
  }

  
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  doc.setName('SEI - Lista de Transferências Sala: ' + salaSelecionada + '. Período: ' + dataInicio + ' a ' + dataFim); //Nome do documento.
  doc.saveAndClose();
  var blob = doc.getAs('application/pdf');
  DriveApp.createFile(blob); //COnversão do documento em pdf.
  //E-mail de confirmação.
  MailApp.sendEmail(emailUser,
                    "Lista de Transferências - SiPat",
                    "Sala: " + salaSelecionada + "\n" +
                    "Lista de Transferências gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  

  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1oSuV2X5QifI7vyJaQdyH47sPuG-aq0xz4WpIsyADhCw/edit');
  doc.setName('Documento Modelo - SEI - Lista de Transferências');
  doc.saveAndClose();
  
//  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1BijHKfhv6A6vuUdulR81Yr-WrVxc5bHT8-fMFeQrUkk/edit');
SpreadsheetApp.getUi().alert('Relatório gerado com sucesso! Confira sua caixa de e-mail!'); //Alerta de confirmação.
  
}