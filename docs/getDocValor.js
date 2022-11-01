//Documento filtrado por valor de parimônio.
function getDocValor(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  //Escolha para o usuário de apenas uma sala ou todas.
  var salaSelecionada = Browser.inputBox('Gerar Lista de Itens por Valor','Digite o código da sala que deseja filtrar os itens por valor. '+
                                         'Se deseja filtrar para todas as salas, digite Todas', Browser.Buttons.OK_CANCEL);
  if (salaSelecionada == 'cancel') return;
  //Caso seja escolhida uma sala e ela não exista, o usuário é alertado.
  else if (salaSelecionada != 'Todas') {
    salaSelecionada = findSala(salaSelecionada);
    if (salaSelecionada == null) {
      Browser.msgBox('Sala inexistente. Por favor, tente novamente.', Browser.Buttons.OK);
      return
    }
  }
  
  var valorMin = Browser.inputBox('Gerar Lista de Itens por Valor','Digite o valor mínimo para o filtro.', Browser.Buttons.OK_CANCEL); //Usuário é solicitado para digitar o valor mínimo para o filtro.
  if (valorMin == 'cancel') return;
  
  var valorMax = Browser.inputBox('Gerar Lista de Itens por Valor','Digite o valor máximo para o filtro.', Browser.Buttons.OK_CANCEL); //Usuário é solicitado para digitar o valor máximo para o filtro.
  if (valorMax == 'cancel') return;
  //Usuário é avisado caso houver erro de lógica nos valores.
  if (valorMin > valorMax) {
    SpreadsheetApp.getUi().alert('Valor mínimo maior que o valor máximo. Por favor, repita o procedimento.');
    return 
  }
  
  var ultimaLinha = lastRowMestre(); //Última linha escrita da planilha mestre.
  
    // Definir estrutura do arquivo.  
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1KAMBD5ZdgWBQnjosOyp7GhqUFrXCwXGHF8oBFiI0TYs/edit');
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
//Preenchimento do cabeçalho.
  cabecalho.getCell(0,1).setText(now).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
  cabecalho.getCell(1,1).setText(valorMin).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
  cabecalho.getCell(2,1).setText(valorMax).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
  cabecalho.getCell(3,1).setText(salaSelecionada).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(10).setBold(false);
 
  var k = 0;
  //Loop que varre os patrimônios.
  for (var i = 2; i <= ultimaLinha; i++) {
    //Dados dos patrimônios obtidos.
    var plaqueta = mestre.getRange(i,1).getValue();
    var patrimonio = mestre.getRange(i,2).getValue();
    var descricaoPat = mestre.getRange(i,3).getValue();
    
    var valorPat = mestre.getRange(i,6).getValue();
    if (valorPat == null) valorPat = 0;
    var salaAtual = mestre.getRange(i,8).getValue(); //Sala atual do patrimônio.
    //Caso haja uma sala específica, apenas os patrimônios daquela determinada sala dentro dos valores solicitados serão inseridos no documento.
    if (salaSelecionada != 'Todas') {
      if (salaSelecionada == salaAtual && valorPat >= valorMin && valorPat <= valorMax) {
        move.getCell(k+2,0).setText(plaqueta).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,1).setText(patrimonio).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,2).setText(descricaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,3).setText(valorPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,4).setText(salaAtual).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        k = k+1;
      }
      else continue;
    }
    //Caso não haja sala específica, todos patrimônios dentro desse valor são inseridos no documento.
    else {
      if (valorPat >= valorMin && valorPat <= valorMax) {
        cabecalho.getCell(0,1).setText(now);
        move.getCell(k+2,0).setText(plaqueta).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,1).setText(patrimonio).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,2).setText(descricaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,3).setText(valorPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(k+2,4).setText(salaAtual).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        k= k+1;
      }
      else continue;
    }
  }
  
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  doc.setName('SEI - Lista de Itens por Valor - Sala: ' + salaSelecionada + '. Valor: ' + valorMin + ' a ' + valorMax + ' reais.'); //Nome do documento.
  doc.saveAndClose();
  var blob = doc.getAs('application/pdf');
  DriveApp.createFile(blob); //Documento é convertido em pdf.
  //E-mail de confirmação.
  MailApp.sendEmail(emailUser,
                    "Lista de Itens Por Valor - SiPat",
                    "Sala: " + salaSelecionada + "\n" +
                    "Lista de itens por valor gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  

  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1oSuV2X5QifI7vyJaQdyH47sPuG-aq0xz4WpIsyADhCw/edit');
  doc.setName('Documento Modelo - SEI - Lista de Itens por Valor');
  doc.saveAndClose();
  
//  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1BijHKfhv6A6vuUdulR81Yr-WrVxc5bHT8-fMFeQrUkk/edit');
SpreadsheetApp.getUi().alert('Relatório gerado com sucesso! Confira sua caixa de e-mail!'); //Alerta de confirmação.
  
}