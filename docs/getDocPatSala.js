//Função para criar o documento de uma determinada sala
function getDocPatSala(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lista = ss.getSheetByName('Lista de transferências'); //Planilha da lista de transferências.
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  var salaSelecionada = Browser.inputBox('Gerar Nota de Movimentação','Digite o código da sala que deseja gerar a lista de patrimônio.', Browser.Buttons.OK_CANCEL); //É solicitado para o usuário o código da sala.
  if (salaSelecionada == 'cancel') return;
  salaSelecionada = findSala(salaSelecionada); //Procura essa sala.
  var salaSelecionadaSheet = ss.getSheetByName(salaSelecionada); //Planilha dessa sala.
  //Se a sala não existir, o usuário será avisado.
  if (salaSelecionadaSheet == null) {
    SpreadsheetApp.getUi().alert('Sala inexistente! Por favor, tente novamente!', Browser.Buttons.OK);
    return;
  }
  var ultimaLinha = salaSelecionadaSheet.getLastRow(); //Última linha escrita da sala.
  var nomeResponsavel = getNomeResponsavel(salaSelecionada); //Nome do responsável pela sala.
  var setorResponsavel = getSetorResponsavel(salaSelecionada); //Setor da sala.
  var orgaoResponsavel = getOrgaoResponsavel(salaSelecionada); //Órgão da sala.
  
  
    // Definir estrutura do arquivo.  
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1BijHKfhv6A6vuUdulR81Yr-WrVxc5bHT8-fMFeQrUkk/edit');
  var body = doc.getBody();
  var tabelas = body.getTables();
  var cabecalho = tabelas[0];
  var discriminacao = tabelas[1];
  var move = tabelas[2];
  var numTotal = tabelas[3];
  var numLinhasMove = move.getNumRows();

  // Garante a limpeza do documento antes de colocar os dados
  for (var j = 0; j <= 3; j++) { // Limpa o campo de discriminação de identificação
    discriminacao.getCell(j,1).clear();
  }
  for (var j = 1; j < numLinhasMove-1; j++) {
    move.getCell(j,0).clear();
    move.getCell(j,1).clear(); 
    move.getCell(j,2).clear(); 
    move.getCell(j,3).clear(); 
    move.getCell(j,4).clear(); 
    move.getCell(j,5).clear(); 
    move.getCell(j,6).clear(); 
  }
  
  cabecalho.getCell(0,1).setText(now).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
  discriminacao.getCell(0,1).setText(orgaoResponsavel).setFontSize(10).setBold(false);
  discriminacao.getCell(1,1).setText(setorResponsavel).setFontSize(10).setBold(false);
  discriminacao.getCell(2,1).setText(salaSelecionada).setFontSize(10).setBold(false);
  discriminacao.getCell(3,1).setText(nomeResponsavel).setFontSize(10).setBold(false);
  
  var k = 0;
  var indices = indexColums(salaSelecionadaSheet);
  //Varre os patrimônios, pegando seus dados e inserindo-os no documento.
  for (var i = 2; i <= ultimaLinha; i++) {
    var plaqueta = salaSelecionadaSheet.getRange(i,indices[0]).getValue();
    var patrimonio = salaSelecionadaSheet.getRange(i,indices[1]).getValue();
    var descricaoPat = salaSelecionadaSheet.getRange(i,indices[2]).getValue();
    var conservacaoPat = salaSelecionadaSheet.getRange(i,indices[3]).getValue();
    var usoPat = salaSelecionadaSheet.getRange(i,indices[4]).getValue();
    var valorPat = salaSelecionadaSheet.getRange(i,indices[5]).getValue();
   
    move.getCell(i-1,0).setText(plaqueta).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-1,1).setText(patrimonio).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-1,2).setText(descricaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-1,4).setText(conservacaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-1,5).setText(valorPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-1,6).setText(usoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
  }
  
  numTotal.getCell(0,1).setText(ultimaLinha-1);
  
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  doc.setName('SEI - Lista de Patrimônio - Sala: ' + salaSelecionada); //Nome do documento.
  doc.saveAndClose();
  var blob = doc.getAs('application/pdf'); //Arquivo é convertido para pdf para o drive do usuário.
  DriveApp.createFile(blob);
  MailApp.sendEmail(emailUser,
                    "Lista de Patrimônio - SiPat - Sala " + salaSelecionada,
                    "Sala: " + salaSelecionada + "\n" +
                    "Lista de Patrimônio gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  //E-mail contendo o link para o documento. 
  
  
//  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1BijHKfhv6A6vuUdulR81Yr-WrVxc5bHT8-fMFeQrUkk/edit');
SpreadsheetApp.getUi().alert('Relatório gerado com sucesso! Confira sua caixa de e-mail!'); //Aviso do sucesso do relatório gerado.
  
}