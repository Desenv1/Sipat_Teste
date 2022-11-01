//Documento da lista de salas.
function getDocSalas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets(); //Vetor de salas.
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1GXwHNkSv0Xtwxmv2lOtHPKgcu0vAzN9j9MbnUHx0vZU/edit'); //Link do esqueleto do documento.
  //Partes do documento separadas para edição.
  var body = doc.getBody();
  var tabelas = body.getTables();
  var cabecalho = tabelas[0];
  var move = tabelas[1];
  var numLinhasMove = move.getNumRows();

  // Garante a limpeza do documento antes de colocar os dados
  for (var j = 1; j < numLinhasMove-1; j++) {
    move.getCell(j,0).clear();
    move.getCell(j,1).clear(); 
    move.getCell(j,2).clear();
    move.getCell(j,3).clear(); 
    move.getCell(j,4).clear(); 
  }
  
  cabecalho.getCell(0,1).setText(now).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER); //Adiciona a data atual.
  
  //Primeiro loop, percorrendo as páginas slaves da Spreadsheet.
  for(var i=5; i<sheets.length; i++){
    var slave = ss.getSheets()[i]; //Pega uma das páginas das salas.
    var nomeSala = slave.getSheetName(); //Nome da sala.
    var setorResponsavel = getSetorResponsavel(nomeSala); //Setor da sala.
    var orgaoResponsavel = getOrgaoResponsavel(nomeSala); //Órgão da sala.
    var nomeResponsavel = getNomeResponsavel(nomeSala); //Responsável pela sala.
    var emailResponsavel = getEmailResponsavel(nomeSala); //E-mail do responsável pela sala.
    //Dados obtidos são inseridos no documento.
    move.getCell(i-3,0).setText(nomeSala).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-3,1).setText(setorResponsavel).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-3,2).setText(orgaoResponsavel).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-3,3).setText(nomeResponsavel).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    move.getCell(i-3,4).setText(emailResponsavel).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
  }
  
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  doc.setName('Sipat - Lista de Salas - COLTEC'); //Nome do documento.
  doc.saveAndClose();
  var blob = doc.getAs('application/pdf');
  DriveApp.createFile(blob); //Converter para pdf e salvar no drive.
  //E-mail de confirmação.
  MailApp.sendEmail(emailUser,
                    "Sipat - Lista de Salas - COLTEC",
                    "Lista de Salas gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  
  
  Browser.msgBox('Lista de salas gerada com sucesso! Confira seu e-mail!', Browser.Buttons.OK); //Alerta de confirmação.
}