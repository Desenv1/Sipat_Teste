//Função para criar o documento por setor.
function getDocPatSetor(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lista = ss.getSheetByName('Lista de transferências'); //Planilha da lista de transferências.
  var dados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');
  var setorSelecionado = Browser.inputBox('Gerar Lista de Itens por Setor','Digite o nome do setor que deseja gerar a lista de patrimônio.', Browser.Buttons.OK_CANCEL); //É solicitado para o usuário o setor que deseja.
  if (setorSelecionado == 'cancel') return;
  var checkSetor = findSetor(setorSelecionado);
  //Se o setor não existe, o usuário é avisado.
  if (checkSetor == null) {
    SpreadsheetApp.getUi().alert('Setor não encontrado. Por favor, tente novamente.');
    return
  }
  var ultimaLinhaDados = dados.getLastRow(); //Última linha escrita dos dados.
  
  
    // Definir estrutura do arquivo.  
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1DcGlFJiBSK2ecmGE8B5mZ1UuD4rHzqs4AXAxh6UhUEc/edit');
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
  discriminacao.getCell(0,1).setText('-').setFontSize(10).setBold(false);
  discriminacao.getCell(1,1).setText(setorSelecionado).setFontSize(10).setBold(false);
  discriminacao.getCell(2,1).setText('-').setFontSize(10).setBold(false);
  discriminacao.getCell(3,1).setText('-').setFontSize(10).setBold(false);
  
  var k = 0;
  var numPatrimonios = 0;
  
  //Loop que varre as salas da planilha de dados dos responsáveis.
  for (var j = 2; j <= ultimaLinhaDados; j++) {
    
    var setor = dados.getRange(j, colSetor).getValue(); //Obtém o setor de determinada sala.
    // Checa se o setor é o mesmo que o solicitado.
    if (setor == setorSelecionado) {
      var salaSelecionadaName = dados.getRange(j, colSala).getValue(); //Nome da sala.
      var salaSelecionadaSheet = ss.getSheetByName(salaSelecionadaName); // Planilha da sala.
      var ultimaLinha = salaSelecionadaSheet.getLastRow(); //Última linha escrita da sala
      var indices = indexColums(salaSelecionadaSheet);
      indices.splice(indices.length - 5, 5);
      //Loop que varre os patrimônios da sala.
      for (var i = 2; i <= ultimaLinha; i++) {
        //Obtém os dados dos patrimônios e insere no documento.
        var valPat = getValuesByIndex(salaSelecionadaSheet,indices,i);
        var plaqueta = valPat[0];
        var patrimonio = valPat[1];
        var descricaoPat = valPat[2];
        var conservacaoPat = valPat[3];
        var usoPat = valPat[4];
        var valorPat = valPat[5];
       
        move.getCell(i+numPatrimonios-1,0).setText(plaqueta).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(i+numPatrimonios-1,1).setText(patrimonio).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(i+numPatrimonios-1,2).setText(descricaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(i+numPatrimonios-1,4).setText(conservacaoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(i+numPatrimonios-1,5).setText(valorPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(i+numPatrimonios-1,6).setText(usoPat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
        move.getCell(i+numPatrimonios-1,7).setText(salaSelecionadaName).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
      }
      numPatrimonios = numPatrimonios + ultimaLinha - 1; //Obtém o número de patrimônios.
    }
  }
  
  numTotal.getCell(0,1).setText(numPatrimonios); //Insere o número de patrimônios.
  
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  doc.setName('SEI - Lista de Patrimônio - Setor: ' + setorSelecionado); //Nome do documento.
  doc.saveAndClose();
  var blob = doc.getAs('application/pdf');
  DriveApp.createFile(blob); //Converte o documento para pdf e coloca no drive do usuário.
  //E-mail confirmando e enviando o relatório.
  MailApp.sendEmail(emailUser,
                    "Lista de Patrimônio - SiPat - Setor " + setorSelecionado,
                    "Setor: " + setorSelecionado + "\n" +
                    "Lista de Patrimônio gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  
  
  
//  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1BijHKfhv6A6vuUdulR81Yr-WrVxc5bHT8-fMFeQrUkk/edit');
SpreadsheetApp.getUi().alert('Relatório gerado com sucesso! Confira sua caixa de e-mail!'); //Alerta confirmando o sucesso da ação.
  
}