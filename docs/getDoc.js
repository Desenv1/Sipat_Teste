// Função para criar o documento em PDF de Nota de Movimentação de Material Permanente.
function getDoc(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetDestino = ss.getActiveSheet(); //Planilha da sala destino do patrimônio.
  var destinoName = sheetDestino.getName(); //Nome da sala destino.
  //O usuário é solicitado para digitar as linhas dos patrimônios que serão registrados no documento.
  var linhas = Browser.inputBox('Gerar Nota de Movimentação','Digite as linhas dos patrimônios que serão registrados no documento separadas por vírgula. ' +
                                'Selecione apenas patrimônios que serão transferidos para uma mesma sala ' +
                                'e com a mesma finalidade.', Browser.Buttons.OK_CANCEL)
  if (linhas == 'cancel' || linhas == 0 || linhas == 1) return; //Apenas a partir da linha 2.
  var linhasVector = getLines(linhas); //Vetor de linhas.
  var numPatrimonios = linhasVector.length; //Tamanho desse vetor.
  var sheetDados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var indices = indexColums(sheetDestino);

  var salaDestino = sheetDestino.getSheetName(); //Nome da sala destino.
  var salaOrigem = sheetDestino.getRange(linhasVector[linhasVector.length-1],indices[7]).getValue(); //Nome da sala de origem.
  var opcao = sheetDestino.getRange(linhasVector[linhasVector.length-1],indices[8]).getValue(); //Opção de transferência.
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  var devolucao = sheetDestino.getRange(linhasVector[linhasVector.length-1],indices[9]).getValue(); //Data de devolução, caso tenha uma.
  
  var emailDestino = getEmailResponsavel(salaDestino); //E-mail do responsável da sala destino.
  //Caso não tenha permissão, o usuário será avisado.
  if (emailDestino != emailUser) {
    var verificacao = Browser.msgBox('Gerar Nota de Movimentação', 'A sala que você está não é de sua responsabilidade. Deseja continuar para a geração do documento?',
                                     Browser.Buttons.YES_NO); 
    if (verificacao == 'no') {
      return 
    }
  }
  
  var nomeResponsavelDestino = getNomeResponsavel(salaDestino); //Nome do responsável da sala destino.
  var nomeResponsavelOrigem = getNomeResponsavel(salaOrigem); //Nome do responsável da sala de origem.
  var setorDestino = getSetorResponsavel(salaDestino); //Nome do setor da sala destino.
  var setorOrigem = getSetorResponsavel(salaOrigem); //Nome do setor da sala de origem.
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1EzMUMzET7LvfuUyDe7IP4E6faout-jkkR9zZRmD-_zs/edit'); //Link do esqueleto do documento de transferência.
  var body = doc.getBody(); //Corpo do documento.
  var tabelas = body.getTables();
  var finalidade = tabelas[0]; //Primeira parte do corpo do documento.
  var discriminacao = tabelas[1]; //Segunda parte do corpo do documento.
  var move = tabelas[2]; //Terceira parte do corpo do documento.
  
  // Garante a limpeza do documento antes de colocar os dados
  finalidade.getCell(1,0).clear(); // Limpa o campo de transferência
  for (var j = 2; j <= 12; j++) { // Limpa o campo dos patrimônios
    discriminacao.getCell(j,0).clear();
    discriminacao.getCell(j,1).clear();
    discriminacao.getCell(j,2).clear();
    discriminacao.getCell(j,3).clear();
    discriminacao.getCell(j,4).clear();
  }
  move.getCell(1,1).clear(); // Limpa a sala de origem
  move.getCell(1,3).clear(); // Limpa a sala de destino
  move.getCell(3,1).clear(); // Limpa o nome do responsável da origem
  move.getCell(3,3).clear(); // Limpa o nome do responsável do destino
  
  //Estruturas para cada finalidade.
  var transf = '(X) TRANSFERÊNCIA       ESTA NOTA SERÁ SUBSTITUÍDA PELO TERMO DE RESPONSABILIDADE - TR\n'+
    '(  ) EMPRÉSTIMO             DATA ESTIMADA DE DEVOLUÇÃO:\n'+
      '(  ) MANUTENÇÃO            DATA ESTIMADA DE DEVOLUÇÃO: ';
  var emprestimo = '(  ) TRANSFERÊNCIA       ESTA NOTA SERÁ SUBSTITUÍDA PELO TERMO DE RESPONSABILIDADE - TR\n'+
    '(X) EMPRÉSTIMO             DATA ESTIMADA DE DEVOLUÇÃO: ' + devolucao +'\n'+
      '(  ) MANUTENÇÃO            DATA ESTIMADA DE DEVOLUÇÃO: ';
  var manutencao = '(  ) TRANSFERÊNCIA       ESTA NOTA SERÁ SUBSTITUÍDA PELO TERMO DE RESPONSABILIDADE - TR\n'+
    '(  ) EMPRÉSTIMO             DATA ESTIMADA DE DEVOLUÇÃO:\n'+
      '(X) MANUTENÇÃO            DATA ESTIMADA DE DEVOLUÇÃO: ';
  var vazio = '(  ) TRANSFERÊNCIA       ESTA NOTA SERÁ SUBSTITUÍDA PELO TERMO DE RESPONSABILIDADE - TR\n'+
    '(  ) EMPRÉSTIMO             DATA ESTIMADA DE DEVOLUÇÃO:\n'+
      '(  ) MANUTENÇÃO            DATA ESTIMADA DE DEVOLUÇÃO: ';
  
  //Confere qual é a finalidade e se adapta à mesma.
  if (opcao == 'Transferência')finalidade.getCell(1,0).setText(transf).setFontSize(10).setBold(false);
  else if (opcao == 'Empréstimo')finalidade.getCell(1,0).setText(emprestimo).setFontSize(10).setBold(false);
  else if (opcao == 'Manutenção')finalidade.getCell(1,0).setText(manutencao).setFontSize(10).setBold(false);
  else finalidade.getCell(1,0).setText(vazio).setFontSize(10);
  
  var plaqueta = [];
  // Varre os patrimônios, pega seus dados e coloca na parte de discriminação do documento.
  for(var j = 0; j< linhasVector.length; j++){
    var pat = sheetDestino.getRange(linhasVector[j], indices[1]).getValue();
    plaqueta[j] = sheetDestino.getRange(linhasVector[j], indices[0]).getValue();
    var descricao = sheetDestino.getRange(linhasVector[j], indices[2]).getValue();
    discriminacao.getCell(j+2,0).setText(pat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    discriminacao.getCell(j+2,1).setText(plaqueta[j]).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    discriminacao.getCell(j+2,2).setText(descricao).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
  }
  
  //Os dados das salas são adicionadas à terceira parte do corpo do documento.
  move.getCell(1,1).setText(setorOrigem + ' - ' + salaOrigem).setFontSize(10);
  move.getCell(1,3).setText(setorDestino + ' - ' + salaDestino).setFontSize(10);
  move.getCell(2,1).setText('Colégio Técnico da UFMG').setFontSize(10);
  move.getCell(2,3).setText('Colégio Técnico da UFMG').setFontSize(10);
  move.getCell(3,1).setText(nomeResponsavelOrigem).setFontSize(10);
  move.getCell(3,3).setText(nomeResponsavelDestino).setFontSize(10);
  
  doc.setName('SEI - Movimentação ' + salaOrigem + ' para ' + salaDestino); //Nome do documento.
  doc.saveAndClose(); //Salva e fecha o documento.
  MailApp.sendEmail(emailUser,
                    "Transferência de Patrimônio - SiPat - Nota de Movimentação de Material Permanente",
                    "Patrimônios: " + plaqueta + "\n" +
                    "Origem: " + salaOrigem + "\n" +
                    "Destino: " + salaDestino + "\n" +
                    "Responsável da destino: " + emailDestino + "\n" +
                    "Nota de Movimentação de Material Permanente gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  //E-mail confirmando a transferência.
  
  
  var blob = doc.getAs('application/pdf');
  DriveApp.createFile(blob); //Cria um pdf do documento e salva no drive do usuário.
  
  
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1EzMUMzET7LvfuUyDe7IP4E6faout-jkkR9zZRmD-_zs/edit');
  var body = doc.getBody();
  var tabelas = body.getTables();
  var finalidade = tabelas[0];
  var discriminacao = tabelas[1];
  var move = tabelas[2];
  
  // Garante a limpeza do documento depois de colocar os dados
  finalidade.getCell(1,0).clear(); // Limpa o campo de transferência
  for (var j = 2; j <= 12; j++) { // Limpa o campo dos patrimônios
    discriminacao.getCell(j,0).clear();
    discriminacao.getCell(j,1).clear();
    discriminacao.getCell(j,2).clear();
    discriminacao.getCell(j,3).clear();
    discriminacao.getCell(j,4).clear();
  }
  move.getCell(1,1).clear(); // Limpa a sala de origem
  move.getCell(1,3).clear(); // Limpa a sala de destino
  move.getCell(3,1).clear(); // Limpa o nome do responsável da origem
  move.getCell(3,3).clear(); // Limpa o nome do responsável do destino
  doc.setName('SEI - Arquivo Geral');
  doc.saveAndClose();
  SpreadsheetApp.getUi().alert('Relatório gerado com sucesso! Confira sua caixa de e-mail!'); //Alerta o usuário do sucesso da criação do documento.
  
}