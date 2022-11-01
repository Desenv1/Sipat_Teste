//Função de criar o documento de transferência automaticamente.
function autoDoc(origem, destino, patrimonios) {
  var salaDestino = destino.getSheetName(); //Nome da sala destino.
  var salaOrigem = origem.getSheetName(); //Nome da sala origem.
  
  var opcao = patrimonios[0][10]; //Opção de transferência.
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  var devolucao = patrimonios[0][11]; //Data de devolução (caso seja empréstimo).
  
  var emailDestino = getEmailResponsavel(salaDestino); //E-mail do usuário da sala destino.
  
  var nomeResponsavelDestino = getNomeResponsavel(salaDestino); //Nome do usuário da sala destino.
  var nomeResponsavelOrigem = getNomeResponsavel(salaOrigem); //Nome do usuário da sala de origem.
  var setorDestino = getSetorResponsavel(salaDestino); //Setor da sala destino.
  var setorOrigem = getSetorResponsavel(salaOrigem); //Setor da sala origem.
  
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
  for(var j = 0; j< patrimonios.length; j++){
    var pat = patrimonios[j][1];
    plaqueta[j] = patrimonios[j][0];
    var descricao = patrimonios[j][2];
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
                    "Nota de Movimentação de Material Permanente gerada com sucesso!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  //E-mail confirmando a transferência.


var blob = doc.getAs('application/pdf');
var aux = DriveApp.createFile(blob); //Cria um pdf do documento e salva no drive do usuário.

  MailApp.sendEmail("desenv1@teiacoltec.org",
                    "Transferência de Patrimônio - SiPat - Nota de Movimentação de Material Permanente",
                    "Patrimônios: " + plaqueta + "\n" +
                    "Origem: " + salaOrigem + "\n" +
                    "Destino: " + salaDestino + "\n" +
                    "Responsável da destino: " + emailDestino + "\n"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});  //E-mail confirmando a transferência.

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
SpreadsheetApp.getUi().alert('Transferência realizada! Relatório gerado com sucesso! Confira sua caixa de e-mail!');  //Alerta o usuário do sucesso da criação do documento.
}

