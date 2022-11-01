//Função para criar o documento de manutenção.
function getDocManut(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lista = ss.getActiveSheet(); //Planilha ativa 
  //Se não estiver na lista de transferências, a ação não pode ser concluída.
  if(lista != 'Lista de transferências'){
    SpreadsheetApp.getUi().alert('Página errada! Você deve realizar essa ação apenas na página "Lista de transferências".');
    return
  }
  //var destinoName = sheetDestino.getName();
  //O usuário é solicitado para digitar as linhas dos patrimônios que serão registrados no documento.
  var linhas = Browser.inputBox('Gerar Nota de Movimentação','Digite as linhas dos patrimônios que serão registrados no documento separadas por vírgula. ' +
                                'Selecione apenas patrimônios que serão transferidos para uma mesma sala ' +
                                'e para manutenção.', Browser.Buttons.OK_CANCEL)
  if (linhas == 'cancel' || linhas == 0 || linhas == 1) return; //Apenas a partir da linha 2.
  var linhasVector = linhas.split(','); //Vetor de linhas.
  var numPatrimonios = linhasVector.length; //Tamanho desse vetor.
  var sheetDados = ss.getSheetByName('Dados dos responsáveis'); //Planilha de dados dos responsáveis.
  var salaOrigem = lista.getRange(linhasVector[linhasVector.length-1],8).getValue(); //Nome da sala de origem.
  var salaDestino = lista.getRange(linhasVector[linhasVector.length-1],9).getValue(); //Nome da sala destino.
  var opcao = lista.getRange(linhasVector[linhasVector.length-1],11).getValue(); //Opção de transferência.
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  var devolucao = lista.getRange(linhasVector[linhasVector.length-1],12).getValue(); //Data de devolução.
  
  var nomeResponsavelOrigem = getNomeResponsavel(salaOrigem); //Nome do responsável da sala de origem.
  
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
  
//Estruturas para a finalidade de manutenção ou nenhuma.
  var manutencao = '(  ) TRANSFERÊNCIA       ESTA NOTA SERÁ SUBSTITUÍDA PELO TERMO DE RESPONSABILIDADE - TR\n'+
    '(  ) EMPRÉSTIMO             DATA ESTIMADA DE DEVOLUÇÃO:\n'+
      '(X) MANUTENÇÃO            DATA ESTIMADA DE DEVOLUÇÃO: '+devolucao;
  var vazio = '(  ) TRANSFERÊNCIA       ESTA NOTA SERÁ SUBSTITUÍDA PELO TERMO DE RESPONSABILIDADE - TR\n'+
    '(  ) EMPRÉSTIMO             DATA ESTIMADA DE DEVOLUÇÃO:\n'+
      '(  ) MANUTENÇÃO            DATA ESTIMADA DE DEVOLUÇÃO: ';
  
//Confere qual é a finalidade e se adapta à mesma.
  if (opcao == 'Manutenção')finalidade.getCell(1,0).setText(manutencao).setFontSize(10).setBold(false);
  else finalidade.getCell(1,0).setText(vazio).setFontSize(10).setBold(false);
  
  var plaqueta = [];
  // Varre os patrimônios, pega seus dados e coloca na parte de discriminação do documento.
  for(var j = 0; j< linhasVector.length; j++){
    var pat = lista.getRange(linhasVector[j], 2).getValue();
    plaqueta[j] = lista.getRange(linhasVector[j], 1).getValue();
    var descricao = lista.getRange(linhasVector[j], 3).getValue();
    discriminacao.getCell(j+2,0).setText(pat).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER).setFontSize(8).setBold(false);
    discriminacao.getCell(j+2,1).setText(plaqueta[j]).setFontSize(8).setBold(false);
    discriminacao.getCell(j+2,2).setText(descricao).setFontSize(8).setBold(false);
  }
  
  //Os dados das salas são adicionadas à terceira parte do corpo do documento.
  move.getCell(1,1).setText(salaOrigem).setFontSize(10);
  move.getCell(1,3).setText(salaDestino).setFontSize(10);
  move.getCell(2,1).setText('Colégio Técnico da UFMG').setFontSize(10).setBold(false);
  move.getCell(3,1).setText(nomeResponsavelOrigem).setFontSize(10).setBold(false);
  move.getCell(3,3).setText(nomeResponsavelOrigem).setFontSize(10).setBold(false);
  
  doc.setName('SEI - Movimentação ' + salaOrigem + ' para ' + salaDestino); //Nome do documento.
  doc.saveAndClose(); //Salva e fecha o documento.
  MailApp.sendEmail(emailUser,
                    "Transferência de Patrimônio - SiPat - Nota de Movimentação de Material Permanente",
                    "Patrimônios: " + plaqueta + "\n" +
                    "Origem: " + salaOrigem + "\n" +
                    "Destino: " + salaDestino + "\n" +
                    "Nota de Movimentação de Material gerada com sucesso! Confira o anexo!"
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec",{attachments: [doc.getAs(MimeType.PDF)]});   //E-mail confirmando o documento.
  
  
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