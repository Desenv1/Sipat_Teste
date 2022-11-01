//Função para criar sala chamada na função de Adicionar Usuário.
function createSala(salaCadastro) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets(); //Vetor de todas as páginas.
  var numSheets = sheets.length; //Tamanho desse vetor.
  var newSala = ss.insertSheet(salaCadastro,numSheets+1); //Adiciona a planilha da nova sala.
  newSala.deleteRows(200,800);
  newSala.deleteColumns(12,14);
  var titleRange = newSala.getRange('A1:L1');
  var leftTitleRange = newSala.getRange('A1:G1');
  var rightTitleRange = newSala.getRange('H1:L1');
  
  //Todas colunas são nomeadas e editadas conforme o padrão.
  var titleString = ['Nº Patrimônio','Nº ATM', 'Descrição', 'Conservação', 'Situação', 'Valor', 'Não Emplaquetado',
                     'Última Transferência', 'Sala Anterior', 'Finalidade', 'Data Estimada de Devolução', 'Sincronizado (UFMG)'];
  titleRange.setValues([titleString]);
  titleRange.setFontWeight("bold").setFontColor('black');
  leftTitleRange.setBackground('#03BB85');
  rightTitleRange.setBackground('#FF6961');
  autoFortmatacao(newSala);

}


