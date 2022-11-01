
//Função que organiza as salas em ordem alfabética.
function sortSheets () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetNameArray = []; 
  var sheets = ss.getSheets(); //Vetor de planilhas.

  ss.setActiveSheet(ss.getSheetByName("Mestre"));
  ss.moveActiveSheet(1);
  ss.setActiveSheet(ss.getSheetByName("Dados dos responsáveis"));
  ss.moveActiveSheet(2);
  ss.setActiveSheet(ss.getSheetByName("Lista de transferências"));
  ss.moveActiveSheet(3);
  ss.setActiveSheet(ss.getSheetByName("Lista de desfazimento"));
  ss.moveActiveSheet(4); 
  ss.setActiveSheet(ss.getSheetByName("Histórico de Transferência"));
  ss.moveActiveSheet(sheets.length);
  ss.setActiveSheet(ss.getSheetByName("Valores estados dos patrimonios"));
  ss.moveActiveSheet(sheets.length); 

  

   //Loop que forma um vetor com os nomes das salas.
  for (var i = 0; i < sheets.length; i++) {
    sheetNameArray.push(sheets[i].getName());
  }
  
  sheetNameArray = sheetNameArray.filter((!reservado));

  sheetNameArray.sort(); //Vetor é colocado em ordem alfabética.
  //Loop que move as planilhas para ordem alfabética.
  for( var j = 0; j < sheetNameArray.length; j++ ) {
    var nameSheet = sheetNameArray[j];
    ss.setActiveSheet(ss.getSheetByName(nameSheet));
    ss.moveActiveSheet(j+5);
  }
  
  Browser.msgBox("Salas organizadas em ordem alfabética! Obrigado!"); //Alerta de conclusão da ação.
}
