function removerCopia() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var mestreLC = mestre.getLastColumn(); //Última coluna da planilha mestre.
  var plaquetas = mestre.getRange("A:A");
  // Verifica qual é a última linha com algo escrito.
  var mestreLR = lastRowMestre();
  for(var j=2; j <= mestreLR; j++){
    var plaqueta = mestre.getRange(j,1).getValue();
    var procurar = plaquetas.createTextFinder(plaqueta); // Cria uma busca da plaqueta.
    var cellPlaqueta = procurar.findNext();
    
    if(cellPlaqueta == null) continue;
   
     else {
       mestre.deleteRow(j)
     }
  }
}
