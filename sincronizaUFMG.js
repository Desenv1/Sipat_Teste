//Função que atualiza as sincronizações dos patrimônios da planilha mestre com a UFMG, compartilhando essa situação com as salas do patrimônio.
function sincronizaUFMG() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var mestreLC = mestre.getLastColumn(); //Última coluna escrita da planilha mestre.
  var mestreLR = lastRowMestre(); //Última linha escrita da planilha mestre.
  
  //Loop que varre os patrimônios da planilha mestre.
  for (var i = 2; i < mestreLR; i++) {
    var plaqueta = mestre.getRange(i,1).getValue(); //Plaqueta do patrimônio.
    var salaAtualPat = mestre.getRange(i,8).getValue(); //Sala atual do patrimônio.
    var sincronismo = mestre.getRange(i,10).getValue(); //Sincronismo do patrimônio.
    var sheet = ss.getSheetByName(salaAtualPat); //Planilha da sala do patrimônio.
    if (sheet == null) continue;
    var plaquetasSalaAtual = sheet.getRange(1,1,1000,1); //Coluna de plaquetas da sala atual.
    var procurar = plaquetasSalaAtual.createTextFinder(plaqueta); // Cria uma busca da plaqueta.
    var cellPlaqueta = procurar.findNext();
    if (cellPlaqueta == null) continue;
    var linhaPlaqueta = cellPlaqueta.getRow(); //Pega a linha do patrimônio
    sheet.getRange(linhaPlaqueta,colSincroniza).setValue(sincronismo); //Sincroniza o patrimônio de acordo com a planilha mestre.
  }
  
  Browser.msgBox('Aviso de sincronismo realizado com sucesso para os usuários! Obrigado!', Browser.Buttons.OK); //Alerta de confrimação.
}