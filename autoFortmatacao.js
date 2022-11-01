function getRangeOfColumn(numberColumn){
  var abcdario = "ABCDEFGHIJKLMNOPQRSTUVWXZ";
  var column = String(abcdario.charAt(numberColumn-1));
  column = column + "2:" + column;
  return column;
}


function autoFortmatacao(sheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var vep = ss.getSheetByName("Valores estados dos patrimonios");
  var aux = getRangeOfColumn(sheet.getLastColumn());
  aux = aux.substring(3,4);
  var allTable = sheet.getRange("A2:" + aux);
  var indices = indexColums(sheet);

  var valValidations = Array();
  var rule = Array();
  valValidations.push(vep.getRange(2,1,3,1));
  valValidations.push(vep.getRange(2,2,2,1));
  rule.push(SpreadsheetApp.newDataValidation().
                           requireValueInRange(valValidations[0]).
                           setAllowInvalid(false).
                           build());
  rule.push(SpreadsheetApp.newDataValidation().
                           requireValueInRange(valValidations[1]).
                           setAllowInvalid(false).
                           build());
  
  allTable.setFontSize(11);
  allTable.setFontFamily("Arial");
  allTable.setBorder(true,true,true,true,true,true);
  neColumn = getColumnNE(sheet);

  var patrimonio   = sheet.getRange(getRangeOfColumn(indices[0]));
  var atm          = sheet.getRange(getRangeOfColumn(indices[1]));
  var descricao    = sheet.getRange(getRangeOfColumn(indices[2]));
  var conservacao  = sheet.getRange(getRangeOfColumn(indices[3]));
  var situacao     = sheet.getRange(getRangeOfColumn(indices[4]));
  var valor        = sheet.getRange(getRangeOfColumn(indices[5]));
  var ne           = sheet.getRange(getRangeOfColumn(neColumn));
  
  var ultimaTrans  = sheet.getRange(getRangeOfColumn(indices[6]));
  var salaAnteior  = sheet.getRange(getRangeOfColumn(indices[7]));
  var finalidade   = sheet.getRange(getRangeOfColumn(indices[8]));
  var devolucao    = sheet.getRange(getRangeOfColumn(indices[9]));
  var sincronizado = sheet.getRange(getRangeOfColumn(indices[10]));

  var debug = getRangeOfColumn(indices[2]);
  
  sheet.setColumnWidth(indices[0],120);
  sheet.setColumnWidth(indices[1],130);
  sheet.setColumnWidth(indices[2],400);
  sheet.setColumnWidth(indices[3],95);
  sheet.setColumnWidth(indices[4],105);
  sheet.setColumnWidth(indices[5],120);
  sheet.setColumnWidth(neColumn,  100);

  sheet.setColumnWidth(indices[6],160);
  sheet.setColumnWidth(indices[7],145);
  sheet.setColumnWidth(indices[8],120);
  sheet.setColumnWidth(indices[9],160);
  sheet.setColumnWidth(indices[10],90);

  patrimonio.setHorizontalAlignment("center");
  patrimonio.setVerticalAlignment("middle");

  atm.setHorizontalAlignment("center");
  atm.setVerticalAlignment("middle");

  descricao.setHorizontalAlignment("left");
  descricao.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  conservacao.setHorizontalAlignment("center");
  conservacao.setVerticalAlignment("middle");

  situacao.setHorizontalAlignment("center");
  situacao.setVerticalAlignment("middle");
  situacao.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  valor.setHorizontalAlignment("center");
  valor.setVerticalAlignment("middle");

  ne.setHorizontalAlignment("center")
  ne.setVerticalAlignment("middle")
  var validation = ne.getDataValidations();
  if(!(validation[0][0] != null // verificando se a coluna tem checkbox
        && validation[0][0].getCriteriaType() === SpreadsheetApp.DataValidationCriteria.CHECKBOX))
  {
    ne.insertCheckboxes();
  }

  ultimaTrans.setHorizontalAlignment("center");
  ultimaTrans.setVerticalAlignment("bottom");

  salaAnteior.setHorizontalAlignment("center");
  salaAnteior.setVerticalAlignment("bottom");
  salaAnteior.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  finalidade.setHorizontalAlignment("center");
  finalidade.setVerticalAlignment("bottom");

  devolucao.setHorizontalAlignment("center");
  devolucao.setVerticalAlignment("bottom");

  sincronizado.setHorizontalAlignment("center");
  sincronizado.setVerticalAlignment("bottom");

  conservacao.setDataValidation(rule[0]);
  situacao.setDataValidation(rule[1]);
  
  var lastColumn = Math.max.apply(null, indices);

  var title = sheet.getRange(1,1,1,lastColumn);
  title.setHorizontalAlignment("center");
  title.setVerticalAlignment("middle");
  title.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  title.setBorder(true,true,true,true,true,true);
  title.setFontSize(10);
}



