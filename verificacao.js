// exporta toda a planilha para a verificação
function exportToCheck(){

  var Ssheet = SpreadsheetApp.getActiveSpreadsheet();
  var year = new Date().getFullYear();
  var newFolder = DriveApp.createFolder("Planilhas de verificação de patrimônio de " + String(year));

  var salaByUser = spacesByUser();
  var esqueleto = SpreadsheetApp.openById("1EHtJiWkSld23MZX-NYlNsXs_sNPp91Uq4uRMhXnIQ_k");

  for(let i of salaByUser){
	var copy = DriveApp.getFileById(esqueleto.getId())
	.makeCopy("Planilha de verificação - " + i[1],newFolder);
    var copySheet = SpreadsheetApp.openById(copy.getId());
	i.splice(2,0,copySheet);

  }

  for(let i of salaByUser){
    var templateSheet = i[2].getSheetByName('Patrimônios');
    for(let [value,index] of i[3].entries()){
      var aux = copySheet.insertSheet(value, index+4, {template: templateSheet});
      createSpacesAndPutPatris(aux,mySheets[i]);
    }
    copySheet.deleteSheet(templateSheet);
  }
}

//
function spacesByUser(){
  var Ssheet = SpreadsheetApp.getActiveSpreadsheet();
  var respSheet = Ssheet.getSheetByName("Dados dos responsáveis");
  var lastRow = respSheet.getLastRow();
  var dados = respSheet.getRange(2,4,lastRow-1,3).getValues();

  var resps = Array();
  var salaByResp = Array();
  var nome = Array();
  for(let i of dados){
	let pos = resps.indexOf(i[2]);
	if(pos == -1){
		resps.push(i[2]);
		nome.push(i[1]);
		salaByResp.push([i[0]]);
	}
	else{
		salaByResp[pos].push(i[0]);
	}
  }
  return transpose([resps, nome, salaByResp]);
}

//
function createSpacesAndPutPatris(copySheet, sipatSheet){
	
	var sheetLR = lastRow(sipatSheet);
	var indices = indexColums(sipatSheet);
  if(sheetLR == 1){
    copySheet.insertRowsAfter(3,sheetLR+6);
	  copySheet.getRange(3,1,sheetLR+7,9).setBorder(true,true,true,true,true,true);
    return
  }
	var numPatri = sipatSheet.getRange(2,indices[0],sheetLR-1,1).getValues();
	var atm = sipatSheet.getRange(2,indices[1],sheetLR-1,1).getValues();
	var descricao = sipatSheet.getRange(2,indices[2],sheetLR-1,1).getValues();

	copySheet.insertRowsAfter(3,sheetLR+6);
	copySheet.getRange(3,1,sheetLR-1,1).setValues(numPatri);
	copySheet.getRange(3,2,sheetLR-1,1).setValues(atm);
	copySheet.getRange(3,3,sheetLR-1,1).setValues(descricao);
	copySheet.getRange(3,1,sheetLR+7,9).setBorder(true,true,true,true,true,true);
}
