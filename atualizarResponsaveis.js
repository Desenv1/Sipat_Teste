
// Atualiza a lista de responsáveis
function atualizarResponsaveis() {
	var emailUser = Session.getActiveUser().getEmail(); //Pega o e-mail do usuário.
	if (emailUser != 'desenv1@teiacoltec.org' && emailUser != 'patrimonio@teiacoltec.org' && emailUser != 'wagnercoltec@teiacoltec.org') {
	  SpreadsheetApp.getUi().alert('Você não está autorizado a realizar esse tipo de operação!');  //Se o usuário não tiver a permissão, ele não conseguirá adicionar usuário novo.
	  return;
	}
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var dr = ss.getSheetByName("Dados dos responsáveis");
	var mestreP = ss.getSheetByName("Mestre");
	var mestreProtection = mestreP.protect();
	var mestreEditors = mestreProtection.getEditors();
	var lr = dr.getLastRow();
	var salas = dr.getRange(2, 4, lr - 1, 1).getValues();
	var email_responsaveis = dr.getRange(2, 6, lr - 1, 1).getValues();
  
	email_responsaveis = transpose(email_responsaveis);
	email_responsaveis = email_responsaveis[0];
	email_responsaveis = email_responsaveis.filter(function (el) {
	  return el != null;
	});
	salas = transpose(salas);
	salas = salas[0];
	salas = salas.filter(function (el) {
	  return el != null;
	});
	if(salas.length != email_responsaveis.length){
	  console.log(" tamanhos diferentes do vetor de salas e do vetor de emails");
	  return;
	}
	var ne = Array();
	for (var i = 0; i < salas.length; i++) {
	  var current = ss.getSheetByName(salas[i]);
	  if (current == null) {
		ne.push(salas[i]);
		continue;
	  }
	  var protection = current.protect();
	  var editors = protection.getEditors();
	  if(editors.indexOf(email_responsaveis[i])<0){
		protection.addEditor(emailUser);
		protection.removeEditors(protection.getEditors());
		if (protection.canDomainEdit()) { // Verifica se a sala pode ser editada por qualquer usuário do domínio. Se sim, remove essa possibilidade.
		  protection.setDomainEdit(false);
		}
		protection.addEditor(email_responsaveis[i]);
	  }
  
	  protection.addEditor('wagnercoltec@teiacoltec.org');
	  
	  if(mestreEditors.indexOf(email_responsaveis[i])<0){
		mestreProtection.addEditor(email_responsaveis[i]);
	  }
	}
	console.log(ne);
  }