
// função feita para evitar edição e consulta 
// simultânea na tabela de transferência por usuários diferentes
function manageTrans(lista, command, patrimonios, lines){
    var lock = LockService.getScriptLock();
	try {
		lock.waitLock(10000);
	  } catch (e) {
		Logger.log('Could not obtain lock after 10 seconds.');
	  }
    // para adicionar patrimônios, lines é ignorado
    // ***Recomendação: chamar a função com lines = 0
    if(command == 'add patris'){ 
        lista.insertRows(2,patrimonios.length);
        lista.getRange(2,1,patrimonios.length,9).setBackground('#ffffff');
        lista.getRange(2,1,patrimonios.length,11).setValues(patrimonios);
    }
    
    // para obter dados da tabela, patrimonios é ignorado
    // mas sua variável é usada
    // ***Recomendação: chamar a função com patrimonios = 0
    if(command == 'get data'){
      patrimonios = Array();
      for(let i of lines){
          patrimonios.push(lista.getRange(i,1,1,12).getValues()[0]);
      }
      lock.releaseLock();
      return patrimonios;
    }

	if(command == 'get background'){
		patrimonios = Array();
    for(let i of lines){
        patrimonios.push(lista.getRange(i,1,1,9).getBackground() == "#e6b8af"); // fundo igual a vermelho
    }
    lock.releaseLock();
    return patrimonios;
        
	}
    
    // para remover os patrimônios, lines é ignorado
    // mas sua variável é usada
    // ***Recomendação: chamar a função com lines = 0
    if(command == 'remove lines'){
        var plaqsPatri = transpose(patrimonios)[0];
        transLR = lastRow(lista);
        var plaqsTrans = lista.getRange(2,1,transLR,1).getValues();
        plaqsTrans = transpose(plaqsTrans)[0];
        lines = Array();
        for(let j of plaqsPatri){
            lines.push(plaqsTrans.indexOf(j)+2);
        }
        // reodernando as linhas para que 
        // a remoção seja feita de baixo pra cima
        lines.sort(function(a,b){return a-b});
        lines.reverse();

        // removendo linha por linha
        for (let k of lines){
            if(k<2){
                continue;
            }
            lista.deleteRows(k,1);
        }
    }

    // para , lines é ignorado
    // mas sua variável é usada
    // ***Recomendação: chamar a função com lines = 0
    if(command == 'deny trans'){
        var plaqsPatri = transpose(patrimonios)[0];
        var data = new Date().getTime();
        var now = new Date(data).toLocaleString('pt-BR');
        transLR = lastRow(lista);
        var plaqsTrans = lista.getRange(2,1,transLR,1).getValues();
        plaqsTrans = transpose(plaqsTrans)[0];
        lines = Array();
        for(let j of plaqsPatri){
            lines.push(plaqsTrans.indexOf(j)+2);
        }
        // reodernando as linhas para que 
        // a remoção seja feita de baixo pra cima
        lines.sort(function(a,b){return a-b});
        lines.reverse();

        // removendo linha por linha
        for (let k of lines){
            if(k<2){
              continue;
            }
            lista.getRange(k,7).setValue(now);
            lista.getRange(k,11).setValue('-');
            lista.getRange(k,1,1,9).setBackground("#e6b8af");
        }
    }
	lock.releaseLock();
}

function manageHistory(history, patrimonios){
var lock = LockService.getScriptLock();
  try {
  	lock.waitLock(10000);
  	} catch (e) {
  	Logger.log('Could not obtain lock after 10 seconds.');
  	}
  var historyLR = lastRow(history);
  var nPat = patrimonios.length;
  history.getRange(historyLR+1, 1, nPat, 13).setValues(patrimonios);
  lock.releaseLock();
}


function manageMestre(mestre,command, patrimonios){
  var lock = LockService.getScriptLock();
 try {
 	lock.waitLock(20000);
 	} catch (e) {
 	Logger.log('Could not obtain lock after 10 seconds.');
 	}
  if(command == 'sort'){
	  // ordenando a planilha mestre
	  var rangeTotal = mestre.getRange("A2:J");
	  rangeTotal.sort(SORT_ORDER);
    lock.releaseLock();
    return;
  }
  var mestreLR = lastRow(mestre);
  if(command == 'update'){
	  var newPats = Array();
  
	  	//Loop que varre os patrimônios a serem transferidos.
	  var plaquetas = mestre.getRange(1,1,mestreLR,1).getValues(); //Coluna de plaquetas.
	  plaquetas = transpose(plaquetas)[0];
  
	  // obtendo o numeros das linhas na planilha mestre de cada patrimonio
	  // se não tiver, ele adiciona
	  for(let [index, value] of patrimonios.entries()){
	  	let pos = plaquetas.indexOf(value[0]); //Procurar pela plaqueta na coluna.
	  	if (pos == -1) {
	  	newPats.push(index);
	  	}
	  	//Se ela for localizada, a sua linha é obtida.
	  	else {
	  	  mestre.getRange(pos + 1,1,1,10).setValues([value]);
	  	}
	  }
	  var newPats = patrimonios.filter(function(value,index){
	  	return !(newPats.indexOf(index) < 0);
	  });
	  // colocando os novos patrimônios
    if(newPats.length > 0){
      mestre.getRange(mestreLR+1,1,newPats.length,10).setValues(newPats);
    }
  }
  if(command == 'remove'){
	  var linhas = Array();
  
	  	//Loop que varre os patrimônios a serem transferidos.
	  var plaquetas = mestre.getRange(1,1,mestreLR,1).getValues(); //Coluna de plaquetas.
	  plaquetas = transpose(plaquetas)[0];
  
	  // obtendo o numeros das linhas na planilha mestre de cada patrimonio
	  // se não tiver, ele adiciona
	  for(value of patrimonios){
	  	let pos = plaquetas.indexOf(value[0]); //Procurar pela plaqueta na coluna.
	  	if (pos != -1) {
	  	  linhas.push(pos);
	  	}
	  }
    linhas.reverse();
    for (let k of linhas){
      if(k<0){
          continue;
      }
      mestre.deleteRows(k+1,1);
    }

  }
  lock.releaseLock();
}


// verifica se o nome não é de sala
function reservado(value){
	var nomesReservados = ["Mestre", 
						   "Dados dos responsáveis", 
						   "Lista de desfazimento", 
						   "Lista de transferências",  
						   "Histórico de Transferência", 
						   "Valores estados dos patrimonios"];
	return nomesReservados.indexOf(value) > -1;
  }

