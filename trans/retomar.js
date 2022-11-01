
//Função para retomar patrimônios da lista de transferências.
function retomar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre'); //Planilha mestre.
  var mestreLR = lastRowMestre(); //Última linha escrita da planilha mestre.
  var lista = ss.getActiveSheet(); //Planilha ativa.
  var history = ss.getSheetByName('Histórico de Transferência'); //Planilha mestre.
  var responsaveis = ss.getSheetByName('Dados dos responsáveis');
  
  var ui = SpreadsheetApp.getUi(); 
  
  var data = new Date().getTime(); //Data atual;
  var now = new Date(data).toLocaleString('pt-BR');

  if(lista.getSheetName() != 'Lista de transferências'){
    SpreadsheetApp.getUi().alert('Página errada! Você deve realizar essa ação apenas na página "Lista de transferências".');
    return;
  }

  //Usuário é solicitado para digitar as linhas dos patrimônios que deseja retomar.
  var linhaInput = Browser.inputBox('Retomada de Patrimônio', 'Digite o(s) número(s) da(s) linha(s) do(s) patrimônio(s) a ser retomados(s). Caso seja mais de um, separe as linhas por vírgulas e defina intervalos por "-". Não coloque espaço. Exemplo: "5,6-8,13-16,17"', Browser.Buttons.OK_CANCEL);
  if (linhaInput == 'cancel' || linhaInput == 0) return;

  // pegando as linhas e colocando em ordem
  var linhas = getLines(linhaInput); //Vetor com os numéros das linhas.
  if (linhas == 0 || linhas.length == 0) {
    Browser.msgBox("Nenhuma linha foi inserida, ou somente o cabeçalho foi selecionado.\n"+
                   "Encerrando a retomada. Tente novamente\n");
    return;
  } 

  var numLinhas = linhas.length; //Tamanho do vetor.

  //pegando todas as salas em que o usuário é responsável
  var emailUser = Session.getActiveUser().getEmail(); //E-mail do usuário.
  var dadosResponsaveis = responsaveis.getRange(2,4,responsaveis.getLastRow()-1,3).getValues();
  var salasResp = dadosResponsaveis;
  salasResp = salasResp.filter(function(elem){
    return elem[2] == emailUser;
  })
  salasResp = transpose(salasResp)[0];


  // pegando os patrimônios,
  var patrimonios = manageTrans(lista,'get data', 0, linhas);
  // removendo a última coluna, 
  var patTrans = transpose(patrimonios);
  // colocando a data da retomada,
  patTrans[6].fill(now);
  // e transpondo a planilha novamente
  patrimonios = transpose(patTrans);
  
  // filtrando por transferências negadas e 
  // pelas salas de origem pertencentes ao usuário
  var background = manageTrans(lista,'get background', 0, linhas);
  patrimonios = patrimonios.filter(function(elem){
    return salasResp.indexOf(elem[6]) != -1 || emailUser == 'desenv1@teiacoltec.org'; // sala de origem pertence ao usuário
  })

  patrimonios = patrimonios.filter(function(elem,index){
    return background[index]; // verifica se o patrimônio foi recusado
  });


  //
  var patTranspose = transpose(patrimonios);
  var salasOrigem = patTranspose[7];
  //var salasDestino = patTranspose[8];
  var uniqueOrigem = uniq_fast(salasOrigem).sort();

  var patBySala = Array();
  for(let i of uniqueOrigem){
    patBySala.push([i,Array()]);
  }

  for(let i of patrimonios){
    let pos = uniqueOrigem.indexOf(i[7]);
    patBySala[pos][1].push(i);
  }

  var permissao = Array();
  patrimonios = Array();
  var msgSala = String();
  for(let i of patBySala){
    var plaquetas = String();
    let aux = 1;
    for(let j of i[1]){
      plaquetas += aux + ' - ' + j[0] + '   :   ' + j[2] + '\n';
      aux++;
    }

    var msg = 'Você deseja realizar a retomada dos seguintes patrimônio para a sala:\n';
    msg += i[0] + '?\n\n'+ plaquetas.slice(0, -1);
    var retomada = ui.alert('Retomada de Patrimônio', msg, ui.ButtonSet.YES_NO_CANCEL);
    if(retomada == ui.Button.YES){
      permissao.push(true);
      patrimonios = patrimonios.concat(i[1]);
      patTrans = transpose(i[1]);
      patTrans.splice(8,1);
      i[1] = transpose(patTrans);
      msgSala += "Sala " + i[0] + ":\n" + plaquetas + "\n";
    }
    if(retomada == ui.Button.NO){
      permissao.push(false);
    }
    if(retomada == ui.Button.CANCEL){
      return;
    }
  }
  patBySala = patBySala.filter(function(elem, index){
    return permissao[index];
  });  
  
  var dadosResponsaveisT = transpose(dadosResponsaveis);
  var destinatarios = Array();
  var patByUserDest = Array();
  
  for(let i of patrimonios){
    let pos = dadosResponsaveisT[0].indexOf(i[8]);
    destinatarios.push(dadosResponsaveisT[2][pos]);
  }
  var uniqueUserDest = uniq_fast(destinatarios).sort();
  for(let i of uniqueUserDest){
    patByUserDest.push([i,Array()]);
  }

  for(let [index, value] of patrimonios.entries()){
    let pos = uniqueUserDest.indexOf(destinatarios[index]);
    patByUserDest[pos][1].push(value);
  }

  patTrans = transpose(patrimonios);
  patTrans.splice(9,0,Array(numLinhas).fill("Retomado"));
  patrimonios = transpose(patTrans);
  manageHistory(history,patrimonios);

  //falta modificar pra entrar pra planilha mestre
  var salaDest = patTrans[8];
  patTrans.splice(8,4);
  patTrans.splice(7,0,salaDest);
  patrimonios = transpose(patTrans);
  manageMestre(mestre, 'update', patrimonios);
  manageMestre(mestre, 'sort', 0);

  for (let i of patBySala){
    var sheetOrigem = ss.getSheetByName(i[0]);
    var indices = indexColums(sheetOrigem);
    insertPatrisByIndex(sheetOrigem, indices,i[1]);
  }

  manageTrans(lista, 'remove lines', patrimonios,0);

  MailApp.sendEmail(emailUser,
    "Transferência de Patrimônio - SiPat - Patrimônio Recuperado",
    "Recuperação de patrimônios feito com sucesso.\n" + 
    "Abaixo estará a lista com todos os patrimônios retomados.\n" + msgSala +
    "\n-------------------------------------------------------\n" +
    "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");

  for(let i of patByUserDest){
    msgSala = String();
    for(let j of i[1]){
        msgSala += j[0] + '   :   ' + j[2] + '\n';
    }
    MailApp.sendEmail(i[0],
        "Transferência de Patrimônio - SiPat - Patrimônio Recuperado",
        "Patrimônios: \n" + msgSala + "\n" +
        "Situação: Patrimônio recuperado pelo responsável da sala de origem."
        + "\n-------------------------------------------------------\n"
        + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  
  }
}

function uniq_fast(a) {
    var seen = {};
    var out = [];
    var len = a.length;
    var j = 0;
    for(var i = 0; i < len; i++) {
         var item = a[i];
         if(seen[item] !== 1) {
               seen[item] = 1;
               out[j++] = item;
         }
    }
    return out;
}

