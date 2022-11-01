// Função para atualizar os patrimônios de uma sala na planilha mestre.
function atualizarSala(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mestre = ss.getSheetByName('Mestre');  //Planilha mestre.
  var slave = ss.getActiveSheet(); //Planilha ativa no momento.
  var nomeSala = slave.getSheetName(); //Pega o nome da planilha ativa.
  var emailUser = Session.getActiveUser(); //E-mail do usuário que chamou essa função.
  
  var emailResponsavel = getEmailResponsavel(nomeSala);
  if (emailResponsavel != emailUser &&
      emailUser != 'desenv1@teiacoltec.org' &&
      emailUser != 'desenv2@teiacoltec.org' &&
      emailUser != 'wagnercoltec@teiacoltec.org') {
      Browser.msgBox('Você não tem permissão para atualizar os patrimônios dessa sala na planilha mestre. ' +
      'Por favor, execute a função em uma sala de sua responsabilidade.', Browser.Buttons.OK);
      return
    } //Confere se a pessoa está autorizada a realizar ações nessa sala.
  
  var indices = indexColums(slave);
  indices.splice(8,2);
  var patrimonios = getValuesByIndexSala(slave,indices);
  if(patrimonios.length == 0){ // não tem nenhum patrimônio na planilha
    return;
  }
  var patT = transpose(patrimonios);
  patT.splice(7,0,Array(patT[0].length).fill(nomeSala));
  patrimonios = transpose(patT);
  manageMestre(mestre,'update', patrimonios);
}

function atualizarTodasAsSalas(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mestre = ss.getSheetByName('Mestre');  //Planilha mestre.
    var emailUser = Session.getActiveUser(); //E-mail do usuário que chamou essa função.
    var salas = ss.getSheets();

    if (emailUser != 'desenv1@teiacoltec.org' &&
        emailUser != 'desenv2@teiacoltec.org' &&
        emailUser != 'wagnercoltec@teiacoltec.org') {
        
        Browser.msgBox('Você não tem permissão para atualizar os patrimônios de todas as salas na planilha mestre. ' +
        'Por favor, execute a função em uma sala de sua responsabilidade.', Browser.Buttons.OK);
        return
      } //Confere se a pessoa está autorizada a realizar ações nessa sala.

    var patrimonios = Array();
    for(let slave of salas){
        var nomeSala = slave.getSheetName();
        if(reservado(nomeSala))
        { continue; } // pula se a planilha não for uma sala

        var indices = indexColums(slave);
        indices.splice(8,2);
        var pat = getValuesByIndexSala(slave,indices);
        if(pat.length == 0){ // planilha sem patrimônio
          continue; // pula a planilha
        }
        pat = transpose(pat);
        pat.splice(7,0,Array(pat[0].length).fill(nomeSala));
        patrimonios.concat(transpose(pat));
    }
   manageMestre(mestre,'update', patrimonios);
   manageMestre(mestre, 'sort', 0);
  }
