function verificarNE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var current = ss.getActiveSheet();
  var sala = current.getSheetName();
  var user = Session.getActiveUser();
  var userSala = getNomeResponsavel(sala);
  var currentLR = lastRow(current);
  var patrimonios = current.getRange(2,1,currentLR-1,7).getValues();
  patrimonios = transpose(patrimonios);
  patrimonios.splice(3,3);
  patrimonios = transpose(patrimonios);
  var ne = patrimonios.filter(function naoEmplaq(line){ return line[3]; });
  var neString = agruparPatris(ne);
  MailApp.sendEmail('desenv1@teiacoltec.org',
                    "Nota de Não-Emplaquetamento de patrimônios",
                    "Patrimônios:\n" + neString +
                    "Sala: " + sala + "\n" +
                    "Responsável da sala: " + userSala + "\n" +
                    "Usuário que notificou a falta da plaqueta: " + user +
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");  //E-mail notificando quais patrimonios nap tem plaqueta
}

function agruparPatris(patris){
  var retorno = String();
  for (let i = 0; i < patris.length; i++){
    retorno += "\"" + patris[i][0] + "\"" + " - " + patris[i][2] + "\n";
  }
}
