//Função para desfazer-se de um patrimônio.
function desfazer(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var slave = ss.getActiveSheet(); //Planilha ativa no momento.
  var lista = ss.getSheetByName('Lista de desfazimento'); //Planilha da lista de desfazimento.
  var dados = ss.getSheetByName('Dados dos responsáveis'); //Planilha dos dados dos responsáveis.
  var origem = slave.getName(); //Nome da planilha ativa.
  var linha = Browser.inputBox('Desfazer de um Patrimônio - Patrimônio', 'Digite o número da linha do patrimônio a ser desfeito.', Browser.Buttons.OK_CANCEL); //O Usuário é solicitado a digitar a linha do patrimônio que será desfeito.
  if (linha == 'cancel') {
    return
  }
  else if (linha == 1) {
    SpreadsheetApp.getUi().alert('Linha 1 bloqueada! Por favor, repita o procedimento.'); //Não pode ser a linha do cabeçalho. O usuário é alertado.
    return
  }
  
  var imagem = Browser.inputBox('Desfazer de um Patrimônio - Foto', 'Digite o link com as fotos do patrimônio a ser desfeito.', Browser.Buttons.OK_CANCEL); //O Usuário é solicitado a digitar o link da foto do patrimônio que será desfeito.
  if (imagem == 'cancel') {
    return
  }
  
  var justificativa = Browser.inputBox('Desfazer de um Patrimônio - Justificativa', 'POSSÍVEIS JUSTIFICATIVAS PARA O NÃO APROVEITAMENTO ATUAL DE UM BEM: \n' +
                                       '1 – Bem está velho e obsoleto e foi substituído por outro mais novos, modernos e/ou de qualidade superior;\n' +
                                       '2 – Bem é antigo não sendo mais útil devido às obras de melhoria de infraestrutura/ do Colégio;\n' +
                                       '3 – Bem não tem mais utilidade na rotina atual de atividades do laboratório;\n' +
                                       '4 – Bem não tem mais utilidade devido a informatização do setor;\n' +
                                       '5 – Bem perdeu suas características e funcionalidades devido ao uso prolongado ou em razão de ter atingido seu tempo de vida útil;\n' +
                                       '6 – Bem não tem mais utilidade devido a adoção de novas tecnologias e estratégias na rotina escolar e/ou administrativa;\n' +
                                       '7 – Bem não tem mais utilidade devido a digitalização de documentos e virtualização de processos;\n' +
                                       '8 – Bem não se encontra mais em condições de uso e necessita ser substituído;\n' +
                                       '9 – Outros (descrever a razão).\n' +
                                       'Copiar e colar a sua justificativa. Caso seja outra, descreva-a abaixo.', Browser.Buttons.OK_CANCEL); //O Usuário é solicitado a digitar a justificativa para o desfazimento do patrimônio.
  if (justificativa == 'cancel') {
    return
  }
  var indices = indexColums(slave);

  var valoresPat = getValuesByIndex(slave,indices,linha); //Pega todos os valores do patrimônio a ser transferido.
  valoresPat.splice(indices.length - 3, 3);
  var plaqueta = valoresPat[0]; //Pega o valor da plaqueta.
  
  var listaLR = lastRow(lista); //Última linha escrita da lista.
  lista.getRange(listaLR+1,1,1,8).setValues([valoresPat]); //Adiciona o patrimônio na planilha. 
  lista.getRange(listaLR+1,8).setValue(origem); //Coloca a sala de origem do patrimônio.
  lista.getRange(listaLR+1,7).setValue('Disponível'); //Atualiza a situação do patrimônio na lista.
  lista.getRange(listaLR+1,9).setValue(imagem); //Coloca o link da imagem do patrimônio.
  lista.getRange(listaLR+1,10).setValue(justificativa); //Coloca a justificativa do desfazimento.
  lista.setRowHeight(listaLR, 130);
  
  var emailRemetente = getEmailResponsavel(origem); //E-mail do responsável pela sala de origem.
  
  MailApp.sendEmail(emailRemetente,
                    "Transferência de Patrimônio - SiPat - Notificação de Envio",
                    "Patrimônio: " + plaqueta + "\n" +
                    "Origem: " + origem + "\n" +
                    "Situação: Disponível para transferência na Lista de Desfazimento. "
                    + "\n-------------------------------------------------------\n"
                    + "Mensagem auto-enviada pelo SiPat: Sistema de Patrimônio do Coltec");    //Um e-mail é enviado confirmando o desfazimento.
  
  slave.deleteRow(linha); //O patrimônio é deletado da sala de origem.
}