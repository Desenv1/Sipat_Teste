//Interface do usuário.
function onOpen(){
  var ui = SpreadsheetApp.getUi();
  //Funções do administrador.
  ui.createMenu("Administrador") 
  .addItem("Adicionar Usuário", "adicionar")
  .addItem("Remover Usuário", "remover")
  .addItem("Gerar Lista de Transferências", "getDocSeletivo")
  .addItem("Gerar Lista de Patrimônio por Sala", "getDocPatSala")
  .addItem("Gerar Lista de Patrimônio por Valores", "getDocValor")
  .addItem("Gerar Lista de Patrimônio por Setor", "getDocPatSetor")
  .addItem("Gerar Lista de Salas do Sipat", "getDocSalas")
  .addItem("Aviso de sincronismo com SicPat - UFMG", "sincronizaUFMG")
  .addItem("Organizar Salas em Ordem Alfabética", "sortSheets")
  .addItem("Atualizar Patrimônio de Todas as Salas", "atualizarTodasAsSalas")
  .addItem("Atualizar Proteção de salas","atualizarResponsaveis")
  .addToUi();
  //Funções de atualização dos patrimônios.
  ui.createMenu("Atualizar Patrimônios")
  .addItem("Atualizar patrimônio da sala atual na Planilha Mestre", "atualizarSala")
  .addItem("Remover Patrimônio da Sala Atual", "removerPatrimonio")
  .addToUi();
  //Funções de transferência para o usuário.
  ui.createMenu("Tranferência")
  .addItem("Transferência de Patrimônio","transferir")
  .addItem("Autorização de Transferência", "permissoes")
  .addItem("Cancelar Transferência", "cancelTransfer")
  .addItem("Desfazer-se de um Patrimônio", "desfazer")
  .addItem("Obter Patrimônio Desfeito", "obter")
  .addItem("Retomada de Patrimônio", "retomar")
  .addItem("Gerar Documento de Transferência", "getDoc")
  .addItem("Gerar Documento de Manutenção", "getDocManut")
  .addToUi();  
}
