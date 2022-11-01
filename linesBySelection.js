
// retorna as linhas por seleção do usuário
function bySelection(){
  var selection = SpreadsheetApp.getActiveSpreadsheet().getSelection(); // Pegando a seleção da tabela
  var rangeList = selection.getActiveRangeList(); // Pega a lista de intervalos selecionados
  var ranges = rangeList.getRanges(); // Pega um vetor contendo os ranges seecionados
  var rows = String();
  for(let i = 0; i < ranges.length; i++){ // Nesse loop ele pega as linhas e gera uma string
    let posR = ranges[i].getRowIndex();
    let numberR = ranges[i].getNumRows();
    rows += posR;
    if(numberR > 1){
      rows += "-" + (posR + numberR - 1);
    }
    rows += ","
  }
  rows = rows.slice(0, -1);
  return rows;
}

