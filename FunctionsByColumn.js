// retorna os índices das principais colunas de uma sala
function indexColums(slave){
  var slaveLC = slave.getLastColumn();
  var cabecalho = slave.getRange(1,1,1,slaveLC).getValues();
  cabecalho = cabecalho[0];
  var indices = Array();
  indices.push(cabecalho.indexOf("Nº Patrimônio")+1);
  indices.push(cabecalho.indexOf("Nº ATM")+1);
  indices.push(cabecalho.indexOf("Descrição")+1);
  indices.push(cabecalho.indexOf("Conservação")+1);
  indices.push(cabecalho.indexOf("Situação")+1);
  indices.push(cabecalho.indexOf("Valor")+1);
  indices.push(cabecalho.indexOf("Última Transferência")+1);
  indices.push(cabecalho.indexOf("Sala Anterior")+1);
  indices.push(cabecalho.indexOf("Finalidade")+1);
  indices.push(cabecalho.indexOf("Data Estimada de Devolução")+1);
  indices.push(cabecalho.indexOf("Sincronizado (UFMG)")+1);
  return indices;
}

// retorna a coluna de não  emplaquetado
function getColumnNE(slave){
  var slaveLC = slave.getLastColumn();
  var cabecalho = slave.getRange(1,1,1,slaveLC).getValues();
  cabecalho = cabecalho[0];
  return cabecalho.indexOf("Não Emplaquetado")+1;
}


// retorna os patrimônios especificados pelas linhas
function getValuesByIndex(slave, indexColums, rows){
  var aux = Array.isArray(rows);
  if (!Array.isArray(rows)){
    rows = [rows];
  }
  var slaveLR = lastRow(slave);
  if(slaveLR == 1){
    return [];
  }
  var slaveLC = slave.getLastColumn();
  var patrimonios = slave.getRange(2,1,slaveLR-1,slaveLC).getValues();
  patrimonios = patrimonios.filter(function (value, it){
    return rows.indexOf(it+1) != -1;
  });
  var patTrans  = transpose(patrimonios);
  patrimonios = patTrans.filter(function (value, it){
    return indexColums.indexOf(it+1) != -1;
  });
  return transpose(patrimonios);
}

// retorna os patrimônios de toda a sala por índices
function getValuesByIndexSala(slave, index){
  var slaveLR = lastRow(slave);
  if(slaveLR == 1){
    return [];
  }
  var slaveLC = slave.getLastColumn();
  var patrimonios = slave.getRange(2,1,slaveLR-1,slaveLC).getValues();
  var patTrans  = transpose(patrimonios);
  patrimonios = patTrans.filter(function (value, it){
    return index.indexOf(it+1) != -1;
  });
  return transpose(patrimonios);
}

// 
function setValueByIndex(dest, indexColums, nRow, patrimonio){
  var debug;
  for(var i = 0; i < patrimonio.length; i++){
    debug = patrimonio[i];
    dest.getRange(nRow,indexColums[i]).setValue(patrimonio[i]);
  }
}

//
function insertPatrisByIndex(dest, indexColums, patris){
  patris = transpose(patris);
  var nRow = patris[0].length;
  var nCol = patris.length;
  var destLR = lastRow(dest);
  for(var i = 0; i < nCol; i++){
    dest.getRange(destLR+1,indexColums[i],nRow,1).setValues(transpose([patris[i]]));
  }
}


// através de intervalos, retorna um vetor com todas as linhas ordenadas
function getLines(linha){
  var linhas = linha.split(','); //Vetor com os numéros das linhas.
  for(var i = 0; i < linhas.length; i++){
    if(linhas[i].indexOf('-') > 0){
      let interval = linhas[i].split('-');
      if(interval.length != 2){
        SpreadsheetApp.getUi().alert('Intervalo inválido');
        return 0;
      }
      var begin = parseInt(interval[0]);
      var end = parseInt(interval[1]);
      if (begin > end){
          end = begin;
          begin = parseInt(interval[1]);
      }
      linhas.splice(i,1);
      while(begin <= end){
        linhas.splice(i, 0, begin);
        i++;
        begin++;
      }
      i--;
    }
    else if(linhas[i] == 0){
      linhas.splice(i,1);
      i--;
    }
    else{
      linhas[i] = parseInt(linhas[i]);
    }
  }

  linhas = [...new Set(linhas)];

  linhas.sort(function(a,b){
    return a - b;
  }); // ordenando o array de linhas
  //Se a linha do cabeçalho for escolhida, o usuário é alertado.
  if(linhas[0] == 1){
    linhas.shift();
    SpreadsheetApp.getUi().alert('Linha do cabeçalho bloqueada! Continuando a operação de transferência sem a linha 1');
  }
  return linhas;
}



// retorna a matriz transposta
function transpose(a)
{
  return a[0].map(function (_, c) { return a.map(function (r) { return r[c]; }); });
  // or in more modern dialect
  // return a[0].map((_, c) => a.map(r => r[c]));
}

