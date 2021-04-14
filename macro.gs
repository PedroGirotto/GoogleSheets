/** @OnlyCurrentDoc */

function ComparaçãoAB(){
  // Definindo as variáveis responsável por pegar as informações da planilha e da folha selecionada;
  var planilha = SpreadsheetApp.getActive();
  var folha = planilha.getActiveSheet();

  // Definindo variáveis responáveis por pegar todos os valores da coluna A e da coluna B;
  // Ele pega todas as celulas, estando vazias ou não;
  var cA = folha.getRange(2, 1, folha.getMaxRows(), 1);
  var cB = folha.getRange(2, 2, folha.getMaxRows(), 2);
  
  // Variáveis para armazenar os valores das colunas A e B
  var valor_A = cA.getValues();
  var valor_B = cB.getValues();
  var cont = 0;

  // Pegando o tamanho da coluna A e coluna B;
  var tamanho_cA = cA.getValues().length;
  var tamanho_cB = cB.getValues().length;

  // Como ele tá pegando a quantidade máxima de células, muitos estão vazias fazendo o laço fazer calculo desnecessários;
  // Então esses laços é para verificar onde está acabando as colunas A e B, indentificando a primeira célula vazia;
  valor_A.every(function(valor){
    if(valor[0] == ""){
      tamanho_cA = cont;
      cont = 0;
      return false;
    }
    else{
      cont++;
      return true;
    }
  })
  valor_B.every(function(valor){
    if(valor[0] == ""){
      tamanho_cB = cont;
      cont = 0;
      return false;
    }
    else{
      cont++;
      return true;
    }
  })

  // Mudando o tamanho das colunas para salvar memória;
  cA = folha.getRange(2, 1, tamanho_cA, 1);
  cB = folha.getRange(2, 2, tamanho_cB, 2);

  valor_A = cA.getValues();
  valor_B = cB.getValues();

  // Compara todos os valores de A com B, caso sejam iguais a celula da coluna B será pintada de uma cor
  var igual = false;
  valor_A.forEach(function(valorA, i){
    valor_B.forEach(function(valorB, j){
      if(valorA[0] == valorB[0]){
        folha.getRange(j+2, 2).setBackground("#90EE90");
        igual = true;
      }
    })
    if(igual){
      folha.getRange(i+2, 1).setBackground("#ffb27a");
      igual = false;
    }
  })

// fim do código
}
