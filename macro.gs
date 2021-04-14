/** @OnlyCurrentDoc */

function ComparaçãoAB(){
  // Definindo as variáveis responsável por pegar as informações da planilha e da folha selecionada;
  var planilha = SpreadsheetApp.getActive();
  var folha = planilha.getActiveSheet();

  // Definindo variáveis responáveis por pegar todos os valores da coluna A e da coluna B;
  // Ele pega todas as celulas, estando vazias ou não;
  var cA = folha.getRange(2, 1, folha.getMaxRows(), 1);
  var cB = folha.getRange(2, 2, folha.getMaxRows(), 2);
  
  // Pegando o tamanho da coluna A e coluna B;
  var tamanho_cA = cA.getValues().length;
  var tamanho_cB = cB.getValues().length;

  // Como ele tá pegando a quantidade máxima de células, muitos estão vazias fazendo o laço fazer calculo desnecessários;
  // Então esses laços é para verificar onde está acabando as colunas A e B, indentificando a primeira célula vazia;
  for(var i = 0; i < tamanho_cA; i++){
    if(cA.getValues()[i][0] == ""){
      tamanho_cA = i;
      break;
    }
  }
  for(var i = 0; i < tamanho_cB; i++){
    if(cB.getValues()[i][0] == ""){
      tamanho_cB = i;
      break;
    }
  }

  // Mudando o tamanho das colunas para salvar memória;
  var cA = folha.getRange(2, 1, tamanho_cA, 1);
  var cB = folha.getRange(2, 2, tamanho_cB, 2);

  // Compara todos os valores de  A com B, caso sejam iguais a celula da coluna B será pintada de uma cor
  for(var i = 0; i < tamanho_cA; i++){
    for(var j = 0; j < tamanho_cB; j++){
      if(cA.getValues()[i][0] == cB.getValues()[j][0]){
        // como o valor de j começa no 0 e a primeira coluna na segunda linha, precisa somar com +2 para ajustar cara a célula correta
        folha.getRange(j+2, 2).setBackground("#90EE90");
      }
    }
  }

}