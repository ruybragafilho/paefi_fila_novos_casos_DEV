"use strict";

/**
 * Módulo:    atualizarCalculoTempoEsperaBE.gs
 * Objetivo:  Atualiza o cálculo do tempo de espera dos casos designados
 */


/**
 * Função que atualiza o cálculo do tempo de espera
 * dos casos designados
 */
function atualizarCalculoTempoEsperaBE() {

  try {

    console.log( "atualizarCalculoTempoEsperaBE - Início" );
 
    // Calcula e insere, na tabela, o tempo de espera em meses  
    // A função getRange endereça as linhas e colunas começando do indice 1
    // Por isso os +1 na linha e na coluna  
    let idCaso;

    for( idCaso=1; idCaso<=NUM_CASOS; ++idCaso ) {

      if( BUFFER_CASOS[idCaso-1][DATA_DE_DESIGNACAO] != "" ) {

        const dataDeChegadaNoCREAS = TABELA_CASOS.getRange( idCaso+1, DATA_DE_CHEGADA_NO_CREAS+1 ).getA1Notation();  
        const dataDesignacaoA1 = TABELA_CASOS.getRange( idCaso+1, DATA_DE_DESIGNACAO+1 ).getA1Notation();  
        TABELA_CASOS.getRange( idCaso+1, TEMPO_DE_ESPERA+1 ).setFormula(`=DATEDIF(${dataDeChegadaNoCREAS};${dataDesignacaoA1};"M")`);
      } 
    }    

    console.log( "Último idCaso: " + idCaso );
    
    PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
    SpreadsheetApp.flush();  

    console.log( "atualizarCalculoTempoEsperaBE - Fim" );

  } catch( error ) {
    throw( "atualizarCalculoTempoEsperaBE: " + error.message );
  }      

} // Fim da função atualizarCalculoTempoEsperaBE



/**
 * ##### FIM DO MÓDULO atualizarCalculoTempoEsperaBE.gs #####
 */




