"use strict";

/**
 * Módulo:    reativarCasoBE.gs
 * Objetivo:  Reativa um caso da tabela CASOS
 */


/**
 * Função que reativa um caso da tabela CASOS
 *  
 * @param {String} id: id do caso de será reativado 
 *  
 */
function reativarCasoBE( id ) {

  // Se id inválido, retorna uma exceção
  if( id < 1  ||  id > NUM_CASOS ) {
    throw( new Error( "ID Inválido" ) );
  }  

  // Verifica se o usuário do app tem permissão para reativar o caso
  let usuarioLogado;
  try {
    usuarioLogado = JSON.parse( autenticarUsuario() );
  } catch( error ) {
    throw( "reativarCasoBE: " + error.message );
  }    
  if( usuarioLogado.tipo == "2" || usuarioLogado.regional != BUFFER_CASOS[id-1][REGIONAL] ) {
    throw( new Error( "Usuário sem permissão para reativar o caso" ) );
  }  

  // TENTA PEGAR O LOCK
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);  

  // SE PEGAR O LOCK, PROSSEGUE COM A REATIVAÇÃO
  if( lock.hasLock() ) {

    // Converte o id para Integer
    const idCaso = parseInt(id);
        
    // Grava null na data de designação do caso    
    const data = TABELA_CASOS.getRange( idCaso+1, DATA_DE_DESIGNACAO+1 );
    data.setValue( "" );    

    // Grava null no motivo de designação do caso
    const idMotivo = TABELA_CASOS.getRange( idCaso+1, MOTIVO_DE_DESIGNACAO+1 );
    idMotivo.setValue( "" );        

    // Grava null no id do técnico do PAEFI
    const idTecnico = TABELA_CASOS.getRange( idCaso+1, ID_TECNICO_PAEFI+1 );
    idTecnico.setValue( "" );     
    
    // Calcula e insere, na tabela, o tempo de espera em meses  
    // A função getRange endereça as linhas e colunas começando do indice 1
    // Por isso os +1 na linha e na coluna  
    try {
      const dataDeChegadaNoCREAS = TABELA_CASOS.getRange( idCaso+1, DATA_DE_CHEGADA_NO_CREAS+1 ).getA1Notation();  
      TABELA_CASOS.getRange( idCaso+1, TEMPO_DE_ESPERA+1 ).setFormula(`=DATEDIF(${dataDeChegadaNoCREAS};Today();"M")`);
    } catch( error ) {
      throw( error.message );
    }    

    // SOLTA O LOCK
    lock.releaseLock();

    return true;

  } else {

    // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
    throw( new Error( "Nao foi possivel pegar o LOCK" ) );
  }

} // Fim da função reativarCasoBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */



/**
 * Função para testar a função principal reativarCaso
 */
function teste_reativarCasoBE() {

  let id = 1;  

  try {
    reativarCasoBE( id );
  } catch( error ) {
    console.log( error.message );
  }

} // Fim da função teste_reativarCasoBE




/**
 * ##### FIM DO MÓDULO reativarCaso.gs #####
 */




