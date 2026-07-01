"use strict";

/**
 * Módulo:    designarCasoBE.gs
 * Objetivo:  designa um caso da tabela CASOS
 */


/**
 * Função que designa um caso da tabela CASOS
 *  
 * @param {String} id: id do caso de será designado 
 * @param {String} idMotivoDeDesignacao: id do motivo de designação 
 * @param {String} idTecnicoPAEFI: ID do técnico do PAEFI, caso o tipo de designação seja Inserção no PAEFI
 *                                 "", para os outros tipos de designação            
 */
function designarCasoBE( id,                       
                         idMotivoDeDesignacao,
                         idTecnicoPAEFI ) {


  // Se id inválido, retorna uma exceção
  if( id < 1  ||  id > NUM_CASOS ) {
    throw( new Error( "ID Inválido" ) );
  }  

  // Verifica se o usuário do app tem permissão para designar o caso  
  let usuarioLogado;
  try {
    usuarioLogado = JSON.parse( autenticarUsuario() );
  } catch( error ) {
    throw( "designarCasoBE: " + error.message );
  }    
  if( usuarioLogado.tipo == "2" || usuarioLogado.regional != BUFFER_CASOS[id-1][REGIONAL] ) {
    throw( new Error( "Usuário sem permissão para designar o caso" ) );
  }  

  // TENTA PEGAR O LOCK
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);  

  // SE PEGAR O LOCK, PROSSEGUE COM A DESIGNAÇÃO
  if( lock.hasLock() ) {

    // Converte o id para Integer
    const idCaso = parseInt(id);
        
    // Gera, formata e grava a data de designação do caso
    const dataDeDesignacao = new Date().toLocaleString("pt-BR", {dateStyle: "short"});
    const data = TABELA_CASOS.getRange( idCaso+1, DATA_DE_DESIGNACAO+1 );
    data.setValue( dataDeDesignacao );    

    // Grava o motivo de designação do caso
    const idMotivo = TABELA_CASOS.getRange( idCaso+1, MOTIVO_DE_DESIGNACAO+1 );
    idMotivo.setValue( idMotivoDeDesignacao );        

    // Grava o id do técncio do PAEFI
    const idTecnico = TABELA_CASOS.getRange( idCaso+1, ID_TECNICO_PAEFI+1 );    
    idTecnico.setValue( idTecnicoPAEFI );


    // Calcula e insere, na tabela, o tempo de espera em meses  
    // A função getRange endereça as linhas e colunas começando do indice 1
    // Por isso os +1 na linha e na coluna  
    try {
      const dataDeChegadaNoCREAS = TABELA_CASOS.getRange( idCaso+1, DATA_DE_CHEGADA_NO_CREAS+1 ).getA1Notation();  
      const dataDesignacaoA1 = TABELA_CASOS.getRange( idCaso+1, DATA_DE_DESIGNACAO+1 ).getA1Notation();  
      TABELA_CASOS.getRange( idCaso+1, TEMPO_DE_ESPERA+1 ).setFormula(`=DATEDIF(${dataDeChegadaNoCREAS};${dataDesignacaoA1};"M")`);
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

} // Fim da função designarCasoBE



/**
 *  #####  TESTES PARA AS FUNÇÕES DESSE MÓDULO  #####
 */



/**
 * Função para testar a função principal designarCasoBE
 */
function teste_designarCasoBE() {

  let id = 1;
  let idMotivoDeDesignacao = "3";   

  try {
    designarCasoBE( id,
                    idMotivoDeDesignacao,
                    "" );
  } catch( error ) {
    console.log( error.message );
  }

} // Fim da função teste_designarCasoBE




/**
 * ##### FIM DO MÓDULO designarCasoBE.gs #####
 */




