/**
 * Planilha RELATORIO
 */
const PLANILHA_RELATORIO_ID     =  "1dKDF-AFro9pju-U97LfH6pw5bVsHH-6pvdUzXwq_Wtg";
const PLANILHA_RELATORIO        =  SpreadsheetApp.openById(PLANILHA_RELATORIO_ID);

const TABELA_RELATORIO          =  PLANILHA_RELATORIO.getSheetByName('RELATORIO');
let BUFFER_RELATORIO            =  TABELA_RELATORIO.getDataRange().getDisplayValues().splice(1);
let NUM_RELATORIOS              =  BUFFER_RELATORIO.length;
const NUM_COLUNAS_TABELA_RELATORIO  =  24;



/**
 * Função que limpa a planilha RELATORIO
 */
function limparRelatorio() {

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE COM A DESIGNAÇÃO
    if( lock.hasLock() ) {

      console.log( "limparRelatorio - Início" );
  
      // Caso nulo
      let casoNulo = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");
      let bufferCasosNulos = [];
    
      // Limpa a fila
      let range;
      for( let linha=2; linha<=NUM_RELATORIOS+1; ++linha ) {
        bufferCasosNulos.push(casoNulo);
      }    
      
      
      // Grava o buffer de casos nulos na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, bufferCasosNulos.length, NUM_COLUNAS_TABELA_RELATORIO ).setValues( bufferCasosNulos );  
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  
  
      console.log( "limparRelatorio - Fim" );

    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "limparRelatorio: " + error.message );

  } finally {

    // SOLTA O LOCK
    lock.releaseLock();    
  }    

} // Fim da função limparRelatorio



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * gravando-os na planilha RELATORIO
 */
function gerarRelatorio() {

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE COM A DESIGNAÇÃO
    if( lock.hasLock() ) {
      
      console.log( "gerarRelatorioFila - Início" );
  
      let caso;
    
      let bufferRelatorioCaso;

      const relatorio = [];
  
      // Percorre todos os casos da fila, gerando o relatório de cada caso
      for( let idCaso=1; idCaso<=NUM_CASOS; ++idCaso ) {
      
        caso = BUFFER_CASOS[idCaso - 1];
    
        bufferRelatorioCaso = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");

        bufferRelatorioCaso[0] = caso[ ID ];
        bufferRelatorioCaso[1] = caso[ REFERENCIA_FAMILIAR ];
        bufferRelatorioCaso[2] = caso[ TIPO_LOGRADOURO ];
        bufferRelatorioCaso[3] = caso[ NOME_LOGRADOURO ];
        bufferRelatorioCaso[4] = caso[ NUMERO ];
        bufferRelatorioCaso[5] = caso[ COMPLEMENTO ];
        bufferRelatorioCaso[6] = caso[ BAIRRO ];
        bufferRelatorioCaso[7] = idsToNomes( caso[ REGIONAL ], "REGIONAIS" );
        bufferRelatorioCaso[8] = caso[ CEP ];
      
        bufferRelatorioCaso[9] = caso[ TPSA ];
        bufferRelatorioCaso[10] = caso[ DATA_DE_CHEGADA_NO_CREAS ];
      
        let orgaosEncaminhadores =  idsToNomes( caso[ ORGAOS_ENCAMINHADORES ], "ORGAOS_ENCAMINHADORES" );
        if( orgaosEncaminhadores != "" ) {
          orgaosEncaminhadores = orgaosEncaminhadores.replaceAll( ";", ";\n" );
        } else {
          orgaosEncaminhadores = "SEM INFORMAÇÃO";
        }  
        bufferRelatorioCaso[11] = orgaosEncaminhadores;
      
        bufferRelatorioCaso[12] = caso[ DATA_PREVISTA_PARA_RESPOSTA ];
        bufferRelatorioCaso[13] = caso[ DATA_DA_ULTIMA_RESPOSTA ];  
        bufferRelatorioCaso[14] = caso[ DATA_DE_DESIGNACAO ];
    
        bufferRelatorioCaso[15] = caso[ MOTIVO_DE_DESIGNACAO ] != "" ? 
                            idsToNomes( caso[ MOTIVO_DE_DESIGNACAO ], "MOTIVOS_DE_DESIGNACAO" ) : 
                            "NÃO DESIGNADO";
      
        bufferRelatorioCaso[16] = caso[ TOTAL_DE_PONTOS ];
        bufferRelatorioCaso[17] = caso[ TEMPO_DE_ESPERA ];
      
        let violacoes =  idsToNomes( caso[ VIOLACOES_CASO ], "VIOLACOES" );
        if( violacoes != "" ) {
          violacoes = violacoes.replaceAll( ";", ";\n" );
        } else {
          violacoes = "SEM INFORMAÇÃO";
        }
        bufferRelatorioCaso[18] = violacoes;
      
        let categorias =  idsToNomes( caso[ CATEGORIAS_CASO ], "CATEGORIAS" );
        if( categorias != "" ) {
          categorias = categorias.replaceAll( ";", ";\n" );
        } else {
          categorias = "SEM INFORMAÇÃO";
        }  
        bufferRelatorioCaso[19] = categorias;
      
        bufferRelatorioCaso[20] = caso[ PONTUACAO_PARAMETROS_CASO ];  
      
        let parametros =  idsToNomes( caso[ PARAMETROS_CASO ], "PARAMETROS" );
        if( parametros != "" ) {
          parametros = parametros.replaceAll( ";", ";\n" );
        } else {
          parametros = "SEM INFORMAÇÃO";
        }
        bufferRelatorioCaso[21] = parametros;  
      
        bufferRelatorioCaso[22] = caso[ OBSERVACAO ];    
        bufferRelatorioCaso[23] = caso[ ID_TECNICO_PAEFI ] != "" ? 
                            idsToNomes( caso[ ID_TECNICO_PAEFI ], "TECNICOS" ) : 
                            "SEM INFORMAÇÃO";
        
        relatorio.push( bufferRelatorioCaso );
        
  
      } // Fim do for que percorre todos os casos da fila
  

      // Grava o buffer do relatório na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, relatorio.length, NUM_COLUNAS_TABELA_RELATORIO ).setValues( relatorio );
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  
  
      console.log( "gerarRelatorioFila - Fim" );

    } else {

      // SE NAO CONSEGUIR PEGAR O LOCK, LANCA UMA EXCESSAO
      throw( new Error( "Nao foi possivel pegar o LOCK" ) );
    }

  } catch( error ) {

    throw( "gerarRelatorioFila: " + error.message );

  } finally {

    // SOLTA O LOCK
    lock.releaseLock();    
  }   

} // Fim da função gerarRelatorio



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * e o retorna em excel
 */
function getRelatorioExel() {


  // Verifica se o usuário do app tem permissão para obter o relatório
  let usuarioLogado;
  try {
    usuarioLogado = JSON.parse( autenticarUsuario() );
  } catch( error ) {
    throw( "getRelatorioExel: " + error.message );
  }    
  if( usuarioLogado.tipo != "2" && usuarioLogado.tipo != "0" ) {
    throw( new Error( "Usuário sem permissão para gerar o relatório" ) );
  }    


  try {

    console.log( "getRelatorioExel - Início" );
    limparRelatorio();  
    gerarRelatorio();
    console.log( "getRelatorioExel - Fim" );

    return `https://docs.google.com/spreadsheets/d/${PLANILHA_RELATORIO_ID}/export?format=xlsx`;

  } catch( error ) {
    throw( "getRelatorioExel: " + error.message );
  }    

} // Fim da função getRelatorioExel



