/**
 * Planilha RELATORIO
 */
const PLANILHA_RELATORIO_ID     =  "1dKDF-AFro9pju-U97LfH6pw5bVsHH-6pvdUzXwq_Wtg";
const PLANILHA_RELATORIO        =  SpreadsheetApp.openById(PLANILHA_RELATORIO_ID);

const TABELA_RELATORIO          =  PLANILHA_RELATORIO.getSheetByName('RELATORIO');
const TABELA_QUANTITATIVOS      =  PLANILHA_RELATORIO.getSheetByName('QUANTITATIVOS');
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

    // SE PEGAR O LOCK, PROSSEGUE A EXCLUSÃO DOS DADOS DO RELATÓRIO
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
function gerarRelatorio( regional ) {

  // Lock
  let lock;

  try {

    // TENTA PEGAR O LOCK
    lock = LockService.getScriptLock();
    lock.waitLock(10000);  

    // SE PEGAR O LOCK, PROSSEGUE COM A GERAÇÃO DOS DADOS DO RELATÓRIO
    if( lock.hasLock() ) {
      
      console.log( "gerarRelatorioFila - Início" );
  
      let caso;
    
      let bufferRelatorioCaso;

      const relatorio = [];
  
      // Percorre todos os casos da fila, gerando o relatório de cada caso
      for( let idCaso=1; idCaso<=NUM_CASOS; ++idCaso ) {
      
        caso = BUFFER_CASOS[idCaso - 1];

        if( caso[REGIONAL] == regional  ||  regional == "0" ) {

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

        }
        
  
      } // Fim do for que percorre todos os casos da fila
  

      // Grava o buffer do relatório na planilha RELATORIO
      TABELA_RELATORIO.getRange( 2, 1, relatorio.length, NUM_COLUNAS_TABELA_RELATORIO ).setValues( relatorio );
      PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
      SpreadsheetApp.flush();  

      calcularQuantitativos( relatorio );
  
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
 * Função que calcula dados quantitativos do relatório
 */
function calcularQuantitativos( relatorio ) {


  console.log( "calcularQuantitativos - Início" );  

  // Números de casos
  let numTotalDeCasos  = relatorio.length;
  let numCasosAtivos   = relatorio.filter( caso => { return caso[15] == "NÃO DESIGNADO" } ).length;
  let numCasosInativos = numTotalDeCasos - numCasosAtivos;
   
   
  // Data do caso ativo mais antigo
  let casosAtivos = relatorio.filter( caso => { return caso[15] == "NÃO DESIGNADO" } );
  let datasCasosAtivos = casosAtivos.map( caso => {
                                             let auxData = caso[10].split("/");
                                             let data = new Date( auxData[2], parseInt(auxData[1])-1, auxData[0] );
                                             return data; 
                                         });
  let dataCasoAtivoMaisAntigo = datasCasosAtivos.reduce( (maisAntiga, atual) => {
                                                          return atual < maisAntiga ? atual : maisAntiga;
                                                       });   
  //let dataFormatada = `${dataCasoAtivoMaisAntigo.getDate()}/${parseInt(dataCasoAtivoMaisAntigo.getMonth())+1}/${dataCasoAtivoMaisAntigo.getFullYear()} `; 
  let dataFormatada = new Date(dataCasoAtivoMaisAntigo).toLocaleString("pt-BR", {dateStyle: "short"}); 
   
   
  // Médias dos tempos de espera
  let mediaTempoEsperaTodosOsCasos  = relatorio.reduce((acumulador, caso) => acumulador + parseInt(caso[17]), 0) / numTotalDeCasos; 
  let mediaTempoEsperaCasosAtivos   = relatorio.filter( caso => { return caso[15] == "NÃO DESIGNADO" } ).reduce((acumulador, caso) => acumulador + parseInt(caso[17]), 0) / numCasosAtivos;
  let mediaTempoEsperaCasosInativos = relatorio.filter( caso => { return caso[15] != "NÃO DESIGNADO" } ).reduce((acumulador, caso) => acumulador + parseInt(caso[17]), 0) / numCasosInativos;
   
   
  // Grava os dados quantitativos na tabela QUANTITATIVOS   
  TABELA_QUANTITATIVOS.getRange( 'B2' ).setValue( numTotalDeCasos );
  TABELA_QUANTITATIVOS.getRange( 'B3' ).setValue( numCasosInativos );
  TABELA_QUANTITATIVOS.getRange( 'B4' ).setValue( numCasosAtivos );
  TABELA_QUANTITATIVOS.getRange( 'B5' ).setValue( dataFormatada );
  TABELA_QUANTITATIVOS.getRange( 'B6' ).setValue( mediaTempoEsperaTodosOsCasos.toFixed(2) );
  TABELA_QUANTITATIVOS.getRange( 'B7' ).setValue( mediaTempoEsperaCasosInativos.toFixed(2) );
  TABELA_QUANTITATIVOS.getRange( 'B8' ).setValue( mediaTempoEsperaCasosAtivos.toFixed(2) );
   
  PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
  SpreadsheetApp.flush();  


  console.log( "calcularQuantitativos - Fim" );

   
} // Fim da função calcularQuantitativos



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

  // Gera e retorna o relatório
  try {

    console.log( "getRelatorioExel - Início" );
    limparRelatorio();  
    gerarRelatorio( usuarioLogado.regional );    
    console.log( "getRelatorioExel - Fim" );

    return `https://docs.google.com/spreadsheets/d/${PLANILHA_RELATORIO_ID}/export?format=xlsx`;

  } catch( error ) {
    throw( "getRelatorioExel: " + error.message );
  }    

} // Fim da função getRelatorioExel



