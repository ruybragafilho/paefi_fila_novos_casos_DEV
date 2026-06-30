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

  try {

    console.log( "limparRelatorio - Início" );

    // Caso nulo
    let casoNulo = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");
  
    // Limpa a fila
    let range;
    for( let linha=2; linha<=NUM_RELATORIOS+1; ++linha ) {
      range = TABELA_RELATORIO.getRange( linha, 1, 1, NUM_COLUNAS_TABELA_RELATORIO );
      range.setValues([casoNulo]);
    }    
    
    PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
    SpreadsheetApp.flush();  

    console.log( "limparRelatorio - Fim" );

  } catch( error ) {
    throw( "limparRelatorio: " + error.message );
  }    

} // Fim da função limparRelatorio



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * gravando-os na planilha RELATORIO
 */
function gerarRelatorioFila(idInicio, idFim) {

  try {

    console.log( "gerarRelatorioFila - Início" );

    let caso;
  
    const relatorioCaso = new Array(NUM_COLUNAS_TABELA_RELATORIO).fill("");

    // Percorre todos os casos da fila, gerando o relatório de cada caso
    for( let idCaso=idInicio; idCaso<idFim; ++idCaso ) {
    
      caso = BUFFER_CASOS[idCaso - 1];
  
      relatorioCaso[0] = caso[ ID ];
      relatorioCaso[1] = caso[ REFERENCIA_FAMILIAR ];
      relatorioCaso[2] = caso[ TIPO_LOGRADOURO ];
      relatorioCaso[3] = caso[ NOME_LOGRADOURO ];
      relatorioCaso[4] = caso[ NUMERO ];
      relatorioCaso[5] = caso[ COMPLEMENTO ];
      relatorioCaso[6] = caso[ BAIRRO ];
      relatorioCaso[7] = idsToNomes( caso[ REGIONAL ], "REGIONAIS" );
      relatorioCaso[8] = caso[ CEP ];
    
      relatorioCaso[9] = caso[ TPSA ];
      relatorioCaso[10] = caso[ DATA_DE_CHEGADA_NO_CREAS ];
    
      let orgaosEncaminhadores =  idsToNomes( caso[ ORGAOS_ENCAMINHADORES ], "ORGAOS_ENCAMINHADORES" );
      if( orgaosEncaminhadores != "" ) {
        orgaosEncaminhadores = orgaosEncaminhadores.replaceAll( ";", ";\n" );
      } else {
        orgaosEncaminhadores = "SEM INFORMAÇÃO";
      }  
      relatorioCaso[11] = orgaosEncaminhadores;
    
      relatorioCaso[12] = caso[ DATA_PREVISTA_PARA_RESPOSTA ];
      relatorioCaso[13] = caso[ DATA_DA_ULTIMA_RESPOSTA ];  
      relatorioCaso[14] = caso[ DATA_DE_DESIGNACAO ];
  
      relatorioCaso[15] = caso[ MOTIVO_DE_DESIGNACAO ] != "" ? 
                          idsToNomes( caso[ MOTIVO_DE_DESIGNACAO ], "MOTIVOS_DE_DESIGNACAO" ) : 
                          "NÃO DESIGNADO";
    
      relatorioCaso[16] = caso[ TOTAL_DE_PONTOS ];
      relatorioCaso[17] = caso[ TEMPO_DE_ESPERA ];
    
      let violacoes =  idsToNomes( caso[ VIOLACOES_CASO ], "VIOLACOES" );
      if( violacoes != "" ) {
        violacoes = violacoes.replaceAll( ";", ";\n" );
      } else {
        violacoes = "SEM INFORMAÇÃO";
      }
      relatorioCaso[18] = violacoes;
    
      let categorias =  idsToNomes( caso[ CATEGORIAS_CASO ], "CATEGORIAS" );
      if( categorias != "" ) {
        categorias = categorias.replaceAll( ";", ";\n" );
      } else {
        categorias = "SEM INFORMAÇÃO";
      }  
      relatorioCaso[19] = categorias;
    
      relatorioCaso[20] = caso[ PONTUACAO_PARAMETROS_CASO ];  
    
      let parametros =  idsToNomes( caso[ PARAMETROS_CASO ], "PARAMETROS" );
      if( parametros != "" ) {
        parametros = parametros.replaceAll( ";", ";\n" );
      } else {
        parametros = "SEM INFORMAÇÃO";
      }
      relatorioCaso[21] = parametros;  
    
      relatorioCaso[22] = caso[ OBSERVACAO ];    
      relatorioCaso[23] = caso[ ID_TECNICO_PAEFI ] != "" ? 
                          idsToNomes( caso[ ID_TECNICO_PAEFI ], "TECNICOS" ) : 
                          "SEM INFORMAÇÃO";
      
      TABELA_RELATORIO.appendRow( relatorioCaso );

    } // Fim do for que percorre todos os casos da fila

    PLANILHA_RELATORIO.waitForAllDataExecutionsCompletion(2);      
    SpreadsheetApp.flush();  

    console.log( "gerarRelatorioFila - Fim" );


  } catch( error ) {
    throw( "gerarRelatorioFila: " + error.message );
  }    


} // Fim da função gerarRelatorioFila


function gerarRelatorio_1_500() {
  gerarRelatorioFila( 1, 500 );
}

function gerarRelatorio_501_1000() {
  gerarRelatorioFila( 500, 1000 );
}

function gerarRelatorio_1001_1500() {
  gerarRelatorioFila( 1000, 1500 );
}

function gerarRelatorio_1501_NUM_CASOS() {
  gerarRelatorioFila( 1500, NUM_CASOS );
}



/**
 * Função que gera os relatórios de todos os casos da fila, 
 * e o retorna em excel
 */
function getRelatorioExel() {

  try {

    console.log( "getRelatorioExel - Início" );
    limparRelatorio();  
    gerarRelatorioFilaBE();
    console.log( "getRelatorioExel - Fim" );

  } catch( error ) {
    throw( "getRelatorioExel: " + error.message );
  }    

} // Fim da função getRelatorioExel