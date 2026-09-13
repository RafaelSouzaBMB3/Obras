/*******************************************************
 * DASHBOARD DE OBRAS
 * GOOGLE DRIVE + EXCEL + HTML
 *******************************************************/


/*******************************************************
 * CONFIGURAÇÕES
 *******************************************************/

// NOVA PASTA DO GOOGLE DRIVE
const ID_PASTA_DRIVE =
  '1XtRXl2Tqf_Lg78iWIFd1n_r1SrbFnAx2';


// Nome da aba dentro do Excel
const NOME_ABA =
  'Planilha1';


/*******************************************************
 * ABRE O INDEX.HTML
 *******************************************************/

function doGet(e) {

  const parametros =
    e && e.parameter
      ? e.parameter
      : {};

  /*
   * API JSONP para a versão externa do Dashboard Obras.
   *
   * O HTML de apresentação roda no GitHub Pages e, por isso,
   * não consegue usar google.script.run.
   */
  if (String(parametros.api || '').toLowerCase() === 'obras') {

    const callbackBruto =
      String(parametros.callback || 'callback');

    /*
     * Aceita somente nomes de callback JavaScript simples,
     * evitando inserir conteúdo arbitrário na resposta.
     */
    const callback =
      /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callbackBruto)
        ? callbackBruto
        : 'callback';

    try {

      const dados =
        obterObras();

      return ContentService
        .createTextOutput(
          callback + '(' + JSON.stringify(dados) + ');'
        )
        .setMimeType(
          ContentService.MimeType.JAVASCRIPT
        );

    } catch (erro) {

      console.error(
        'API OBRAS — erro: ' +
        erro.message
      );

      return ContentService
        .createTextOutput(
          callback +
          '(' +
          JSON.stringify({
            erro: erro.message || String(erro)
          }) +
          ');'
        )
        .setMimeType(
          ContentService.MimeType.JAVASCRIPT
        );
    }
  }

  /*
   * Abertura normal do Dashboard Obras no Apps Script.
   * Mantida para não quebrar o acesso atual.
   */
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Controle de Obras')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/*******************************************************
 * LOCALIZA O EXCEL
 *
 * Procura o arquivo .xlsx dentro da pasta
 * e também dentro das subpastas.
 *******************************************************/

function localizarArquivo() {

  const pasta =
    DriveApp.getFolderById(
      ID_PASTA_DRIVE
    );


  const arquivo =
    procurarExcelNaPasta(
      pasta
    );


  if (!arquivo) {

    throw new Error(
      'Nenhum arquivo Excel (.xlsx) foi encontrado na pasta do Google Drive.'
    );

  }


  console.log(
    'Arquivo encontrado: ' +
    arquivo.getName()
  );


  console.log(
    'ID do arquivo: ' +
    arquivo.getId()
  );


  return arquivo;

}


/*******************************************************
 * PROCURA EXCEL RECURSIVAMENTE
 *******************************************************/

function procurarExcelNaPasta(
  pasta
) {

  /*
   * Primeiro procura arquivos
   * diretamente nesta pasta.
   */

  const arquivos =
    pasta.getFiles();


  while (
    arquivos.hasNext()
  ) {

    const arquivo =
      arquivos.next();


    const nome =
      arquivo
        .getName()
        .toLowerCase();


    if (
      nome.endsWith('.xlsx')
    ) {

      return arquivo;

    }

  }


  /*
   * Depois procura nas subpastas.
   */

  const subpastas =
    pasta.getFolders();


  while (
    subpastas.hasNext()
  ) {

    const subpasta =
      subpastas.next();


    const resultado =
      procurarExcelNaPasta(
        subpasta
      );


    if (
      resultado
    ) {

      return resultado;

    }

  }


  return null;

}


/*******************************************************
 * LÊ O EXCEL
 *******************************************************/

function obterObras() {

  const arquivo =
    localizarArquivo();


  let arquivoTemporario =
    null;


  try {

    console.log(
      'Lendo arquivo: ' +
      arquivo.getName()
    );


    /*
     * Obtém o conteúdo do Excel
     */

    const blob =
      arquivo.getBlob();


    /*
     * Cria uma conversão temporária
     * para Google Sheets.
     *
     * O Excel original NÃO é alterado.
     */

    const recurso = {

      name:
        '[TEMP] ' +
        arquivo.getName(),

      mimeType:
        'application/vnd.google-apps.spreadsheet'

    };


    arquivoTemporario =
      Drive.Files.create(
        recurso,
        blob,
        {
          fields:
            'id,name'
        }
      );


    console.log(
      'Arquivo temporário criado: ' +
      arquivoTemporario.id
    );


    /*
     * Abre a planilha temporária
     */

    const planilha =
      SpreadsheetApp.openById(
        arquivoTemporario.id
      );


    /*
     * Localiza a aba
     */

    const aba =
      planilha.getSheetByName(
        NOME_ABA
      );


    if (!aba) {

      throw new Error(
        'A aba "' +
        NOME_ABA +
        '" não foi encontrada no Excel.'
      );

    }


    /*
     * Lê todos os dados
     */

    const dados =
      aba
        .getDataRange()
        .getValues();


    console.log(
      'Linhas encontradas: ' +
      dados.length
    );


    if (
      dados.length < 3
    ) {

      return [];

    }


    /***************************************************
     * CABEÇALHO
     *
     * Linha 3 = cabeçalho
     * Linha 4 em diante = dados
     ***************************************************/

    const cabecalho =
      dados[2].map(
        function(valor) {

          return normalizar(
            valor
          );

        }
      );


    console.log(
      'Cabeçalho: ' +
      JSON.stringify(
        cabecalho
      )
    );


    /***************************************************
     * LOCALIZA AS COLUNAS
     ***************************************************/

    const colunas = {

      cs:
        cabecalho.indexOf(
          'CS'
        ),

      tipo:
        cabecalho.indexOf(
          'TIPO'
        ),

      codigo:
        cabecalho.indexOf(
          'CODIGO'
        ),

      local:
        cabecalho.indexOf(
          'LOCAL'
        ),

      etapa:
        cabecalho.indexOf(
          'ETAPA'
        ),

      responsavel:
        cabecalho.indexOf(
          'RESPONSAVEL'
        ),

      status:
        cabecalho.indexOf(
          'STATUS'
        )

    };


    /***************************************************
     * CONFERE AS COLUNAS
     ***************************************************/

    for (
      const campo in colunas
    ) {

      if (
        colunas[campo] === -1
      ) {

        throw new Error(
          'A coluna "' +
          campo +
          '" não foi encontrada na linha 3 do Excel.'
        );

      }

    }


    console.log(
      'Colunas encontradas com sucesso.'
    );


    /***************************************************
     * CRIA LISTA
     ***************************************************/

    const obras = [];


    /***************************************************
     * LÊ AS LINHAS
     ***************************************************/

    for (
      let linha = 3;
      linha < dados.length;
      linha++
    ) {

      const registro =
        dados[linha];


      /*
       * Ignora linha vazia
       */

      if (
        registro.every(
          function(valor) {

            return (
              valor === '' ||
              valor === null ||
              valor === undefined
            );

          }
        )
      ) {

        continue;

      }


      /*************************************************
       * MONTA OBJETO DA OBRA
       *************************************************/

      const obra = {

        cs:
          texto(
            registro[
              colunas.cs
            ]
          ),

        tipo:
          texto(
            registro[
              colunas.tipo
            ]
          ),

        codigo:
          texto(
            registro[
              colunas.codigo
            ]
          ),

        local:
          texto(
            registro[
              colunas.local
            ]
          ),

        etapa:
          texto(
            registro[
              colunas.etapa
            ]
          ),

        responsavel:
          texto(
            registro[
              colunas.responsavel
            ]
          ),

        status:
          normalizar(
            registro[
              colunas.status
            ]
          )

      };


      /*************************************************
       * FINALIZADO
       *
       * NÃO aparece no painel.
       *************************************************/

      if (
        obra.status ===
        'FINALIZADO'
      ) {

        continue;

      }


      /*************************************************
       * SOMENTE STATUS VÁLIDOS
       *************************************************/

      if (
        obra.status !==
        'EM EXECUCAO'
        &&
        obra.status !==
        'EM ORCAMENTO'
      ) {

        continue;

      }


      /*************************************************
       * ADICIONA À LISTA
       *************************************************/

      obras.push(
        obra
      );

    }


    console.log(
      'Total de obras enviadas ao HTML: ' +
      obras.length
    );


    return obras;


  }

  finally {

    /***************************************************
     * REMOVE A CÓPIA TEMPORÁRIA
     ***************************************************/

    if (
      arquivoTemporario &&
      arquivoTemporario.id
    ) {

      try {

        Drive.Files.delete(
          arquivoTemporario.id
        );


        console.log(
          'Arquivo temporário excluído.'
        );

      }

      catch (erro) {

        console.log(
          'Não foi possível excluir o temporário: ' +
          erro.message
        );

      }

    }

  }

}


/*******************************************************
 * NORMALIZA TEXTO
 *******************************************************/

function normalizar(
  valor
) {

  return String(
    valor === null ||
    valor === undefined
      ? ''
      : valor
  )
  .normalize('NFD')
  .replace(
    /[\u0300-\u036f]/g,
    ''
  )
  .toUpperCase()
  .trim();

}


/*******************************************************
 * CONVERTE PARA TEXTO
 *******************************************************/

function texto(
  valor
) {

  if (
    valor === null ||
    valor === undefined
  ) {

    return '';

  }


  return String(
    valor
  ).trim();

}


/*******************************************************
 * TESTE DE LEITURA
 *******************************************************/

function testeObterObras() {

  console.log(
    '========================================'
  );

  console.log(
    'INICIANDO TESTE'
  );

  console.log(
    '========================================'
  );


  const obras =
    obterObras();


  console.log(
    '========================================'
  );

  console.log(
    'TOTAL: ' +
    obras.length
  );

  console.log(
    '========================================'
  );


  obras.forEach(
    function(
      obra,
      indice
    ) {

      console.log(
        (indice + 1) +
        ' | ' +
        obra.cs +
        ' | ' +
        obra.tipo +
        ' | ' +
        obra.codigo +
        ' | ' +
        obra.local +
        ' | ' +
        obra.etapa +
        ' | ' +
        obra.responsavel +
        ' | ' +
        obra.status
      );

    }
  );


  return obras;

}