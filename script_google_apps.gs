// ================================================================
// ACOMPANHAMENTO VD MOSSORÓ — Google Apps Script
// Cole este código em: Planilha > Extensões > Apps Script
// Depois publique como Web App (instruções no final do arquivo)
// ================================================================

const ABA_NOME = 'Registros VD';

const CABECALHO = [
  'ID do Pedido',
  'Data do Acompanhamento',
  'Data/Hora do Registro',
  'Código ER',
  'Nome da Loja',
  'Nome da Caixa',
  'Nº do Pedido',
  'Hora do Pedido',
  'Com Make',
  'Com Cabelo',
  'Tipo de Incrementação',
  'Meta Make (%)',
  'Meta Cabelo (%)',
];

// ----------------------------------------------------------------
// Ponto de entrada — chamado pelo app HTML via JSONP (GET)
// ----------------------------------------------------------------
function doGet(e) {
  const params = e.parameter;
  const action = params.action;
  const callback = params.callback || '';

  let resultado;
  try {
    if (action === 'registrar') {
      resultado = registrarPedido(params);
    } else if (action === 'zerar') {
      resultado = zerarRegistros(params);
    } else if (action === 'ping') {
      resultado = { status: 'ok', mensagem: 'Conexão estabelecida com sucesso!' };
    } else {
      resultado = { status: 'erro', mensagem: 'Ação não reconhecida: ' + action };
    }
  } catch (err) {
    resultado = { status: 'erro', mensagem: err.message };
  }

  // JSONP: envolve o JSON numa função de callback se solicitado
  const json = JSON.stringify(resultado);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// Garante que a aba existe e tem cabeçalho
// ----------------------------------------------------------------
function garantirAba() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba = ss.getSheetByName(ABA_NOME);

  if (!aba) {
    aba = ss.insertSheet(ABA_NOME);
    const cabLine = aba.getRange(1, 1, 1, CABECALHO.length);
    cabLine.setValues([CABECALHO]);
    cabLine.setFontWeight('bold');
    cabLine.setBackground('#1F4E79');
    cabLine.setFontColor('#FFFFFF');
    aba.setFrozenRows(1);
    aba.setColumnWidth(1, 180);  // ID
    aba.setColumnWidth(2, 140);  // Data Acompanhamento
    aba.setColumnWidth(3, 160);  // Data/Hora Registro
    aba.setColumnWidth(4, 100);  // Código ER
    aba.setColumnWidth(5, 180);  // Nome Loja
    aba.setColumnWidth(6, 140);  // Nome Caixa
    aba.setColumnWidth(7, 120);  // Nº Pedido
    aba.setColumnWidth(8, 100);  // Hora Pedido
    aba.setColumnWidth(9, 100);  // Com Make
    aba.setColumnWidth(10, 100); // Com Cabelo
    aba.setColumnWidth(11, 160); // Tipo Incrementação
    aba.setColumnWidth(12, 110); // Meta Make
    aba.setColumnWidth(13, 110); // Meta Cabelo
  }

  return aba;
}

// ----------------------------------------------------------------
// Registra um pedido — adiciona nova linha
// ----------------------------------------------------------------
function registrarPedido(p) {
  const aba = garantirAba();

  const dataRegistro = Utilities.formatDate(
    new Date(), 'America/Fortaleza', 'dd/MM/yyyy HH:mm:ss'
  );

  const linha = [
    p.id            || '',
    p.dataAcomp     || '',
    dataRegistro,
    p.er            || '',
    p.loja          || '',
    p.caixa         || '',
    p.numPedido     || '',
    p.hora          || '',
    p.make === 'true' ? 'SIM' : 'NÃO',
    p.cabelo === 'true' ? 'SIM' : 'NÃO',
    p.incrementacao || '',
    p.metaMake      || '',
    p.metaCabelo    || '',
  ];

  aba.appendRow(linha);

  // Colorir linha conforme incrementação
  const ultimaLinha = aba.getLastRow();
  const range = aba.getRange(ultimaLinha, 1, 1, CABECALHO.length);

  const tipo = p.incrementacao || '';
  if (tipo === 'Make+Cabelo') {
    range.setBackground('#EAD1FA');
  } else if (tipo === 'Make') {
    range.setBackground('#FCE4D6');
  } else if (tipo === 'Cabelo') {
    range.setBackground('#E2EFDA');
  }

  return { status: 'ok', mensagem: 'Pedido registrado na planilha.' };
}

// ----------------------------------------------------------------
// Zera registros de um ER+data específico
// ----------------------------------------------------------------
function zerarRegistros(p) {
  const aba = garantirAba();
  const er = p.er || '';
  const dataAcomp = p.dataAcomp || '';

  if (!er || !dataAcomp) {
    return { status: 'erro', mensagem: 'ER e data são obrigatórios para zerar.' };
  }

  const dados = aba.getDataRange().getValues();
  // Percorre de baixo pra cima para deletar sem deslocar índices
  let removidos = 0;
  for (let i = dados.length - 1; i >= 1; i--) {
    const linhaER = String(dados[i][3]);       // coluna D — Código ER
    const linhaData = String(dados[i][1]);     // coluna B — Data Acompanhamento
    if (linhaER === er && linhaData === dataAcomp) {
      aba.deleteRow(i + 1); // +1 porque array é 0-based mas sheet é 1-based
      removidos++;
    }
  }

  return {
    status: 'ok',
    mensagem: `${removidos} registro(s) removido(s) para ER ${er} em ${dataAcomp}.`
  };
}

// ----------------------------------------------------------------
// Monta resposta JSON com header CORS
// ----------------------------------------------------------------
function resposta(obj) {
  const output = ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ----------------------------------------------------------------
// Handler OPTIONS para preflight CORS (navegadores modernos)
// ----------------------------------------------------------------
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ================================================================
// COMO PUBLICAR ESTE SCRIPT COMO WEB APP
// ================================================================
// 1. Na planilha, clique em "Extensões" > "Apps Script"
// 2. Cole TODO este código substituindo o que tiver lá
// 3. Clique em "Salvar" (ícone de disquete ou Ctrl+S)
// 4. Clique em "Implantar" > "Nova implantação"
// 5. Clique no ícone de engrenagem ao lado de "Tipo" > "App da Web"
// 6. Configure:
//    - Descrição: Acompanhamento VD Mossoró
//    - Executar como: EU MESMO (sua conta Google)
//    - Quem pode acessar: QUALQUER PESSOA
//      (a URL longa já serve como senha — só quem tiver ela acessa)
// 7. Clique "Implantar" > autorize as permissões
// 8. COPIE a URL gerada (começa com https://script.google.com/macros/s/...)
// 9. Cole essa URL no arquivo HTML no campo SCRIPT_URL
// ================================================================
