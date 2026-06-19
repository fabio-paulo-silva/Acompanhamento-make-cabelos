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
  'Com Multimarca',
  'Tipo de Incrementação',
  'Origem',
];

// ----------------------------------------------------------------
// Ponto de entrada — chamado pelo app HTML via JSONP (GET)
// ----------------------------------------------------------------
function doGet(e) {
  // Proteção caso e ou e.parameter seja null
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || '';
  const callback = params.callback || '';

  let resultado;
  try {
    if (action === 'registrar') {
      resultado = registrarPedido(params);
    } else if (action === 'zerar') {
      resultado = zerarRegistros(params);
    } else if (action === 'deletar') {
      resultado = deletarPedido(params);
    } else if (action === 'getCaixas') {
      resultado = getCaixas(params);
    } else if (action === 'getPedidosDia') {
      resultado = getPedidosDia(params);
    } else if (action === 'ping') {
      resultado = { status: 'ok', mensagem: 'Conexão estabelecida com sucesso!' };
    } else if (action === 'debug') {
      // Endpoint de diagnóstico — retorna tudo que o script recebeu
      resultado = { status: 'ok', params: params, keys: Object.keys(params) };
    } else {
      resultado = { status: 'erro', mensagem: 'Ação não reconhecida: "' + action + '"', params_recebidos: params };
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
    aba.setColumnWidth(11, 120); // Com Multimarca
    aba.setColumnWidth(12, 160); // Tipo Incrementação
    aba.setColumnWidth(13, 180); // Origem
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
    p.multimarca === 'true' ? 'SIM' : 'NÃO',
    p.incrementacao || '',
    p.origem       || '',
  ];

  aba.appendRow(linha);

  // Colorir linha conforme incrementação
  const ultimaLinha = aba.getLastRow();
  const range = aba.getRange(ultimaLinha, 1, 1, CABECALHO.length);

  const tipo = p.incrementacao || '';
  if (tipo.includes('Make') && tipo.includes('Cabelo')) {
    range.setBackground('#EAD1FA');
  } else if (tipo.includes('Multimarca')) {
    range.setBackground('#CCFBF1');
  } else if (tipo.includes('Make')) {
    range.setBackground('#FCE4D6');
  } else if (tipo.includes('Cabelo')) {
    range.setBackground('#E2EFDA');
  }

  return { status: 'ok', mensagem: 'Pedido registrado na planilha.' };
}

// ----------------------------------------------------------------
// Deleta um pedido específico pelo ID
// ----------------------------------------------------------------
function deletarPedido(p) {
  const aba = garantirAba();
  const pedidoId = p.pedidoId || '';

  if (!pedidoId) {
    return { status: 'erro', mensagem: 'ID do pedido é obrigatório para deletar.' };
  }

  const dados = aba.getDataRange().getValues();
  for (let i = dados.length - 1; i >= 1; i--) {
    if (String(dados[i][0]) === pedidoId) {
      aba.deleteRow(i + 1);
      return { status: 'ok', mensagem: `Pedido ${pedidoId} removido.` };
    }
  }

  return { status: 'erro', mensagem: `Pedido ${pedidoId} não encontrado.` };
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
// Retorna todos os pedidos de um ER+data a partir de "Registros VD"
// ----------------------------------------------------------------
function getPedidosDia(p) {
  const er = String(p.er || '').trim();
  const dataAcomp = String(p.dataAcomp || '').trim();
  if (!er || !dataAcomp) return { status: 'erro', mensagem: 'ER e dataAcomp são obrigatórios.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(ABA_NOME);
  if (!aba) return { status: 'ok', pedidos: [] };

  const dados = aba.getDataRange().getValues();
  const pedidos = [];
  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    // Coluna B (data) pode ser objeto Date quando a célula tem formato de data no Sheets
    const rowData = row[1] instanceof Date
      ? Utilities.formatDate(row[1], 'America/Fortaleza', 'dd/MM/yyyy')
      : String(row[1]).trim();
    const rowER = String(row[3]).trim();
    if (rowER === er && rowData === dataAcomp) {
      pedidos.push({
        id:          String(row[0]),
        caixa:       String(row[5]),
        numPedido:   String(row[6]),
        hora:        String(row[7]),
        make:        row[8] === 'SIM',
        cabelo:      row[9] === 'SIM',
        multimarca:  row[10] === 'SIM',
        origem:      String(row[12]),
      });
    }
  }
  return { status: 'ok', pedidos, debug_er: er, debug_data: dataAcomp };
}

// ----------------------------------------------------------------
// Retorna lista de caixas de um ER a partir da aba "Caixas"
// ----------------------------------------------------------------
function getCaixas(p) {
  const er = String(p.er || '').trim();
  if (!er) return { status: 'erro', mensagem: 'ER obrigatório.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName('Caixas');
  if (!aba) return { status: 'erro', mensagem: 'Aba "Caixas" não encontrada na planilha.' };

  const dados = aba.getDataRange().getValues();
  const caixas = [];
  for (let i = 1; i < dados.length; i++) {
    const linhaER = String(dados[i][0]).trim();
    const nome    = String(dados[i][1]).trim();
    if (linhaER === er && nome) caixas.push(nome);
  }

  return { status: 'ok', caixas };
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
