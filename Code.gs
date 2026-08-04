/**
 * VEGAS VIGILÂNCIA E SEGURANÇA
 * Backend — Solicitação de Aprovação de Manutenção
 *
 * COMO PUBLICAR:
 * 1. Abra script.google.com/home > Novo Projeto (ou vincule a uma planilha:
 *    Planilha Google > Extensões > Apps Script).
 * 2. Apague o conteúdo padrão do arquivo Code.gs e cole este arquivo inteiro.
 * 3. Rode a função "configurarInicial" uma vez (menu Executar > configurarInicial)
 *    para criar a aba "Solicitacoes" e definir a senha padrão.
 * 4. Implantar > Nova implantação > Tipo: App da Web.
 *    - Executar como: Eu
 *    - Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL do App da Web gerada e cole em API_URL no arquivo index.html.
 * 6. Para trocar a senha do dashboard, rode "definirSenha('novaSenha')" no editor
 *    (Executar > selecionar a função > digitar a senha entre aspas antes de rodar),
 *    ou simplesmente rode configurarInicial novamente e edite a linha da senha.
 */

var SHEET_NAME = 'Solicitacoes';
var HEADERS = [
  'id', 'numero', 'timestamp', 'data', 'solicitante', 'setor', 'veiculo',
  'descricaoJSON', 'pecasJSON', 'maoObraJSON', 'obs',
  'status', 'diretor', 'dataAprovacao', 'assinatura',
  'pecaEscolhidaJSON', 'maoObraEscolhidaJSON'
];

// ---------- SETUP ----------

function configurarInicial() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    migrarCabecalho_(sheet);
  }
  // Trava as colunas de data/hora como TEXTO puro, para o Sheets não converter
  // "2026-07-29" automaticamente num valor de data (o que bagunça fuso horário
  // e formatação quando lido de volta pelo Apps Script).
  sheet.getRange('C2:D1000').setNumberFormat('@');
  sheet.getRange('N2:N1000').setNumberFormat('@');
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DASHBOARD_PASSWORD')) {
    props.setProperty('DASHBOARD_PASSWORD', 'vegas2026');
  }
  var counter = props.getProperty('CONTADOR');
  if (!counter) props.setProperty('CONTADOR', '0');
}

// Garante que planilhas já implantadas anteriormente ganhem as colunas novas
// (ex.: pecaEscolhidaJSON / maoObraEscolhidaJSON) sem apagar dados existentes.
function migrarCabecalho_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var atuais = headerRange.getValues()[0];
  HEADERS.forEach(function (h, i) {
    if (atuais[i] !== h) {
      sheet.getRange(1, i + 1).setValue(h);
    }
  });
}

function definirSenha(novaSenha) {
  PropertiesService.getScriptProperties().setProperty('DASHBOARD_PASSWORD', novaSenha);
}

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  // Se este script estiver vinculado a uma planilha (container-bound), usa ela.
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  // Caso contrário cria uma nova planilha na primeira execução e memoriza o ID.
  var ss = SpreadsheetApp.create('Vegas - Solicitacoes de Manutencao');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// ---------- ENTRY POINTS ----------

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  try {
    if (action === 'list') return respond_(listSolicitacoes_());
    return respond_({ ok: false, error: 'Ação GET desconhecida.' });
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    switch (action) {
      case 'create': return respond_(criarSolicitacao_(data));
      case 'update': return respond_(editarSolicitacao_(data));
      case 'checkPassword': return respond_(checarSenha_(data));
      case 'approve': return respond_(aprovarSolicitacao_(data));
      case 'delete': return respond_(excluirSolicitacao_(data));
      default: return respond_({ ok: false, error: 'Ação POST desconhecida.' });
    }
  } catch (err) {
    return respond_({ ok: false, error: String(err) });
  }
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- HELPERS ----------

function sheet_() {
  configurarInicial();
  return getSpreadsheet_().getSheetByName(SHEET_NAME);
}

function proximoNumero_() {
  var props = PropertiesService.getScriptProperties();
  var atual = parseInt(props.getProperty('CONTADOR') || '0', 10) + 1;
  props.setProperty('CONTADOR', String(atual));
  return ('0000' + atual).slice(-4);
}

function rowToObj_(row, headers) {
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = row[i]; });
  ['descricaoJSON', 'pecasJSON', 'maoObraJSON'].forEach(function (key) {
    var outKey = key.replace('JSON', '');
    try { obj[outKey] = obj[key] ? JSON.parse(obj[key]) : []; }
    catch (e) { obj[outKey] = []; }
  });
  ['pecaEscolhidaJSON', 'maoObraEscolhidaJSON'].forEach(function (key) {
    var outKey = key.replace('JSON', '');
    try { obj[outKey] = obj[key] ? JSON.parse(obj[key]) : null; }
    catch (e) { obj[outKey] = null; }
  });
  return obj;
}

function listSolicitacoes_() {
  var sh = sheet_();
  var values = sh.getDataRange().getValues();
  var headers = values.shift();
  var list = values.map(function (row) { return rowToObj_(row, headers); });
  return { ok: true, items: list };
}

function criarSolicitacao_(data) {
  var sh = sheet_();
  var id = Utilities.getUuid();
  var numero = proximoNumero_();
  var now = new Date();
  var newRow = sh.getLastRow() + 1;

  // Formata a linha inteira como texto ANTES de escrever, para as colunas de
  // data (C = timestamp, D = data) não serem convertidas em valores de data
  // nativos do Sheets (o que causava o bug de fuso horário / texto quebrado).
  sh.getRange(newRow, 3, 1, 2).setNumberFormat('@');

  sh.getRange(newRow, 1, 1, HEADERS.length).setValues([[
    id,
    numero,
    now.toISOString(),
    data.data || '',
    data.solicitante || '',
    data.setor || '',
    data.veiculo || '',
    JSON.stringify(data.descricao || []),
    JSON.stringify(data.pecas || []),
    JSON.stringify(data.maoObra || []),
    data.obs || '',
    'Pendente',
    '',
    '',
    '',
    '',
    ''
  ]]);
  return { ok: true, id: id, numero: numero };
}

// Permite que o próprio solicitante edite a solicitação enquanto ela ainda
// não tiver sido aprovada. Depois de aprovada, os dados ficam travados (o
// registro passa a ser um documento assinado pela diretoria).
function editarSolicitacao_(data) {
  var sh = sheet_();
  var row = findRow_(sh, data.id);
  if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };

  var status = sh.getRange(row, 12).getValue(); // coluna 'status'
  if (status !== 'Pendente') {
    return { ok: false, error: 'Esta solicitação já foi aprovada e não pode mais ser editada.' };
  }

  sh.getRange(row, 4, 1, 1).setNumberFormat('@'); // coluna 'data' como texto
  sh.getRange(row, 4, 1, 4).setValues([[
    data.data || '',
    data.solicitante || '',
    data.setor || '',
    data.veiculo || ''
  ]]);
  sh.getRange(row, 8, 1, 3).setValues([[
    JSON.stringify(data.descricao || []),
    JSON.stringify(data.pecas || []),
    JSON.stringify(data.maoObra || [])
  ]]);
  sh.getRange(row, 11).setValue(data.obs || '');
  return { ok: true };
}

function checarSenha_(data) {
  var senha = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
  return { ok: data.senha === senha };
}

function findRow_(sh, id) {
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === id) return i + 1; // linha real na planilha (1-indexed)
  }
  return -1;
}

function aprovarSolicitacao_(data) {
  var sh = sheet_();
  var row = findRow_(sh, data.id);
  if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };

  var now = new Date();
  sh.getRange(row, 14).setNumberFormat('@'); // trava dataAprovacao como texto
  sh.getRange(row, 12).setValue('Aprovado');       // status
  sh.getRange(row, 13).setValue(data.diretor || ''); // diretor
  sh.getRange(row, 14).setValue(now.toISOString());  // dataAprovacao
  sh.getRange(row, 15).setValue(data.assinatura || ''); // assinatura (dataURL base64)
  sh.getRange(row, 16).setValue(JSON.stringify(data.pecaEscolhida || null));    // fornecedor de peça escolhido
  sh.getRange(row, 17).setValue(JSON.stringify(data.maoObraEscolhida || null)); // fornecedor de mão de obra escolhido
  return { ok: true };
}

function excluirSolicitacao_(data) {
  var sh = sheet_();
  var row = findRow_(sh, data.id);
  if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };

  var status = sh.getRange(row, 12).getValue(); // coluna 'status'
  if (status === 'Aprovado') {
    // Só exige a senha da diretoria quando a solicitação já foi assinada/aprovada.
    var senha = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
    if (data.senha !== senha) return { ok: false, error: 'Senha incorreta.' };
  }

  sh.deleteRow(row);
  return { ok: true };
}
