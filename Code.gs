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
  'status', 'diretor', 'dataAprovacao', 'assinatura'
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
  }
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DASHBOARD_PASSWORD')) {
    props.setProperty('DASHBOARD_PASSWORD', 'vegas2026');
  }
  var counter = props.getProperty('CONTADOR');
  if (!counter) props.setProperty('CONTADOR', '0');
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
  sh.appendRow([
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
    ''
  ]);
  return { ok: true, id: id, numero: numero };
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
  var senha = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
  if (data.senha !== senha) return { ok: false, error: 'Senha incorreta.' };

  var sh = sheet_();
  var row = findRow_(sh, data.id);
  if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };

  var now = new Date();
  sh.getRange(row, 12).setValue('Aprovado');       // status
  sh.getRange(row, 13).setValue(data.diretor || ''); // diretor
  sh.getRange(row, 14).setValue(now.toISOString());  // dataAprovacao
  sh.getRange(row, 15).setValue(data.assinatura || ''); // assinatura (dataURL base64)
  return { ok: true };
}

function excluirSolicitacao_(data) {
  var senha = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
  if (data.senha !== senha) return { ok: false, error: 'Senha incorreta.' };

  var sh = sheet_();
  var row = findRow_(sh, data.id);
  if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };
  sh.deleteRow(row);
  return { ok: true };
}
