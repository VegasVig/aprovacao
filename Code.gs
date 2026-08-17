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
  'pecaEscolhidaJSON', 'maoObraEscolhidaJSON', 'tratativasJSON',
  'placa', 'km'
];

// Estados possíveis:
//   'Pendente'  -> bola com a diretoria (aguardando análise/assinatura)
//   'Em Ajuste' -> diretoria devolveu com apontamento, bola com o solicitante
//   'Aprovado'  -> assinado; vai para o dashboard e fica travado
var ST_PENDENTE = 'Pendente';
var ST_AJUSTE = 'Em Ajuste';
var ST_APROVADO = 'Aprovado';

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
      case 'devolver': return respond_(devolverSolicitacao_(data));
      case 'responder': return respond_(responderSolicitacao_(data));
      case 'delete': return respond_(excluirSolicitacao_(data));
      case 'arquivarPdf': return respond_(arquivarPdf_(data));
      case 'uploadComprovante': return respond_(uploadComprovante_(data));
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
  try { obj.tratativas = obj.tratativasJSON ? JSON.parse(obj.tratativasJSON) : []; }
  catch (e) { obj.tratativas = []; }
  return obj;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function lerTratativas_(sh, row) {
  var raw = sh.getRange(row, 18).getValue();
  try { return raw ? JSON.parse(raw) : []; } catch (e) { return []; }
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
    ST_PENDENTE,
    '',
    '',
    '',
    '',
    '',
    '[]',
    data.placa || '',
    data.km || ''
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
  if (status === ST_APROVADO) {
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
  if (data.placa !== undefined) sh.getRange(row, 19).setValue(data.placa || '');
  if (data.km !== undefined) { sh.getRange(row, 20).setNumberFormat('@'); sh.getRange(row, 20).setValue(data.km || ''); }
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
  return withLock_(function () {
    var sh = sheet_();
    var row = findRow_(sh, data.id);
    if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };
    if (sh.getRange(row, 12).getValue() === ST_APROVADO) {
      return { ok: false, error: 'Esta solicitação já foi aprovada e assinada.' };
    }

    var now = new Date();
    var trilha = lerTratativas_(sh, row);
    trilha.push({
      autor: 'diretor',
      tipo: 'aprovacao',
      nome: data.diretor || '',
      texto: data.textoAprovacao || 'De acordo. Orçamento aprovado e assinado.',
      timestamp: now.toISOString()
    });

    sh.getRange(row, 14).setNumberFormat('@'); // trava dataAprovacao como texto
    sh.getRange(row, 12).setValue(ST_APROVADO);        // status
    sh.getRange(row, 13).setValue(data.diretor || ''); // diretor
    sh.getRange(row, 14).setValue(now.toISOString());  // dataAprovacao
    sh.getRange(row, 15).setValue(data.assinatura || ''); // assinatura (dataURL base64)
    sh.getRange(row, 16).setValue(JSON.stringify(data.pecaEscolhida || null));
    sh.getRange(row, 17).setValue(JSON.stringify(data.maoObraEscolhida || null));
    sh.getRange(row, 18).setValue(JSON.stringify(trilha));
    return { ok: true };
  });
}

// Diretoria em desacordo: registra o apontamento e devolve a bola ao solicitante.
function devolverSolicitacao_(data) {
  var texto = String(data.texto || '').trim();
  if (!texto) return { ok: false, error: 'Descreva o motivo do desacordo.' };
  if (!String(data.diretor || '').trim()) return { ok: false, error: 'Informe o nome do diretor(a).' };

  return withLock_(function () {
    var sh = sheet_();
    var row = findRow_(sh, data.id);
    if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };
    if (sh.getRange(row, 12).getValue() === ST_APROVADO) {
      return { ok: false, error: 'Solicitação já aprovada — não pode mais ser devolvida.' };
    }

    var trilha = lerTratativas_(sh, row);
    trilha.push({
      autor: 'diretor',
      tipo: 'apontamento',
      nome: data.diretor,
      texto: texto,
      timestamp: new Date().toISOString()
    });

    sh.getRange(row, 12).setValue(ST_AJUSTE);
    sh.getRange(row, 13).setValue(data.diretor);
    sh.getRange(row, 18).setValue(JSON.stringify(trilha));
    notificar_(rowToObj_(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0], HEADERS), 'apontamento');
    return { ok: true, rodada: trilha.length };
  });
}

// Solicitante responde ao apontamento e devolve a bola à diretoria.
// Pode anexar novos orçamentos na mesma ação (caso clássico: "traga outra cotação").
function responderSolicitacao_(data) {
  var texto = String(data.texto || '').trim();
  if (!texto) return { ok: false, error: 'Escreva a resposta à diretoria.' };

  return withLock_(function () {
    var sh = sheet_();
    var row = findRow_(sh, data.id);
    if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };
    var status = sh.getRange(row, 12).getValue();
    if (status === ST_APROVADO) return { ok: false, error: 'Solicitação já aprovada.' };
    if (status !== ST_AJUSTE) return { ok: false, error: 'Esta solicitação já está com a diretoria.' };

    var novos = Array.isArray(data.novos) ? data.novos : [];
    var anexados = [];
    if (novos.length) {
      var pecas = parseArr_(sh.getRange(row, 9).getValue());
      var mo = parseArr_(sh.getRange(row, 10).getValue());
      novos.forEach(function (n) {
        if (!n || !String(n.nome || '').trim()) return;
        var item = { nome: String(n.nome).trim(), valor: n.valor || 0, status: n.status || 'ok', comprovante: String(n.comprovante || '').trim() };
        if (n.tipo === 'maoObra') mo.push(item); else pecas.push(item);
        anexados.push((n.tipo === 'maoObra' ? 'Mão de obra' : 'Peça') + ': ' + item.nome + ' — R$ ' + item.valor);
      });
      sh.getRange(row, 9).setValue(JSON.stringify(pecas));
      sh.getRange(row, 10).setValue(JSON.stringify(mo));
    }

    var trilha = lerTratativas_(sh, row);
    trilha.push({
      autor: 'solicitante',
      tipo: 'resposta',
      nome: String(data.solicitante || sh.getRange(row, 5).getValue() || ''),
      texto: texto,
      anexos: anexados,
      timestamp: new Date().toISOString()
    });

    sh.getRange(row, 12).setValue(ST_PENDENTE);
    sh.getRange(row, 18).setValue(JSON.stringify(trilha));
    notificar_(rowToObj_(sh.getRange(row, 1, 1, HEADERS.length).getValues()[0], HEADERS), 'resposta');
    return { ok: true, rodada: trilha.length };
  });
}

function parseArr_(raw) {
  try { var a = raw ? JSON.parse(raw) : []; return Array.isArray(a) ? a : []; }
  catch (e) { return []; }
}

// ---------- ARQUIVAMENTO DO PDF NO DRIVE ----------
// O WhatsApp (wa.me) não aceita anexo — só texto. Então o PDF vai para uma pasta
// do Drive e o que viaja na mensagem é o link. O nível de compartilhamento é
// controlado pela propriedade PDF_ACESSO:
//   'LINK'     -> qualquer pessoa com o link vê (padrão; necessário para fornecedor externo)
//   'DOMINIO'  -> só quem tem e-mail do domínio Workspace da empresa
//   'PRIVADO'  -> só quem você compartilhar manualmente (o link exige login autorizado)
function pastaPdf_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PASTA_PDF_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* pasta apagada: recria */ }
  }
  var nome = 'Vegas - Solicitacoes Aprovadas (PDF)';
  var it = DriveApp.getFoldersByName(nome);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(nome);
  props.setProperty('PASTA_PDF_ID', folder.getId());
  return folder;
}

function arquivarPdf_(data) {
  if (!data.base64) return { ok: false, error: 'PDF vazio.' };
  var nome = data.filename || ('solicitacao_' + (data.numero || 'sn') + '.pdf');
  var folder = pastaPdf_();

  // Substitui a versão anterior do mesmo número em vez de acumular duplicatas.
  var antigos = folder.getFilesByName(nome);
  while (antigos.hasNext()) antigos.next().setTrashed(true);

  var blob = Utilities.newBlob(Utilities.base64Decode(data.base64), 'application/pdf', nome);
  var file = folder.createFile(blob);

  var acesso = PropertiesService.getScriptProperties().getProperty('PDF_ACESSO') || 'LINK';
  try {
    if (acesso === 'LINK') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    else if (acesso === 'DOMINIO') file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* conta pessoal não tem DOMAIN_WITH_LINK; mantém privado */ }

  // Registra o envio na trilha — quem mandou, para qual número, com qual chave.
  if (data.id && data.envio) {
    try {
      withLock_(function () {
        var sh = sheet_();
        var row = findRow_(sh, data.id);
        if (row === -1) return;
        var trilha = lerTratativas_(sh, row);
        var e = data.envio;
        trilha.push({
          autor: 'diretor',
          tipo: 'envio',
          nome: e.remetente || sh.getRange(row, 13).getValue() || '',
          texto: 'PDF enviado por WhatsApp para ' + (e.whatsapp || '—') +
                 '. Favorecido PIX: ' + (e.favorecido || '—') +
                 ' | chave: ' + mascararChave_(e.chave) +
                 ' | valor: R$ ' + (e.valor || '0'),
          anexos: [file.getUrl()],
          timestamp: new Date().toISOString()
        });
        sh.getRange(row, 18).setValue(JSON.stringify(trilha));
      });
    } catch (err) { /* não impede a devolução do link */ }
  }

  return { ok: true, url: file.getUrl(), id: file.getId(), acesso: acesso };
}

// Nunca gravar chave PIX inteira em log: mantém só o suficiente para conferência.
function mascararChave_(chave) {
  var s = String(chave || '');
  if (s.length <= 6) return s ? s.charAt(0) + '***' : '—';
  return s.slice(0, 3) + '***' + s.slice(-3);
}

// ---------- COMPROVANTE DE ORÇAMENTO (foto ou upload do solicitante) ----------
// O solicitante pode anexar uma foto do orçamento ao lado do valor de cada
// fornecedor (peça ou mão de obra). A foto vai para uma pasta própria do Drive
// e o que fica salvo no item é só o link — o mesmo campo aceita colar um link
// já pronto (ex.: PDF de orçamento no Drive/WhatsApp) em vez de subir foto.
function pastaComprovantes_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PASTA_COMPROVANTES_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* pasta apagada: recria */ }
  }
  var nome = 'Vegas - Comprovantes de Orcamento';
  var it = DriveApp.getFoldersByName(nome);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(nome);
  props.setProperty('PASTA_COMPROVANTES_ID', folder.getId());
  return folder;
}

function uploadComprovante_(data) {
  if (!data.base64) return { ok: false, error: 'Arquivo vazio.' };
  var mimeType = data.mimeType || 'image/jpeg';
  var nome = data.filename || ('comprovante_' + Utilities.getUuid() + '.jpg');
  var folder = pastaComprovantes_();

  var blob;
  try {
    blob = Utilities.newBlob(Utilities.base64Decode(data.base64), mimeType, nome);
  } catch (e) {
    return { ok: false, error: 'Não foi possível processar o arquivo enviado.' };
  }
  var file = folder.createFile(blob);

  var acesso = PropertiesService.getScriptProperties().getProperty('PDF_ACESSO') || 'LINK';
  try {
    if (acesso === 'LINK') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    else if (acesso === 'DOMINIO') file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* conta pessoal sem domínio: mantém padrão */ }

  return { ok: true, url: file.getUrl(), id: file.getId() };
}

// ---------- NOTIFICAÇÃO (opcional) ----------
// Preencha as propriedades do script EMAIL_DIRETORIA e/ou EMAIL_SOLICITANTES
// (Configurações do projeto > Propriedades do script) para ativar o aviso por
// e-mail. Sem isso o ciclo depende de alguém abrir a página — ver README.
function notificar_(reg, evento) {
  try {
    var props = PropertiesService.getScriptProperties();
    var para = evento === 'apontamento'
      ? props.getProperty('EMAIL_SOLICITANTES')
      : props.getProperty('EMAIL_DIRETORIA');
    if (!para) return;
    var url = props.getProperty('APP_URL') || '';
    var assunto = evento === 'apontamento'
      ? '[Vegas] Solicitação Nº ' + reg.numero + ' devolvida para ajuste'
      : '[Vegas] Solicitação Nº ' + reg.numero + ' respondida — aguarda diretoria';
    var ultima = (reg.tratativas || [])[(reg.tratativas || []).length - 1] || {};
    MailApp.sendEmail({
      to: para,
      subject: assunto,
      body: [
        'Solicitação Nº ' + reg.numero + ' — ' + reg.veiculo,
        'Solicitante: ' + reg.solicitante + ' | Setor: ' + reg.setor,
        '',
        (ultima.nome || '') + ' escreveu:',
        ultima.texto || '',
        '',
        url
      ].join('\n')
    });
  } catch (e) { /* nunca deixa a notificação derrubar a gravação */ }
}

function excluirSolicitacao_(data) {
  var sh = sheet_();
  var row = findRow_(sh, data.id);
  if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };

  var status = sh.getRange(row, 12).getValue(); // coluna 'status'
  if (status === ST_APROVADO) {
    // Só exige a senha da diretoria quando a solicitação já foi assinada/aprovada.
    var senha = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
    if (data.senha !== senha) return { ok: false, error: 'Senha incorreta.' };
  }

  sh.deleteRow(row);
  return { ok: true };
}
