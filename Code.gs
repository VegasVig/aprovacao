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
var SHEET_VEICULOS = 'Veiculos';
var HEADERS_VEICULOS = ['veiculo', 'placa', 'timestamp'];
var HEADERS = [
  'id', 'numero', 'timestamp', 'data', 'solicitante', 'setor', 'veiculo',
  'descricaoJSON', 'pecasJSON', 'maoObraJSON', 'obs',
  'status', 'diretor', 'dataAprovacao', 'assinatura',
  'pecaEscolhidaJSON', 'maoObraEscolhidaJSON', 'tratativasJSON',
  'placa', 'km', 'unidade',
  'realizado', 'dataRealizado',
  'cancelado', 'dataCancelado', 'motivoCancelado'
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

  // Aba de veículos cadastrados (para o solicitante escolher da lista)
  var shV = ss.getSheetByName(SHEET_VEICULOS);
  if (!shV) {
    shV = ss.insertSheet(SHEET_VEICULOS);
    shV.appendRow(HEADERS_VEICULOS);
    shV.setFrozenRows(1);
  }
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
    if (action === 'listVeiculos') return respond_(listVeiculos_());
    if (action === 'painelManutencao') return respond_(painelManutencao_());
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
      case 'marcarRealizado': return respond_(marcarRealizado_(data));
      case 'marcarCancelado': return respond_(marcarCancelado_(data));
      case 'devolver': return respond_(devolverSolicitacao_(data));
      case 'responder': return respond_(responderSolicitacao_(data));
      case 'delete': return respond_(excluirSolicitacao_(data));
      case 'arquivarPdf': return respond_(arquivarPdf_(data));
      case 'uploadComprovante': return respond_(uploadComprovante_(data));
      case 'listVeiculos': return respond_(listVeiculos_());
      case 'salvarVeiculo': return respond_(salvarVeiculo_(data));
      // ----- Manutenção preventiva por quilometragem -----
      case 'painelManutencao': return respond_(painelManutencao_());
      case 'registrarManutencao': return respond_(registrarManutencao_(data));
      case 'atualizarHodometro': return respond_(atualizarHodometro_(data));
      case 'excluirManutencao': return respond_(excluirManutencao_(data));
      case 'salvarViaturaManutencao': return respond_(salvarViaturaManutencao_(data));
      case 'salvarConfigManutencao': return respond_(salvarConfigManutencao_(data));
      case 'historicoHodometro': return respond_(historicoHodometro_(data));
      case 'uploadComprovanteManutencao': return respond_(uploadComprovanteManutencao_(data));
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

  var linha = [
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
    data.km || '',
    data.unidade || ''
  ];
  // Completa com vazios até o total de colunas do cabeçalho. Assim, sempre que
  // uma coluna nova for adicionada em HEADERS, a gravação não quebra mais.
  while (linha.length < HEADERS.length) linha.push('');
  sh.getRange(newRow, 1, 1, HEADERS.length).setValues([linha]);
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
  if (data.unidade !== undefined) sh.getRange(row, 21).setValue(data.unidade || '');
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

// Marca (ou desmarca) uma solicitação APROVADA como "serviço realizado".
// Grava nas colunas 22 (realizado) e 23 (dataRealizado).
function marcarRealizado_(data) {
  return withLock_(function () {
    var sh = sheet_();
    var row = findRow_(sh, data.id);
    if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };
    if (sh.getRange(row, 12).getValue() !== ST_APROVADO) {
      return { ok: false, error: 'Só é possível marcar como realizado uma solicitação já aprovada.' };
    }
    var feito = (data.realizado === true || data.realizado === 'true');
    if (feito && sh.getRange(row, 24).getValue() === 'Sim') {
      return { ok: false, error: 'Esta solicitação está cancelada. Desfaça o cancelamento antes de marcar como realizada.' };
    }
    sh.getRange(row, 23).setNumberFormat('@'); // dataRealizado como texto
    if (feito) {
      sh.getRange(row, 22).setValue('Sim');
      sh.getRange(row, 23).setValue(new Date().toISOString());
    } else {
      sh.getRange(row, 22).setValue('');
      sh.getRange(row, 23).setValue('');
    }
    return { ok: true };
  });
}

// Marca (ou desmarca) uma solicitação APROVADA como "cancelada".
// Grava nas colunas 24 (cancelado), 25 (dataCancelado) e 26 (motivoCancelado).
// Cancelar limpa a marcação de "realizado" — as duas são excludentes.
function marcarCancelado_(data) {
  return withLock_(function () {
    var sh = sheet_();
    var row = findRow_(sh, data.id);
    if (row === -1) return { ok: false, error: 'Solicitação não encontrada.' };
    if (sh.getRange(row, 12).getValue() !== ST_APROVADO) {
      return { ok: false, error: 'Só é possível cancelar uma solicitação já aprovada.' };
    }
    var cancelar = (data.cancelado === true || data.cancelado === 'true');
    sh.getRange(row, 25).setNumberFormat('@'); // dataCancelado como texto
    if (cancelar) {
      sh.getRange(row, 22, 1, 5).setValues([[
        '', '',                          // limpa realizado / dataRealizado
        'Sim',
        new Date().toISOString(),
        data.motivo || ''
      ]]);
    } else {
      sh.getRange(row, 24, 1, 3).setValues([['', '', '']]);
    }
    return { ok: true };
  });
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

// ---------- VEÍCULOS CADASTRADOS ----------

function sheetVeiculos_() {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(SHEET_VEICULOS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_VEICULOS);
    sh.appendRow(HEADERS_VEICULOS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// Retorna a lista de veículos cadastrados, em ordem alfabética.
function listVeiculos_() {
  var sh = sheetVeiculos_();
  var last = sh.getLastRow();
  if (last < 2) return { ok: true, veiculos: [] };
  var values = sh.getRange(2, 1, last - 1, 2).getValues();
  var list = values
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) { return { veiculo: String(r[0]).trim(), placa: String(r[1]).trim() }; });
  // ordem alfabética pelo nome do veículo (ignorando maiúsc/minúsc e acentos)
  list.sort(function (a, b) {
    return a.veiculo.localeCompare(b.veiculo, 'pt-BR', { sensitivity: 'base' });
  });
  return { ok: true, veiculos: list };
}

// Salva um veículo novo se ainda não existir (compara por placa; se não houver
// placa, compara por nome). Idempotente: chamar de novo com o mesmo carro não duplica.
function salvarVeiculo_(data) {
  return withLock_(function () {
    var veiculo = String(data.veiculo || '').trim();
    var placa = String(data.placa || '').trim().toUpperCase();
    if (!veiculo) return { ok: false, error: 'Nome do veículo vazio.' };

    var sh = sheetVeiculos_();
    var last = sh.getLastRow();
    var existentes = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
    var jaTem = existentes.some(function (r) {
      var v = String(r[0]).trim().toLowerCase();
      var p = String(r[1]).trim().toUpperCase();
      if (placa) return p === placa;      // se tem placa, a placa é a chave
      return v === veiculo.toLowerCase();  // senão, o nome
    });
    if (!jaTem) {
      sh.appendRow([veiculo, placa, new Date().toISOString()]);
    }
    return listVeiculos_();
  });
}

// =====================================================================
//  MANUTENÇÃO PREVENTIVA POR QUILOMETRAGEM
// ---------------------------------------------------------------------
//  Abas usadas (criadas automaticamente na primeira chamada do módulo,
//  sem tocar nas abas existentes):
//    Veiculos            -> ganha colunas novas à direita (id, marca, modelo,
//                           tipo, ano, motor, comando, kmAtual, ...). As três
//                           colunas originais (veiculo, placa, timestamp)
//                           continuam iguais.
//    Manutencoes         -> cada troca registrada (histórico permanente;
//                           exclusão é lógica, a linha nunca é apagada).
//    Hodometro           -> todas as leituras de quilometragem (auditoria).
//    AlertasManutencao   -> alertas abertos/encerrados por viatura e item.
//    ConfigManutencao    -> intervalos padrão e tabela correia/corrente
//                           por modelo (editáveis pela diretoria).
//  Cada registro é ligado ao "id" único da viatura (coluna id da aba
//  Veiculos), nunca só ao nome, para não misturar dados entre carros.
// =====================================================================

var SHEET_MANUT = 'Manutencoes';
var SHEET_HODOMETRO = 'Hodometro';
var SHEET_ALERTAS = 'AlertasManutencao';
var SHEET_CONFIG_MANUT = 'ConfigManutencao';

// Colunas acrescentadas à aba Veiculos (depois de veiculo / placa / timestamp).
var HEADERS_VEICULOS_MNT = [
  'id', 'marca', 'modelo', 'tipo', 'ano', 'motor', 'comando',
  'kmAtual', 'kmAtualizadoEm', 'kmOrigem', 'kmRegistradoEm',
  'intervalosJSON', 'atualizadoEm'
];
var TEXTO_VEICULOS_MNT = ['timestamp', 'kmAtualizadoEm', 'kmRegistradoEm', 'atualizadoEm', 'intervalosJSON', 'placa', 'ano'];

var HEADERS_MANUT = [
  'id', 'veiculoId', 'placa', 'veiculo', 'item', 'posicoes', 'descricao',
  'data', 'km', 'valor', 'oficina', 'obs', 'comprovante',
  'registradoEm', 'tipoVeiculo', 'intervaloAplicado', 'proximaKm',
  'excluido', 'excluidoEm', 'excluidoPor', 'motivoExclusao'
];
var TEXTO_MANUT = ['data', 'registradoEm', 'excluidoEm', 'posicoes', 'placa'];

var HEADERS_HODOMETRO = [
  'id', 'veiculoId', 'placa', 'data', 'km', 'kmAnterior', 'origem', 'ref',
  'responsavel', 'obs', 'regressao', 'autorizadoPor', 'registradoEm'
];
var TEXTO_HODOMETRO = ['data', 'registradoEm', 'placa'];

var HEADERS_ALERTAS = [
  'id', 'veiculoId', 'placa', 'item', 'status', 'kmPrevisto', 'kmAtual',
  'abertoEm', 'atualizadoEm', 'encerradoEm', 'motivoEncerramento', 'manutencaoId'
];
var TEXTO_ALERTAS = ['abertoEm', 'atualizadoEm', 'encerradoEm', 'placa'];

var HEADERS_CONFIG_MANUT = ['chave', 'valor', 'descricao'];

var ITENS_PREVENTIVOS = ['oleo', 'comando', 'pneuDianteiro', 'pneuTraseiro'];
var ITENS_VALIDOS = ['oleo', 'comando', 'pneuDianteiro', 'pneuTraseiro', 'outro'];
var POSICOES_PNEU = {
  carro: { pneuDianteiro: ['DE', 'DD'], pneuTraseiro: ['TE', 'TD'] },
  moto:  { pneuDianteiro: ['D'],        pneuTraseiro: ['T'] }
};
var NOMES_POSICAO = { DE: 'Dianteiro esquerdo', DD: 'Dianteiro direito', TE: 'Traseiro esquerdo', TD: 'Traseiro direito', D: 'Dianteiro', T: 'Traseiro' };

// Valores iniciais. Depois da primeira carga, valem os da aba ConfigManutencao.
var CONFIG_MANUT_PADRAO = {
  oleoCarro: 10000,
  oleoMoto: 1500,
  pneuDianteiro: 50000,
  pneuTraseiro: 50000,
  alertaKm: 1000,
  comandoModelos: [
    { modelo: 'Mobi',  marca: 'Fiat',       km: 200000 },
    { modelo: 'Kwid',  marca: 'Renault',    km: 200000 },
    { modelo: 'Onix',  marca: 'Chevrolet',  km: 200000 },
    { modelo: 'Palio', marca: 'Fiat',       km: 50000 },
    { modelo: 'Gol',   marca: 'Volkswagen', km: 50000 }
  ]
};
var DESCRICAO_CONFIG = {
  oleoCarro: 'Troca de óleo - carros (km)',
  oleoMoto: 'Troca de óleo - motos (km)',
  pneuDianteiro: 'Pneus dianteiros (km)',
  pneuTraseiro: 'Pneus traseiros (km)',
  alertaKm: 'Alerta antecipado: avisar quando faltarem (km)'
};

// Sugestão automática de marca/modelo/tipo a partir do nome já cadastrado
// (ex.: "Palio Fire" -> Fiat / Palio / carro). Só é usada enquanto a
// diretoria não preencher o cadastro da viatura; nunca é gravada sozinha.
var MODELOS_CONHECIDOS = [
  ['mobi', 'Fiat', 'Mobi', 'carro'], ['palio', 'Fiat', 'Palio', 'carro'], ['uno', 'Fiat', 'Uno', 'carro'],
  ['strada', 'Fiat', 'Strada', 'carro'], ['argo', 'Fiat', 'Argo', 'carro'], ['cronos', 'Fiat', 'Cronos', 'carro'],
  ['siena', 'Fiat', 'Siena', 'carro'], ['fiorino', 'Fiat', 'Fiorino', 'carro'], ['toro', 'Fiat', 'Toro', 'carro'],
  ['kwid', 'Renault', 'Kwid', 'carro'], ['sandero', 'Renault', 'Sandero', 'carro'], ['logan', 'Renault', 'Logan', 'carro'],
  ['duster', 'Renault', 'Duster', 'carro'], ['onix', 'Chevrolet', 'Onix', 'carro'], ['prisma', 'Chevrolet', 'Prisma', 'carro'],
  ['celta', 'Chevrolet', 'Celta', 'carro'], ['spin', 'Chevrolet', 'Spin', 'carro'], ['montana', 'Chevrolet', 'Montana', 'carro'],
  ['s10', 'Chevrolet', 'S10', 'carro'], ['gol', 'Volkswagen', 'Gol', 'carro'], ['voyage', 'Volkswagen', 'Voyage', 'carro'],
  ['fox', 'Volkswagen', 'Fox', 'carro'], ['saveiro', 'Volkswagen', 'Saveiro', 'carro'], ['polo', 'Volkswagen', 'Polo', 'carro'],
  ['hb20', 'Hyundai', 'HB20', 'carro'], ['ka', 'Ford', 'Ka', 'carro'], ['fiesta', 'Ford', 'Fiesta', 'carro'],
  ['cg', 'Honda', 'CG', 'moto'], ['titan', 'Honda', 'CG Titan', 'moto'], ['biz', 'Honda', 'Biz', 'moto'],
  ['bros', 'Honda', 'NXR Bros', 'moto'], ['xre', 'Honda', 'XRE', 'moto'], ['factor', 'Yamaha', 'Factor', 'moto'],
  ['fazer', 'Yamaha', 'Fazer', 'moto'], ['crosser', 'Yamaha', 'Crosser', 'moto'], ['lander', 'Yamaha', 'Lander', 'moto'],
  ['ybr', 'Yamaha', 'YBR', 'moto']
];

// Rode uma vez no editor (opcional): cria/atualiza as abas do módulo.
function configurarManutencao() {
  sheetVeiculosMnt_();
  sheetManut_();
  sheetHodometro_();
  sheetAlertas_();
  lerConfigManut_();
}

// ---------- utilidades do módulo ----------

function tz_() {
  try { return Session.getScriptTimeZone() || 'America/Sao_Paulo'; } catch (e) { return 'America/Sao_Paulo'; }
}
function agoraIso_() { return new Date().toISOString(); }
function hojeStr_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }

// Valor de célula -> texto (Sheets às vezes devolve Date mesmo em coluna de texto)
function txt_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    if (isNaN(v.getTime())) return '';
    return Utilities.formatDate(v, tz_(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(v);
}
function num_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var n = Number(String(v).trim().replace(',', '.'));
  return isFinite(n) ? n : null;
}
// Quilometragem digitada: aceita 85000, "85000", "85.000" (milhar com ponto).
function parseKm_(v) {
  if (v === null || v === undefined) return null;
  var s = String(v).trim();
  if (!s) return null;
  if (typeof v === 'number') {
    if (!isFinite(v) || v < 0 || Math.floor(v) !== v) return null;
    return v <= 3000000 ? v : null;
  }
  s = s.replace(/\s/g, '').replace(/\.(?=\d{3}(\D|$))/g, '');
  if (!/^\d+$/.test(s)) return null;
  var n = parseInt(s, 10);
  return n <= 3000000 ? n : null;
}
function parseValor_(v) {
  if (v === null || v === undefined || String(v).trim() === '') return '';
  var n = num_(v);
  if (n === null || n < 0 || n > 10000000) return null;
  return Math.round(n * 100) / 100;
}
function dataValida_(s) {
  var d = String(s || '').trim().split('T')[0];
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return null;
  var p = d.split('-').map(Number);
  var dt = new Date(p[0], p[1] - 1, p[2]);
  if (dt.getFullYear() !== p[0] || dt.getMonth() !== p[1] - 1 || dt.getDate() !== p[2]) return null;
  return d;
}
function amanhaStr_() {
  return Utilities.formatDate(new Date(Date.now() + 86400000), tz_(), 'yyyy-MM-dd');
}
function normTxtMnt_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}
function tokens_(s) { return normTxtMnt_(s).split(/[^a-z0-9]+/).filter(Boolean); }
function placaNorm_(p) { return String(p || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function senhaValida_(senha) {
  var s = PropertiesService.getScriptProperties().getProperty('DASHBOARD_PASSWORD');
  return !!s && String(senha || '') === s;
}

// Cria a aba se não existir e garante os cabeçalhos (colunas faltantes vão
// para o fim — nada existente é sobrescrito nem reordenado).
function garantirAba_(nome, headers, colsTexto) {
  var ss = getSpreadsheet_();
  var sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    (colsTexto || []).forEach(function (c) {
      var i = headers.indexOf(c);
      if (i >= 0) sh.getRange(2, i + 1, 2000, 1).setNumberFormat('@');
    });
    return sh;
  }
  garantirCabecalhos_(sh, headers);
  return sh;
}
function garantirCabecalhos_(sh, headers) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var atuais = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var faltam = headers.filter(function (h) { return atuais.indexOf(h) === -1; });
  if (faltam.length) {
    // ocupa colunas vazias no fim do cabeçalho antes de criar novas
    var fim = atuais.length;
    while (fim > 0 && atuais[fim - 1] === '') fim--;
    sh.getRange(1, fim + 1, 1, faltam.length).setValues([faltam]);
  }
}
function cabecalho_(sh) {
  return sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
}
function lerTabela_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var head = cabecalho_(sh);
  var vals = sh.getRange(2, 1, last - 1, head.length).getValues();
  return vals.map(function (row, i) {
    var o = { _row: i + 2 };
    head.forEach(function (h, j) { if (h) o[h] = row[j]; });
    return o;
  });
}
function anexarLinha_(sh, obj, colsTexto) {
  var head = cabecalho_(sh);
  var row = sh.getLastRow() + 1;
  (colsTexto || []).forEach(function (c) {
    var i = head.indexOf(c);
    if (i >= 0) sh.getRange(row, i + 1).setNumberFormat('@');
  });
  var linha = head.map(function (h) { return obj.hasOwnProperty(h) && obj[h] !== null && obj[h] !== undefined ? obj[h] : ''; });
  sh.getRange(row, 1, 1, head.length).setValues([linha]);
  return row;
}
function atualizarCelulas_(sh, row, campos, colsTexto) {
  var head = cabecalho_(sh);
  Object.keys(campos).forEach(function (k) {
    var i = head.indexOf(k);
    if (i < 0) return;
    var cel = sh.getRange(row, i + 1);
    if ((colsTexto || []).indexOf(k) >= 0) cel.setNumberFormat('@');
    var v = campos[k];
    cel.setValue(v === null || v === undefined ? '' : v);
  });
}

// ---------- abas ----------

function sheetManut_() { return garantirAba_(SHEET_MANUT, HEADERS_MANUT, TEXTO_MANUT); }
function sheetHodometro_() { return garantirAba_(SHEET_HODOMETRO, HEADERS_HODOMETRO, TEXTO_HODOMETRO); }
function sheetAlertas_() { return garantirAba_(SHEET_ALERTAS, HEADERS_ALERTAS, TEXTO_ALERTAS); }

// Aba Veiculos existente + colunas do módulo + id único para cada viatura.
function sheetVeiculosMnt_() {
  var sh = sheetVeiculos_();
  garantirCabecalhos_(sh, HEADERS_VEICULOS.concat(HEADERS_VEICULOS_MNT));
  var last = sh.getLastRow();
  if (last >= 2) {
    var head = cabecalho_(sh);
    var cId = head.indexOf('id'), cNome = head.indexOf('veiculo');
    var vals = sh.getRange(2, 1, last - 1, head.length).getValues();
    var mudou = false;
    var ids = vals.map(function (r) {
      var id = String(r[cId] || '').trim();
      if (!id && String(r[cNome] || '').trim()) { id = Utilities.getUuid(); mudou = true; }
      return [id];
    });
    if (mudou) sh.getRange(2, cId + 1, ids.length, 1).setValues(ids);
  }
  return sh;
}

function inferirVeiculo_(nome) {
  var tk = tokens_(nome);
  var r = { marca: '', modelo: '', tipo: '' };
  for (var i = 0; i < MODELOS_CONHECIDOS.length; i++) {
    var m = MODELOS_CONHECIDOS[i];
    if (tk.indexOf(m[0]) >= 0) { r.marca = m[1]; r.modelo = m[2]; r.tipo = m[3]; break; }
  }
  if (tk.indexOf('moto') >= 0 || tk.indexOf('motocicleta') >= 0) r.tipo = 'moto';
  return r;
}

function lerVeiculosMnt_(sh) {
  return lerTabela_(sh)
    .filter(function (r) { return String(r.veiculo || '').trim() !== ''; })
    .map(function (r) {
      var inf = inferirVeiculo_(r.veiculo);
      var intervalos = {};
      try { intervalos = r.intervalosJSON ? JSON.parse(txt_(r.intervalosJSON)) : {}; } catch (e) { intervalos = {}; }
      var marca = String(r.marca || '').trim(), modelo = String(r.modelo || '').trim(), tipo = String(r.tipo || '').trim();
      return {
        _row: r._row,
        id: String(r.id || ''),
        veiculo: String(r.veiculo || '').trim(),
        placa: String(r.placa || '').trim().toUpperCase(),
        marca: marca || inf.marca,
        modelo: modelo || inf.modelo,
        tipo: (tipo === 'moto' || tipo === 'carro') ? tipo : (inf.tipo || 'carro'),
        inferido: { marca: !marca && !!inf.marca, modelo: !modelo && !!inf.modelo, tipo: !tipo },
        ano: txt_(r.ano),
        motor: String(r.motor || ''),
        comando: String(r.comando || ''),
        kmAtual: num_(r.kmAtual),
        kmAtualizadoEm: txt_(r.kmAtualizadoEm),
        kmOrigem: String(r.kmOrigem || ''),
        kmRegistradoEm: txt_(r.kmRegistradoEm),
        intervalos: intervalos || {}
      };
    });
}

function lerManutencoes_() {
  return lerTabela_(sheetManut_()).filter(function (r) { return String(r.id || '') !== ''; }).map(function (r) {
    return {
      id: String(r.id), veiculoId: String(r.veiculoId || ''), placa: String(r.placa || ''), veiculo: String(r.veiculo || ''),
      item: String(r.item || ''), posicoes: String(r.posicoes || ''), descricao: String(r.descricao || ''),
      data: txt_(r.data).split('T')[0], km: num_(r.km), valor: num_(r.valor),
      oficina: String(r.oficina || ''), obs: String(r.obs || ''), comprovante: String(r.comprovante || ''),
      registradoEm: txt_(r.registradoEm), tipoVeiculo: String(r.tipoVeiculo || ''),
      intervaloAplicado: num_(r.intervaloAplicado), proximaKm: num_(r.proximaKm),
      excluido: String(r.excluido || '') === 'Sim' ? 'Sim' : '',
      excluidoEm: txt_(r.excluidoEm), excluidoPor: String(r.excluidoPor || ''), motivoExclusao: String(r.motivoExclusao || '')
    };
  });
}

// ---------- configuração (intervalos padrão) ----------

function lerConfigManut_() {
  var sh = garantirAba_(SHEET_CONFIG_MANUT, HEADERS_CONFIG_MANUT, []);
  var linhas = lerTabela_(sh);
  if (!linhas.length) {
    escreverConfigManut_(sh, CONFIG_MANUT_PADRAO);
    linhas = lerTabela_(sh);
  }
  var cfg = {
    oleoCarro: CONFIG_MANUT_PADRAO.oleoCarro, oleoMoto: CONFIG_MANUT_PADRAO.oleoMoto,
    pneuDianteiro: CONFIG_MANUT_PADRAO.pneuDianteiro, pneuTraseiro: CONFIG_MANUT_PADRAO.pneuTraseiro,
    alertaKm: CONFIG_MANUT_PADRAO.alertaKm, comandoModelos: []
  };
  linhas.forEach(function (r) {
    var chave = String(r.chave || '').trim();
    var v = num_(r.valor);
    if (chave.indexOf('comando:') === 0) {
      var modelo = chave.slice(8).trim();
      if (modelo && v > 0) cfg.comandoModelos.push({ modelo: modelo, marca: String(r.descricao || '').trim(), km: v });
    } else if (cfg.hasOwnProperty(chave) && v !== null && v >= 0) {
      cfg[chave] = v;
    }
  });
  return cfg;
}

function escreverConfigManut_(sh, cfg) {
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 3).clearContent();
  var linhas = ['oleoCarro', 'oleoMoto', 'pneuDianteiro', 'pneuTraseiro', 'alertaKm'].map(function (k) {
    return [k, cfg[k], DESCRICAO_CONFIG[k]];
  });
  (cfg.comandoModelos || []).forEach(function (m) {
    linhas.push(['comando: ' + m.modelo, m.km, m.marca || '']);
  });
  sh.getRange(2, 1, linhas.length, 3).setValues(linhas);
}

// ---------- regras de cálculo ----------

function modeloComando_(v, cfg) {
  var tk = tokens_(v.modelo + ' ' + v.veiculo);
  var lista = (cfg.comandoModelos || []).slice().sort(function (a, b) {
    return tokens_(b.modelo).length - tokens_(a.modelo).length;
  });
  for (var i = 0; i < lista.length; i++) {
    var mt = tokens_(lista[i].modelo);
    if (mt.length && mt.every(function (t) { return tk.indexOf(t) >= 0; })) return lista[i];
  }
  return null;
}

function intervaloItem_(v, item, cfg) {
  var ov = num_((v.intervalos || {})[item]);
  if (ov !== null && ov > 0) return { km: ov, origem: 'veiculo' };
  if (item === 'oleo') return v.tipo === 'moto'
    ? { km: cfg.oleoMoto, origem: 'padrao_moto' }
    : { km: cfg.oleoCarro, origem: 'padrao_carro' };
  if (item === 'pneuDianteiro') return { km: cfg.pneuDianteiro, origem: 'padrao' };
  if (item === 'pneuTraseiro') return { km: cfg.pneuTraseiro, origem: 'padrao' };
  if (item === 'comando') {
    var m = modeloComando_(v, cfg);
    return m ? { km: m.km, origem: 'modelo', modelo: m.modelo } : { km: null, origem: '' };
  }
  return { km: null, origem: '' };
}

function alertaKm_(v, cfg) {
  var ov = num_((v.intervalos || {}).alertaKm);
  return ov !== null && ov >= 0 ? ov : cfg.alertaKm;
}

function nomeItem_(item, v) {
  if (item === 'oleo') return 'Óleo do motor';
  if (item === 'comando') {
    if (v && v.comando === 'correia') return 'Correia dentada';
    if (v && v.comando === 'corrente') return 'Corrente de comando';
    return 'Correia dentada / corrente de comando';
  }
  if (item === 'pneuDianteiro') return 'Pneus dianteiros';
  if (item === 'pneuTraseiro') return 'Pneus traseiros';
  return 'Outro serviço';
}

function posicoesItem_(v, item) {
  var t = v.tipo === 'moto' ? 'moto' : 'carro';
  return (POSICOES_PNEU[t] || {})[item] || null;
}

function cobrePosicao_(reg, pos, tipo) {
  var ps = String(reg.posicoes || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!ps.length || tipo === 'moto') return true;
  return ps.indexOf(pos) >= 0 || ps.indexOf(pos.charAt(0)) >= 0;
}

// Última troca = maior quilometragem registrada (desempate: data mais recente).
function maisRecente_(regs) {
  var r = null;
  regs.forEach(function (x) {
    if (x.km === null) return;
    if (!r || x.km > r.km || (x.km === r.km && (x.data + x.registradoEm) > (r.data + r.registradoEm))) r = x;
  });
  return r;
}

function statusPorRestante_(restante, alerta) {
  if (restante <= 0) return 'vencida';
  if (restante <= alerta) return 'proxima';
  return 'em_dia';
}

function resumoReg_(r) {
  return r ? { id: r.id, data: r.data, km: r.km, valor: r.valor, oficina: r.oficina, posicoes: r.posicoes, descricao: r.descricao } : null;
}

function calcularItem_(v, item, regsVeic, cfg) {
  var regs = regsVeic.filter(function (r) { return r.item === item && r.excluido !== 'Sim'; });
  var iv = intervaloItem_(v, item, cfg);
  var alerta = alertaKm_(v, cfg);
  var maxReg = maisRecente_(regs);
  var kmRef = v.kmAtual !== null ? v.kmAtual : (maxReg ? maxReg.km : null);
  var res = {
    item: item, nome: nomeItem_(item, v), intervalo: iv.km, intervaloOrigem: iv.origem,
    intervaloModelo: iv.modelo || '', alertaKm: alerta, kmRef: kmRef,
    ultima: null, proximaKm: null, restante: null, status: 'sem_registro',
    posicoes: [], incompleto: false, totalRegistros: regs.length
  };
  var posLista = posicoesItem_(v, item);
  if (posLista) {
    var maisUrgente = null;
    posLista.forEach(function (pos) {
      var ult = maisRecente_(regs.filter(function (r) { return cobrePosicao_(r, pos, v.tipo); }));
      var p = { pos: pos, nome: NOMES_POSICAO[pos] || pos, ultima: resumoReg_(ult), proximaKm: null, restante: null, status: 'sem_registro' };
      if (ult && iv.km) {
        p.proximaKm = ult.km + iv.km;
        if (kmRef !== null) { p.restante = p.proximaKm - kmRef; p.status = statusPorRestante_(p.restante, alerta); }
        if (!maisUrgente || p.proximaKm < maisUrgente.proximaKm) maisUrgente = p;
      } else if (ult) {
        p.status = 'sem_intervalo';
      }
      if (!ult) res.incompleto = true;
      res.posicoes.push(p);
    });
    if (maisUrgente) {
      res.ultima = maisUrgente.ultima;
      res.proximaKm = maisUrgente.proximaKm;
      res.restante = maisUrgente.restante;
      res.status = maisUrgente.status;
    } else if (maxReg) {
      res.ultima = resumoReg_(maxReg);
    }
    if (posLista.length === 1) res.incompleto = false;
  } else {
    res.ultima = resumoReg_(maxReg);
    if (maxReg && iv.km) {
      res.proximaKm = maxReg.km + iv.km;
      if (kmRef !== null) { res.restante = res.proximaKm - kmRef; res.status = statusPorRestante_(res.restante, alerta); }
    }
  }
  if (!iv.km) res.status = 'sem_intervalo';
  // Progresso dentro do intervalo (0 a 1+) para a barra visual.
  if (res.ultima && iv.km && kmRef !== null && res.proximaKm !== null) {
    res.progresso = Math.max(0, (kmRef - (res.proximaKm - iv.km)) / iv.km);
  }
  return res;
}

function situacaoGeral_(itens) {
  var st = itens.map(function (i) { return i.status; });
  if (st.indexOf('vencida') >= 0) return 'vencida';
  if (st.indexOf('proxima') >= 0) return 'proxima';
  if (st.indexOf('sem_registro') >= 0 || st.indexOf('sem_intervalo') >= 0) return 'incompleta';
  return 'em_dia';
}

// Média de km/dia pelas leituras do hodômetro dos últimos 180 dias.
function mediaKmDia_(leituras) {
  var lim = Date.now() - 180 * 86400000;
  var pts = leituras
    .filter(function (h) { return h.regressao !== 'Sim' && h.km !== null; })
    .map(function (h) { return { t: new Date(String(h.data).length <= 10 ? h.data + 'T12:00:00' : h.data).getTime(), km: h.km }; })
    .filter(function (p) { return !isNaN(p.t) && p.t >= lim; })
    .sort(function (a, b) { return a.t - b.t; });
  if (pts.length < 2) return null;
  var dias = (pts[pts.length - 1].t - pts[0].t) / 86400000;
  var delta = pts[pts.length - 1].km - pts[0].km;
  if (dias < 7 || delta <= 0) return null;
  return Math.round(delta / dias * 10) / 10;
}

// ---------- integração com a quilometragem das solicitações ----------
// A quilometragem informada em cada solicitação (campo já existente) alimenta
// o hodômetro da viatura. Só sobe o valor — nunca reduz — e só considera
// solicitações criadas depois do último registro de km da viatura, para que
// uma correção autorizada não seja desfeita por um lançamento antigo.
function sincronizarKmSolicitacoes_(shV, veics) {
  var sh = getSpreadsheet_().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return false;
  var ncol = Math.min(sh.getLastColumn(), HEADERS.length);
  if (ncol < 20) return false;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, ncol).getValues();
  var lidas = vals.map(function (r) {
    return {
      id: String(r[0] || ''), numero: String(r[1] || ''), timestamp: txt_(r[2]), data: txt_(r[3]).split('T')[0],
      veiculo: String(r[6] || ''), placa: placaNorm_(r[18]), km: parseKm_(txt_(r[19]))
    };
  }).filter(function (s) { return s.km !== null && s.km > 0; });
  if (!lidas.length) return false;
  var shH = null, mudou = false;
  veics.forEach(function (v) {
    var pv = placaNorm_(v.placa), nv = normTxtMnt_(v.veiculo);
    var corte = v.kmRegistradoEm || '';
    var melhor = null;
    lidas.forEach(function (s) {
      var casa = s.placa ? (pv && s.placa === pv) : (normTxtMnt_(s.veiculo) === nv);
      if (!casa) return;
      if (corte && s.timestamp && s.timestamp <= corte) return;
      if (!melhor || s.km > melhor.km) melhor = s;
    });
    if (!melhor || (v.kmAtual !== null && melhor.km <= v.kmAtual)) return;
    var agora = agoraIso_();
    var quando = melhor.timestamp || (melhor.data ? melhor.data + 'T12:00:00' : agora);
    atualizarCelulas_(shV, v._row, {
      kmAtual: melhor.km, kmAtualizadoEm: quando, kmOrigem: 'Solicitação Nº ' + melhor.numero, kmRegistradoEm: agora
    }, TEXTO_VEICULOS_MNT);
    shH = shH || sheetHodometro_();
    anexarLinha_(shH, {
      id: Utilities.getUuid(), veiculoId: v.id, placa: v.placa, data: quando, km: melhor.km,
      kmAnterior: v.kmAtual === null ? '' : v.kmAtual, origem: 'solicitacao', ref: melhor.id,
      responsavel: '', obs: 'Quilometragem informada na solicitação Nº ' + melhor.numero,
      regressao: '', autorizadoPor: '', registradoEm: agora
    }, TEXTO_HODOMETRO);
    v.kmAtual = melhor.km; v.kmAtualizadoEm = quando; v.kmOrigem = 'Solicitação Nº ' + melhor.numero; v.kmRegistradoEm = agora;
    mudou = true;
  });
  return mudou;
}

// ---------- alertas ----------
// Abre um alerta quando o item entra em "próxima" ou "vencida", atualiza
// enquanto continuar pendente e encerra quando a manutenção é registrada
// (ou a previsão muda). Linhas encerradas ficam como histórico.
function recalcularAlertas_(veiculosMontados, ctx) {
  var sh = sheetAlertas_();
  var linhas = lerTabela_(sh);
  var abertos = {};
  linhas.forEach(function (a) {
    if (!txt_(a.encerradoEm)) abertos[a.veiculoId + '|' + a.item] = a;
  });
  var agora = agoraIso_();
  var resultado = {};
  veiculosMontados.forEach(function (v) {
    v.itens.forEach(function (it) {
      var k = v.id + '|' + it.item;
      var aberto = abertos[k];
      var pendente = it.status === 'vencida' || it.status === 'proxima';
      var stTxt = it.status === 'vencida' ? 'Vencida' : 'Próxima';
      var mesmoItemCtx = ctx && ctx.veiculoId === v.id && ctx.item === it.item;
      if (aberto && (!pendente || num_(aberto.kmPrevisto) !== it.proximaKm)) {
        atualizarCelulas_(sh, aberto._row, {
          encerradoEm: agora,
          motivoEncerramento: mesmoItemCtx ? (ctx.motivo || 'Manutenção registrada') : (pendente ? 'Nova previsão de troca' : 'Previsão recalculada'),
          manutencaoId: mesmoItemCtx ? (ctx.manutencaoId || '') : '',
          atualizadoEm: agora
        }, TEXTO_ALERTAS);
        aberto = null;
      }
      if (!pendente) return;
      if (aberto) {
        if (String(aberto.status) !== stTxt || num_(aberto.kmAtual) !== it.kmRef) {
          atualizarCelulas_(sh, aberto._row, { status: stTxt, kmAtual: it.kmRef, atualizadoEm: agora }, TEXTO_ALERTAS);
        }
        resultado[k] = { abertoEm: txt_(aberto.abertoEm), status: stTxt };
      } else {
        anexarLinha_(sh, {
          id: Utilities.getUuid(), veiculoId: v.id, placa: v.placa, item: it.item, status: stTxt,
          kmPrevisto: it.proximaKm, kmAtual: it.kmRef, abertoEm: agora, atualizadoEm: agora
        }, TEXTO_ALERTAS);
        resultado[k] = { abertoEm: agora, status: stTxt };
      }
    });
  });
  return resultado;
}

// ---------- painel (leitura completa do módulo) ----------

function painelManutencao_() {
  return withLock_(function () { return painelInterno_(null); });
}

function painelInterno_(ctx) {
  var cfg = lerConfigManut_();
  var shV = sheetVeiculosMnt_();
  var veics = lerVeiculosMnt_(shV);
  try { sincronizarKmSolicitacoes_(shV, veics); } catch (e) { /* nunca derruba o painel */ }
  var regs = lerManutencoes_();
  var hodo = lerTabela_(sheetHodometro_()).map(function (h) {
    return { veiculoId: String(h.veiculoId || ''), data: txt_(h.data), km: num_(h.km), regressao: String(h.regressao || '') };
  });
  var porVeic = {};
  regs.forEach(function (r) { (porVeic[r.veiculoId] = porVeic[r.veiculoId] || []).push(r); });
  var hodoPorVeic = {};
  hodo.forEach(function (h) { (hodoPorVeic[h.veiculoId] = hodoPorVeic[h.veiculoId] || []).push(h); });

  var montados = veics.map(function (v) {
    var itens = ITENS_PREVENTIVOS.map(function (item) { return calcularItem_(v, item, porVeic[v.id] || [], cfg); });
    return {
      id: v.id, veiculo: v.veiculo, placa: v.placa, marca: v.marca, modelo: v.modelo, tipo: v.tipo,
      inferido: v.inferido, ano: v.ano, motor: v.motor, comando: v.comando,
      kmAtual: v.kmAtual, kmAtualizadoEm: v.kmAtualizadoEm, kmOrigem: v.kmOrigem,
      intervalos: v.intervalos, alertaKm: alertaKm_(v, cfg), kmDia: mediaKmDia_(hodoPorVeic[v.id] || []),
      itens: itens, situacao: situacaoGeral_(itens)
    };
  });
  var alertas = recalcularAlertas_(montados, ctx);
  montados.forEach(function (v) {
    v.itens.forEach(function (it) {
      var a = alertas[v.id + '|' + it.item];
      it.alertaDesde = a ? a.abertoEm : '';
    });
  });
  montados.sort(function (a, b) { return a.veiculo.localeCompare(b.veiculo, 'pt-BR', { sensitivity: 'base' }); });
  return { ok: true, config: cfg, veiculos: montados, manutencoes: regs, geradoEm: agoraIso_() };
}

function acharVeiculo_(veics, id) {
  for (var i = 0; i < veics.length; i++) if (veics[i].id === String(id || '')) return veics[i];
  return null;
}

// ---------- registrar troca / manutenção ----------

function registrarManutencao_(data) {
  return withLock_(function () {
    var shV = sheetVeiculosMnt_();
    var veics = lerVeiculosMnt_(shV);
    var v = acharVeiculo_(veics, data.veiculoId);
    if (!v) return { ok: false, error: 'Viatura não encontrada. Atualize a página e tente de novo.' };

    var item = String(data.item || '');
    if (ITENS_VALIDOS.indexOf(item) < 0) return { ok: false, error: 'Escolha o tipo de manutenção.' };
    var dataStr = dataValida_(data.data);
    if (!dataStr) return { ok: false, error: 'Informe a data da troca.' };
    if (dataStr > amanhaStr_()) return { ok: false, error: 'A data da troca não pode estar no futuro.' };
    var km = parseKm_(data.km);
    if (km === null) return { ok: false, error: 'Informe a quilometragem do veículo no momento da troca (somente números).' };
    var valor = parseValor_(data.valor);
    if (valor === null) return { ok: false, error: 'Valor gasto inválido.' };
    var descricao = String(data.descricao || '').trim().slice(0, 300);
    if (item === 'outro' && !descricao) return { ok: false, error: 'Descreva o serviço realizado.' };

    var posicoes = '';
    var posLista = posicoesItem_(v, item);
    if (posLista) {
      var sel = (Array.isArray(data.posicoes) ? data.posicoes : String(data.posicoes || '').split(','))
        .map(function (p) { return String(p).trim(); })
        .filter(function (p) { return posLista.indexOf(p) >= 0; });
      if (!sel.length) sel = posLista.slice();
      posicoes = posLista.filter(function (p) { return sel.indexOf(p) >= 0; }).join(',');
    }

    var cfg = lerConfigManut_();
    var regs = lerManutencoes_().filter(function (r) { return r.veiculoId === v.id; });
    var antes = item !== 'outro' ? calcularItem_(v, item, regs, cfg) : null;
    var iv = intervaloItem_(v, item, cfg);
    var agora = agoraIso_();
    var id = Utilities.getUuid();
    anexarLinha_(sheetManut_(), {
      id: id, veiculoId: v.id, placa: v.placa, veiculo: v.veiculo, item: item, posicoes: posicoes,
      descricao: descricao, data: dataStr, km: km, valor: valor,
      oficina: String(data.oficina || '').trim().slice(0, 200), obs: String(data.obs || '').trim().slice(0, 1000),
      comprovante: String(data.comprovante || '').trim().slice(0, 500),
      registradoEm: agora, tipoVeiculo: v.tipo,
      intervaloAplicado: item !== 'outro' && iv.km ? iv.km : '',
      proximaKm: item !== 'outro' && iv.km ? km + iv.km : '',
      excluido: ''
    }, TEXTO_MANUT);

    // Troca registrada com km acima do hodômetro atual atualiza o hodômetro.
    if (v.kmAtual === null || km > v.kmAtual) {
      var quando = dataStr === hojeStr_() ? agora : dataStr + 'T12:00:00';
      atualizarCelulas_(shV, v._row, { kmAtual: km, kmAtualizadoEm: quando, kmOrigem: 'Registro de manutenção', kmRegistradoEm: agora }, TEXTO_VEICULOS_MNT);
      anexarLinha_(sheetHodometro_(), {
        id: Utilities.getUuid(), veiculoId: v.id, placa: v.placa, data: quando, km: km,
        kmAnterior: v.kmAtual === null ? '' : v.kmAtual, origem: 'manutencao', ref: id,
        responsavel: String(data.oficina || ''), obs: 'Quilometragem informada no registro de ' + nomeItem_(item, v).toLowerCase(),
        regressao: '', autorizadoPor: '', registradoEm: agora
      }, TEXTO_HODOMETRO);
    }

    var aviso = '';
    if (antes && antes.ultima && antes.ultima.km !== null && km < antes.ultima.km) {
      aviso = 'Registro salvo no histórico. Como a quilometragem é menor que a da última troca (' + antes.ultima.km + ' km), a previsão continua baseada na troca mais recente.';
    }
    return {
      ok: true, id: id, aviso: aviso,
      proximaKm: item !== 'outro' && iv.km ? km + iv.km : null,
      painel: painelInterno_({ veiculoId: v.id, item: item, manutencaoId: id, motivo: 'Manutenção registrada' })
    };
  });
}

// ---------- hodômetro ----------

function atualizarHodometro_(data) {
  return withLock_(function () {
    var shV = sheetVeiculosMnt_();
    var veics = lerVeiculosMnt_(shV);
    var v = acharVeiculo_(veics, data.veiculoId);
    if (!v) return { ok: false, error: 'Viatura não encontrada. Atualize a página e tente de novo.' };
    var km = parseKm_(data.km);
    if (km === null) return { ok: false, error: 'Informe a quilometragem atual (somente números).' };
    var agora = agoraIso_();
    var quando = String(data.data || '').trim();
    if (quando) {
      var d = dataValida_(quando);
      if (!d) return { ok: false, error: 'Data da leitura inválida.' };
      if (d > amanhaStr_()) return { ok: false, error: 'A data da leitura não pode estar no futuro.' };
      // Leitura de hoje guarda a hora real; data passada fica ao meio-dia.
      if (quando.length <= 10) quando = d === hojeStr_() ? agora : d + 'T12:00:00';
    } else {
      quando = agora;
    }
    var anterior = v.kmAtual;
    var regressao = anterior !== null && km < anterior;
    var motivo = String(data.motivo || '').trim();
    if (regressao) {
      if (!(data.confirmarRegressao === true || data.confirmarRegressao === 'true')) {
        return { ok: false, regressao: true, kmAtual: anterior, error: 'A nova quilometragem (' + km + ' km) é menor que a última registrada (' + anterior + ' km). Confirme a correção.' };
      }
      if (!senhaValida_(data.senha)) return { ok: false, regressao: true, kmAtual: anterior, error: 'Senha da diretoria incorreta. A redução do hodômetro precisa de autorização.' };
      if (!motivo) return { ok: false, regressao: true, kmAtual: anterior, error: 'Informe o motivo da correção.' };
    }
    atualizarCelulas_(shV, v._row, {
      kmAtual: km, kmAtualizadoEm: quando, kmOrigem: regressao ? 'Correção autorizada' : 'Leitura do hodômetro', kmRegistradoEm: agora
    }, TEXTO_VEICULOS_MNT);
    anexarLinha_(sheetHodometro_(), {
      id: Utilities.getUuid(), veiculoId: v.id, placa: v.placa, data: quando, km: km,
      kmAnterior: anterior === null ? '' : anterior, origem: regressao ? 'correcao' : 'manual', ref: '',
      responsavel: String(data.responsavel || '').trim().slice(0, 120),
      obs: (regressao ? 'Motivo: ' + motivo + (data.obs ? ' | ' : '') : '') + String(data.obs || '').trim().slice(0, 500),
      regressao: regressao ? 'Sim' : '', autorizadoPor: regressao ? (String(data.autorizadoPor || '').trim() || 'Diretoria (senha)') : '',
      registradoEm: agora
    }, TEXTO_HODOMETRO);
    return { ok: true, painel: painelInterno_(null) };
  });
}

function historicoHodometro_(data) {
  var id = String(data.veiculoId || '');
  var lista = lerTabela_(sheetHodometro_())
    .filter(function (h) { return String(h.veiculoId || '') === id; })
    .map(function (h) {
      return {
        data: txt_(h.data), km: num_(h.km), kmAnterior: num_(h.kmAnterior), origem: String(h.origem || ''),
        responsavel: String(h.responsavel || ''), obs: String(h.obs || ''), regressao: String(h.regressao || ''),
        autorizadoPor: String(h.autorizadoPor || ''), registradoEm: txt_(h.registradoEm)
      };
    })
    .sort(function (a, b) { return (b.registradoEm || b.data).localeCompare(a.registradoEm || a.data); });
  return { ok: true, leituras: lista };
}

// ---------- exclusão (somente diretoria; exclusão lógica) ----------

function excluirManutencao_(data) {
  if (!senhaValida_(data.senha)) return { ok: false, error: 'Senha da diretoria incorreta.' };
  var motivo = String(data.motivo || '').trim();
  if (!motivo) return { ok: false, error: 'Informe o motivo da exclusão.' };
  return withLock_(function () {
    var sh = sheetManut_();
    var linhas = lerTabela_(sh);
    var alvo = null;
    for (var i = 0; i < linhas.length; i++) if (String(linhas[i].id) === String(data.id || '')) { alvo = linhas[i]; break; }
    if (!alvo) return { ok: false, error: 'Registro não encontrado.' };
    if (String(alvo.excluido) === 'Sim') return { ok: false, error: 'Este registro já foi excluído.' };
    atualizarCelulas_(sh, alvo._row, {
      excluido: 'Sim', excluidoEm: agoraIso_(), excluidoPor: String(data.nome || '').trim() || 'Diretoria', motivoExclusao: motivo.slice(0, 500)
    }, TEXTO_MANUT);
    return { ok: true, painel: painelInterno_({ veiculoId: String(alvo.veiculoId), item: String(alvo.item), motivo: 'Registro excluído' }) };
  });
}

// ---------- cadastro / configuração da viatura (somente diretoria) ----------

function intervaloValido_(v, min) {
  if (v === null || v === undefined || String(v).trim() === '') return '';
  var n = parseKm_(v);
  if (n === null || n < min || n > 1000000) return null;
  return n;
}

function salvarViaturaManutencao_(data) {
  if (!senhaValida_(data.senha)) return { ok: false, error: 'Senha da diretoria incorreta.' };
  var tipo = String(data.tipo || 'carro');
  if (tipo !== 'carro' && tipo !== 'moto') return { ok: false, error: 'Tipo de veículo inválido.' };
  var comando = String(data.comando || '');
  if (['', 'correia', 'corrente'].indexOf(comando) < 0) return { ok: false, error: 'Sistema de comando inválido.' };
  var ano = String(data.ano || '').trim();
  if (ano && (!/^\d{4}$/.test(ano) || +ano < 1950 || +ano > new Date().getFullYear() + 1)) return { ok: false, error: 'Ano inválido.' };
  var iv = data.intervalos || {};
  var intervalos = {};
  var nomesIv = { oleo: 'óleo', comando: 'correia/corrente', pneuDianteiro: 'pneus dianteiros', pneuTraseiro: 'pneus traseiros', alertaKm: 'alerta antecipado' };
  for (var k in nomesIv) {
    var val = intervaloValido_(iv[k], k === 'alertaKm' ? 0 : 1);
    if (val === null) return { ok: false, error: 'Intervalo de ' + nomesIv[k] + ' inválido.' };
    if (val !== '') intervalos[k] = val;
  }
  var campos = {
    marca: String(data.marca || '').trim().slice(0, 60),
    modelo: String(data.modelo || '').trim().slice(0, 60),
    tipo: tipo, ano: ano, motor: String(data.motor || '').trim().slice(0, 60), comando: comando,
    intervalosJSON: JSON.stringify(intervalos), atualizadoEm: agoraIso_()
  };
  return withLock_(function () {
    var shV = sheetVeiculosMnt_();
    var veics = lerVeiculosMnt_(shV);
    var novoId = '';
    if (data.veiculoId) {
      var v = acharVeiculo_(veics, data.veiculoId);
      if (!v) return { ok: false, error: 'Viatura não encontrada.' };
      atualizarCelulas_(shV, v._row, campos, TEXTO_VEICULOS_MNT);
      novoId = v.id;
    } else {
      var nome = String(data.veiculo || '').trim();
      var placa = String(data.placa || '').trim().toUpperCase();
      if (!nome) return { ok: false, error: 'Informe o nome do veículo.' };
      if (placa && veics.some(function (x) { return placaNorm_(x.placa) === placaNorm_(placa); })) {
        return { ok: false, error: 'Já existe uma viatura com a placa ' + placa + '.' };
      }
      var kmIni = String(data.kmInicial || '').trim() === '' ? null : parseKm_(data.kmInicial);
      if (String(data.kmInicial || '').trim() !== '' && kmIni === null) return { ok: false, error: 'Quilometragem atual inválida.' };
      novoId = Utilities.getUuid();
      var agora = agoraIso_();
      var obj = { veiculo: nome, placa: placa, timestamp: agora, id: novoId };
      for (var c in campos) obj[c] = campos[c];
      if (kmIni !== null) { obj.kmAtual = kmIni; obj.kmAtualizadoEm = agora; obj.kmOrigem = 'Cadastro da viatura'; obj.kmRegistradoEm = agora; }
      anexarLinha_(shV, obj, TEXTO_VEICULOS_MNT);
      if (kmIni !== null) {
        anexarLinha_(sheetHodometro_(), {
          id: Utilities.getUuid(), veiculoId: novoId, placa: placa, data: agora, km: kmIni, kmAnterior: '',
          origem: 'cadastro', ref: '', responsavel: '', obs: 'Quilometragem informada no cadastro', regressao: '', autorizadoPor: '', registradoEm: agora
        }, TEXTO_HODOMETRO);
      }
    }
    return { ok: true, id: novoId, painel: painelInterno_(null) };
  });
}

function salvarConfigManutencao_(data) {
  if (!senhaValida_(data.senha)) return { ok: false, error: 'Senha da diretoria incorreta.' };
  var c = data.config || {};
  var cfg = { comandoModelos: [] };
  var nomes = { oleoCarro: 'óleo (carros)', oleoMoto: 'óleo (motos)', pneuDianteiro: 'pneus dianteiros', pneuTraseiro: 'pneus traseiros', alertaKm: 'alerta antecipado' };
  for (var k in nomes) {
    var n = parseKm_(c[k]);
    if (n === null || (k !== 'alertaKm' && n < 1) || n > 1000000) return { ok: false, error: 'Valor inválido para ' + nomes[k] + '.' };
    cfg[k] = n;
  }
  var vistos = {};
  var lista = Array.isArray(c.comandoModelos) ? c.comandoModelos : [];
  for (var i = 0; i < lista.length; i++) {
    var m = lista[i] || {};
    var modelo = String(m.modelo || '').trim().replace(/:/g, '');
    if (!modelo) continue;
    var km = parseKm_(m.km);
    if (km === null || km < 1 || km > 1000000) return { ok: false, error: 'Intervalo inválido para o modelo ' + modelo + '.' };
    var chave = normTxtMnt_(modelo);
    if (vistos[chave]) return { ok: false, error: 'O modelo ' + modelo + ' aparece duas vezes.' };
    vistos[chave] = true;
    cfg.comandoModelos.push({ modelo: modelo.slice(0, 60), marca: String(m.marca || '').trim().slice(0, 60), km: km });
  }
  return withLock_(function () {
    escreverConfigManut_(garantirAba_(SHEET_CONFIG_MANUT, HEADERS_CONFIG_MANUT, []), cfg);
    return { ok: true, painel: painelInterno_(null) };
  });
}

// ---------- comprovante / nota fiscal da manutenção ----------

function uploadComprovanteManutencao_(data) {
  if (!data.base64) return { ok: false, error: 'Arquivo vazio.' };
  var tipos = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];
  var mimeType = String(data.mimeType || 'image/jpeg');
  if (tipos.indexOf(mimeType) < 0) return { ok: false, error: 'Envie uma foto (JPG/PNG) ou um PDF.' };
  if (String(data.base64).length > 14000000) return { ok: false, error: 'Arquivo muito grande (máximo 10 MB).' };
  var props = PropertiesService.getScriptProperties();
  var folder = null;
  var idPasta = props.getProperty('PASTA_MANUTENCAO_ID');
  if (idPasta) { try { folder = DriveApp.getFolderById(idPasta); } catch (e) { folder = null; } }
  if (!folder) {
    var nomePasta = 'Vegas - Comprovantes de Manutencao Preventiva';
    var it = DriveApp.getFoldersByName(nomePasta);
    folder = it.hasNext() ? it.next() : DriveApp.createFolder(nomePasta);
    props.setProperty('PASTA_MANUTENCAO_ID', folder.getId());
  }
  var ext = mimeType === 'application/pdf' ? '.pdf' : (mimeType === 'image/png' ? '.png' : '.jpg');
  var nome = String(data.filename || ('manutencao_' + Utilities.getUuid() + ext)).replace(/[\\/]/g, '_');
  var blob;
  try { blob = Utilities.newBlob(Utilities.base64Decode(data.base64), mimeType, nome); }
  catch (e) { return { ok: false, error: 'Não foi possível processar o arquivo enviado.' }; }
  var file = folder.createFile(blob);
  var acesso = props.getProperty('PDF_ACESSO') || 'LINK';
  try {
    if (acesso === 'LINK') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    else if (acesso === 'DOMINIO') file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* mantém padrão */ }
  return { ok: true, url: file.getUrl(), id: file.getId() };
}
