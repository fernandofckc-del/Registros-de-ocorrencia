/**
 * REGISTRO DE OCORRÊNCIAS - IRRIGAÇÃO
 * Backend em Google Apps Script.
 *
 * O QUE ISSO FAZ:
 * - Recebe os registros que o app manda (quando o celular está online)
 * - Salva os dados numa aba do Google Sheets (a mesma planilha onde você colar este script)
 * - Salva as fotos (até várias por ocorrência) numa pasta do Google Drive
 * - Guarda a lista de "Tipos de Ocorrência", editável apenas por quem sabe a senha de administrador
 * - Deixa você (ou qualquer app) puxar todos os dados já sincronizados a qualquer momento
 *
 * COMO INSTALAR - veja o arquivo INSTRUCOES.md que veio junto com este arquivo.
 */

// ==== TROQUE ESTE TOKEN por uma palavra secreta sua antes de publicar ====
// Esse token TODO ENCARREGADO vai usar no celular dele pra poder sincronizar.
var SECRET = '123456789';

// ==== TROQUE ESTA SENHA por outra, diferente da de cima ====
// Essa senha SÓ VOCÊ deve saber. É o que protege a edição dos "Tipos de Ocorrência".
var ADMIN_SECRET = '987654321';

var SHEET_NAME = 'Ocorrencias';
var FOLDER_NAME = 'Fotos_Ocorrencias_Irrigacao';
var FOTO_SEP = ' || ';

var TIPOS_PROP_KEY = 'TIPOS_OCORRENCIA';
var TIPOS_PADRAO = [
  'Mangueira Desconectada',
  'Vazamento',
  'Conector Quebrado',
  'Mangueira Furada (Espinho/Roseta)',
  'Registro/Válvula com Defeito',
  'Gotejador Entupido',
  'Outro'
];

var COLUNAS = ['id', 'data', 'setor', 'tipo', 'causa', 'encarregado', 'descricao',
               'acao', 'status', 'foto_url', 'excluido', 'criado_em'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUNAS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getFolder_() {
  var folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // linha real na planilha
  }
  return -1;
}

function saveFoto_(dataUrl, id) {
  var match = dataUrl.match(/^data:(image\/\w+);base64,(.*)$/);
  if (!match) return '';
  var mime = match[1];
  var base64 = match[2];
  var bytes = Utilities.base64Decode(base64);
  var blob = Utilities.newBlob(bytes, mime, id + '.jpg');
  var folder = getFolder_();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

// Recebe uma lista (array) onde cada item é OU uma foto nova (base64 "data:image/...")
// OU uma URL que já existia. Faz upload só das novas, e devolve a lista final de URLs.
function processFotos_(fotos, baseId) {
  if (!fotos || !fotos.length) return [];
  var urls = [];
  fotos.forEach(function (item, i) {
    if (!item) return;
    if (String(item).indexOf('data:image') === 0) {
      var url = saveFoto_(item, baseId + '_' + i + '_' + new Date().getTime());
      if (url) urls.push(url);
    } else {
      urls.push(item); // já era uma URL, mantém
    }
  });
  return urls;
}

function getTipos_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(TIPOS_PROP_KEY);
  if (!raw) {
    props.setProperty(TIPOS_PROP_KEY, JSON.stringify(TIPOS_PADRAO));
    return TIPOS_PADRAO;
  }
  try {
    var arr = JSON.parse(raw);
    return (arr && arr.length) ? arr : TIPOS_PADRAO;
  } catch (e) {
    return TIPOS_PADRAO;
  }
}

// GET -> devolve todos os registros + a lista atual de tipos de ocorrência
function doGet(e) {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var rows = data.map(function (r) {
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = r[i]; });
    obj.fotos = obj.foto_url ? String(obj.foto_url).split(FOTO_SEP).filter(function (x) { return x; }) : [];
    return obj;
  }).filter(function (r) { return !r.excluido; });
  return jsonOut_({ ok: true, rows: rows, tipos: getTipos_() });
}

// POST -> recebe um registro novo, edição completa, mudança de status, exclusão, ou atualização dos tipos
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'create';

    // Ação especial de administrador: não usa o token normal, usa a senha de admin
    if (action === 'update_tipos') {
      if (body.adminToken !== ADMIN_SECRET) {
        return jsonOut_({ ok: false, error: 'senha de administrador inválida' });
      }
      var novosTipos = (body.tipos || []).map(function (t) { return String(t).trim(); }).filter(function (t) { return t; });
      if (!novosTipos.length) {
        return jsonOut_({ ok: false, error: 'lista de tipos vazia' });
      }
      PropertiesService.getScriptProperties().setProperty(TIPOS_PROP_KEY, JSON.stringify(novosTipos));
      return jsonOut_({ ok: true, tipos: novosTipos });
    }

    if (body.token !== SECRET) {
      return jsonOut_({ ok: false, error: 'token inválido' });
    }
    var sheet = getSheet_();

    if (action === 'create') {
      // evita duplicar se o app reenviar o mesmo registro por engano
      if (findRowById_(sheet, body.id) > -1) {
        return jsonOut_({ ok: true, duplicado: true });
      }
      var urls = processFotos_(body.fotos, body.id);
      sheet.appendRow([
        body.id, body.data, body.setor, body.tipo, body.causa,
        body.encarregado, body.descricao, body.acao || '', body.status || 'Aberto',
        urls.join(FOTO_SEP), false, new Date()
      ]);
      return jsonOut_({ ok: true, fotoUrls: urls });
    }

    if (action === 'update') {
      var rowIdxU = findRowById_(sheet, body.id);
      if (rowIdxU === -1) return jsonOut_({ ok: false, error: 'registro não encontrado' });
      sheet.getRange(rowIdxU, 1, 1, 9).setValues([[
        body.id, body.data, body.setor, body.tipo, body.causa,
        body.encarregado, body.descricao, body.acao || '', body.status || 'Aberto'
      ]]);
      var updatedUrls;
      if (body.fotos !== undefined) {
        updatedUrls = processFotos_(body.fotos, body.id);
        sheet.getRange(rowIdxU, 10).setValue(updatedUrls.join(FOTO_SEP));
      }
      return jsonOut_({ ok: true, fotoUrls: updatedUrls });
    }

    if (action === 'update_status') {
      var rowIdx = findRowById_(sheet, body.id);
      if (rowIdx > -1) sheet.getRange(rowIdx, 9).setValue(body.status); // coluna 9 = status
      return jsonOut_({ ok: true });
    }

    if (action === 'delete') {
      var rowIdx2 = findRowById_(sheet, body.id);
      if (rowIdx2 > -1) sheet.getRange(rowIdx2, 11).setValue(true); // coluna 11 = excluido
      return jsonOut_({ ok: true });
    }

    return jsonOut_({ ok: false, error: 'ação desconhecida: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}
