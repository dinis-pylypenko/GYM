/**
 * Синхронізація PWA «Треную» → Google Таблиця.
 * Встав цей код у таблицю: Розширення → Apps Script, заміни все, збережи,
 * потім Deploy → New deployment → Web app → Execute as: Me,
 * Who has access: Anyone → Deploy → скопіюй URL (…/exec) в апку.
 */
var LOG_NAME = 'Журнал';
var BW_NAME  = 'Вага тіла';
var LOG_HEADERS = ['Дата','Вправа','Підхід №','Вага, кг','Повтори','Тоннаж, кг','Оцінка 1ПМ','Коліно','Нотатка'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var res;
    if (data.action === 'replaceAll')      res = replaceAll_(ss, data);
    else if (data.action === 'append')     res = { ok: true, added: append_(ss, data.rows || []) };
    else                                   res = { ok: false, error: 'unknown action' };
    return out_(res);
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return out_({ ok: true, ping: 'gym-sync працює 💪' }); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}

function toDate_(iso) { var p = iso.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }

function firstEmpty_(sh) {
  var vals = sh.getRange('A2:A' + Math.max(sh.getLastRow(), 2)).getValues();
  for (var i = 0; i < vals.length; i++) if (vals[i][0] === '' || vals[i][0] === null) return i + 2;
  return vals.length + 2;
}

function formulas_(sh, row) {
  if (!sh.getRange(row, 6).getFormula())
    sh.getRange(row, 6).setFormula('=IF(E' + row + '="","",D' + row + '*E' + row + ')');
  if (!sh.getRange(row, 7).getFormula())
    sh.getRange(row, 7).setFormula('=IF(E' + row + '="","",ROUND(D' + row + '*(1+E' + row + '/30),1))');
}

/* швидкий допис нових записів */
function append_(ss, rows) {
  var log = sheet_(ss, LOG_NAME, LOG_HEADERS);
  var bw  = sheet_(ss, BW_NAME, ['Дата','Вага, кг']);
  var n = 0;
  rows.forEach(function (r) {
    if (r.t === 'set') {
      var row = firstEmpty_(log);
      log.getRange(row, 1, 1, 5).setValues([[toDate_(r.d), r.ex, r.i, r.w, r.r]]);
      log.getRange(row, 1).setNumberFormat('DD.MM.YYYY');
      if (r.knee) log.getRange(row, 8).setValue(r.knee);
      if (r.note) log.getRange(row, 9).setValue(r.note);
      formulas_(log, row);
      n++;
    } else if (r.t === 'bw') {
      var row2 = firstEmpty_(bw);
      bw.getRange(row2, 1, 1, 2).setValues([[toDate_(r.d), r.kg]]);
      bw.getRange(row2, 1).setNumberFormat('DD.MM.YYYY');
      n++;
    }
  });
  return n;
}

/* повна пересинхронізація: перезаписує Журнал і Вагу тіла даними з апки */
function replaceAll_(ss, data) {
  var log = sheet_(ss, LOG_NAME, LOG_HEADERS);
  var last = Math.max(log.getLastRow(), 2);
  log.getRange(2, 1, last - 1, 5).clearContent();
  log.getRange(2, 8, last - 1, 2).clearContent();
  var sets = data.sets || [];
  if (sets.length) {
    var a = sets.map(function (r) { return [toDate_(r.d), r.ex, r.i, r.w, r.r]; });
    var b = sets.map(function (r) { return [r.knee || '', r.note || '']; });
    log.getRange(2, 1, a.length, 5).setValues(a);
    log.getRange(2, 8, b.length, 2).setValues(b);
    log.getRange(2, 1, a.length, 1).setNumberFormat('DD.MM.YYYY');
    for (var i = 0; i < a.length; i++) formulas_(log, i + 2);
  }
  var bw = sheet_(ss, BW_NAME, ['Дата','Вага, кг']);
  var lastB = Math.max(bw.getLastRow(), 2);
  bw.getRange(2, 1, lastB - 1, 2).clearContent();
  var bws = data.bw || [];
  if (bws.length) {
    var c = bws.map(function (r) { return [toDate_(r.d), r.kg]; });
    bw.getRange(2, 1, c.length, 2).setValues(c);
    bw.getRange(2, 1, c.length, 1).setNumberFormat('DD.MM.YYYY');
  }
  return { ok: true, sets: sets.length, bw: bws.length };
}
