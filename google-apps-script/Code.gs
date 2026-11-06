/**
 * Jednokratno, u Google nalogu koji poseduje tabelu.
 * Ne deli se tabela sa nikim drugim: skripta se izvršava kao vlasnik.
 *
 * 1. Otvori tabelu koju sajt već čita (File → Share → Publish to web, CSV).
 * 2. Extensions → Apps Script. Obriši prazan fajl i nalepi ovaj kod.
 * 3. Deploy → New deployment → Type: Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 4. Kopiraj Web app URL koji se završava na /exec.
 * 5. Na sajtu otvori "Ubaci Hevy tekst", otvori "Jednokratno: adresa za upis"
 *    i nalepi tu adresu. Ostaje u ovom pregledaču.
 *
 * Svaki sledeći nalepak ide POST-om ovde. Isti datum + naziv treninga
 * zamenjuje stare redove, pa ponovni uvoz ne duplira serije.
 */
var HEADER = ['Date', 'Workout', 'Exercise', 'Set', 'Weight_kg', 'Reps', 'Distance_km', 'Duration_sec', 'Notes'];

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, service: 'hevy-sheet-append' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var body = readBody_(e);
    var rows = Array.isArray(body.rows) ? body.rows : [];
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    ensureHeader_(sheet);
    var written = replaceWorkouts_(sheet, rows);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, written: written }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message || err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function readBody_(e) {
  var raw = '';
  if (e && e.parameter && e.parameter.payload) raw = e.parameter.payload;
  else if (e && e.postData && e.postData.contents) raw = e.postData.contents;
  if (!raw) return {};
  return JSON.parse(raw);
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    return;
  }
  var first = sheet.getRange(1, 1, 1, HEADER.length).getValues()[0];
  var same = HEADER.every(function(name, i) {
    return String(first[i] || '').trim().toLowerCase() === name.toLowerCase();
  });
  if (!same) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
}

function replaceWorkouts_(sheet, rows) {
  var clean = rows.filter(function(row) {
    return row && row.date && row.workout && row.exercise;
  });
  var keys = {};
  clean.forEach(function(row) {
    keys[String(row.date) + '||' + String(row.workout)] = true;
  });
  var tz = sheet.getParent().getSpreadsheetTimeZone();
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    var key = cellDate_(values[i][0], tz) + '||' + String(values[i][1] || '');
    if (keys[key]) sheet.deleteRow(i + 1);
  }
  if (!clean.length) return 0;
  var matrix = clean.map(function(row) {
    return [
      String(row.date),
      String(row.workout),
      String(row.exercise),
      row.set === undefined || row.set === null || row.set === '' ? '' : Number(row.set),
      row.weight === undefined || row.weight === null || row.weight === '' ? '' : Number(row.weight),
      row.reps === undefined || row.reps === null || row.reps === '' ? '' : Number(row.reps),
      row.distance === undefined || row.distance === null || row.distance === '' ? '' : Number(row.distance),
      row.duration === undefined || row.duration === null || row.duration === '' ? '' : Number(row.duration),
      row.notes ? String(row.notes) : ''
    ];
  });
  var start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, matrix.length, HEADER.length).setValues(matrix);
  return matrix.length;
}

function cellDate_(value, tz) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var match = String(value || '').trim().match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : String(value || '').trim();
}
