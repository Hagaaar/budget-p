const SHEET_NAME    = 'Données';
const ARCHIVE_SHEET = 'Archives';
const JOURNAL_SHEET = 'Journal';
const SYNC_KEY       = 'dc8d372a5da56a7620f68626dc2fdb0d';

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cb = e.parameter.callback || null;

  function respond(text, mime) {
    return ContentService.createTextOutput(text).setMimeType(mime);
  }
  function jsonp(obj) {
    const s = JSON.stringify(obj);
    return cb
      ? respond(cb + '(' + s + ');', ContentService.MimeType.JAVASCRIPT)
      : respond(s, ContentService.MimeType.JSON);
  }

  const isWrite = e.parameter.action === 'save' || e.parameter.action === 'archive' || e.parameter.action === 'logEntry';
  if (isWrite && e.parameter.key !== SYNC_KEY) {
    return jsonp({ ok: false, error: 'unauthorized' });
  }

  if (e.parameter.action === 'logEntry') {
    let journal = ss.getSheetByName(JOURNAL_SHEET);
    if (!journal) {
      journal = ss.insertSheet(JOURNAL_SHEET);
      journal.appendRow(['Horodatage', 'Mois', 'Action', 'ID', 'Montant', 'Description', 'Catégorie']);
    }
    journal.appendRow([
      new Date().toISOString(),
      e.parameter.mois,
      e.parameter.entryAction,
      e.parameter.id,
      e.parameter.montant,
      e.parameter.desc,
      e.parameter.cat
    ]);
    return jsonp({ ok: true });
  }

  if (e.parameter.action === 'save') {
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue(e.parameter.data);
    return jsonp({ ok: true });
  }

  if (e.parameter.action === 'archive') {
    let arch = ss.getSheetByName(ARCHIVE_SHEET);
    if (!arch) arch = ss.insertSheet(ARCHIVE_SHEET);
    arch.appendRow([e.parameter.mois, new Date().toISOString(), e.parameter.data]);
    return jsonp({ ok: true });
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  const raw = sheet ? (sheet.getRange('A1').getValue() || '{}') : '{}';
  return cb
    ? respond(cb + '(' + raw + ');', ContentService.MimeType.JAVASCRIPT)
    : respond(raw, ContentService.MimeType.JSON);
}

function doPost(e) {
  return ContentService.createTextOutput('ok').setMimeType(ContentService.MimeType.TEXT);
}
