// ניר הנדסה - Apps Script Web App

var SHEET_ID = '1rOFSo4vz8BoGBydWGgTTMM5xT6Bpm81NLd_1fTTm37Q';

function doGet(e) {
  return respond({ok:true, msg:'API active'});
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    if (action === 'update')           return respond(updateRow(data));
    if (action === 'delete')           return respond(deleteRow(data));
    if (action === 'archive')          return respond(archiveRow(data));
    if (action === 'restore')          return respond(restoreRow(data));
    if (action === 'sendNotification') return respond(sendNotification(data));
    throw new Error('unknown action: ' + action);
  } catch(err) {
    debugLog('ERROR: ' + err.message);
    return respond({ok:false, error: err.message});
  }
}

function cleanStr(s) {
  if (!s) return '';
  var result = String(s);
  var chars = ['\u200f','\u200e','\u202a','\u202b','\u202c','\ufeff','\u200b','\u200c','\u200d'];
  for (var i = 0; i < chars.length; i++) {
    result = result.split(chars[i]).join('');
  }
  return result.trim();
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var names = ['\u05de\u05e1\u05de\u05db\u05d9\u05dd', '\u05d2\u05d9\u05dc\u05d9\u05d5\u05df1', 'Sheet1'];
  for (var i = 0; i < names.length; i++) {
    var s = ss.getSheetByName(names[i]);
    if (s) return s;
  }
  return ss.getSheets()[0];
}

function getHeaders(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function colIdx(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] && cleanStr(String(headers[i])).indexOf(name) >= 0) return i;
  }
  return -1;
}

function debugLog(msg) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var log = ss.getSheetByName('Debug');
    if (!log) log = ss.insertSheet('Debug');
    log.appendRow([new Date(), msg]);
  } catch(e) {}
}

function findRow(sheet, filename, client) {
  var cleanFile   = cleanStr(filename);
  var cleanClient = cleanStr(client);
  var headers = getHeaders(sheet);
  var iFile = colIdx(headers, '\u05e9\u05dd \u05e7\u05d5\u05d1\u05e5');
  var iCli  = colIdx(headers, '\u05dc\u05e7\u05d5\u05d7');
  if (iFile < 0 || iCli < 0) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var data = sheet.getRange(2, 1, lastRow-1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < data.length; i++) {
    var rowFile   = cleanStr(String(data[i][iFile] || ''));
    var rowClient = cleanStr(String(data[i][iCli]  || ''));
    if (rowFile === cleanFile && rowClient === cleanClient) {
      return i + 2;
    }
  }
  return -1;
}

function updateRow(data) {
  var sheet = getSheet();
  var headers = getHeaders(sheet);
  var rowNum = findRow(sheet, data.filename, data.originalClient || data.client);
  if (rowNum < 0) throw new Error('row not found: ' + data.filename);
  var colNames = ['\u05dc\u05e7\u05d5\u05d7','\u05de\u05d9\u05e7\u05d5\u05dd','\u05e1\u05d5\u05d2','\u05ea\u05d0\u05e8\u05d9\u05da \u05de\u05e1\u05de\u05da','\u05ea\u05d5\u05e7\u05e3','\u05e7\u05d9\u05e9\u05d5\u05e8'];
  var colVals  = [data.client, data.location, data.docType, data.docDate, data.expiry, data.link];
  for (var i = 0; i < colNames.length; i++) {
    if (colVals[i] === undefined) continue;
    var ci = colIdx(headers, colNames[i]);
    if (ci >= 0) sheet.getRange(rowNum, ci+1).setValue(colVals[i] || '');
  }
  return {ok:true};
}

function deleteRow(data) {
  var sheet = getSheet();
  var rowNum = findRow(sheet, data.filename, data.client);
  if (rowNum < 0) throw new Error('row not found: ' + data.filename);
  sheet.deleteRow(rowNum);
  return {ok:true};
}

function archiveRow(data) {
  debugLog('archiveRow: ' + data.filename);
  var sheet = getSheet();
  debugLog('sheet name: ' + sheet.getName());
  var headers = getHeaders(sheet);
  debugLog('headers count: ' + headers.length);
  debugLog('headers: ' + JSON.stringify(headers));
  var rowNum = findRow(sheet, data.filename, data.client);
  debugLog('rowNum: ' + rowNum);
  if (rowNum < 0) throw new Error('row not found: ' + data.filename);
  var iArc = colIdx(headers, '\u05d0\u05e8\u05db\u05d9\u05d5\u05df');
  debugLog('iArc: ' + iArc);
  if (iArc < 0) {
    sheet.getRange(1, headers.length+1).setValue('\u05d0\u05e8\u05db\u05d9\u05d5\u05df');
    iArc = headers.length;
    debugLog('created archive col at: ' + (iArc+1));
  }
  sheet.getRange(rowNum, iArc+1).setValue('\u05db\u05df');
  debugLog('archive done!');
  return {ok:true};
}

function restoreRow(data) {
  var sheet = getSheet();
  var headers = getHeaders(sheet);
  var rowNum = findRow(sheet, data.filename, data.client);
  if (rowNum < 0) throw new Error('row not found: ' + data.filename);
  var iArc = colIdx(headers, '\u05d0\u05e8\u05db\u05d9\u05d5\u05df');
  if (iArc >= 0) sheet.getRange(rowNum, iArc+1).setValue('');
  return {ok:true};
}

// ── שליחת התראות מייל על מסמכים פוקעים ──
function sendNotification(data) {
  var emails   = data.emails   || [];
  var docs     = data.docs     || [];
  var critDocs = data.critDocs || [];
  var daysWarn = data.daysWarn || 90;

  if (emails.length === 0) throw new Error('no email recipients');
  if (docs.length === 0 && critDocs.length === 0) return {ok:true, sent:0};

  debugLog('sendNotification: ' + emails.length + ' recipients, ' + docs.length + ' docs, ' + critDocs.length + ' crit');

  // ── נושא המייל ──
  var subject = critDocs.length > 0
    ? '\uD83D\uDD34 דחוף! ניר הנדסה — ' + critDocs.length + ' מסמכים קריטיים פוקעים'
    : '\u26A0\uFE0F ניר הנדסה — ' + docs.length + ' מסמכים פוקעים תוך ' + daysWarn + ' ימים';

  // ── גוף המייל (טקסט רגיל) ──
  var lines = [];
  lines.push('שלום,');
  lines.push('');
  lines.push('להלן רשימת המסמכים הדורשים טיפול:');
  lines.push('');

  if (critDocs.length > 0) {
    lines.push('🔴 קריטי — פוקעים בקרוב מאוד:');
    for (var i = 0; i < critDocs.length; i++) {
      var d = critDocs[i];
      var when = d.remaining === 0 ? 'היום!' : 'בעוד ' + d.remaining + ' ימים';
      lines.push('  • ' + d.client + ' — ' + d.docType + ' | תוקף: ' + d.expiry + ' (' + when + ')');
      if (d.location) lines.push('    מיקום: ' + d.location);
    }
    lines.push('');
  }

  if (docs.length > 0) {
    lines.push('🟠 פוקעים תוך ' + daysWarn + ' ימים:');
    for (var j = 0; j < docs.length; j++) {
      var doc = docs[j];
      var docWhen = doc.remaining === 0 ? 'היום!' : 'בעוד ' + doc.remaining + ' ימים';
      lines.push('  • ' + doc.client + ' — ' + doc.docType + ' | תוקף: ' + doc.expiry + ' (' + docWhen + ')');
      if (doc.location) lines.push('    מיקום: ' + doc.location);
    }
  }

  lines.push('');
  lines.push('— מערכת מעקב מסמכים, ניר הנדסה');

  var body = lines.join('\n');

  // ── גוף HTML (נראה טוב יותר בגוגל/אאוטלוק) ──
  var htmlLines = [];
  htmlLines.push('<div dir="rtl" style="font-family:Arial,sans-serif;font-size:14px;color:#1e293b">');
  htmlLines.push('<p>שלום,</p>');
  htmlLines.push('<p>להלן רשימת המסמכים הדורשים טיפול:</p>');

  if (critDocs.length > 0) {
    htmlLines.push('<h3 style="color:#ef4444">🔴 קריטי — פוקעים בקרוב מאוד</h3>');
    htmlLines.push('<table style="border-collapse:collapse;width:100%;margin-bottom:16px">');
    htmlLines.push('<tr style="background:#fef2f2"><th style="padding:6px 10px;border:1px solid #fecaca;text-align:right">לקוח</th><th style="padding:6px 10px;border:1px solid #fecaca;text-align:right">סוג מסמך</th><th style="padding:6px 10px;border:1px solid #fecaca;text-align:right">תוקף</th><th style="padding:6px 10px;border:1px solid #fecaca;text-align:right">ימים נותרו</th></tr>');
    for (var ci = 0; ci < critDocs.length; ci++) {
      var cd = critDocs[ci];
      var cdWhen = cd.remaining === 0 ? '<strong style="color:#ef4444">היום!</strong>' : cd.remaining + ' ימים';
      htmlLines.push('<tr><td style="padding:6px 10px;border:1px solid #fecaca">' + cd.client + '</td><td style="padding:6px 10px;border:1px solid #fecaca">' + cd.docType + '</td><td style="padding:6px 10px;border:1px solid #fecaca">' + cd.expiry + '</td><td style="padding:6px 10px;border:1px solid #fecaca;font-weight:bold;color:#ef4444">' + cdWhen + '</td></tr>');
    }
    htmlLines.push('</table>');
  }

  if (docs.length > 0) {
    htmlLines.push('<h3 style="color:#f97316">🟠 פוקעים תוך ' + daysWarn + ' ימים</h3>');
    htmlLines.push('<table style="border-collapse:collapse;width:100%;margin-bottom:16px">');
    htmlLines.push('<tr style="background:#fff7ed"><th style="padding:6px 10px;border:1px solid #fed7aa;text-align:right">לקוח</th><th style="padding:6px 10px;border:1px solid #fed7aa;text-align:right">סוג מסמך</th><th style="padding:6px 10px;border:1px solid #fed7aa;text-align:right">תוקף</th><th style="padding:6px 10px;border:1px solid #fed7aa;text-align:right">ימים נותרו</th></tr>');
    for (var di = 0; di < docs.length; di++) {
      var dd = docs[di];
      var ddWhen = dd.remaining === 0 ? '<strong style="color:#f97316">היום!</strong>' : dd.remaining + ' ימים';
      htmlLines.push('<tr><td style="padding:6px 10px;border:1px solid #fed7aa">' + dd.client + '</td><td style="padding:6px 10px;border:1px solid #fed7aa">' + dd.docType + '</td><td style="padding:6px 10px;border:1px solid #fed7aa">' + dd.expiry + '</td><td style="padding:6px 10px;border:1px solid #fed7aa;font-weight:bold;color:#f97316">' + ddWhen + '</td></tr>');
    }
    htmlLines.push('</table>');
  }

  htmlLines.push('<p style="color:#94a3b8;font-size:12px;margin-top:24px">— מערכת מעקב מסמכים, ניר הנדסה</p>');
  htmlLines.push('</div>');
  var htmlBody = htmlLines.join('');

  // ── שליחה ──
  var sent = 0;
  for (var ei = 0; ei < emails.length; ei++) {
    try {
      MailApp.sendEmail({
        to:       emails[ei],
        subject:  subject,
        body:     body,
        htmlBody: htmlBody
      });
      sent++;
      debugLog('mail sent to: ' + emails[ei]);
    } catch(mailErr) {
      debugLog('mail error to ' + emails[ei] + ': ' + mailErr.message);
    }
  }

  return {ok: true, sent: sent};
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
