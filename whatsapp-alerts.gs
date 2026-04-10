// ═══════════════════════════════════════════════════════════════════
//  WhatsApp Alerts — ניר הנדסה
//  ───────────────────────────────────────────────────────────────────
//  הוסף קובץ זה לפרויקט Apps Script הקיים (זה שמאחורי APPS_SCRIPT_URL)
//
//  הגדרה ראשונית (פעם אחת בלבד):
//  1. מלא את WA_PHONE ו-WA_API_KEY למטה
//  2. הפעל את createDailyTrigger() פעם אחת מהעורך → Triggers יווצר אוטומטית
//  3. הפעל את testWhatsApp() כדי לוודא שהחיבור עובד
// ═══════════════════════════════════════════════════════════════════

// ── הגדרות — מלא כאן ──────────────────────────────────────────────
const WA_PHONE   = 'XXXXXXXXXXX';   // מספר שלך עם קידומת מדינה, ללא + (דוגמה: 972501234567)
const WA_API_KEY = 'XXXXXXX';       // המפתח שתקבל מ-CallMeBot (ראה הוראות הגדרה למטה)
const SHEET_NAME = '';              // שם הגיליון — השאר ריק לשימוש בגיליון הראשון
const ALERT_HOUR = 8;               // שעת שליחה יומית (0–23), ברירת מחדל: 8 בבוקר
// ─────────────────────────────────────────────────────────────────

// ימים לפני תפוגה שבהם תישלח התרעה
const ALERT_DAYS = [90, 60, 30, 14, 7, 3, 1];

// ═══════════════════════════════════════════════════════════════════
//  הפונקציה הראשית — מופעלת יומית ע"י ה-Trigger
// ═══════════════════════════════════════════════════════════════════
function sendWhatsAppAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];

  if (!sh) {
    Logger.log('שגיאה: לא נמצא הגיליון "' + SHEET_NAME + '"');
    return;
  }

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;

  // מצא עמודות לפי כותרת
  const H    = data[0];
  const iClient  = H.findIndex(h => String(h).includes('לקוח'));
  const iType    = H.findIndex(h => String(h).includes('סוג'));
  const iExpiry  = H.findIndex(h => String(h).includes('תוקף'));
  const iArchive = H.findIndex(h => String(h).includes('ארכיון'));

  if (iClient < 0 || iExpiry < 0) {
    Logger.log('שגיאה: לא נמצאו עמודות "לקוח" / "תוקף" בגיליון');
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const alerts = { urgent: [], warning: [], notice: [] };

  for (let i = 1; i < data.length; i++) {
    const row      = data[i];
    const client   = clean(row[iClient]);
    const docType  = iType    >= 0 ? clean(row[iType])    : '';
    const expiry   = iExpiry  >= 0 ? clean(row[iExpiry])  : '';
    const archived = iArchive >= 0 && clean(row[iArchive]) === 'כן';

    if (!client || !expiry || archived) continue;

    const expiryDate = parseDate(expiry);
    if (!expiryDate) continue;

    const daysLeft = Math.floor((expiryDate - today) / 86400000);

    if (ALERT_DAYS.includes(daysLeft)) {
      const item = { client, docType, expiry, daysLeft };
      if      (daysLeft <= 7)  alerts.urgent.push(item);
      else if (daysLeft <= 30) alerts.warning.push(item);
      else                     alerts.notice.push(item);
    }
  }

  const total = alerts.urgent.length + alerts.warning.length + alerts.notice.length;

  if (total === 0) {
    Logger.log('אין מסמכים הדורשים התרעה היום (' + today.toLocaleDateString('he-IL') + ')');
    return;
  }

  const msg = buildMessage(alerts, today);
  sendWhatsApp(msg);
  Logger.log('נשלחה התרעה עם ' + total + ' מסמכים');
}

// ═══════════════════════════════════════════════════════════════════
//  בניית ההודעה
// ═══════════════════════════════════════════════════════════════════
function buildMessage(alerts, today) {
  const dateStr = today.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
  let msg = '⚠️ *ניר הנדסה — תזכורת תוקף מסמכים*\n';
  msg    += '📅 ' + dateStr + '\n';
  msg    += '─────────────────────\n\n';

  if (alerts.urgent.length) {
    msg += '🔴 *דחוף — פוקע תוך ' + Math.max(...alerts.urgent.map(a => a.daysLeft)) + ' ימים או פחות:*\n';
    alerts.urgent.forEach(a => {
      msg += '  • ' + a.client + (a.docType ? ' — ' + a.docType : '') + '\n';
      msg += '    תוקף: ' + a.expiry + ' (עוד ' + a.daysLeft + ' ' + dayWord(a.daysLeft) + ')\n';
    });
    msg += '\n';
  }

  if (alerts.warning.length) {
    msg += '🟠 *קרוב לפקיעה (עד 30 ימים):*\n';
    alerts.warning.forEach(a => {
      msg += '  • ' + a.client + (a.docType ? ' — ' + a.docType : '') + '\n';
      msg += '    תוקף: ' + a.expiry + ' (עוד ' + a.daysLeft + ' ' + dayWord(a.daysLeft) + ')\n';
    });
    msg += '\n';
  }

  if (alerts.notice.length) {
    msg += '🟡 *לתשומת לב (עד 90 ימים):*\n';
    alerts.notice.forEach(a => {
      msg += '  • ' + a.client + (a.docType ? ' — ' + a.docType : '') + '\n';
      msg += '    תוקף: ' + a.expiry + ' (עוד ' + a.daysLeft + ' ' + dayWord(a.daysLeft) + ')\n';
    });
    msg += '\n';
  }

  const total = alerts.urgent.length + alerts.warning.length + alerts.notice.length;
  msg += '─────────────────────\n';
  msg += 'סה"כ: ' + total + ' מסמכים דורשים טיפול';
  return msg;
}

function dayWord(n) { return n === 1 ? 'יום' : 'ימים'; }

// ═══════════════════════════════════════════════════════════════════
//  שליחה ל-CallMeBot
// ═══════════════════════════════════════════════════════════════════
function sendWhatsApp(text) {
  if (!WA_PHONE || WA_PHONE === 'XXXXXXXXXXX') {
    Logger.log('שגיאה: WA_PHONE לא הוגדר');
    return;
  }
  if (!WA_API_KEY || WA_API_KEY === 'XXXXXXX') {
    Logger.log('שגיאה: WA_API_KEY לא הוגדר');
    return;
  }

  const url = 'https://api.callmebot.com/whatsapp.php'
    + '?phone='  + encodeURIComponent(WA_PHONE)
    + '&text='   + encodeURIComponent(text)
    + '&apikey=' + encodeURIComponent(WA_API_KEY);

  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (code === 200) {
      Logger.log('✅ ווצאפ נשלח בהצלחה');
    } else {
      Logger.log('❌ שגיאת CallMeBot: קוד ' + code + ' — ' + res.getContentText().slice(0, 200));
    }
  } catch (e) {
    Logger.log('❌ שגיאת רשת: ' + e.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
//  כלי עזר
// ═══════════════════════════════════════════════════════════════════
function parseDate(s) {
  if (!s) return null;
  // פורמט DD/MM/YYYY
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  // פורמט YYYY-MM-DD
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function clean(s) {
  if (!s) return '';
  let r = String(s);
  for (const c of '\u200f\u200e\u202a\u202b\u202c\ufeff\u200b') r = r.split(c).join('');
  return r.trim();
}

// ═══════════════════════════════════════════════════════════════════
//  הגדרת Trigger יומי — הפעל פעם אחת בלבד!
// ═══════════════════════════════════════════════════════════════════
function createDailyTrigger() {
  // מחק triggers קיימים של אותה פונקציה (למניעת כפילויות)
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendWhatsAppAlerts')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendWhatsAppAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(ALERT_HOUR)
    .create();

  Logger.log('✅ Trigger יומי נוצר — יופעל כל יום בשעה ' + ALERT_HOUR + ':00');
}

// ═══════════════════════════════════════════════════════════════════
//  בדיקת חיבור — הפעל מהעורך לאחר מילוי WA_PHONE ו-WA_API_KEY
// ═══════════════════════════════════════════════════════════════════
function testWhatsApp() {
  sendWhatsApp('✅ בדיקת חיבור — מערכת ניר הנדסה פועלת!\nהתרעות ווצאפ מוגדרות בהצלחה 🎉');
}

// ═══════════════════════════════════════════════════════════════════
//  הרצה ידנית מיידית (לבדיקת כל הלוגיקה עם הנתונים האמיתיים)
// ═══════════════════════════════════════════════════════════════════
function runNow() {
  sendWhatsAppAlerts();
}
