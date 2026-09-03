/**
 * אקספו הסופרים — קליטת רישומים לאלפון התושבים
 * ----------------------------------------------
 * גרסה 3 (3.9.26). שינוי מהותי מול הגרסה הקודמת:
 * הסקריפט **עצמאי** ולא יושב בתוך הגיליון.
 *
 * למה: בגרסה הקודמת הסקריפט היה כבול לגיליון (getActiveSpreadsheet),
 * וכשגוגל חסמה את הגיליון היא חסמה איתו את כתובת הקליטה — הטופס באתר
 * החזיר 403 והקליטה מתה. עכשיו הגיליון הוא רק יעד כתיבה. אם הוא ייחסם
 * שוב, הכתובת תמשיך לענות והרישום עדיין ייתפס במייל.
 *
 * ארבע שורות שאפשר לשנות נמצאות מיד למטה. כל השאר — לא לגעת.
 */

/* מזהה הגיליון — החלק הארוך בכתובת שלו, בין /d/ לבין /edit */
var SHEET_ID = 'PASTE-SHEET-ID-HERE';

/* המייל שיקבל התראה על כל רישום. השאירו ריק ('') כדי לבטל התראות. */
var NOTIFY_EMAIL = 'eh600601@gmail.com';

/* שם הלשונית בגיליון שאליה נכנסים הרישומים. */
var SHEET_NAME = 'רישומים';

/* סיסמת אבטחה. חייבת להיות זהה לזו שבאתר — אל תשנו. */
var TOKEN = 'sofrim-alfon-5787';


/* ==== מכאן והלאה אין מה לשנות ==== */

var COLUMNS = [
  'תאריך קבלה',
  'סוג הפנייה',
  'שם משפחה',
  'שם האיש',
  'שם האשה',
  'רחוב',
  'בניין',
  'כניסה',
  'קומה',
  'דירה',
  'טלפון בבית',
  'נייד הבעל',
  'נייד האשה',
  'מעמד'
];

/* עמודות שחייבות להישאר טקסט, כדי שאפס בהתחלה לא ייעלם */
var TEXT_COLUMNS_FROM = 7;
var TEXT_COLUMNS_TO   = 13;


function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents) {
      return json({ ok: false, error: 'empty request' });
    }

    var data = JSON.parse(e.postData.contents);

    if (data.token !== TOKEN) {
      return json({ ok: false, error: 'bad token' });
    }

    /* הכתיבה לגיליון והמייל הם שני נתיבים נפרדים בכוונה.
       מספיק שאחד מהם הצליח כדי שהרישום לא ילך לאיבוד. */
    var wrote = false;
    var sheetError = '';
    try {
      getSheet().appendRow([
        new Date(),
        pick(data, 'סוג'),
        pick(data, 'שם משפחה'),
        pick(data, 'שם האיש'),
        pick(data, 'שם האשה'),
        pick(data, 'רחוב'),
        pickText(data, 'בניין'),
        pickText(data, 'כניסה'),
        pickText(data, 'קומה'),
        pickText(data, 'דירה'),
        pickText(data, 'טלפון בית'),
        pickText(data, 'טלפון איש'),
        pickText(data, 'טלפון אשה'),
        pick(data, 'מעמד')
      ]);
      wrote = true;
    } catch (err) {
      sheetError = String(err);
    }

    var mailed = notify(data, sheetError);

    /* רק אם שניהם נפלו נכשל הרישום — ואז הטופס באתר נופל לוואטסאפ */
    if (!wrote && !mailed) {
      return json({ ok: false, error: 'sheet and mail both failed: ' + sheetError });
    }
    return json({ ok: true, sheet: wrote, mail: mailed });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* פתיחת הכתובת בדפדפן מחזירה מצב — כך בודקים שהחיבור חי */
function doGet() {
  var reach = 'unknown';
  try {
    getSheet();
    reach = 'ok';
  } catch (err) {
    reach = 'unreachable';
  }
  return json({ ok: true, msg: 'Expo HaSofrim alfon endpoint is live', sheet: reach });
}

function pick(data, key) {
  var v = data[key];
  return (v === undefined || v === null) ? '' : String(v);
}

/* כמו pick, אבל מוסיף גרש מוביל כדי שגוגל תשמור את הערך כטקסט */
function pickText(data, key) {
  var v = pick(data, key);
  return v === '' ? '' : "'" + v;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS);
    sh.getRange(1, 1, 1, COLUMNS.length)
      .setFontWeight('bold')
      .setBackground('#f2efe8');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);

    var width = TEXT_COLUMNS_TO - TEXT_COLUMNS_FROM + 1;
    sh.getRange(1, TEXT_COLUMNS_FROM, sh.getMaxRows(), width).setNumberFormat('@');
    sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.setColumnWidths(1, 6, 118);
    sh.setColumnWidths(7, 4, 65);
  }
  return sh;
}

/* מחזיר true אם המייל יצא. הוא רשת הביטחון של הרישום, לא קישוט. */
function notify(data, sheetError) {
  if (!NOTIFY_EMAIL) return false;
  try {
    var name = (pick(data, 'שם האיש') || pick(data, 'שם האשה')) + ' ' + pick(data, 'שם משפחה');
    var lines = [];

    if (sheetError) {
      lines.push('⚠️ הרישום הזה לא נכנס לגיליון — שמרו את המייל הזה.');
      lines.push('הסיבה: ' + sheetError);
      lines.push('');
    }

    lines = lines.concat([
      pick(data, 'סוג') + ' — אלפון תושבים, מתחם הסופרים',
      '',
      'שם משפחה: ' + pick(data, 'שם משפחה'),
      'שם האיש: ' + (pick(data, 'שם האיש') || '—'),
      'שם האשה: ' + (pick(data, 'שם האשה') || '—'),
      'כתובת: ' + pick(data, 'רחוב') + ', בניין ' + pick(data, 'בניין') +
        (pick(data, 'כניסה') ? ', כניסה ' + pick(data, 'כניסה') : '') +
        ', קומה ' + pick(data, 'קומה') + ', דירה ' + pick(data, 'דירה'),
      'טלפון בבית: ' + (pick(data, 'טלפון בית') || '—'),
      'נייד הבעל: ' + (pick(data, 'טלפון איש') || '—'),
      'נייד האשה: ' + (pick(data, 'טלפון אשה') || '—'),
      'מעמד: ' + pick(data, 'מעמד')
    ]);

    MailApp.sendEmail(NOTIFY_EMAIL,
      (sheetError ? '⚠️ ' : '') + 'רישום לאלפון — ' + name,
      lines.join('\n'));
    return true;
  } catch (ignore) {
    return false;
  }
}
