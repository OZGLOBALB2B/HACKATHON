/**
 * OZ האקתון — דירוג רעיונות
 * מדביקים את הקובץ הזה ב־Apps Script של הגיליון ומפרסמים כ־Web app.
 */

// הסיסמה לעמוד התוצאות. אפשר לשנות.
const RESULTS_PASSWORD = 'oz2026';

const SHEET = 'תשובות';
const IDS = ["5ruqbpts", "m1m20ou9", "9dh05ntf", "w9hlsg5x", "i94tz50c", "dx02mr3t", "pvti4f31", "0wn6hdm1", "18mt1u8q", "0i4o39m6", "0bu0l31f", "fib9d7q6", "srljsp2e", "6w1883op", "meheby9r", "ihnkj9dz", "w228fogn", "ppsz6hos", "z2zeccht", "4su3r9yn", "do4t1t6b", "o1qp462m"];
const TITLES = ["מערכת תמחור", "ניהול הצעות מחיר", "רישום שעות מדויק", "השוואת עלויות קבלני משנה", "Hub לתשתיות כספים", "דשבורד לקוח", "ממשק לקוח לרייק", "ניהול פרויקט אתר", "QA לסטודיו", "יצירה והנפשה של אייקונים", "פלטפורמת 3D לעיצוב משרדים", "מחקר מתחרים", "סושיאל ליסנינג", "ניהול קמפיינים בסושיאל", "סוכני tone of voice", "מיקוד מסרים", "תמלול פגישות ושיחות", "זיהוי לקוחות פוטנציאליים", "מערכת Outreach", "היכרות עם לקוחות חדשים", "ניהול כלים, רישיונות וסיסמאות", "GUSTAVOZ"];
const N5 = 4, N3 = 7;

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) sh = ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0 || sh.getRange(1, 2).getValue() !== 'שם') {
    sh.getRange(1, 1, 1, 3 + IDS.length).setValues([['זמן', 'שם', 'מחלקה'].concat(TITLES)]);
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);
    sh.getRange(1, 1, 1, 3 + IDS.length).setFontWeight('bold');
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function norm_(s) { return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase(); }

// שמירת דירוג. אותו שם שוב = עדכון של השורה הקיימת.
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const p = JSON.parse(e.postData.contents);
    const name = String(p.name || '').trim();
    const dept = String(p.dept || '').trim();
    const sc = p.scores || {};
    if (!name) return json_({ ok: false, error: 'חסר שם' });
    let c5 = 0, c3 = 0;
    const row = IDS.map(id => {
      const v = Number(sc[id]);
      if (v === 5) c5++; else if (v === 3) c3++;
      return [1, 3, 5].indexOf(v) >= 0 ? v : 1;
    });
    if (c5 !== N5 || c3 !== N3) return json_({ ok: false, error: 'צריך בדיוק ' + N5 + ' ציוני 5 ו־' + N3 + ' ציוני 3' });

    const sh = sheet_();
    const data = [new Date(), name, dept].concat(row);
    const last = sh.getLastRow();
    let at = -1;
    if (last > 1) {
      const names = sh.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < names.length; i++) if (norm_(names[i][0]) === norm_(name)) { at = i + 2; break; }
    }
    if (at > 0) sh.getRange(at, 1, 1, data.length).setValues([data]);
    else sh.appendRow(data);
    return json_({ ok: true, updated: at > 0 });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// קריאת התוצאות לעמוד הניהול (דורש סיסמה)
function doGet(e) {
  const q = (e && e.parameter) || {};
  if (q.action === 'ping') return json_({ ok: true });
  if (q.key !== RESULTS_PASSWORD) return json_({ ok: false, error: 'סיסמה שגויה' });
  const sh = sheet_();
  const last = sh.getLastRow();
  const out = [];
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, 3 + IDS.length).getValues();
    vals.forEach(r => {
      if (!r[1]) return;
      const s = {};
      IDS.forEach((id, i) => s[id] = Number(r[3 + i]) || 1);
      out.push({ t: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), name: r[1], dept: r[2], scores: s });
    });
  }
  return json_({ ok: true, responses: out, sheetUrl: SpreadsheetApp.getActive().getUrl() });
}
