// AWSAA · ربط الموقع بجدول أماكن المنتجات
// يتحط في: جدول قوقل ← الإضافات ← Apps Script
// الأعمدة: A رقم · B المنتج · C الصورة · D الرف · E الدرج · F البوكس · G ملاحظة

const PIN = 'ضع_الرمز_هنا'; // الرمز السري للتعديل (ما ترفعهوش على GitHub)

function sheet_() {
  return SpreadsheetApp.getActive().getSheets()[0];
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// قراءة كل الأماكن
function doGet() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return json_({ ok: true, rows: [] });
  const v = sh.getRange(2, 1, n, 7).getDisplayValues();
  const rows = v
    .filter(r => r[0] && r[3])
    .map(r => ({ id: r[0], rack: r[3], drawer: r[4], box: r[5], note: r[6] }));
  return json_({ ok: true, rows: rows });
}

// حفظ مكان منتج واحد
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad' }); }
  if (String(d.pin) !== String(PIN)) return json_({ ok: false, error: 'pin' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const all = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getDisplayValues();
    const i = all.findIndex(r => Number(r[0]) === Number(d.id));
    if (i < 0) return json_({ ok: false, error: 'id' });
    // منع تكرار نفس المكان (الرف-الدرج-البوكس) لمنتجين
    const nk = v => String(v || '').trim().toUpperCase().replace(/\s+/g, '').replace(/^0+(?=\d)/, '');
    const r = nk(d.rack), dr = nk(d.drawer), bx = nk(d.box);
    if (r) {
      for (let k = 0; k < all.length; k++) {
        if (k === i) continue;
        const a = all[k];
        if (nk(a[3]) === r && nk(a[4]) === dr && nk(a[5]) === bx) {
          return json_({ ok: false, error: 'dup', id: a[0], name: a[1], code: [r, dr, bx].filter(x => x).join('-') });
        }
      }
    }
    const clean = s => String(s || '').slice(0, 200);
    const range = sh.getRange(i + 2, 4, 1, 4);
    range.setNumberFormat('@');
    range.setValues([[clean(d.rack).toUpperCase(), clean(d.drawer), clean(d.box), clean(d.note)]]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
