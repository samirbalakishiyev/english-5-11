/* English 5–11 · Google Cədvəl skripti (Apps Script)
   Bu cədvəldə yalnız istifadəçilər (qeydiyyat, son giriş, blok) və müəllim düzəlişləri saxlanılır.
   Nəticələr cədvələ YAZILMIR — şagird onları WhatsApp ilə müəllimə göndərir.
   1) Aşağıdakı KEY-i öz gizli açarınızla dəyişin (məs: Samir-2026-xyz).
   2) Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone (Все). */
const KEY = 'BURAYA-OZ-ACARINIZI-YAZIN';

const UH = ['tel', 'ad', 'soyad', 'rol', 'qeyd', 'son', 'blok', 'sayt'];
const CH = ['g', 'yenilenib', 'json'];
const PART = 45000; // bir xanaya sığan hissə

function sh(name, head) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); s.setFrozenRows(1); }
  return s;
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function rows(s) {
  const v = s.getDataRange().getValues();
  const h = v.shift();
  return v.map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); o.tel = String(o.tel || ''); o.blok = o.blok === true || String(o.blok).toUpperCase() === 'TRUE'; return o; });
}
function findRow(s, val) {
  const v = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(val)) return i + 1;
  return 0;
}

/* Müəllim düzəlişləri: hər sinif (g = gc5, gcv, si7b, e8, e9, e10, e11, aio) bir sətir */
function cfgAll() {
  const s = sh('Config', CH);
  const v = s.getDataRange().getValues();
  const st = {};
  for (let i = 1; i < v.length; i++) {
    const g = String(v[i][0] || ''); if (!g) continue;
    const js = v[i].slice(2).join('');
    try { st[g] = JSON.parse(js); } catch (e) {}
  }
  return st;
}
function cfgSet(g, state) {
  const s = sh('Config', CH);
  const js = JSON.stringify(state || {});
  const parts = [];
  for (let i = 0; i < js.length; i += PART) parts.push(js.slice(i, i + PART));
  if (!parts.length) parts.push('{}');
  let r = findRow(s, g);
  if (!r) { r = s.getLastRow() + 1; }
  const old = s.getLastColumn();
  if (old > 2) s.getRange(r, 3, 1, old - 2).clearContent();
  s.getRange(r, 1, 1, 2).setValues([[g, new Date()]]);
  s.getRange(r, 3, 1, parts.length).setNumberFormat('@').setValues([parts]);
}

/* Oxu: ?a=cfg (hamı üçün) · ?a=users&key=… (müəllim) */
function doGet(e) {
  const p = e.parameter || {};
  if (p.a === 'cfg') return out({ state: cfgAll() });
  if (p.key !== KEY) return out({ error: 'key' });
  if (p.a === 'users') return out({ rows: rows(sh('Users', UH)) });
  return out({ error: 'a' });
}

/* Yaz: ping (qeydiyyat + son giriş + blok yoxlaması), block (müəllim), cfg (müəllim) */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents || '{}');
    const now = Date.now();
    if (o.a === 'ping') {
      const s = sh('Users', UH);
      const r = findRow(s, o.tel);
      if (!r) {
        s.appendRow(["'" + (o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', o.qeyd || now, now, false, o.sayt || '']);
        return out({ ok: true, blok: false });
      }
      const val = s.getRange(r, 1, 1, 8).getValues()[0];
      s.getRange(r, 2, 1, 3).setValues([[o.ad || val[1], o.soyad || val[2], o.rol || val[3]]]);
      s.getRange(r, 6).setValue(now);
      return out({ ok: true, blok: val[6] === true || String(val[6]).toUpperCase() === 'TRUE' });
    }
    if (o.key !== KEY) return out({ ok: false, error: 'key' });
    if (o.a === 'block') {
      const s = sh('Users', UH);
      const r = findRow(s, o.tel);
      if (!r) return out({ ok: false, error: 'tel' });
      s.getRange(r, 7).setValue(!!o.v);
      return out({ ok: true });
    }
    if (o.a === 'cfg') {
      if (!o.g) return out({ ok: false, error: 'g' });
      cfgSet(String(o.g), o.state);
      return out({ ok: true });
    }
    return out({ ok: false, error: 'a' });
  } finally {
    lock.releaseLock();
  }
}
