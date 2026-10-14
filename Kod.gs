/* English 5–11 · Google Cədvəl skripti (Apps Script)
   Bu cədvəldə yalnız istifadəçilər (qeydiyyat, son giriş, blok, əlaqə) və müəllim düzəlişləri saxlanılır.
   Nəticələr cədvələ YAZILMIR — şagird onları WhatsApp ilə müəllimə göndərir.
   1) Aşağıdakı KEY-i öz gizli açarınızla dəyişin (məs: Samir-2026-xyz).
   2) Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone (Все). */
const KEY = 'BURAYA-OZ-ACARINIZI-YAZIN';

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
function findRow(s, val) {
  const v = s.getRange(1, 1, s.getLastRow(), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(val)) return i + 1;
  return 0;
}

/* ---------- Users vərəqi (oxunaqlı başlıqlar, tarixlər, blok qutucuqları) ---------- */
const UL = ['Telefon', 'Ad', 'Soyad', 'Rol', 'Qeydiyyat tarixi', 'Son giriş', 'Bloklanıb', 'Sayt', 'Əlaqə saxlanılıb'];
const DF = 'dd.mm.yyyy hh:mm';
const isTrue = x => x === true || String(x).toUpperCase() === 'TRUE';
const ms = x => x instanceof Date ? x.getTime() : (+x || '');
const asDate = x => (typeof x === 'number' && x > 1e11) ? new Date(x) : x;

function usersSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Users');
  if (!s) { s = ss.insertSheet('Users'); s.setFrozenRows(1); }
  const h = s.getRange(1, 1, 1, UL.length);
  if (String(h.getValues()[0].join('|')) !== UL.join('|')) h.setValues([UL]).setFontWeight('bold');
  return s;
}
/* köhnə sətirləri düzəldir: uzun rəqəmlər → tarix, TRUE/FALSE → qutucuq */
function usersFix(s) {
  const n = s.getLastRow() - 1;
  if (n < 1) return;
  const rg = s.getRange(2, 1, n, UL.length), v = rg.getValues();
  let ch = false;
  v.forEach(r => { [4, 5, 8].forEach(j => { const d = asDate(r[j]); if (d !== r[j]) { r[j] = d; ch = true; } }); const b = isTrue(r[6]); if (r[6] !== b) { r[6] = b; ch = true; } });
  if (ch) rg.setValues(v);
  s.getRange(2, 5, n, 2).setNumberFormat(DF);
  s.getRange(2, 9, n, 1).setNumberFormat(DF);
  s.getRange(2, 7, n, 1).insertCheckboxes();
}
function usersRows(s) {
  usersFix(s);
  const v = s.getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== '').map(r => ({ tel: String(r[0]), ad: r[1], soyad: r[2], rol: r[3], qeyd: ms(r[4]), son: ms(r[5]), blok: isTrue(r[6]), sayt: r[7], elaqe: ms(r[8]) }));
}
function userRow(s, tel) {
  const n = s.getLastRow();
  if (n < 2) return 0;
  const v = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(tel)) return i + 2;
  return 0;
}
function userPing(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) {
    s.appendRow(["'" + (o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', new Date(+o.qeyd || now), new Date(now), false, o.sayt || '', '']);
    const lr = s.getLastRow();
    s.getRange(lr, 5, 1, 2).setNumberFormat(DF);
    s.getRange(lr, 7).insertCheckboxes();
    return { ok: true, blok: false };
  }
  const val = s.getRange(r, 1, 1, 7).getValues()[0];
  s.getRange(r, 2, 1, 3).setValues([[o.ad || val[1], o.soyad || val[2], o.rol || val[3]]]);
  s.getRange(r, 6).setValue(new Date(now)).setNumberFormat(DF);
  return { ok: true, blok: isTrue(val[6]) };
}
function userSet(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) return { ok: false, error: 'tel' };
  if (o.a === 'block') { s.getRange(r, 7).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  s.getRange(r, 9).setValue(o.v ? new Date(now) : '').setNumberFormat(DF);
  return { ok: true, t: o.v ? now : '' };
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
  if (p.a === 'users') return out({ rows: usersRows(usersSheet()) });
  return out({ error: 'a' });
}

/* Yaz: ping (qeydiyyat + son giriş + blok yoxlaması), block, elaqe, cfg (müəllim) */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents || '{}');
    const now = Date.now();
    if (o.a === 'ping') return out(userPing(o, now));
    if (o.key !== KEY) return out({ ok: false, error: 'key' });
    if (o.a === 'block' || o.a === 'elaqe') return out(userSet(o, now));
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
