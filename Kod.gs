/* English 5–11 · Google Cədvəl skripti (Apps Script)
   Bu cədvəldə yalnız istifadəçilər (qeydiyyat, son giriş, blok, ödəniş, 👑 premium) və müəllim düzəlişləri saxlanılır.
   Nəticələr cədvələ YAZILMIR — şagird onları WhatsApp ilə müəllimə göndərir.
   1) Aşağıdakı KEY-i öz gizli açarınızla dəyişin (məs: Samir-2026-xyz).
   2) Deploy → New deployment → Web app → Execute as: Me, Who has access: Anyone (Все). */
const KEY = 'BURAYA-OZ-ACARINIZI-YAZIN';
const ADMIN_PASS = 'BURAYA-PANEL-SIFRESI'; /* 🔐 Müəllim panelinin şifrəsi — özünüz yazın, heç kimə deməyin */

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
const UL = ['Telefon', 'Ad', 'Soyad', 'Rol', 'Qeydiyyat tarixi', 'Son giriş', 'Bloklanıb', 'Sayt', 'Əlaqə saxlanılıb', 'Premium'];
const DF = 'dd.mm.yyyy hh:mm';
const isTrue = x => x === true || /^(true|bəli|beli|yes|1|blok)$/i.test(String(x).trim());
const ms = x => x instanceof Date ? x.getTime() : (+x || '');
const asDate = x => (typeof x === 'number' && x > 1e11) ? new Date(x) : x;

/* başlıq adlarından sütun yerlərini tapır (köhnə və yeni adlar) */
function colMap(hv) {
  const h = hv.map(x => String(x).trim().toLowerCase().replace(/\(.*?\)/g, '').trim());
  const T = [
    k => /^(tel|telefon|nömrə|nomre|phone)/.test(k),
    k => k === 'ad' || k === 'name',
    k => k === 'soyad' || k === 'surname',
    k => k === 'rol' || k === 'role' || k.includes('müəllim'),
    k => k === 'qeyd' || k.startsWith('qeydiyyat') || k === 'at',
    k => k === 'son' || k.startsWith('son giri') || k === 'last',
    k => k === 'blok' || k.startsWith('blok') || k === 'block',
    k => k === 'sayt' || k === 'site',
    k => k === 'elaqe' || k.startsWith('əlaqə') || k.startsWith('elaqe'),
    k => k === 'prem' || k.startsWith('premium')];
  return T.map(f => h.findIndex(f));
}
/* başqa vərəqlərdə qalmış köhnə istifadəçiləri Users-ə köçürür (təkrarlamadan) */
function usersImport(s) {
  const ss = SpreadsheetApp.getActive();
  const have = new Set(s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const add = [];
  ss.getSheets().forEach(sh => {
    const nm = sh.getName();
    if (/^users/i.test(nm) || /config|results|rəy|düzəliş/i.test(nm) || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return;
    const v = sh.getDataRange().getValues(), idx = colMap(v[0]);
    if (idx[0] < 0 || idx[1] < 0) return;
    v.slice(1).forEach(r => {
      const tel = String(r[idx[0]] || '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '');
      if (!tel || have.has(tel)) return;
      have.add(tel);
      add.push(idx.map((i, j) => j === 0 ? "'" + tel : i < 0 ? (j === 6 || j === 9 ? false : '') : r[i]));
    });
  });
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, UL.length).setValues(add);
}
function usersSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Users');
  if (!s) { s = ss.insertSheet('Users'); s.setFrozenRows(1); }
  const w = Math.max(s.getLastColumn(), 1);
  const hv = s.getRange(1, 1, 1, w).getValues()[0].map(String);
  if (hv.slice(0, UL.length).join('|') !== UL.join('|')) {
    /* köhnə başlıqlar (tel, ad, soyad, rol, qeyd, son, blok, sayt, elaqe) → yeni düzülüş */
    const idx = colMap(hv);
    const n = s.getLastRow() - 1;
    if (n > 0) { /* 🛟 sütunları düzməzdən əvvəl HƏR DƏFƏ tarixli nüsxə */ const nm = 'Users yedək ' + Utilities.formatDate(new Date(), 'Asia/Baku', 'dd.MM HH:mm:ss'); if (!ss.getSheetByName(nm)) s.copyTo(ss).setName(nm); }
    if (n > 0 && idx[0] < 0) return s; /* tanınmayan format: heç nəyə toxunma */
    if (n > 0 && idx.some((i, j) => i !== j)) {
      const all = s.getRange(2, 1, n, w).getValues();
      const nv = all.map(r => idx.map((i, j) => i < 0 ? (j === 9 ? false : '') : r[i]));
      s.getRange(2, 1, n, Math.max(w, UL.length)).clearContent();
      s.getRange(2, 1, n, UL.length).setValues(nv);
    }
    s.getRange(1, 1, 1, Math.max(w, UL.length)).clearContent();
    s.getRange(1, 1, 1, UL.length).setValues([UL]).setFontWeight('bold');
  }
  return s;
}
/* köhnə sətirləri düzəldir: uzun rəqəmlər → tarix, TRUE/FALSE → qutucuq */
function usersFix(s) {
  const n = s.getLastRow() - 1;
  if (n < 1) return;
  const rg = s.getRange(2, 1, n, UL.length), v = rg.getValues();
  let ch = false;
  /* yalnız dəyişən xanaları yazır (bütün cədvəli yenidən yazmır — eyni anda edilən blok/ödəniş itməsin) */
  v.forEach((r, i) => {
    [4, 5, 8].forEach(j => { const d = asDate(r[j]); if (d !== r[j]) s.getRange(i + 2, j + 1).setValue(d); });
    [6, 9].forEach(j => { const b = isTrue(r[j]); if (r[j] !== b && r[j] !== '') s.getRange(i + 2, j + 1).setValue(b); });
  });
  s.getRange(2, 5, n, 2).setNumberFormat(DF);
  s.getRange(2, 9, n, 1).setNumberFormat(DF);
  s.getRange(2, 7, n, 1).insertCheckboxes();
  s.getRange(2, 10, n, 1).insertCheckboxes();
}
/* 🎂 Yaş: "Yaşlar" vərəqi (Telefon | Yaş | Tarix). Users vərəqinə toxunmur — yeni qeydiyyatda yazılan yaş buraya düşür */
function ageSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Yaşlar');
  if (!s) { s = ss.insertSheet('Yaşlar'); s.appendRow(['Telefon', 'Yaş', 'Tarix']); s.getRange(1, 1, 1, 3).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function ageSet(tel, y, now) {
  y = +y; tel = String(tel || '');
  if (!tel || !(y >= 4 && y <= 99)) return;
  const s = ageSheet(), n = s.getLastRow();
  if (n > 1) {
    const v = s.getRange(2, 1, n - 1, 2).getValues();
    for (let i = 0; i < v.length; i++) if (String(v[i][0]).replace(/^'/, '') === tel) { if (+v[i][1] !== y) s.getRange(i + 2, 2).setValue(y); return; }
  }
  s.appendRow(["'" + tel, y, new Date(now)]); s.getRange(s.getLastRow(), 3).setNumberFormat(DF);
}
function ageMap() {
  const s = SpreadsheetApp.getActive().getSheetByName('Yaşlar'), m = {};
  if (!s || s.getLastRow() < 2) return m;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '') m[String(r[0]).replace(/^'/, '')] = +r[1] || ''; });
  return m;
}
function usersRows(s) { const m = ageMap(); return usersRows0(s).map(u => (m[u.tel] ? Object.assign(u, { yas: m[u.tel] }) : u)); }
function usersRows0(s) {
  usersImport(s);
  usersFix(s);
  const v = s.getDataRange().getValues(); v.shift();
  const fx = fixMap();
  return v.filter(r => r[0] !== '').map(r => { const t = String(r[0]), f = fx[t] || ['', '']; return { tel: t, ad: r[1], soyad: r[2], rol: r[3], qeyd: ms(r[4]), son: ms(r[5]), blok: isTrue(r[6]), sayt: r[7], elaqe: ms(r[8]), prem: isTrue(r[9]), nm: f[0], tel2: f[1] }; });
}
function userRow(s, tel) {
  const n = s.getLastRow();
  if (n < 2) return 0;
  const v = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(tel)) return i + 2;
  return 0;
}
function userPing(o, now) {
  if (o.yas) ageSet(o.tel, o.yas, now); /* 🎂 */
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) {
    s.appendRow(["'" + (o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', new Date(+o.qeyd || now), new Date(now), false, o.sayt || '', '', false]);
    const lr = s.getLastRow();
    s.getRange(lr, 5, 1, 2).setNumberFormat(DF);
    s.getRange(lr, 7).insertCheckboxes();
    s.getRange(lr, 10).insertCheckboxes();
    return { ok: true, blok: false, prem: false };
  }
  const val = s.getRange(r, 1, 1, 10).getValues()[0];
  s.getRange(r, 2, 1, 3).setValues([[o.ad || val[1], o.soyad || val[2], o.rol || val[3]]]);
  s.getRange(r, 6).setValue(new Date(now)).setNumberFormat(DF);
  return { ok: true, blok: isTrue(val[6]), prem: isTrue(val[9]) };
}
function userSet(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) return { ok: false, error: 'tel' };
  if (o.a === 'block') { s.getRange(r, 7).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  if (o.a === 'premium') { s.getRange(r, 10).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  s.getRange(r, 9).setValue(o.v ? new Date(now) : '').setNumberFormat(DF);
  return { ok: true, t: o.v ? now : '' };
}

/* ✏️ Müəllimin düzəltdiyi ad və nömrə: ayrıca "Düzəlişlər" vərəqi (Users vərəqinə toxunmur) */
const XL = ['Telefon', 'Düzəldilmiş ad', 'Düzgün nömrə'];
function fixSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Düzəlişlər');
  if (!s) { s = ss.insertSheet('Düzəlişlər'); s.appendRow(XL); s.getRange(1, 1, 1, XL.length).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function fixMap() {
  const s = SpreadsheetApp.getActive().getSheetByName('Düzəlişlər');
  if (!s || s.getLastRow() < 2) return {};
  const o = {};
  s.getRange(2, 1, s.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0] !== '') o[String(r[0]).replace(/^'/, '')] = [String(r[1] || ''), String(r[2] || '')]; });
  return o;
}
function fixSet(o) {
  const tel = String(o.tel || ''), t = String(o.v || '').trim().replace(/\s+/g, ' ').slice(0, 80);
  const t2 = String(o.t2 || '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '').slice(0, 9);
  const s = fixSheet(), n = s.getLastRow();
  let r = 0;
  if (n > 1) { const v = s.getRange(2, 1, n - 1, 1).getValues(); for (let i = 0; i < v.length; i++) if (String(v[i][0]).replace(/^'/, '') === tel) { r = i + 2; break; } }
  if (!r) { r = n + 1; s.getRange(r, 1).setNumberFormat('@').setValue(tel); }
  s.getRange(r, 2, 1, 2).setNumberFormat('@').setValues([[t, t2]]);
  return { ok: true, nm: t, tel2: t2 };
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

/* ✍️ Rəy və təkliflər: "Rəylər" vərəqi. Status: Yeni / Oxundu / Cavablandı / Arxiv.
   "Şagird gizlədib" = şagird öz tərəfində silib (müəllim yenə görür). ID = telefon|vaxt. */
const FL = ['Tarix', 'Telefon', 'Ad', 'Soyad', 'Rol', 'Rəy', 'Status', 'Cavab', 'Cavab tarixi', 'Şagird gizlədib', 'ID'];
const FST = ['Yeni', 'Oxundu', 'Cavablandı', 'Arxiv'];
function fbSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Rəylər');
  if (!s) { s = ss.insertSheet('Rəylər'); s.setFrozenRows(1); s.setColumnWidth(6, 420); s.setColumnWidth(8, 360); }
  const w = Math.max(s.getLastColumn(), 1), hv = s.getRange(1, 1, 1, w).getValues()[0].map(String);
  if (hv.slice(0, FL.length).join('|') !== FL.join('|')) s.getRange(1, 1, 1, FL.length).setValues([FL]).setFontWeight('bold'); /* köhnə 6 sütun eynidir, yenilər sağa əlavə olunur */
  return s;
}
function fbAll() {
  const s = fbSheet();
  if (s.getLastRow() < 2) return { s, v: [] };
  const rg = s.getRange(2, 1, s.getLastRow() - 1, FL.length), v = rg.getValues();
  let ch = false;
  v.forEach(r => { if (!r[10] && r[1] !== '') { r[10] = String(r[1]).replace(/^'/, '') + '|' + ms(r[0]); ch = true; } if (!r[6] && r[1] !== '') { r[6] = 'Yeni'; ch = true; } });
  if (ch) { s.getRange(2, 11, v.length, 1).setNumberFormat('@'); rg.setValues(v); }
  return { s, v };
}
const fbObj = r => ({ id: String(r[10]), at: ms(r[0]), tel: String(r[1]), ad: r[2], soyad: r[3], rol: r[4], x: String(r[5]), st: String(r[6] || 'Yeni'), rep: String(r[7] || ''), rat: ms(r[8]), hid: isTrue(r[9]) });
function fbRows() { return fbAll().v.filter(r => r[1] !== '').map(fbObj); }
/* müəllim: status / cavab / silmə */
function fbSet(o, now) {
  const { s, v } = fbAll(), i = v.findIndex(r => String(r[10]) === String(o.id));
  if (i < 0) return { ok: false, error: 'id' };
  const row = i + 2;
  if (o.del) { s.deleteRow(row); return { ok: true, del: true }; }
  if (typeof o.rep === 'string') {
    const t = o.rep.trim().slice(0, 2000);
    s.getRange(row, 7, 1, 3).setValues([[t ? 'Cavablandı' : 'Oxundu', t, t ? new Date(now) : '']]);
    if (t) s.getRange(row, 9).setNumberFormat(DF);
    return { ok: true, st: t ? 'Cavablandı' : 'Oxundu', rat: t ? now : '' };
  }
  if (FST.includes(o.st)) { s.getRange(row, 7).setValue(o.st); return { ok: true, st: o.st }; }
  return { ok: false, error: 'a' };
}


/* Oxu: ?a=cfg (hamı üçün) · ?a=users&key=… · ?a=plist&key=… (müəllim) */
function doGet(e) {
  const lock = LockService.getScriptLock(); /* oxuma da növbə ilə */
  lock.waitLock(10000);
  try { return doGet_(e); } finally { lock.releaseLock(); }
}
function doGet_(e) {
  const p = e.parameter || {};
  if (p.a === 'cfg') return out({ state: cfgAll() });
  if (p.key !== KEY) return out({ error: 'key' });
  if (p.a === 'users') return out({ rows: usersRows(usersSheet()) });
  if (p.a === 'fbs') return out({ rows: fbRows() });
  if (p.a === 'plist') return out({ tels: usersRows(usersSheet()).filter(u => u.prem).map(u => u.tel) });
  return out({ error: 'a' });
}

/* Yaz: ping (qeydiyyat + son giriş + blok yoxlaması), pq (premium yoxlaması), block, elaqe (ödəniş), premium, cfg (müəllim) */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const o = JSON.parse(e.postData.contents || '{}');
    const now = Date.now();
    if (o.a === 'tlogin') { /* 🔐 müəllim girişi: 5 səhv cəhddən sonra 15 dəqiqə bağlanır */
      const c = CacheService.getScriptCache(), n = +(c.get('tl_fail') || 0);
      if (ADMIN_PASS === 'BURAYA-PANEL-SIFRESI' || !ADMIN_PASS) return out({ ok: false, error: 'nopass' });
      if (n >= 5) return out({ ok: false, error: 'wait' });
      if (String(o.p || '') !== ADMIN_PASS) { c.put('tl_fail', String(n + 1), 900); return out({ ok: false, error: 'pass', left: 4 - n }); }
      c.remove('tl_fail');
      return out({ ok: true, key: KEY });
    }
    if (o.a === 'ping') return out(userPing(o, now));
    if (o.a === 'fb') { /* şagirdin rəyi (açar lazım deyil) */
      const x = String(o.x || '').trim().slice(0, 2000);
      if (!x) return out({ ok: false, error: 'x' });
      const s = fbSheet(), at = +o.at || now, id = String(o.tel || '') + '|' + at;
      if (s.getLastRow() > 1 && s.getRange(2, 11, s.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext()) return out({ ok: true, id });
      s.appendRow([new Date(at), "'" + String(o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', x, 'Yeni', '', '', false, "'" + id]);
      s.getRange(s.getLastRow(), 1).setNumberFormat(DF);
      return out({ ok: true, id });
    }
    if (o.a === 'fbmine') { /* şagird: öz rəyləri və müəllimin cavabları */
      const tel = String(o.tel || '');
      if (!tel) return out({ ok: true, rows: [] });
      return out({ ok: true, rows: fbRows().filter(f => f.tel === tel && !f.hid).map(f => ({ id: f.id, at: f.at, x: f.x, rep: f.rep, rat: f.rat })) });
    }
    if (o.a === 'fbhide') { /* şagird öz tərəfində silir: müəllimdə qalır */
      const { s, v } = fbAll(), i = v.findIndex(r => String(r[10]) === String(o.id) && String(r[1]).replace(/^'/, '') === String(o.tel || ''));
      if (i < 0) return out({ ok: false, error: 'id' });
      s.getRange(i + 2, 10).insertCheckboxes().setValue(true);
      return out({ ok: true });
    }
    if (o.a === 'pq') { /* şagirdin premium yoxlaması (Missions) */
      const s = usersSheet(), r = userRow(s, o.tel);
      return out({ ok: true, prem: r ? isTrue(s.getRange(r, 10).getValue()) : false });
    }
    if (o.key !== KEY) return out({ ok: false, error: 'key' });
    if (o.a === 'block' || o.a === 'elaqe' || o.a === 'premium') return out(userSet(o, now));
    if (o.a === 'ren') return out(fixSet(o));
    if (o.a === 'fbset') return out(fbSet(o, now));
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
