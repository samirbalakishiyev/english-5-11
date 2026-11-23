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
  const DM = delMap();
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
      if (!tel || have.has(tel) || DM[tel]) return;
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
function usersRows(s) { usersPurge_(s); const m = ageMap(), ab = abMap(), tg = tagMap(); return usersRows0(s).map(u => Object.assign(u, m[u.tel] ? { yas: m[u.tel] } : {}, ab[u.tel] ? { ab: ab[u.tel] } : {}, tg[u.tel] ? { tg: tg[u.tel] } : {})); }
/* 🏷️ Müəllim etiketi (məs: CK = "Ceyhun müəllimin şagirdlərinə özəl…"): ayrıca "Etiketlər" vərəqi (Users vərəqinə toxunmur).
   Müəllim panelində CK düyməsi ilə qoyulur/götürülür; şagirdin tətbiqində yuxarıda yazı kimi görünür. */
function tagSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Etiketlər');
  if (!s) { s = ss.insertSheet('Etiketlər'); s.appendRow(['Telefon', 'Etiket', 'Tarix']); s.getRange(1, 1, 1, 3).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function tagMap() {
  const s = SpreadsheetApp.getActive().getSheetByName('Etiketlər'), m = {};
  if (!s || s.getLastRow() < 2) return m;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '' && r[1] !== '') m[String(r[0]).replace(/^'/, '')] = String(r[1]); });
  return m;
}
function tagSet(o, now) {
  const tel = String(o.tel || ''), v = String(o.v || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 6).toUpperCase();
  if (!tel) return { ok: false, error: 'tel' };
  const s = tagSheet(), n = s.getLastRow();
  if (n > 1) {
    const t = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < t.length; i++) if (String(t[i][0]).replace(/^'/, '') === tel) { s.getRange(i + 2, 2, 1, 2).setValues([[v, new Date(now)]]); s.getRange(i + 2, 3).setNumberFormat(DF); return { ok: true, tg: v }; }
  }
  if (v) { s.appendRow(["'" + tel, v, new Date(now)]); s.getRange(s.getLastRow(), 3).setNumberFormat(DF); }
  return { ok: true, tg: v };
}

/* ⏳ Avto-blok: 31.10.2026-dan sonra qoşulan MÜƏLLİMLƏR və 25 YAŞDAN BÖYÜKLƏR 5 gün pulsuz istifadə edir.
   💳 "Ödəniş edib" qeyd olunmayıbsa, 5 gündən sonra növbəti girişdə avtomatik bloklanır.
   Müəllim blokunu əl ilə açsa, həmin şəxs bir daha avtomatik bloklanmır ("Avtoblok" vərəqi: avto / istisna). */
const AB_FROM = 1793390400000, AB_DAYS = 5, AB_AGE = 25, AB_PAY_FROM = 1792436563700;
/* ⏳ 02.11.2026 20:40-dan (Bakı) sonra qoşulan HAMI (şagird və müəllim) 5 gün pulsuz istifadə edir, sonra avtomatik bloklanır.
   Müəllim panelində 💳 "Ödədi" basılanda avtomatik blok açılır və bir daha bloklanmır. */
const AB_ALL = 1793637600000;
/* 👑 Premium: əl ilə 👑 qoyulub, və ya 02.11.2026-dan sonra qoşulub, və ya 💳 ödəniş edib → Missions task-ları açıqdır */
const premOf = v => isTrue(v[9]) || (ms(v[4]) || 0) >= AB_ALL || (ms(v[8]) || 0) >= AB_PAY_FROM;
function abSheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Avtoblok');
  if (!s) { s = ss.insertSheet('Avtoblok'); s.appendRow(['Telefon', 'Status', 'Tarix']); s.getRange(1, 1, 1, 3).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function abMap() {
  const s = SpreadsheetApp.getActive().getSheetByName('Avtoblok'), m = {};
  if (!s || s.getLastRow() < 2) return m;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] !== '') m[String(r[0]).replace(/^'/, '')] = String(r[1]); });
  return m;
}
function abSet(tel, st, now) {
  tel = String(tel || ''); if (!tel) return;
  const s = abSheet(), n = s.getLastRow();
  if (n > 1) {
    const v = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < v.length; i++) if (String(v[i][0]).replace(/^'/, '') === tel) { s.getRange(i + 2, 2, 1, 2).setValues([[st, new Date(now)]]); s.getRange(i + 2, 3).setNumberFormat(DF); return; }
  }
  s.appendRow(["'" + tel, st, new Date(now)]); s.getRange(s.getLastRow(), 3).setNumberFormat(DF);
}
function abCheck(s, r, val, o, now) {
  const tel = String(o.tel || ''), q = ms(val[4]);
  if (!q || q < AB_FROM || isTrue(val[6]) || (ms(val[8]) || 0) >= AB_PAY_FROM) return false;
  if (now - q < AB_DAYS * 864e5) return false;
  if (q < AB_ALL && !/müəllim/i.test(String(o.rol || val[3] || ''))) { const y = +(o.yas || ageMap()[tel] || 0); if (!(y > AB_AGE)) return false; }
  if (abMap()[tel] === 'istisna') return false;
  s.getRange(r, 7).insertCheckboxes().setValue(true); abSet(tel, 'avto', now);
  return true;
}
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
/* 🗑 İstifadəçini silmək (müəllim paneli → 🗑 Sil). Silinən nömrələr BÜTÜN saytlar üçün ortaq cədvələ yazılır:
   "Silinənlər – bütün saytlar" (Google Diskinizdə). Bu siyahıdakı nömrə heç bir saytda qeydiyyatdan keçə bilməz,
   hər saytda onun bütün sətirləri silinir. Birinə yenidən icazə vermək üçün həmin cədvəldən onun sətrini silin. */
const DEL_ID = '1huCwvX18oEAaz9jPPGKvqjHgerei7Tqq70l70rLDlDA';
function icazeVer() { SpreadsheetApp.openById(DEL_ID).getName(); } /* 🔑 bir dəfə: bu funksiyanı seçib ▶ İşə sal → İcazə ver */
function delSheet() {
  let b = null; try { b = SpreadsheetApp.openById(DEL_ID); } catch (e) { b = SpreadsheetApp.getActive(); }
  let s = b.getSheetByName('Silinənlər');
  if (!s) { s = b.insertSheet('Silinənlər'); s.appendRow(['Telefon', 'Silinmə tarixi', 'Harada silinib']); s.getRange(1, 1, 1, 3).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function delMap() {
  const c = CacheService.getScriptCache(), k = c.get('dm');
  if (k) { try { return JSON.parse(k); } catch (e) {} }
  const m = {};
  try { const s = delSheet(); if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().forEach(r => { const t = String(r[0]).replace(/^'/, '').replace(/\D/g, ''); if (t) m[t] = 1; }); } catch (e) {}
  try { c.put('dm', JSON.stringify(m), 60); } catch (e) {}
  return m;
}
/* bu cədvəlin bütün vərəqlərindən həmin nömrənin sətirlərini silir */
function delLocal(tel) {
  let n = 0;
  SpreadsheetApp.getActive().getSheets().forEach(sh => {
    const nm = sh.getName();
    if (nm === 'Silinənlər' || nm === 'Paraqraflar' || /yedək/i.test(nm) || sh.getLastRow() < 2 || sh.getLastColumn() < 1) return;
    const hv = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim().toLowerCase());
    const c = hv.findIndex(h => /^(tel|telefon|phone)/.test(h));
    if (c < 0) return;
    const v = sh.getRange(2, c + 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      const t = String(v[i][0]).replace(/^'/, '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '');
      if (t === tel) { sh.deleteRow(i + 2); n++; }
    }
  });
  return n;
}
function userDel(o, now) {
  const tel = String(o.tel || '').replace(/\D/g, '').replace(/^994/, '').replace(/^0/, '');
  if (!tel) return { ok: false, error: 'tel' };
  const n = delLocal(tel);
  if (!delMap()[tel]) { const s = delSheet(); s.appendRow(["'" + tel, new Date(now), SpreadsheetApp.getActive().getName()]); s.getRange(s.getLastRow(), 2).setNumberFormat(DF); }
  CacheService.getScriptCache().remove('dm');
  return { ok: true, n };
}
/* müəllim paneli siyahısı: başqa saytda silinənləri bu cədvəldən də təmizləyir */
function usersPurge_(s) {
  const dm = delMap(); if (!Object.keys(dm).length || s.getLastRow() < 2) return;
  const v = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues(), seen = {};
  v.forEach(r => { const t = String(r[0]).replace(/^'/, ''); if (dm[t] && !seen[t]) { seen[t] = 1; delLocal(t); } });
}
function userPing(o, now) {
  { const tel = String(o.tel || ''); if (tel && delMap()[tel]) { delLocal(tel); return { ok: true, gone: true, blok: false }; } } /* 🗑 silinib: qeydiyyat yoxdur */
  if (o.yas) ageSet(o.tel, o.yas, now); /* 🎂 */
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) {
    s.appendRow(["'" + (o.tel || ''), o.ad || '', o.soyad || '', o.rol || '', new Date(+o.qeyd || now), new Date(now), false, o.sayt || '', '', false]);
    const lr = s.getLastRow();
    s.getRange(lr, 5, 1, 2).setNumberFormat(DF);
    s.getRange(lr, 7).insertCheckboxes();
    s.getRange(lr, 10).insertCheckboxes();
    return { ok: true, blok: false, tg: tagMap()[String(o.tel || '')] || '', prem: false };
  }
  const val = s.getRange(r, 1, 1, 10).getValues()[0];
  s.getRange(r, 2, 1, 3).setValues([[o.ad || val[1], o.soyad || val[2], o.rol || val[3]]]);
  s.getRange(r, 6).setValue(new Date(now)).setNumberFormat(DF);
  const ab = abCheck(s, r, val, o, now); /* ⏳ */
  return { ok: true, blok: ab || isTrue(val[6]), tg: tagMap()[String(o.tel || '')] || '', prem: premOf(val) };
}
function userSet(o, now) {
  const s = usersSheet();
  const r = userRow(s, o.tel);
  if (!r) return { ok: false, error: 'tel' };
  if (o.a === 'block') { s.getRange(r, 7).insertCheckboxes().setValue(!!o.v); if (!o.v) abSet(o.tel, 'istisna', now); return { ok: true }; }
  if (o.a === 'premium') { s.getRange(r, 10).insertCheckboxes().setValue(!!o.v); return { ok: true }; }
  s.getRange(r, 9).setValue(o.v ? new Date(now) : '').setNumberFormat(DF);
  /* 💳 ödəniş edildi → 5 günlük avtomatik blok açılır */
  let unb = false;
  if (o.v && isTrue(s.getRange(r, 7).getValue()) && abMap()[String(o.tel)] === 'avto') { s.getRange(r, 7).insertCheckboxes().setValue(false); abSet(o.tel, 'ödəyib', now); unb = true; }
  return { ok: true, t: o.v ? now : '', unb };
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
  if (p.a === 'plist') return out({ tels: usersRows(usersSheet()).filter(u => u.prem || (+u.qeyd || 0) >= AB_ALL || (+u.elaqe || 0) >= AB_PAY_FROM).map(u => u.tel) });
  return out({ error: 'a' });
}

/* Yaz: ping (qeydiyyat + son giriş + blok yoxlaması), pq (premium yoxlaması), block, elaqe (ödəniş), premium, cfg (müəllim) */
function doPost(e) {
  { let o0 = {}; try { o0 = JSON.parse(e.postData.contents || '{}'); } catch (x) {} if (o0.a === 'pget' || o0.a === 'pset') return out(psHandle(o0)); } /* ☁️ irəliləyiş: öz kilidi var */
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
      return out({ ok: true, prem: r ? premOf(s.getRange(r, 1, 1, 10).getValues()[0]) : false });
    }
    if (o.key !== KEY) return out({ ok: false, error: 'key' });
    if (o.a === 'tag') return out(tagSet(o, now)); /* 🏷️ müəllim etiketi (CK) */
    if (o.a === 'block' || o.a === 'elaqe' || o.a === 'premium') return out(userSet(o, now));
    if (o.a === 'ren') return out(fixSet(o));
    if (o.a === 'fbset') return out(fbSet(o, now));
    if (o.a === 'udel') return out(userDel(o, now)); /* 🗑 */
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

/* ☁️ İRƏLİLƏYİŞ — şagirdin işlədiyi materiallar telefon nömrəsinə bağlı "İrəliləyiş" vərəqində saxlanır.
   Şagird başqa telefonda / kompüterdə eyni nömrə ilə açanda hamısı geri gəlir. Bu vərəqi əl ilə dəyişməyin.
   Hər açar ayrıca saxlanır: {v: dəyər, t: serverə yazılma vaxtı}. Bir cihaz köhnə versiya üzərində yazırsa,
   heç nə itməsin deyə iki versiya birləşdirilir. */
const PS_NAME = 'İrəliləyiş', PS_CH = 45000, PS_MAXK = 300, PS_MAXV = 1500000;
function psSheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(PS_NAME);
  if (!s) {
    s = ss.insertSheet(PS_NAME);
    s.getRange(1, 1, 1, 4).setValues([['Nömrə', 'Yenilənib', 'Ölçü (KB)', 'Məlumat (avtomatik — dəyişməyin)']]).setFontWeight('bold');
    s.setFrozenRows(1); s.setColumnWidth(1, 110); s.setColumnWidth(2, 140);
  }
  return s;
}
function psRow_(s, tel) {
  if (s.getLastRow() < 2) return 0;
  const c = s.getRange(2, 1, s.getLastRow() - 1, 1).createTextFinder(tel).matchEntireCell(true).findNext();
  return c ? c.getRow() : 0;
}
function psLoad_(s, r) {
  if (!r) return {};
  const w = s.getLastColumn(); if (w < 4) return {};
  const b64 = s.getRange(r, 4, 1, w - 3).getValues()[0].map(x => String(x || '').replace(/^~/, '')).join('');
  if (!b64) return {};
  try { return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString('UTF-8')) || {}; }
  catch (e) { console.error('ps load: ' + e); return null; }
}
function psSave_(s, r, tel, map) {
  const json = JSON.stringify(map);
  const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json', 'p.json')).getBytes());
  const ch = []; for (let i = 0; i < b64.length; i += PS_CH) ch.push('~' + b64.slice(i, i + PS_CH));
  const row = ["'" + tel, new Date(), Math.round(json.length / 1024)].concat(ch);
  const put = r => {
    const old = Math.max(s.getLastColumn() - 3, 0);
    if (old > ch.length) s.getRange(r, 4 + ch.length, 1, old - ch.length).clearContent();
    s.getRange(r, 1, 1, row.length).setValues([row]);
    s.getRange(r, 2).setNumberFormat('dd.MM.yyyy HH:mm');
  };
  if (r) return put(r);
  const lock = LockService.getScriptLock(); lock.waitLock(15000); /* yeni sətir: iki şagird eyni sətrə yazmasın */
  try { put(psRow_(s, tel) || s.getLastRow() + 1); SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
}
const psIsO_ = x => x && typeof x === 'object' && !Array.isArray(x);
function psDeep_(a, b) { /* a üstündür; heç nə itmir: siyahılar birləşir, rəqəmlərdən böyüyü, true qalır */
  if (a === undefined || a === null) return b;
  if (b === undefined || b === null) return a;
  if (Array.isArray(a) && Array.isArray(b)) { const seen = new Set(a.map(x => JSON.stringify(x))), o = a.slice(); b.forEach(x => { const k = JSON.stringify(x); if (!seen.has(k)) { seen.add(k); o.push(x); } }); return o; }
  if (psIsO_(a) && psIsO_(b)) { const o = Object.assign({}, b, a); for (const k in a) if (k in b) o[k] = psDeep_(a[k], b[k]); return o; }
  if (typeof a === 'number' && typeof b === 'number') return Math.max(a, b);
  if (typeof a === 'boolean' && typeof b === 'boolean') return a || b;
  return a;
}
function psMergeRaw_(a, b) { try { return JSON.stringify(psDeep_(JSON.parse(a), JSON.parse(b))); } catch (e) { return a; } }
function psHandle(o) {
  const tel = String(o.tel || '').replace(/\D/g, '');
  if (tel.length !== 9 || !psUserOK_(tel)) return { ok: false, error: 'user' };
  const s = psSheet_(), r = psRow_(s, tel), map = psLoad_(s, r);
  if (map === null) return { ok: false, error: 'read' };
  if (o.a === 'pget') return { ok: true, d: map };
  /* pset: d = {açar: {v: dəyər, b: cihazın son gördüyü vaxt}} */
  const d = psIsO_(o.d) ? o.d : {}, now = Date.now(), res = {};
  Object.keys(d).slice(0, PS_MAXK).forEach(k => {
    const x = d[k]; if (!x || typeof x.v !== 'string' || x.v.length > PS_MAXV) return;
    const cur = map[k];
    if (!cur || !cur.t || (+x.b || 0) >= cur.t) { map[k] = { v: x.v, t: now }; res[k] = now; }
    else if (cur.v === x.v) { res[k] = cur.t; }
    else { map[k] = { v: psMergeRaw_(x.v, cur.v), t: now }; res[k] = -1; } /* başqa cihaz yeniləyib: birləşdirildi */
  });
  if (Object.keys(res).length) psSave_(s, r, tel, map);
  return { ok: true, t: res };
}

function psUserOK_(tel) { if (delMap()[tel]) return false; return !!userRow(usersSheet(), tel); }
