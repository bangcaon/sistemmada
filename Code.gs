/**
 * SISTEM JADUAL PERGERAKAN PEGAWAI MADA
 * Backend: Google Apps Script (Web App) + Google Sheets sebagai pangkalan data.
 *
 * Helaian (tab) yang digunakan - dicipta automatik oleh setup():
 *   Pengguna    : akaun pegawai & penyelia
 *   Pergerakan  : rekod jadual pergerakan
 *   Sesi        : token log masuk (jangan edit)
 */

const SHEET_USERS = 'Pengguna';
const SHEET_MOV = 'Pergerakan';
const SHEET_SES = 'Sesi';
const DEFAULT_SUPERVISOR_PHONE = '0194750194'; // penyelia bagi pegawai yang daftar sendiri
const SESSION_DAYS = 7;
const TZ = 'Asia/Kuala_Lumpur';

const HEAD_USERS = ['id', 'phone', 'name', 'grade', 'empno', 'role', 'supervisor_id', 'password_hash', 'created_at'];
const HEAD_MOV = ['id', 'user_id', 'tarikh', 'masa', 'perkara', 'lokasi', 'status', 'decided_by', 'decided_at', 'created_at'];
const HEAD_SES = ['token', 'user_id', 'expires'];

/* =====================  PEMASANGAN (jalankan SEKALI)  ===================== */

function setup() {
  ensureSheet_(SHEET_USERS, HEAD_USERS, [2, 5, 9]);          // lajur teks: phone, empno, created_at
  ensureSheet_(SHEET_MOV, HEAD_MOV, [3, 4, 5, 6, 9, 10]);    // tarikh & masa kekal teks
  ensureSheet_(SHEET_SES, HEAD_SES, [1]);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid() + Utilities.getUuid());

  if (readTable_(SHEET_USERS).length === 0) {
    const nik = insertUser_({
      phone: '0125639159', name: 'Ir. NIK KUN BIN NIK MAN', grade: 'VU7', empno: '2383',
      role: 'Penyelia', supervisor_id: '', password: 'mada2383'
    });
    insertUser_({
      phone: '0194750194', name: 'MOHD FADZIL BIN MAT DESA', grade: 'J13', empno: '3270',
      role: 'Penyelia', supervisor_id: nik, password: 'mada3270'
    });
  }
  Logger.log('Setup siap. Helaian dan akaun penyelia telah disediakan.');
}

/** Tukar kata laluan seseorang (jalankan manual dari editor jika ada yang terlupa). */
function resetPassword() {
  const phone = '0194750194';      // <-- tukar no. telefon
  const newPassword = 'mada3270';  // <-- tukar kata laluan baharu
  const u = readTable_(SHEET_USERS).find(function (r) { return String(r.phone) === phone; });
  if (!u) throw new Error('Pengguna tidak dijumpai: ' + phone);
  const col = HEAD_USERS.indexOf('password_hash') + 1;
  sheet_(SHEET_USERS).getRange(u._row, col).setValue(hash_(newPassword));
  Logger.log('Kata laluan ' + phone + ' telah ditukar.');
}

/* =====================  PINTU MASUK WEB APP  ===================== */

function doGet() {
  return json_({ ok: true, message: 'API Sistem Jadual Pergerakan MADA aktif.' });
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = route_(body);
  } catch (err) {
    out = { ok: false, error: 'Ralat pelayan: ' + err.message };
  }
  return json_(out);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const WRITE_ACTIONS = ['login', 'register', 'logout', 'add', 'decide'];

function route_(b) {
  const action = String(b.action || '');
  const lock = LockService.getScriptLock();
  const needLock = WRITE_ACTIONS.indexOf(action) >= 0;
  if (needLock) lock.waitLock(20000);
  try {
    switch (action) {
      case 'login': return login_(b);
      case 'register': return register_(b);
      case 'me': { const me = auth_(b.token); return { ok: true, user: me ? userOut_(me) : null }; }
      case 'logout': { dropSession_(b.token); return { ok: true }; }
    }
    const me = auth_(b.token);
    if (!me) return { ok: false, code: 401, error: 'Sesi tamat. Sila log masuk semula.' };
    switch (action) {
      case 'data': return data_(me);
      case 'add': return add_(me, b);
      case 'decide': return decide_(me, b);
      default: return { ok: false, error: 'Tindakan tidak dikenali.' };
    }
  } finally {
    if (needLock) lock.releaseLock();
  }
}

/* =====================  TINDAKAN  ===================== */

function login_(b) {
  const phone = cleanPhone_(b.phone);
  const cache = CacheService.getScriptCache();
  const failKey = 'fail:' + phone;
  if (Number(cache.get(failKey) || 0) >= 5) {
    return { ok: false, error: 'Terlalu banyak percubaan gagal. Cuba lagi selepas 10 minit.' };
  }
  const u = readTable_(SHEET_USERS).find(function (r) { return String(r.phone) === phone; });
  if (!u || u.password_hash !== hash_(String(b.password || ''))) {
    cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 600);
    Utilities.sleep(400);
    return { ok: false, error: 'No. telefon atau kata laluan tidak tepat.' };
  }
  cache.remove(failKey);
  return { ok: true, user: userOut_(u), token: createSession_(u.id) };
}

function register_(b) {
  const phone = cleanPhone_(b.phone);
  const name = clip_(b.name, 120), grade = clip_(b.grade, 20), empno = clip_(b.empno, 20);
  const pw = String(b.password || '');
  if (!/^\d{9,12}$/.test(phone)) return { ok: false, error: 'No. telefon tidak sah.' };
  if (!name || !grade || !empno) return { ok: false, error: 'Sila lengkapkan semua maklumat.' };
  if (pw.length < 6) return { ok: false, error: 'Kata laluan mesti sekurang-kurangnya 6 aksara.' };
  const users = readTable_(SHEET_USERS);
  if (users.some(function (r) { return String(r.phone) === phone; })) {
    return { ok: false, error: 'No. telefon ini telah didaftarkan.' };
  }
  const sup = users.find(function (r) { return String(r.phone) === DEFAULT_SUPERVISOR_PHONE; });
  const id = insertUser_({
    phone: phone, name: name, grade: grade, empno: empno, role: 'Pegawai',
    supervisor_id: sup ? sup.id : '', password: pw
  });
  const u = readTable_(SHEET_USERS).find(function (r) { return Number(r.id) === id; });
  return { ok: true, user: userOut_(u), token: createSession_(id) };
}

function data_(me) {
  const users = readTable_(SHEET_USERS);
  const nameOf = {};
  users.forEach(function (u) { nameOf[Number(u.id)] = u.name; });
  const records = readTable_(SHEET_MOV).map(function (m) {
    return {
      id: Number(m.id), uid: Number(m.user_id), name: nameOf[Number(m.user_id)] || '-',
      date: fmtDate_(m.tarikh), time: fmtTime_(m.masa),
      purpose: String(m.perkara), location: String(m.lokasi), status: String(m.status)
    };
  });
  records.sort(function (a, b) { return (b.date + b.time + ('000000' + b.id).slice(-6)).localeCompare(a.date + a.time + ('000000' + a.id).slice(-6)); });
  const team = users.filter(function (u) { return u.supervisor_id !== '' && Number(u.supervisor_id) === Number(me.id); })
    .map(function (u) { return { id: Number(u.id), phone: String(u.phone), name: u.name, grade: String(u.grade), empno: String(u.empno) }; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
  return { ok: true, records: records.slice(0, 1000), team: team };
}

function add_(me, b) {
  const date = clip_(b.date, 10), time = clip_(b.time, 5);
  const purpose = clip_(b.purpose, 255), location = clip_(b.location, 255);
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  if (!d || d.getMonth() !== Number(m[2]) - 1) return { ok: false, error: 'Tarikh tidak sah.' };
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)) return { ok: false, error: 'Masa tidak sah.' };
  if (!purpose || !location) return { ok: false, error: 'Sila isi perkara dan lokasi.' };
  // Penyelia tertinggi (tiada penyelia atasan) diluluskan terus.
  const status = me.supervisor_id !== '' ? 'Menunggu' : 'Diluluskan';
  const id = nextId_(SHEET_MOV);
  sheet_(SHEET_MOV).appendRow([id, Number(me.id), date, time, purpose, location, status, '', '', now_()]);
  return { ok: true, status: status };
}

function decide_(me, b) {
  const result = String(b.result || '');
  if (result !== 'Diluluskan' && result !== 'Ditolak') return { ok: false, error: 'Keputusan tidak sah.' };
  const mov = readTable_(SHEET_MOV).find(function (r) { return Number(r.id) === Number(b.id); });
  const owner = mov && readTable_(SHEET_USERS).find(function (u) { return Number(u.id) === Number(mov.user_id); });
  if (!mov || mov.status !== 'Menunggu' || !owner || owner.supervisor_id === '' || Number(owner.supervisor_id) !== Number(me.id)) {
    return { ok: false, error: 'Anda tidak dibenarkan mengesahkan rekod ini.' };
  }
  const sh = sheet_(SHEET_MOV);
  sh.getRange(mov._row, HEAD_MOV.indexOf('status') + 1, 1, 3).setValues([[result, Number(me.id), now_()]]);
  return { ok: true };
}

/* =====================  SESI & KATA LALUAN  ===================== */

function createSession_(userId) {
  const sh = sheet_(SHEET_SES);
  // bersihkan sesi tamat
  const now = Date.now();
  const rows = readTable_(SHEET_SES);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (Number(rows[i].expires) < now) sh.deleteRow(rows[i]._row);
  }
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  sh.appendRow([token, Number(userId), now + SESSION_DAYS * 86400000]);
  return token;
}

function auth_(token) {
  token = String(token || '');
  if (token.length < 32) return null;
  const s = readTable_(SHEET_SES).find(function (r) { return String(r.token) === token; });
  if (!s || Number(s.expires) < Date.now()) return null;
  return readTable_(SHEET_USERS).find(function (u) { return Number(u.id) === Number(s.user_id); }) || null;
}

function dropSession_(token) {
  token = String(token || '');
  const s = readTable_(SHEET_SES).find(function (r) { return String(r.token) === token; });
  if (s) sheet_(SHEET_SES).deleteRow(s._row);
}

function hash_(password) {
  const salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/* =====================  PEMBANTU DATA  ===================== */

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Helaian "' + name + '" tiada. Jalankan fungsi setup() dahulu.');
  return sh;
}

function ensureSheet_(name, headers, textCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  textCols.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
}

function readTable_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow(), cols = sh.getLastColumn();
  if (last < 2) return [];
  const vals = sh.getRange(1, 1, last, cols).getValues();
  const head = vals[0];
  const out = [];
  vals.slice(1).forEach(function (row, i) {
    if (row[0] === '') return;                 // abaikan baris kosong
    const o = { _row: i + 2 };
    head.forEach(function (h, j) { o[h] = row[j]; });
    out.push(o);
  });
  return out;
}

function nextId_(name) {
  const rows = readTable_(name);
  let max = 0;
  rows.forEach(function (r) { max = Math.max(max, Number(r.id) || 0); });
  return max + 1;
}

function insertUser_(u) {
  const id = nextId_(SHEET_USERS);
  sheet_(SHEET_USERS).appendRow([id, u.phone, u.name, u.grade, u.empno, u.role, u.supervisor_id, hash_(u.password), now_()]);
  return id;
}

function userOut_(u) {
  let supName = null;
  if (u.supervisor_id !== '') {
    const s = readTable_(SHEET_USERS).find(function (r) { return Number(r.id) === Number(u.supervisor_id); });
    supName = s ? s.name : null;
  }
  return {
    id: Number(u.id), phone: String(u.phone), name: String(u.name), grade: String(u.grade),
    empno: String(u.empno), role: String(u.role),
    supervisor: u.supervisor_id !== '' ? Number(u.supervisor_id) : null, supervisor_name: supName
  };
}

function cleanPhone_(p) { return String(p || '').replace(/\D+/g, ''); }
function clip_(v, max) { return String(v == null ? '' : v).trim().slice(0, max); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function fmtDate_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v); }
function fmtTime_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'HH:mm') : String(v).slice(0, 5); }
