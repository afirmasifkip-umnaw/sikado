/*************************************************************
 *  SIKEDO — Sistem Kehadiran Dosen
 *  Sistem Pencatatan Kehadiran Perkuliahan — PKA 26/27
 *
 *  Terintegrasi via ID:
 *    • Spreadsheet : 1tsyFBzBv5EvvzJLBZHntrjVm8SuT6lwFGHsUT1NbhPI
 *    • Folder SS   : 1uMpx-lT_i3tVVl5h3dxhJMtd6mzXsY5J
 *
 *  Sheet: Form Responses 1 | USERS | LOG_AKTIVITAS
 *         REF_KELAS | REF_MATAKULIAH (daftar pilihan, dikelola admin)
 *************************************************************/

/* ==================== KONFIGURASI ==================== */
var SHEET_ID        = '1tsyFBzBv5EvvzJLBZHntrjVm8SuT6lwFGHsUT1NbhPI';
var DRIVE_FOLDER_ID = '1uMpx-lT_i3tVVl5h3dxhJMtd6mzXsY5J';

var SHEET_DATA  = 'Form Responses 1';
var SHEET_USERS = 'USERS';
var SHEET_LOG   = 'LOG_AKTIVITAS';
var SHEET_KELAS = 'REF_KELAS';
var SHEET_MK    = 'REF_MATAKULIAH';

var SESSION_HOURS = 6;
var SETUP_FLAG    = 'sik_setup_ok';
var NIDN_LENGTH   = 10;

/* Nilai awal daftar Kelas (admin dapat menambah/mengurangi) */
var DEFAULT_KELAS = ['PAUD L','PAUD M','PAUD N','PAUD P','PAUD Q','PAUD R','PAUD S',
                     'PGSD K','PGSD L','PGSD M'];
/* Nilai awal tambahan daftar Mata Kuliah (digabung dengan yang terdapat di data) */
var DEFAULT_MK_EXTRA = ['PENDIDIKAN AGAMA','BAHASA INDONESIA','LAINNYA'];

var BULAN = ['Januari','Februari','Maret','April','Mei','Juni','Juli',
             'Agustus','September','Oktober','November','Desember'];

var DATA_HEADERS = ['Timestamp','NIDN','NAMA DOSEN','JABFUNG','KELAS','MATAKULIAH',
  'PERTEMUAN KE','MATERI','JAM MASUK','SCREANSHOOT AWAL','SCREANSHOOT TENGAH','SCREANSHOOT AKHIR'];

/* ==================== AKSES SPREADSHEET & DRIVE (VIA ID) ==================== */
var SS_CACHE = null;
function getSS_() {
  if (SS_CACHE) return SS_CACHE;
  try { SS_CACHE = SpreadsheetApp.openById(SHEET_ID); }
  catch(e) {
    throw new Error('Spreadsheet (ID: ' + SHEET_ID + ') tidak dapat dibuka. Pastikan ID benar dan akun ini punya akses editor.');
  }
  return SS_CACHE;
}

var TZ_CACHE = null;
function getTZ_() {
  if (TZ_CACHE) return TZ_CACHE;
  try { TZ_CACHE = getSS_().getSpreadsheetTimeZone(); }
  catch(e) { TZ_CACHE = Session.getScriptTimeZone() || 'Asia/Jakarta'; }
  return TZ_CACHE;
}

/* Cari sheet toleran: abaikan spasi berlebih & huruf besar/kecil */
function findSheet_(ss, name) {
  var target = String(name).trim().toLowerCase();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getName()).trim().toLowerCase() === target) return sheets[i];
  }
  return null;
}

function getSheet_(name) {
  var sh = findSheet_(getSS_(), name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan. Login ulang sekali lagi (setup otomatis akan berjalan) atau jalankan setupAwal().');
  return sh;
}

function getScreenshotFolder_() {
  try { return DriveApp.getFolderById(DRIVE_FOLDER_ID); }
  catch(e) {
    throw new Error('Folder screenshot (ID: ' + DRIVE_FOLDER_ID + ') tidak dapat diakses. Pastikan ID benar dan akun ini punya akses editor ke folder.');
  }
}

/* ==================== NORMALISASI NIDN ==================== */
function normNidn_(v) {
  var s = String(v == null ? '' : v).trim();
  if (/^\d+$/.test(s) && s.length > 0 && s.length < NIDN_LENGTH) {
    var pad = '';
    for (var i = 0; i < NIDN_LENGTH - s.length; i++) pad += '0';
    s = pad + s;
  }
  return s;
}

/* ==================== WEB APP ==================== */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('SIKEDO — Sistem Kehadiran Dosen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ==================== SHEET REFERENSI (KELAS & MATA KULIAH) ==================== */
/* Membuat sheet REF_KELAS & REF_MATAKULIAH beserta nilai awalnya jika belum ada. */
function ensureRefSheets_() {
  var ss = getSS_();

  var k = findSheet_(ss, SHEET_KELAS);
  if (!k) {
    k = ss.insertSheet(SHEET_KELAS);
    k.getRange(1,1,1,1).setValue('KELAS')
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    k.setFrozenRows(1);
    DEFAULT_KELAS.forEach(function(v){ k.appendRow([v]); });
  }

  var m = findSheet_(ss, SHEET_MK);
  if (!m) {
    m = ss.insertSheet(SHEET_MK);
    m.getRange(1,1,1,1).setValue('MATAKULIAH')
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    m.setFrozenRows(1);
    // seed: nilai unik dari data + daftar tambahan
    var seen = {};
    var seed = [];
    var sh = findSheet_(ss, SHEET_DATA);
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 6, sh.getLastRow() - 1, 1).getValues();   // kolom F = MATAKULIAH
      rows.forEach(function(r){
        var v = String(r[0] == null ? '' : r[0]).trim().toUpperCase();
        if (v && !seen[v]) { seen[v] = true; seed.push(v); }
      });
    }
    DEFAULT_MK_EXTRA.forEach(function(v){
      if (!seen[v]) { seen[v] = true; seed.push(v); }
    });
    seed.forEach(function(v){ m.appendRow([v]); });
  }
}

/* Daftar nilai referensi (array string, urut sesuai sheet) */
function getRefValues_(sheetName) {
  var sh = findSheet_(getSS_(), sheetName);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] == null ? '' : values[i][0]).trim();
    if (v) out.push(v);
  }
  return out;
}

/* Daftar referensi + nomor baris (untuk hapus) */
function getRefList_(sheetName) {
  var sh = findSheet_(getSS_(), sheetName);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var v = String(values[i][0] == null ? '' : values[i][0]).trim();
    if (v) out.push({ value: v, row: i + 2 });
  }
  return out;
}

/* ==================== API REFERENSI (ADMIN) ==================== */
function apiGetRefs(token) {
  try {
    requireRole_(token, 'admin');
    try { ensureRefSheets_(); } catch(e) {}
    return { success:true, data: { kelas: getRefList_(SHEET_KELAS), matakuliah: getRefList_(SHEET_MK) } };
  } catch(e) { return { success:false, message: e.message }; }
}

function apiAddRef(token, type, value) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var user = requireRole_(token, 'admin');
    type = String(type || '').toLowerCase();
    var sheetName = (type === 'kelas') ? SHEET_KELAS : (type === 'mk' ? SHEET_MK : '');
    if (!sheetName) throw new Error('Tipe referensi tidak dikenal.');

    value = String(value || '').trim().toUpperCase();
    if (!value) throw new Error('Nama wajib diisi.');
    if (value.length > 60) throw new Error('Nama maksimal 60 karakter.');

    try { ensureRefSheets_(); } catch(e) {}
    var sh = getSheet_(sheetName);
    var values = sh.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim().toUpperCase() === value) {
        throw new Error('"' + value + '" sudah ada di daftar.');
      }
    }
    sh.appendRow([value]);
    logActivity_(user, 'TAMBAH REF', (type === 'kelas' ? 'KELAS' : 'MATAKULIAH') + ': ' + value);
    return { success:true, message:'"' + value + '" berhasil ditambahkan.' };
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function apiDeleteRef(token, type, row) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var user = requireRole_(token, 'admin');
    type = String(type || '').toLowerCase();
    var sheetName = (type === 'kelas') ? SHEET_KELAS : (type === 'mk' ? SHEET_MK : '');
    if (!sheetName) throw new Error('Tipe referensi tidak dikenal.');

    var sh = getSheet_(sheetName);
    row = Number(row);
    if (isNaN(row) || row < 2 || row > sh.getLastRow()) throw new Error('Item tidak ditemukan.');
    var val = String(sh.getRange(row, 1).getValue()).trim();
    sh.deleteRow(row);
    logActivity_(user, 'HAPUS REF', (type === 'kelas' ? 'KELAS' : 'MATAKULIAH') + ': ' + val);
    return { success:true, message:'"' + val + '" dihapus dari daftar.' };
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/* ==================== PERBAIKAN NIDN ==================== */
function repairNidnUsers_() {
  var users = getSheet_(SHEET_USERS);
  var lastRow = users.getLastRow();
  if (lastRow < 2) return 0;
  var values = users.getRange(2, 1, lastRow - 1, 5).getValues();
  var fixed = 0;
  for (var i = 0; i < values.length; i++) {
    var raw = String(values[i][0] == null ? '' : values[i][0]).trim();
    if (!raw) continue;
    var pad = normNidn_(raw);
    if (pad !== raw) {
      values[i][0] = pad;
      if (String(values[i][2]) === hashPassword_(raw)) {
        values[i][2] = hashPassword_(pad);
      }
      fixed++;
    }
  }
  if (fixed > 0) users.getRange(2, 1, lastRow - 1, 5).setValues(values);
  return fixed;
}

function repairNidnData_() {
  var sh = getSheet_(SHEET_DATA);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var range = sh.getRange(2, 2, lastRow - 1, 1);
  var values = range.getValues();
  var fixed = 0;
  for (var i = 0; i < values.length; i++) {
    var raw = String(values[i][0] == null ? '' : values[i][0]).trim();
    if (!raw) continue;
    var pad = normNidn_(raw);
    if (pad !== raw) { values[i][0] = pad; fixed++; }
  }
  if (fixed > 0) range.setValues(values);
  return fixed;
}

function perbaikiNidn() {
  ensureSetup_();
  var f1 = repairNidnUsers_();
  var f2 = repairNidnData_();
  try { PropertiesService.getScriptProperties().setProperty('sik_nidn_repaired', String(Date.now())); } catch(e) {}
  var msg = '✅ Perbaikan NIDN selesai!\n' +
    '• Sheet USERS : ' + f1 + ' NIDN diperbaiki\n' +
    '• Sheet data  : ' + f2 + ' NIDN diperbaiki';
  Logger.log(msg);
  return msg;
}

/* ==================== AUTO-SETUP (SELF-HEALING) ==================== */
function ensureSetup_() {
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch(e) {}
  if (cache && cache.get(SETUP_FLAG)) return 0;

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) {}
  try {
    if (cache && cache.get(SETUP_FLAG)) return 0;

    var ss = getSS_();

    // 1) Sheet data
    var sh = findSheet_(ss, SHEET_DATA);
    if (!sh) {
      sh = ss.insertSheet(SHEET_DATA);
      sh.getRange(1,1,1,DATA_HEADERS.length).setValues([DATA_HEADERS])
        .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }

    // 2) Sheet USERS
    var users = findSheet_(ss, SHEET_USERS);
    if (!users) {
      users = ss.insertSheet(SHEET_USERS);
      users.getRange(1,1,1,5).setValues([['NIDN','NAMA','PASSWORD_HASH','ROLE','JABFUNG']])
        .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
      users.setFrozenRows(1);
    }

    // 3) Sheet LOG
    if (!findSheet_(ss, SHEET_LOG)) {
      var log = ss.insertSheet(SHEET_LOG);
      log.getRange(1,1,1,4).setValues([['WAKTU','USER','AKSI','DETAIL']])
        .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    }

    // 4) Sheet referensi Kelas & Mata Kuliah (BARU)
    ensureRefSheets_();

    // 5) Admin default
    var data = users.getDataRange().getValues();
    var hasAdmin = false;
    for (var i = 1; i < data.length; i++)
      if (String(data[i][3]).toLowerCase() === 'admin') hasAdmin = true;
    if (!hasAdmin) {
      users.appendRow(['admin','Administrator', hashPassword_('admin123'), 'admin', '-']);
      logActivity_({ nidn:'auto', nama:'Auto-Setup' }, 'AUTO-SETUP', 'Admin default dibuat');
    }

    // 6) Import dosen (NIDN dinormalisasi 10 digit)
    var existing = {};
    for (var j = 1; j < data.length; j++) existing[normNidn_(data[j][0])] = true;
    var n = 0;
    if (sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, DATA_HEADERS.length).getValues();
      for (var k = 0; k < rows.length; k++) {
        var nidn = normNidn_(rows[k][1]);
        if (nidn && !existing[nidn]) {
          users.appendRow([nidn, String(rows[k][2]), hashPassword_(nidn), 'dosen', String(rows[k][3])]);
          existing[nidn] = true;
          n++;
        }
      }
      if (n > 0) logActivity_({ nidn:'auto', nama:'Auto-Setup' }, 'AUTO-SETUP', n + ' dosen diimport otomatis');
    }

    // 7) Perbaiki NIDN yang kehilangan 0 di depan (sekali saja)
    var repaired = false;
    try { repaired = !!PropertiesService.getScriptProperties().getProperty('sik_nidn_repaired'); } catch(e) {}
    if (!repaired) {
      var fu = repairNidnUsers_();
      var fd = repairNidnData_();
      try { PropertiesService.getScriptProperties().setProperty('sik_nidn_repaired', String(Date.now())); } catch(e) {}
      if (fu + fd > 0) logActivity_({ nidn:'auto', nama:'Auto-Repair' }, 'PERBAIKI NIDN', 'USERS: ' + fu + ', DATA: ' + fd);
    }

    if (cache) { try { cache.put(SETUP_FLAG, '1', 600); } catch(e) {} }
    return n;
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/* ==================== SETUP AWAL (manual — OPSIONAL) ==================== */
function setupAwal() {
  var n = ensureSetup_();
  var ss = getSS_();
  var nk = getRefValues_(SHEET_KELAS).length;
  var nm = getRefValues_(SHEET_MK).length;
  var msg = '✅ Setup SIKEDO selesai!\n' +
    '• Spreadsheet : ' + ss.getName() + '\n' +
    '• Admin       : "admin" / "admin123" (SEGERA GANTI)\n' +
    '• Dosen       : ' + n + ' akun baru diimport (password awal = NIDN 10 digit)\n' +
    '• Ref Kelas   : ' + nk + ' item | Ref Mata Kuliah : ' + nm + ' item\n' +
    '• Folder SS   : https://drive.google.com/drive/folders/' + DRIVE_FOLDER_ID;
  Logger.log(msg);
  return msg;
}

/* ==================== FUNGSI CEK ==================== */
function tesKoneksi() {
  var ss = getSS_();
  var folder = getScreenshotFolder_();
  var dataSheet  = findSheet_(ss, SHEET_DATA);
  var usersSheet = findSheet_(ss, SHEET_USERS);
  var refK = findSheet_(ss, SHEET_KELAS);
  var refM = findSheet_(ss, SHEET_MK);
  var repaired = false;
  try { repaired = !!PropertiesService.getScriptProperties().getProperty('sik_nidn_repaired'); } catch(e) {}
  var msg =
    '✅ Spreadsheet OK : "' + ss.getName() + '"\n' +
    '   • Sheet data  : ' + (dataSheet ? 'DITEMUKAN (' + Math.max(0, dataSheet.getLastRow()-1) + ' baris)' : 'BELUM ADA') + '\n' +
    '   • Sheet USERS : ' + (usersSheet ? 'DITEMUKAN (' + Math.max(0, usersSheet.getLastRow()-1) + ' akun)' : 'BELUM ADA') + '\n' +
    '   • Ref Kelas   : ' + (refK ? (Math.max(0, refK.getLastRow()-1) + ' item') : 'BELUM ADA (dibuat otomatis)') + '\n' +
    '   • Ref Matkul  : ' + (refM ? (Math.max(0, refM.getLastRow()-1) + ' item') : 'BELUM ADA (dibuat otomatis)') + '\n' +
    '   • NIDN repair : ' + (repaired ? 'sudah' : 'otomatis saat login') + '\n' +
    '✅ Folder Drive OK : "' + folder.getName() + '"';
  Logger.log(msg);
  return msg;
}

/* ==================== AUTENTIKASI & SESI ==================== */
function hashPassword_(pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw) + '::sikehadiran::');
  return bytes.map(function(b){ return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function saveSession_(token, user) {
  var payload = JSON.stringify({ user: user, exp: Date.now() + SESSION_HOURS * 3600 * 1000 });
  try { CacheService.getScriptCache().put('tok_' + token, payload, 21600); } catch(e) {}
  try { PropertiesService.getScriptProperties().setProperty('tok_' + token, payload); } catch(e) {}
  if (Math.random() < 0.2) cleanupSessions_();
}

function checkToken_(token) {
  if (!token) return null;
  var key = 'tok_' + String(token);
  var raw = null;
  try { raw = CacheService.getScriptCache().get(key); } catch(e) {}
  if (!raw) {
    try { raw = PropertiesService.getScriptProperties().getProperty(key); } catch(e) {}
    if (raw) { try { CacheService.getScriptCache().put(key, raw, 21600); } catch(e) {} }
  }
  if (!raw) return null;
  var obj = null;
  try { obj = JSON.parse(raw); } catch(e) {}
  if (!obj || !obj.user || !obj.exp || obj.exp < Date.now()) {
    destroySession_(token);
    return null;
  }
  return obj.user;
}

function destroySession_(token) {
  var key = 'tok_' + String(token);
  try { CacheService.getScriptCache().remove(key); } catch(e) {}
  try { PropertiesService.getScriptProperties().deleteProperty(key); } catch(e) {}
}

function cleanupSessions_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    var now = Date.now();
    Object.keys(all).forEach(function(k) {
      if (k.indexOf('tok_') !== 0) return;
      var dead = true;
      try { var o = JSON.parse(all[k]); dead = !o || !o.exp || o.exp < now; } catch(e) {}
      if (dead) props.deleteProperty(k);
    });
  } catch(e) {}
}

function requireRole_(token, role) {
  var user = checkToken_(token);
  if (!user) throw new Error('Sesi berakhir. Silakan login ulang.');
  if (role && user.role !== role) throw new Error('Akses ditolak.');
  return user;
}

function login(username, password) {
  try {
    try { ensureSetup_(); }
    catch(setupErr) {
      return { success:false, message:'Setup otomatis gagal: ' + setupErr.message };
    }

    var uname = normNidn_(username);
    password = String(password || '');
    if (!uname || !password) return { success:false, message:'NIDN dan password wajib diisi.' };

    var rows = null;
    try {
      rows = getSheet_(SHEET_USERS).getDataRange().getValues();
    } catch(e) {
      try { CacheService.getScriptCache().remove(SETUP_FLAG); } catch(e2) {}
      ensureSetup_();
      rows = getSheet_(SHEET_USERS).getDataRange().getValues();
    }

    var hashPw    = hashPassword_(password);
    var hashPwPad = /^\d+$/.test(password) ? hashPassword_(normNidn_(password)) : null;

    for (var i = 1; i < rows.length; i++) {
      if (normNidn_(rows[i][0]) === uname) {
        var stored = String(rows[i][2]);
        if (stored !== hashPw && !(hashPwPad !== null && stored === hashPwPad)) {
          return { success:false, message:'Password salah. Password awal = NIDN Anda 10 digit lengkap (awali angka 0).' };
        }
        var user = {
          nidn: uname,
          nama: String(rows[i][1] || ''),
          role: String(rows[i][3] || 'dosen').toLowerCase(),
          jabfung: String(rows[i][4] || '')
        };
        var token = Utilities.getUuid();
        saveSession_(token, user);
        logActivity_(user, 'LOGIN', 'Login berhasil');
        return { success:true, token:token, user:user };
      }
    }
    return { success:false, message:'NIDN tidak terdaftar. Klik "Daftar sekarang" untuk registrasi, atau hubungi admin.' };
  } catch(e) {
    return { success:false, message: e.message };
  }
}

/* ==================== REGISTRASI MANDIRI ==================== */
function register(nidn, nama, jabfung, password) {
  try {
    try { ensureSetup_(); }
    catch(setupErr) {
      return { success:false, message:'Setup otomatis gagal: ' + setupErr.message };
    }
  } catch(e) {
    return { success:false, message: e.message };
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var nid = normNidn_(nidn);
    nama    = String(nama || '').trim();
    jabfung = String(jabfung || '').trim();
    password = String(password || '').trim();

    if (!nid) return { success:false, message:'NIDN wajib diisi.' };
    if (!/^\d+$/.test(nid)) return { success:false, message:'NIDN hanya boleh berisi angka.' };
    if (!nama) return { success:false, message:'Nama lengkap wajib diisi.' };
    if (password && password.length < 4) return { success:false, message:'Password minimal 4 karakter.' };

    var sh = getSheet_(SHEET_USERS);
    var rows = sh.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (normNidn_(rows[i][0]) === nid) {
        return { success:false, message:'NIDN ini sudah terdaftar. Silakan login langsung (password awal = NIDN Anda 10 digit).' };
      }
    }

    var pw = password || nid;
    sh.appendRow([nid, nama, hashPassword_(pw), 'dosen', jabfung]);
    logActivity_({ nidn:nid, nama:nama }, 'REGISTRASI', 'Akun dosen dibuat melalui registrasi mandiri');
    return { success:true, message:'Registrasi berhasil! Silakan login dengan NIDN Anda' + (password ? '.' : ' (password awal = NIDN 10 digit lengkap).') };
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function logout(token) {
  if (token) destroySession_(String(token));
  return { success:true };
}

/* ==================== HELPER ==================== */
function getProdiFromKelas_(kelas) {
  var s = String(kelas || '').trim().toUpperCase();
  return s ? s.split(/\s+/)[0] : '';
}

function normJam_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, getTZ_(), 'HH:mm');
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(s);
  if (m) {
    var h = parseInt(m[1], 10);
    if (m[4]) { var ap = m[4].toUpperCase(); if (ap==='PM' && h<12) h+=12; if (ap==='AM' && h===12) h=0; }
    return ('0'+h).slice(-2) + ':' + m[2];
  }
  return s;
}

function logActivity_(user, aksi, detail) {
  try {
    var sh = findSheet_(getSS_(), SHEET_LOG);
    if (sh) sh.appendRow([new Date(), user ? (user.nidn + ' — ' + user.nama) : '-', aksi, detail || '']);
  } catch(e) {}
}

/* ==================== BACA DATA ==================== */
function getAllData_(filters) {
  filters = filters || {};
  var sh = getSheet_(SHEET_DATA);
  var lastRow = sh.getLastRow();
  var out = [];
  if (lastRow < 2) return out;
  var values = sh.getRange(2, 1, lastRow - 1, DATA_HEADERS.length).getValues();

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var ts = null;
    if (r[0] instanceof Date) ts = r[0];
    else if (r[0] && String(r[0]).trim() && !isNaN(new Date(r[0]).getTime())) ts = new Date(r[0]);

    var kelas = String(r[4] == null ? '' : r[4]).trim();
    var rec = {
      id: i + 2,
      timestamp: ts ? Utilities.formatDate(ts, getTZ_(), 'yyyy-MM-dd HH:mm') : String(r[0] || ''),
      tgl: ts ? Utilities.formatDate(ts, getTZ_(), 'yyyy-MM-dd') : '',
      nidn: normNidn_(r[1]),
      nama: String(r[2] == null ? '' : r[2]),
      jabfung: String(r[3] == null ? '' : r[3]),
      kelas: kelas,
      prodi: getProdiFromKelas_(kelas),
      matakuliah: String(r[5] == null ? '' : r[5]),
      pertemuan: (r[6] === '' || r[6] === null) ? '' : Number(r[6]),
      materi: String(r[7] == null ? '' : r[7]),
      jamMasuk: normJam_(r[8]),
      ssAwal: String(r[9]  == null ? '' : r[9]),
      ssTengah: String(r[10] == null ? '' : r[10]),
      ssAkhir: String(r[11] == null ? '' : r[11]),
      _ts: ts ? ts.getTime() : 0
    };
    if (matchFilters_(rec, filters)) out.push(rec);
  }
  return out;
}

function matchFilters_(rec, f) {
  if (f.nidn  && f.nidn  !== 'all' && rec.nidn  !== normNidn_(f.nidn))  return false;
  if (f.prodi && f.prodi !== 'all' && rec.prodi !== String(f.prodi).trim()) return false;
  if (f.bulan && f.bulan !== 'all') {
    if (!rec._ts || new Date(rec._ts).getMonth() + 1 !== Number(f.bulan)) return false;
  }
  if (f.tahun && f.tahun !== 'all') {
    if (!rec._ts || String(new Date(rec._ts).getFullYear()) !== String(f.tahun)) return false;
  }
  if (f.search) {
    var q = String(f.search).toLowerCase();
    var hay = (rec.nama + ' ' + rec.nidn + ' ' + rec.matakuliah + ' ' + rec.kelas + ' ' + rec.materi).toLowerCase();
    if (hay.indexOf(q) === -1) return false;
  }
  return true;
}

/* ==================== API FRONTEND ==================== */
function apiGetStats(token) {
  try {
    var user = requireRole_(token);
    var rows = getAllData_(user.role === 'admin' ? {} : { nidn: user.nidn });
    var now = new Date();
    var stats = { total: rows.length, bulanIni: 0, prodi: {}, dosen: {},
                  perBulan: [0,0,0,0,0,0,0,0,0,0,0,0], tahunList: [], recent: [] };
    var years = {};
    rows.forEach(function(r){
      if (r._ts) {
        var d = new Date(r._ts);
        if (d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()) stats.bulanIni++;
        stats.perBulan[d.getMonth()]++;
        years[d.getFullYear()] = 1;
      }
      if (r.prodi) stats.prodi[r.prodi] = (stats.prodi[r.prodi] || 0) + 1;
      stats.dosen[r.nama] = (stats.dosen[r.nama] || 0) + 1;
    });
    stats.tahunList = Object.keys(years).sort().reverse();
    if (user.role === 'admin') {
      var n = 0;
      try {
        var us = getSheet_(SHEET_USERS).getDataRange().getValues();
        for (var i = 1; i < us.length; i++) if (String(us[i][3]).toLowerCase() === 'dosen') n++;
      } catch(e) {}
      stats.totalDosen = n;
    }
    rows.sort(function(a,b){ return b._ts - a._ts; });
    stats.recent = rows.slice(0, 10).map(function(r){
      var c = {}; for (var k in r) if (k !== '_ts') c[k] = r[k]; return c;
    });
    return { success:true, data: stats, user: user };
  } catch(e) { return { success:false, message: e.message }; }
}

function apiGetFilters(token) {
  try {
    var user = requireRole_(token);
    try { ensureRefSheets_(); } catch(e) {}
    var rows = getAllData_(user.role === 'admin' ? {} : { nidn: user.nidn });
    var prodi = {}, tahun = {};
    rows.forEach(function(r){
      if (r.prodi) prodi[r.prodi] = 1;
      if (r._ts) tahun[new Date(r._ts).getFullYear()] = 1;
    });
    var dosenList = [];
    if (user.role === 'admin') {
      var us = getSheet_(SHEET_USERS).getDataRange().getValues();
      for (var i = 1; i < us.length; i++)
        if (String(us[i][3]).toLowerCase() === 'dosen')
          dosenList.push({ nidn:normNidn_(us[i][0]), nama:String(us[i][1]), jabfung:String(us[i][4]||'') });
    } else {
      dosenList.push({ nidn:user.nidn, nama:user.nama, jabfung:user.jabfung || '' });
    }
    return { success:true, data: {
      prodiList: Object.keys(prodi).sort(),
      kelasList: getRefValues_(SHEET_KELAS),     // ← dari sheet REF_KELAS
      mkList: getRefValues_(SHEET_MK),           // ← dari sheet REF_MATAKULIAH
      dosenList: dosenList,
      tahunList: Object.keys(tahun).sort().reverse()
    }};
  } catch(e) { return { success:false, message: e.message }; }
}

function apiGetData(token, filters) {
  try {
    var user = requireRole_(token);
    filters = filters || {};
    if (user.role !== 'admin') filters.nidn = user.nidn;
    var rows = getAllData_(filters);
    rows.sort(function(a,b){ return b._ts - a._ts; });
    rows.forEach(function(r){ delete r._ts; });
    return { success:true, data: rows };
  } catch(e) { return { success:false, message: e.message }; }
}

/* ==================== SIMPAN / UBAH / HAPUS ==================== */
function apiSaveRecord(token, rec) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var user = requireRole_(token);
    var sh = getSheet_(SHEET_DATA);
    rec = rec || {};

    var nidn, nama, jabfung;
    if (user.role === 'admin') {
      nidn = normNidn_(rec.nidn);
      nama = String(rec.nama || '').trim();
      jabfung = String(rec.jabfung || '').trim();
      if (nidn && !nama) {
        var us = getSheet_(SHEET_USERS).getDataRange().getValues();
        for (var i = 1; i < us.length; i++) {
          if (normNidn_(us[i][0]) === nidn) {
            nama = String(us[i][1]);
            jabfung = jabfung || String(us[i][4] || '');
            break;
          }
        }
      }
    } else {
      nidn = user.nidn; nama = user.nama; jabfung = String(rec.jabfung || user.jabfung || '');
    }

    var kelas = String(rec.kelas || '').trim().toUpperCase();
    var matakuliah = String(rec.matakuliah || '').trim();
    var materi = String(rec.materi || '').trim();
    var jam = normJam_(rec.jamMasuk);
    var pertemuan = (rec.pertemuan === '' || rec.pertemuan == null) ? '' : Number(rec.pertemuan);

    if (!kelas) throw new Error('Kelas wajib dipilih.');
    if (!matakuliah) throw new Error('Mata kuliah wajib dipilih.');
    if (!jam) throw new Error('Jam masuk wajib diisi.');
    if (pertemuan !== '' && (isNaN(pertemuan) || pertemuan < 1)) throw new Error('Pertemuan ke- tidak valid.');

    if (rec.id) {
      /* ---------------- UPDATE ---------------- */
      var row = Number(rec.id);
      if (isNaN(row) || row < 2 || row > sh.getLastRow())
        throw new Error('Data tidak ditemukan (baris sudah bergeser). Muat ulang halaman.');
      var old = sh.getRange(row, 1, 1, DATA_HEADERS.length).getValues()[0];
      var oldNidn = normNidn_(old[1]);

      if (user.role !== 'admin' && oldNidn !== user.nidn)
        throw new Error('Anda hanya dapat mengedit data milik sendiri.');
      if (rec.ver && normNidn_(rec.ver) !== oldNidn)
        throw new Error('Data sudah berubah (dimodifikasi pengguna lain). Muat ulang data lalu edit ulang.');

      if (!nama) {
        if (nidn === oldNidn) nama = String(old[2]);
        else throw new Error('Dosen (NIDN) tidak ditemukan di daftar user.');
      }
      if (!jabfung) jabfung = String(old[3]);

      var ts = (old[0] instanceof Date) ? old[0] : (new Date(old[0]).getTime() ? new Date(old[0]) : new Date());
      if (rec.tanggal) {
        var jamStr = jam.length === 5 ? jam + ':00' : jam;
        var nts = new Date(rec.tanggal + 'T' + jamStr);
        if (!isNaN(nts.getTime())) ts = nts;
      }
      sh.getRange(row, 1, 1, DATA_HEADERS.length).setValues([[
        ts, nidn, nama, jabfung, kelas, matakuliah, pertemuan, materi, jam,
        String(rec.ssAwal||''), String(rec.ssTengah||''), String(rec.ssAkhir||'')
      ]]);
      logActivity_(user, 'EDIT DATA', 'Baris ' + row + ' — ' + nidn + ' — ' + matakuliah);
      return { success:true, message:'Data berhasil diperbarui.' };

    } else {
      /* ---------------- TAMBAH (cek duplikat) ---------------- */
      if (!nidn) throw new Error('Dosen wajib dipilih.');
      if (!nama) throw new Error('Dosen (NIDN) tidak ditemukan di daftar user.');

      if (pertemuan !== '') {
        var vals = sh.getDataRange().getValues();
        for (var d = 1; d < vals.length; d++) {
          if (normNidn_(vals[d][1]) === nidn &&
              String(vals[d][4]).trim().toUpperCase() === kelas &&
              String(vals[d][5]).trim().toUpperCase() === matakuliah.toUpperCase() &&
              Number(vals[d][6]) === pertemuan) {
            var tglDup = (vals[d][0] instanceof Date)
              ? Utilities.formatDate(vals[d][0], getTZ_(), 'dd/MM/yyyy') : String(vals[d][0]);
            throw new Error('Pertemuan ke-' + pertemuan + ' untuk ' + matakuliah + ' kelas ' + kelas +
              ' sudah tercatat (' + tglDup + '). Gunakan tombol Edit pada data tersebut.');
          }
        }
      }

      var tanggal = rec.tanggal || Utilities.formatDate(new Date(), getTZ_(), 'yyyy-MM-dd');
      var jamStr2 = jam.length === 5 ? jam + ':00' : (jam || '00:00:00');
      var tsNew = new Date(tanggal + 'T' + jamStr2);
      if (isNaN(tsNew.getTime())) tsNew = new Date();
      sh.appendRow([ tsNew, nidn, nama, jabfung, kelas, matakuliah, pertemuan, materi, jam,
        String(rec.ssAwal||''), String(rec.ssTengah||''), String(rec.ssAkhir||'') ]);
      logActivity_(user, 'TAMBAH DATA', nidn + ' — ' + matakuliah + ' — pertemuan ' + pertemuan);
      return { success:true, message:'Kehadiran berhasil disimpan.' };
    }
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function apiDeleteRecord(token, id, ver) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var user = requireRole_(token);
    var sh = getSheet_(SHEET_DATA);
    var row = Number(id);
    if (isNaN(row) || row < 2 || row > sh.getLastRow())
      throw new Error('Data tidak ditemukan (baris sudah bergeser). Muat ulang halaman.');
    var nidnRow = normNidn_(sh.getRange(row, 2).getValue());
    if (user.role !== 'admin' && nidnRow !== user.nidn) throw new Error('Akses ditolak.');
    if (ver && normNidn_(ver) !== nidnRow)
      throw new Error('Data sudah berubah (dimodifikasi pengguna lain). Muat ulang data lalu coba lagi.');

    var info = sh.getRange(row, 2, 1, 6).getDisplayValues()[0];
    sh.deleteRow(row);
    logActivity_(user, 'HAPUS DATA', 'Baris ' + row + ' — ' + info.join(' | '));
    return { success:true };
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

/* ==================== MANAJEMEN USER (ADMIN) ==================== */
function apiGetUsers(token) {
  try {
    requireRole_(token, 'admin');
    var rows = getSheet_(SHEET_USERS).getDataRange().getValues();
    var out = [];
    for (var i = 1; i < rows.length; i++) {
      out.push({ nidn:normNidn_(rows[i][0]), nama:String(rows[i][1]),
                 role:String(rows[i][3]).toLowerCase(), jabfung:String(rows[i][4]||''), row:i+1 });
    }
    return { success:true, data: out };
  } catch(e) { return { success:false, message: e.message }; }
}

function apiSaveUser(token, u) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var admin = requireRole_(token, 'admin');
    var sh = getSheet_(SHEET_USERS);
    u = u || {};
    var nidn = normNidn_(u.nidn);
    var nama = String(u.nama || '').trim();
    var role = (u.role === 'admin') ? 'admin' : 'dosen';
    if (!nidn) throw new Error('NIDN wajib diisi.');
    if (!nama) throw new Error('Nama wajib diisi.');

    if (u.row) {
      var row = Number(u.row);
      if (isNaN(row) || row < 2 || row > sh.getLastRow()) throw new Error('User tidak ditemukan.');
      var all = sh.getDataRange().getValues();
      for (var j = 1; j < all.length; j++)
        if ((j + 1) !== row && normNidn_(all[j][0]) === nidn) throw new Error('NIDN sudah dipakai user lain.');
      var vals = sh.getRange(row, 1, 1, 5).getValues()[0];
      vals[0] = nidn; vals[1] = nama; vals[3] = role;
      vals[4] = String(u.jabfung || vals[4] || '');
      if (u.password) vals[2] = hashPassword_(u.password);
      sh.getRange(row, 1, 1, 5).setValues([vals]);
      logActivity_(admin, 'EDIT USER', nidn + ' — ' + nama);
      return { success:true, message:'User berhasil diperbarui.' };
    } else {
      var all2 = sh.getDataRange().getValues();
      for (var k = 1; k < all2.length; k++)
        if (normNidn_(all2[k][0]) === nidn) throw new Error('NIDN sudah terdaftar.');
      sh.appendRow([nidn, nama, hashPassword_(u.password || nidn), role, String(u.jabfung || '')]);
      logActivity_(admin, 'TAMBAH USER', nidn + ' — ' + nama);
      return { success:true, message:'User baru ditambahkan (password = NIDN).' };
    }
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function apiDeleteUser(token, row) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {}
  try {
    var admin = requireRole_(token, 'admin');
    var sh = getSheet_(SHEET_USERS);
    row = Number(row);
    if (isNaN(row) || row < 2 || row > sh.getLastRow()) throw new Error('User tidak ditemukan.');
    var nidn = normNidn_(sh.getRange(row, 1).getValue());
    if (nidn === admin.nidn) throw new Error('Tidak dapat menghapus akun Anda sendiri.');
    sh.deleteRow(row);
    logActivity_(admin, 'HAPUS USER', nidn);
    return { success:true, message:'User dihapus.' };
  } catch(e) {
    return { success:false, message: e.message };
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

function apiChangePassword(token, oldPw, newPw) {
  try {
    var user = requireRole_(token);
    if (!newPw || String(newPw).length < 4) throw new Error('Password baru minimal 4 karakter.');
    var sh = getSheet_(SHEET_USERS);
    var rows = sh.getDataRange().getValues();

    var hashOld    = hashPassword_(oldPw);
    var hashOldPad = /^\d+$/.test(String(oldPw||'')) ? hashPassword_(normNidn_(oldPw)) : null;

    for (var i = 1; i < rows.length; i++) {
      if (normNidn_(rows[i][0]) === user.nidn) {
        var stored = String(rows[i][2]);
        if (stored !== hashOld && !(hashOldPad !== null && stored === hashOldPad))
          throw new Error('Password lama salah.');
        sh.getRange(i + 1, 3).setValue(hashPassword_(newPw));
        logActivity_(user, 'GANTI PASSWORD', '-');
        return { success:true };
      }
    }
    throw new Error('User tidak ditemukan.');
  } catch(e) { return { success:false, message: e.message }; }
}

/* ==================== UPLOAD SCREENSHOT ==================== */
function apiUploadScreenshot(token, file) {
  try {
    var user = requireRole_(token);
    file = file || {};
    if (!file.base64) throw new Error('File tidak valid.');
    var mime = String(file.mimeType || '');
    if (mime.indexOf('image/') !== 0) throw new Error('File harus berupa gambar.');
    var bytes = Utilities.base64Decode(file.base64);
    if (bytes.length > 10 * 1024 * 1024) throw new Error('Ukuran file maksimal 10MB.');

    var name = user.nidn + '_' + (String(file.name || 'screenshot.jpg'));
    var blob = Utilities.newBlob(bytes, mime, name);
    var f = getScreenshotFolder_().createFile(blob);
    try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    logActivity_(user, 'UPLOAD SS', name);
    return { success:true, url: 'https://drive.google.com/open?id=' + f.getId() };
  } catch(e) { return { success:false, message: e.message }; }
}

/* ==================== EXPORT EXCEL (ADMIN) ==================== */
function apiExportExcel(token, prodi, bulan, tahun) {
  try {
    var user = requireRole_(token, 'admin');
    prodi = String(prodi || 'all');
    bulan = String(bulan || 'all');
    tahun = String(tahun || 'all');

    var rows = getAllData_({ prodi:prodi, bulan:bulan, tahun:tahun });
    rows.sort(function(a,b){ return a._ts - b._ts; });

    var lblProdi = (prodi === 'all') ? 'SemuaProdi' : ('Prodi-' + prodi);
    var lblPeriode;
    if (bulan !== 'all' && tahun !== 'all') lblPeriode = BULAN[Number(bulan)-1] + '-' + tahun;
    else if (bulan !== 'all') lblPeriode = 'Bulan-' + BULAN[Number(bulan)-1];
    else if (tahun !== 'all') lblPeriode = 'Tahun-' + tahun;
    else lblPeriode = 'SemuaPeriode';

    var headers = ['NO','TANGGAL','NIDN','NAMA DOSEN','JABFUNG','KELAS','MATAKULIAH',
                   'PERTEMUAN KE','MATERI','JAM MASUK','SCREENSHOT AWAL','SCREENSHOT TENGAH','SCREENSHOT AKHIR'];
    var dataArr = rows.map(function(r, i){
      return [i+1, r.tgl || r.timestamp, r.nidn, r.nama, r.jabfung, r.kelas, r.matakuliah,
              r.pertemuan, r.materi, r.jamMasuk, r.ssAwal, r.ssTengah, r.ssAkhir];
    });

    var ss = SpreadsheetApp.create('TEMP_EXPORT_' + Date.now());
    var sh = ss.getSheets()[0];
    sh.setName('Daftar Hadir');
    var n = headers.length;

    sh.getRange(1,1,1,n).merge().setValue('DAFTAR HADIR PERKULIAHAN — PKA 26/27 (SIKEDO)')
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(2,1,1,n).merge().setValue(lblProdi + '  |  ' + lblPeriode + '  |  Total: ' + rows.length + ' pertemuan')
      .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(3,1,1,n).setValues([headers]).setFontWeight('bold')
      .setBackground('#1a73e8').setFontColor('#ffffff').setHorizontalAlignment('center');
    if (dataArr.length) sh.getRange(4,1,dataArr.length,n).setValues(dataArr).setVerticalAlignment('middle');

    sh.getRange(3,1,Math.max(dataArr.length,1)+1,n)
      .setBorder(true,true,true,true,true,true,'#94a3b8',SpreadsheetApp.BorderStyle.SOLID);
    var widths = [40,95,110,220,80,90,200,105,320,80,260,260,260];
    for (var c = 0; c < widths.length; c++) sh.setColumnWidth(c+1, widths[c]);
    sh.setFrozenRows(3);
    SpreadsheetApp.flush();

    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    var blob = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    var b64 = Utilities.base64Encode(blob.getBytes());
    try { DriveApp.getFileById(ss.getId()).setTrashed(true); } catch(e) {}

    var fname = 'DaftarHadir_' + lblProdi + '_' + lblPeriode + '_' +
                Utilities.formatDate(new Date(), getTZ_(), 'yyyyMMdd_HHmm') + '.xlsx';
    logActivity_(user, 'EXPORT EXCEL', lblProdi + ' | ' + lblPeriode + ' | ' + rows.length + ' baris');
    return { success:true, base64:b64, filename:fname, jumlah:rows.length };
  } catch(e) {
    return { success:false, message: e.message };
  }
}