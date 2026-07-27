// ============================================================
// STUDENT IKIGAI DASHBOARD — Google Apps Script
// File: Code.gs — backend | Versi 5.1 (hide dinamis semua sheet + soft-protection)
// Desain & logic dashboard di bawah SAMA PERSIS kayak versi asli —
// yang baru cuma bagian AKTIVASI + HIDE/PROTECT di paling atas.
// ============================================================
const SHEET_IKIGAI  = 'IKIGAI';
const SHEET_SWOT    = 'SWOT ANALYSIS';
const SHEET_DREAM   = 'DREAM PLAN';
const SHEET_COLLEGE = 'COLLEGE PLAN';
const SHEET_ACTION  = 'ACTION PLAN';
const SHEET_TRACKER_STATE = 'TRACKER_STATE';
const SHEET_WELCOME = 'WELCOME';

// ============================================================
// GERBANG AKTIVASI
// ============================================================
const RUMAH_A_URL = 'https://script.google.com/macros/s/AKfycbwSO7yK3eU-8eeOzKoVfx-EN4P0krysLfkSbotbJimBhHFhz0gQSZwSQN-iRcC-4IUoCA/exec';
const WELCOME_CODE_CELL = 'G23'; // cell token di tab WELCOME (merged G23:H24)
const PROP_ACTIVATED = 'ikigai_activated';
const PROP_ACTIVATED_EMAIL = 'ikigai_activated_email';
const PROP_ACTIVATED_CODE = 'ikigai_activated_code';
const WELCOME_BANNER_RANGE = 'A1:L2'; // area yg dipakai buat tampilan "error" pas nonaktif — ganti kalau ternyata dipake konten lain
const PROP_WELCOME_LOCKED_SNAPSHOT = 'ikigai_welcome_locked_snapshot';

function isActivated_() {
  return PropertiesService.getScriptProperties().getProperty(PROP_ACTIVATED) === 'true';
}

function getTokenInSheet_(userSpreadsheet) {
  try {
    if (!userSpreadsheet) return '';
    const sheet = userSpreadsheet.getSheetByName(SHEET_WELCOME);
    if (!sheet) return '';
    return String(sheet.getRange(WELCOME_CODE_CELL).getValue() || '').trim();
  } catch (e) {
    return '';
  }
}

/**
 * Panggilan mentah ke Rumah A buat validasi email+code. Dipakai bareng
 * oleh activateWithToken() (dari input manual) dan checkAndSyncActivation_()
 * (auto-recheck tiap buka dashboard/spreadsheet).
 */
function validateAgainstRumahA_(email, code) {
  Logger.log('[DEBUG] Memulai validateAgainstRumahA_');
  Logger.log('[DEBUG] Parameter - Email: ' + email + ', Code: ' + code);

  if (!email) {
    Logger.log('[DEBUG] Error: Email kosong.');
    return { valid: false, error: 'Email akun Google nggak terdeteksi. Pastikan kamu login pakai akun Google yang sama dengan waktu pembelian.' };
  }
  if (!code) {
    Logger.log('[DEBUG] Error: Code kosong.');
    return { valid: false, error: 'Token aktivasi belum diisi.' };
  }

  const secret = PropertiesService.getScriptProperties().getProperty('VALIDATE_SECRET');
  if (!secret) {
    Logger.log('[DEBUG] Error: VALIDATE_SECRET belum diset.');
    return { valid: false, error: 'Konfigurasi belum lengkap (VALIDATE_SECRET). Hubungi admin.' };
  }

  Logger.log('[DEBUG] Payload yang akan dikirim: ' + JSON.stringify({ action: 'validate', email: email, code: code, secret: '***' }));
  const payload = { action: 'validate', secret: secret, email: email, code: code };
  let response;
  try {
    Logger.log('[DEBUG] URL Tujuan: ' + RUMAH_A_URL);
    response = UrlFetchApp.fetch(RUMAH_A_URL, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    Logger.log('[DEBUG] HTTP Status Code: ' + response.getResponseCode());
    Logger.log('[DEBUG] Raw Response: ' + response.getContentText());
  } catch (err) {
    Logger.log('[DEBUG] Exception saat UrlFetchApp: ' + err.message);
    return { valid: false, error: 'Gagal menghubungi server lisensi: ' + err.message };
  }

  let result;
  try { 
    result = JSON.parse(response.getContentText()); 
    Logger.log('[DEBUG] Parsed JSON Response: ' + JSON.stringify(result));
  }
  catch (err) { 
    Logger.log('[DEBUG] Gagal parse JSON response.');
    return { valid: false, error: 'Respons server lisensi tidak valid.' }; 
  }

  if (result.error) {
    Logger.log('[DEBUG] Ditolak oleh server, error: ' + result.error);
    return { valid: false, error: result.error };
  }
  if (!result.valid) {
    Logger.log('[DEBUG] Ditolak oleh server, invalid reason: ' + (result.reason || 'Token tidak valid'));
    return { valid: false, error: result.reason || 'Token tidak valid untuk akun ini.' };
  }
  
  Logger.log('[DEBUG] Validasi Sukses!');
  return { valid: true, produk: result.produk || '' };
}

/** Set/reset status aktivasi lokal + langsung sembunyiin/tampilin sheet. */
function setActivatedState_(userSpreadsheet, active, email, code) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_ACTIVATED, active ? 'true' : 'false');
  if (active) {
    props.setProperty(PROP_ACTIVATED_EMAIL, email || '');
    props.setProperty(PROP_ACTIVATED_CODE, code || '');
  }
  enforceContentVisibility_(userSpreadsheet);
}

/**
 * Kasih tau Rumah A kalau lisensi ini baru aja nonaktif dari sisi Rumah B
 * (token di sheet WELCOME dihapus). Best-effort buat proses kunci lokal
 * (yang tetap jalan walau ini gagal), TAPI sekarang dicatat ke Logger biar
 * kalau nggak nyampe ke database Rumah A, kelihatan alasannya di
 * Executions log (Apps Script editor > ikon jam di kiri > Executions).
 */
function notifyDeactivation_(email, code) {
  try {
    if (!email || !code) {
      Logger.log('[notifyDeactivation_] dilewatin: email atau code kosong. email="%s" code="%s"', email, code);
      return;
    }
    const secret = PropertiesService.getScriptProperties().getProperty('VALIDATE_SECRET');
    if (!secret) {
      Logger.log('[notifyDeactivation_] dilewatin: VALIDATE_SECRET belum di-set di Script Properties.');
      return;
    }
    const resp = UrlFetchApp.fetch(RUMAH_A_URL, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ action: 'deactivate', secret: secret, email: email, code: code }),
      muteHttpExceptions: true
    });
    Logger.log('[notifyDeactivation_] email="%s" code="%s" -> HTTP %s | body: %s', email, code, resp.getResponseCode(), resp.getContentText());
  } catch (e) {
    Logger.log('[notifyDeactivation_] ERROR: %s', e.message);
  }
}

/**
 * SATU-SATUNYA sumber kebenaran status aktivasi — dipanggil ulang setiap
 * dashboard/spreadsheet dibuka (doGet, openDashboardModal, onOpen) dan
 * tiap ada edit di cell token (onIkigaiEdit_).
 */
function checkAndSyncActivation_(userSpreadsheet) {
  const code = getTokenInSheet_(userSpreadsheet);
  const wasActive = isActivated_();
  const props = PropertiesService.getScriptProperties();

  if (!code) {
    if (wasActive) {
      const lastEmail = props.getProperty(PROP_ACTIVATED_EMAIL) || Session.getActiveUser().getEmail();
      const lastCode = props.getProperty(PROP_ACTIVATED_CODE) || '';
      notifyDeactivation_(lastEmail, lastCode);
    }
    setActivatedState_(userSpreadsheet, false);
    return { activated: false, error: '' };
  }

  const lastCode = props.getProperty(PROP_ACTIVATED_CODE) || '';
  if (wasActive && lastCode && lastCode === code) {
    return { activated: true, error: '' };
  }

  const email = Session.getActiveUser().getEmail();
  const result = validateAgainstRumahA_(email, code);
  if (result.valid) {
    setActivatedState_(userSpreadsheet, true, email, code);
    return { activated: true, produk: result.produk, error: '' };
  }
  setActivatedState_(userSpreadsheet, false);
  return { activated: false, error: result.error };
}

function getActivationStatus(userSpreadsheet) {
  const status = checkAndSyncActivation_(userSpreadsheet);
  return { activated: status.activated, tokenInSheet: getTokenInSheet_(userSpreadsheet), error: status.error || '' };
}

/**
 * Dipanggil dari Sidebar.html (branch aktivasi) waktu user ngetik token di
 * modal & klik "Aktivasi Sekarang".
 */
function activateWithToken(userSpreadsheet, inputToken) {
  Logger.log('[DEBUG] Memulai activateWithToken dengan input: ' + inputToken);
  const email = Session.getActiveUser().getEmail();
  const code = String(inputToken || '').trim().toUpperCase() || getTokenInSheet_(userSpreadsheet);
  Logger.log('[DEBUG] Deteksi Email: ' + email + ', Code: ' + code);

  const result = validateAgainstRumahA_(email, code);
  Logger.log('[DEBUG] Hasil dari validateAgainstRumahA_: ' + JSON.stringify(result));
  
  if (!result.valid) {
    Logger.log('[DEBUG] Aktivasi dibatalkan karena validasi gagal.');
    return { success: false, error: result.error };
  }

  try {
    if (userSpreadsheet) {
      const sheet = userSpreadsheet.getSheetByName(SHEET_WELCOME);
      if (sheet) {
        sheet.getRange(WELCOME_CODE_CELL).setValue(code);
        Logger.log('[DEBUG] Berhasil menulis token ke sheet WELCOME.');
      }
    }
  } catch (e) { 
    Logger.log('[DEBUG] Exception saat menulis ke sheet WELCOME: ' + e.message);
  }

  setActivatedState_(userSpreadsheet, true, email, code);
  Logger.log('[DEBUG] setActivatedState_(true) berhasil dipanggil.');
  return { success: true, produk: result.produk };
}

/**
 * Trigger installable (BUKAN simple onEdit)
 */
function ensureEditTrigger_(userSpreadsheet) {
  try {
    if (!userSpreadsheet) return;
    const ss = userSpreadsheet;
    const already = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'onIkigaiEdit_' && t.getEventType() === ScriptApp.EventType.ON_EDIT;
    });
    if (!already) {
      ScriptApp.newTrigger('onIkigaiEdit_').forSpreadsheet(ss).onEdit().create();
    }
  } catch (e) { /* kalau gagal (mis. belum ada izin), diem — masih ke-cover pas onOpen/doGet */ }
}

function onIkigaiEdit_(e) {
  try {
    if (!e || !e.range) return;
    if (e.range.getSheet().getName() !== SHEET_WELCOME) return;
    if (e.range.getA1Notation() !== WELCOME_CODE_CELL) return;
    checkAndSyncActivation_(e.range.getSheet().getParent());
  } catch (err) { /* jangan sampai nge-block proses edit user */ }
}

/**
 * Sembunyiin/tampilin SEMUA sheet KECUALI WELCOME dan TRACKER_STATE
 */
function enforceContentVisibility_(userSpreadsheet) {
  try {
    if (!userSpreadsheet) return;
    const ss = userSpreadsheet;
    const visible = isActivated_();

    ss.getSheets().forEach(function (sh) {
      const name = sh.getName();
      if (name === SHEET_WELCOME) {
        if (visible) clearLockedAppearance_(sh); else applyLockedAppearance_(sh);
        return;
      }
      if (name === SHEET_TRACKER_STATE) return;

      if (visible) {
        if (sh.isSheetHidden()) sh.showSheet();
        removeWarningProtection_(sh);
      } else {
        if (!sh.isSheetHidden()) sh.hideSheet();
        addWarningProtection_(sh);
      }
    });

    const ts = ss.getSheetByName(SHEET_TRACKER_STATE);
    if (ts && !ts.isSheetHidden()) ts.hideSheet();
  } catch (e) { }
}

/**
 * Nimpa 2 baris teratas WELCOME jadi keliatan kayak error runtime beneran
 * (bukan sekadar dikosongin) — biar orang yang buka file mentahnya nggak
 * ngira ini masih berfungsi. Isi ASLI di area itu di-snapshot dulu ke
 * Script Properties SEBELUM ditimpa, jadi bisa dipulihin persis (bukan
 * cuma di-clear) begitu diaktivasi. Idempotent — kalau udah ke-apply
 * sebelumnya (ada snapshot tersimpan), nggak nyempotin ulang.
 */
function applyLockedAppearance_(sheet) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(PROP_WELCOME_LOCKED_SNAPSHOT)) return;

    const rng = sheet.getRange(WELCOME_BANNER_RANGE);
    const snapshot = {
      values: rng.getValues(),
      backgrounds: rng.getBackgrounds(),
      fontColors: rng.getFontColors(),
      fontWeights: rng.getFontWeights(),
      fontStyles: rng.getFontStyles(),
      fontFamilies: rng.getFontFamilies(),
      fontSizes: rng.getFontSizes(),
      horizontalAlignments: rng.getHorizontalAlignments(),
      mergedA1: rng.getMergedRanges().map(function (r) { return r.getA1Notation(); }),
    };
    props.setProperty(PROP_WELCOME_LOCKED_SNAPSHOT, JSON.stringify(snapshot));

    rng.breakApart();
    sheet.getRange('A1:L1').merge()
      .setValue('⚠️ RUNTIME ERROR 0x8007045D — Workbook gagal dimuat / sesi kedaluwarsa')
      .setBackground('#3C1414').setFontColor('#FF6B6B').setFontFamily('Courier New')
      .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center');
    sheet.getRange('A2:L2').merge()
      .setValue('Masukkan token aktivasi yang valid di bawah untuk memulihkan akses.')
      .setBackground('#3C1414').setFontColor('#FFB4B4').setFontFamily('Courier New')
      .setFontWeight('normal').setFontSize(10).setHorizontalAlignment('center');
    try { sheet.setTabColor('#E62129'); } catch (e2) { }
  } catch (e) { }
}

/** Kebalikan applyLockedAppearance_ — pulihin persis dari snapshot, bukan cuma bersihin. */
function clearLockedAppearance_(sheet) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(PROP_WELCOME_LOCKED_SNAPSHOT);
    if (!raw) return;
    let snap;
    try { snap = JSON.parse(raw); } catch (e2) { props.deleteProperty(PROP_WELCOME_LOCKED_SNAPSHOT); return; }

    const rng = sheet.getRange(WELCOME_BANNER_RANGE);
    rng.breakApart();
    rng.setValues(snap.values);
    rng.setBackgrounds(snap.backgrounds);
    rng.setFontColors(snap.fontColors);
    rng.setFontWeights(snap.fontWeights);
    rng.setFontStyles(snap.fontStyles);
    rng.setFontFamilies(snap.fontFamilies);
    rng.setFontSizes(snap.fontSizes);
    rng.setHorizontalAlignments(snap.horizontalAlignments);
    (snap.mergedA1 || []).forEach(function (a1) {
      try { sheet.getRange(a1).merge(); } catch (e3) { }
    });
    try { sheet.setTabColor(null); } catch (e4) { }
    props.deleteProperty(PROP_WELCOME_LOCKED_SNAPSHOT);
  } catch (e) { }
}

/** Proteksi WARNING-ONLY — gesekan tambahan, bukan pengunci mutlak. */
function addWarningProtection_(sheet) {
  try {
    const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    if (existing.length > 0) return;
    const p = sheet.protect().setDescription('Terkunci sampai token aktivasi tervalidasi — isi lewat dashboard ElevatEd, bukan manual.');
    p.setWarningOnly(true);
  } catch (e) { }
}
function removeWarningProtection_(sheet) {
  try {
    const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    existing.forEach(function (p) {
      if (p.getDescription().indexOf('Terkunci sampai token aktivasi') === 0) p.remove();
    });
  } catch (e) { }
}

// ============================================================
// DASHBOARD ASLI — TIDAK ADA YANG DIUBAH DARI SINI KE BAWAH,
// selain doGet(), onOpen(), openDashboardModal() yang sekarang
// milih layar Aktivasi vs Sidebar berdasarkan isActivated_().
// ============================================================

const DREAM_DATA_START_ROW = 18;
const DREAM_DATA_END_ROW   = 67;
const DREAM_COL = { no:2, goal:3, targetSem:4, smart:5, status:7, priority:9, decision:10, notes:11 };
const DREAM_STATUS_OPTIONS = ['On Progress', 'Tercapai', 'Gagal'];

const COLLEGE_DATA_START_ROW = 5;
const COLLEGE_DATA_END_ROW   = 25;
const COLLEGE_COL = { no:9, misi:10, semester:11, kategori:12, prioritas:13, progress:14 };
const COLLEGE_STATUS_OPTIONS = ['Belum Mulai', 'On Progress', 'Tercapai', 'Gagal'];

const AP_BLOCK_START_ROW = 4;
const AP_BLOCK_SIZE      = 13;
const AP_HEADER_OFFSET   = 0;
const AP_DATA_OFFSET     = 2;
const AP_DATA_ROWS       = 10;
const AP_MAX_ROW         = 1013;
const AP_GROUPS = [
  { key: 'Q1', label: 'Quarter 1', colGoal: 6,  colNo: 6,  colTaktik: 7,  colFrek: 8,  colStatus: 9  },
  { key: 'Q2', label: 'Quarter 2', colGoal: 13, colNo: 13, colTaktik: 14, colFrek: 15, colStatus: 16 },
];
const AP_PLACEHOLDER = 'Pilih Goal dari SMART Dreams di kiri...';
const AP_STATUS_OPTIONS = ['Belum Dimulai', 'Sedang Berjalan', 'Selesai'];
const AP_ARTIFACT_MARKERS = ['=ai(', 'xludf', 'dummyfunction', '#name?', '#ref!', '#error', '#n/a', '#value!', 'generate a', 'prompt:', 'expert level', 'placeholder'];
function looksLikeArtifact_(text) {
  const t = String(text || '').trim().toLowerCase();
  if (!t) return true;
  if (AP_ARTIFACT_MARKERS.some(mk => t.indexOf(mk) >= 0)) return true;
  if (/^(xs|s|m|l|xl)$/i.test(t)) return true;
  if (/\bexpert\b/.test(t) && /\bweeks?\b|\byears?\b/.test(t)) return true;
  return false;
}

const TRACKER_WEEKS_PER_QUARTER = 12;
const DAY_NAMES = ['Senin','Selasa','Rabu','Kamis','Jumat','Sabtu','Minggu'];

const SCORE_DREAM_TERCAPAI = 10;
const SCORE_DREAM_PROGRESS = 2;
const SCORE_COLLEGE_TERCAPAI = 5;
const SCORE_TRACKER_CHECK = 1;
const BADGES = [
  { min: 0,   name: 'Pemula Ikigai',  icon: '🌱' },
  { min: 20,  name: 'Mulai Panas',    icon: '🔥' },
  { min: 50,  name: 'Konsisten',      icon: '⭐' },
  { min: 100, name: 'Achiever',       icon: '🏆' },
  { min: 200, name: 'Master Ikigai',  icon: '👑' },
  { min: 350, name: 'Living Legend',  icon: '💎' },
];

const SS_ID_PROP_KEY = 'ikigai_ss_id';
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Buka Apps Script dari spreadsheet IKIGAI ini, lalu jalankan setupDashboard() sekali.');
  }
  PropertiesService.getScriptProperties().setProperty(SS_ID_PROP_KEY, ss.getId());
  return 'Dashboard tersambung ke: ' + ss.getName();
}
function getSS_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      try { PropertiesService.getScriptProperties().setProperty(SS_ID_PROP_KEY, active.getId()); } catch (e2) { }
      return active;
    }
  } catch (e) { }
  const cachedId = PropertiesService.getScriptProperties().getProperty(SS_ID_PROP_KEY);
  if (cachedId) return SpreadsheetApp.openById(cachedId);
  throw new Error('Spreadsheet belum ke-detect.');
}

/**
 * Satu file HTML doang (Sidebar.html) buat dua kondisi — status aktivasi
 * dikirim sebagai variabel template, JS di dalam Sidebar.html yang milih
 * mau nampilin layar aktivasi atau dashboard.
 */
/**
 * Google Sheets, kalau di-duplikat (File > Make a copy), NGGAK ikut
 * nyalin Script Properties — cuma kode-nya doang yang ke-copy. Jadi tiap
 * copy baru buat pembeli baru, VALIDATE_SECRET selalu kosong dari awal
 * dan bikin akun manapun kena "Konfigurasi belum lengkap" walau kode &
 * email-nya bener. Fungsi ini nge-isi otomatis kalau kosong, dipanggil di
 * SEMUA pintu masuk (doGet, onOpen, openDashboardModal) — jadi copy baru
 * langsung jalan tanpa perlu ada yang buka Project Settings manual.
 *
 * Nilai default di bawah HARUS PERSIS SAMA kayak WEBHOOK_SECRET_VALIDATE
 * di Rumah A (RumahA_DatabasePusat.gs) — kalau salah satu diganti, yang
 * satu lagi ikut diganti juga.
 */
const DEFAULT_VALIDATE_SECRET = 'ikg_valid_z4Tn9wRfB7cJ';
function ensureScriptProperties_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty('VALIDATE_SECRET')) {
      props.setProperty('VALIDATE_SECRET', DEFAULT_VALIDATE_SECRET);
    }
  } catch (e) { }
}

function doGet(e) {
  ensureScriptProperties_();
  const status = checkAndSyncActivation_();
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.activated = status.activated;
  template.activationError = status.error || '';
  template.initialView = 'overview';
  return template.evaluate().setTitle(status.activated ? 'ElevatEd Student Kit' : 'ElevatEd Student Kit — Aktivasi');
}

function onOpen() {
  ensureScriptProperties_();
  try { setupDashboard(); } catch (e) { }
  checkAndSyncActivation_();
  ensureEditTrigger_();
  SpreadsheetApp.getUi().createMenu('🦉 ElevatEd Student Kit').addItem('📊  Buka Dashboard', 'openDashboardModal').addToUi();
}

function openDashboardModal() {
  ensureScriptProperties_();
  setupDashboard();
  const status = checkAndSyncActivation_();
  const tmpl = HtmlService.createTemplateFromFile('Sidebar');
  tmpl.activated = status.activated;
  tmpl.activationError = status.error || '';
  tmpl.initialView = 'overview';
  const html = tmpl.evaluate().setWidth(1320).setHeight(830);
  SpreadsheetApp.getUi().showModalDialog(html, status.activated ? '🦉 ElevatEd Student Kit' : '🔑 Aktivasi Akses');
}

function getWebAppUrl_() {
  try {
    const url = ScriptApp.getService().getUrl();
    return url || '';
  } catch (e) { return ''; }
}

/**
 * Dipanggil dari tombol "Tab Baru" di sidebar — klik pertama. SEBELUMNYA
 * fungsi ini beneran nge-fetch balik URL-nya sendiri lewat UrlFetchApp
 * buat mastiin bisa diakses — ternyata itu nggak reliable, apalagi kalau
 * dipanggil dari dalam konteks modal dialog di Sheets (bukan dari request
 * HTTP asli): fetch server-ke-server ke URL sendiri bisa ke-block/dapet
 * status yang beda dari yang beneran dialamin browser asli, jadinya malah
 * salah nunjukin 404 padahal linknya sehat. Sekarang cukup ambil ULANG
 * URL TERBARU langsung dari deployment (bukan yang ke-cache lama di
 * browser) + cek formatnya masuk akal — itu udah cukup buat nyegah "salah
 * link" tanpa resiko false-negative dari self-fetch.
 */
function verifyWebAppUrl() {
  const url = getWebAppUrl_();
  if (!url) {
    return { ok: false, url: '', error: 'Web app ini belum ke-deploy sebagai URL (baru bisa diakses lewat menu di Sheets).' };
  }
  if (!/^https:\/\/script\.google\.com\/macros\//.test(url)) {
    return { ok: false, url: url, error: 'URL yang kedeteksi kelihatannya nggak valid: ' + url };
  }
  return { ok: true, url: url };
}

function getBundle(userSpreadsheet) {
  const dream = getDreamPlan_(userSpreadsheet);
  const college = getCollegePlan_(userSpreadsheet);
  return {
    personal: getPersonal_(userSpreadsheet),
    dream, college,
    action: getActionPlan_(userSpreadsheet),
    gamification: getGamification_(userSpreadsheet, dream, college),
    trackerConfig: { startDate: getStartDate_().toISOString() },
    webAppUrl: getWebAppUrl_(),
  };
}

function getInitialBundle(userSpreadsheet) {
  const dream = getDreamPlan_(userSpreadsheet);
  const college = getCollegePlan_(userSpreadsheet);
  return {
    personal: getPersonal_(userSpreadsheet),
    dream, college,
    gamification: getGamification_(userSpreadsheet, dream, college),
    trackerConfig: { startDate: getStartDate_().toISOString() },
    webAppUrl: getWebAppUrl_(),
  };
}

function getActionPlanOnly(userSpreadsheet) {
  return { action: getActionPlan_(userSpreadsheet) };
}

function semesterNum_(label) {
  const m = String(label || '').match(/(\d+)/);
  return m ? Number(m[1]) : 999;
}
function sortSemesters_(arr) { return arr.slice().sort((a, b) => semesterNum_(a) - semesterNum_(b)); }

function getValidationOptions_(sheet, sampleRow, col, fallback) {
  try {
    const rule = sheet.getRange(sampleRow, col).getDataValidation();
    if (rule && rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const vals = rule.getCriteriaValues()[0];
      if (vals && vals.length) return vals;
    }
  } catch (e) { }
  return fallback;
}
function cellFromGrid_(grid, a1) {
  const m = a1.match(/^([A-Z]+)(\d+)$/);
  let col = 0;
  for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
  const row = Number(m[2]);
  return (grid[row - 1] && grid[row - 1][col - 1] !== undefined) ? String(grid[row - 1][col - 1]) : '';
}
function isUnresolved_(displayVal, formulaStr) {
  if (!formulaStr) return false;
  const d = String(displayVal || '').trim();
  if (!d) return true;
  if (/^#(REF|NAME|N\/A|VALUE|ERROR|DIV\/0|NULL|NUM)/i.test(d)) return true;
  return false;
}

function getPersonal_(userSpreadsheet) {
  if (!userSpreadsheet) return null;
  const ss = userSpreadsheet;
  const ik = ss.getSheetByName(SHEET_IKIGAI);
  const sw = ss.getSheetByName(SHEET_SWOT);
  if (!ik) return null;

  const ikGrid = ik.getRange(1, 1, ik.getLastRow(), ik.getLastColumn()).getDisplayValues();
  const g = a1 => cellFromGrid_(ikGrid, a1);

  const name = g('D4'), major = g('D5'), semester = g('D6'), linear = g('D7'), mbti = g('C13');
  const via = [g('D14'), g('E14'), g('F14')].filter(Boolean);
  const careerExplorer = [g('G14'), g('H14'), g('I14')].filter(Boolean);
  const ikigaiSpot = g('C23'), sliceOfLife = g('C27'), sweetspot = g('C37');
  const hardskills = [g('D38'), g('E38'), g('F38')].filter(Boolean);
  const softskills = [g('G38'), g('H38'), g('I38')].filter(Boolean);
  const ikigaiSummary = g('F5');

  let swot = { strength: [], weakness: [], opportunity: [], threat: [] };
  let swotSummary = '';
  if (sw) {
    const swGrid = sw.getRange(1, 1, sw.getLastRow(), sw.getLastColumn()).getDisplayValues();
    const gs = a1 => cellFromGrid_(swGrid, a1);
    swot.strength    = [gs('D23'), gs('D24'), gs('D25')].filter(Boolean);
    swot.weakness    = [gs('F23'), gs('F24'), gs('F25')].filter(Boolean);
    swot.opportunity = [gs('H23'), gs('H24'), gs('H25')].filter(Boolean);
    swot.threat      = [gs('J23'), gs('J24'), gs('J25')].filter(Boolean);
    swotSummary = gs('F6');
  }
  return { name, major, semester, semesterNum: semesterNum_(semester), linear, mbti, via, careerExplorer, ikigaiSpot, sliceOfLife, sweetspot, hardskills, softskills, swot, ikigaiSummary, swotSummary };
}

function getDreamPlan_(userSpreadsheet) {
  if (!userSpreadsheet) return { semesters: [], goals: [], bySemester: {}, summary: {} };
  const ss = userSpreadsheet;
  const s = ss.getSheetByName(SHEET_DREAM);
  if (!s) return { semesters: [], goals: [], bySemester: {}, summary: {} };

  const nRows = DREAM_DATA_END_ROW - DREAM_DATA_START_ROW + 1;
  const data = s.getRange(DREAM_DATA_START_ROW, 1, nRows, 11).getDisplayValues();

  const goals = [];
  data.forEach((row, i) => {
    const goal = row[DREAM_COL.goal - 1];
    if (!goal) return;
    goals.push({
      row: DREAM_DATA_START_ROW + i, no: row[DREAM_COL.no - 1], goal: goal,
      targetSem: row[DREAM_COL.targetSem - 1] || 'Belum Ditentukan',
      smart: row[DREAM_COL.smart - 1], status: row[DREAM_COL.status - 1] || 'Belum Diisi',
      priority: row[DREAM_COL.priority - 1] || '', decision: row[DREAM_COL.decision - 1] || '',
      notes: row[DREAM_COL.notes - 1] || '',
    });
  });

  const bySemester = {};
  goals.forEach(g => { if (!bySemester[g.targetSem]) bySemester[g.targetSem] = []; bySemester[g.targetSem].push(g); });
  const semesters = sortSemesters_(Object.keys(bySemester));

  const summary = { total: goals.length, onProgress: 0, tercapai: 0, gagal: 0, lainnya: 0 };
  goals.forEach(g => {
    const st = String(g.status).toLowerCase();
    if (st.indexOf('progress') >= 0) summary.onProgress++;
    else if (st.indexOf('tercapai') >= 0) summary.tercapai++;
    else if (st.indexOf('gagal') >= 0) summary.gagal++;
    else summary.lainnya++;
  });
  return { semesters, goals, bySemester, summary, statusOptions: getValidationOptions_(s, DREAM_DATA_START_ROW, DREAM_COL.status, DREAM_STATUS_OPTIONS) };
}
function updateDreamStatus(userSpreadsheet, row, status) {
  try {
    if (!userSpreadsheet) throw new Error('userSpreadsheet tidak ditemukan');
    const sheet = userSpreadsheet.getSheetByName(SHEET_DREAM);
    sheet.getRange(row, DREAM_COL.status).setValue(status);
    SpreadsheetApp.flush();
    return { success: true, value: sheet.getRange(row, DREAM_COL.status).getDisplayValue() };
  } catch (e) { return { success: false, error: e.message }; }
}

function getCollegePlan_(userSpreadsheet) {
  if (!userSpreadsheet) return { semesters: [], missions: [], bySemester: {} };
  const ss = userSpreadsheet;
  const s = ss.getSheetByName(SHEET_COLLEGE);
  if (!s) return { semesters: [], missions: [], bySemester: {} };

  const nRows = COLLEGE_DATA_END_ROW - COLLEGE_DATA_START_ROW + 1;
  const data = s.getRange(COLLEGE_DATA_START_ROW, 1, nRows, 14).getDisplayValues();

  const missions = [];
  data.forEach((row, i) => {
    const misi = row[COLLEGE_COL.misi - 1];
    if (!misi) return;
    missions.push({
      row: COLLEGE_DATA_START_ROW + i, no: row[COLLEGE_COL.no - 1], misi: misi,
      semester: row[COLLEGE_COL.semester - 1] || 'Belum Ditentukan',
      kategori: row[COLLEGE_COL.kategori - 1] || 'Lainnya',
      prioritas: row[COLLEGE_COL.prioritas - 1] || '', progress: row[COLLEGE_COL.progress - 1] || 'Belum Mulai',
    });
  });
  const bySemester = {};
  missions.forEach(m => { if (!bySemester[m.semester]) bySemester[m.semester] = []; bySemester[m.semester].push(m); });
  const semesters = sortSemesters_(Object.keys(bySemester));
  return { semesters, missions, bySemester, statusOptions: getValidationOptions_(s, COLLEGE_DATA_START_ROW, COLLEGE_COL.progress, COLLEGE_STATUS_OPTIONS) };
}
function updateCollegeProgress(userSpreadsheet, row, status) {
  try {
    if (!userSpreadsheet) throw new Error('userSpreadsheet tidak ditemukan');
    const sheet = userSpreadsheet.getSheetByName(SHEET_COLLEGE);
    sheet.getRange(row, COLLEGE_COL.progress).setValue(status);
    SpreadsheetApp.flush();
    return { success: true, value: sheet.getRange(row, COLLEGE_COL.progress).getDisplayValue() };
  } catch (e) { return { success: false, error: e.message }; }
}

function updateActionStatus(userSpreadsheet, row, col, status) {
  try {
    if (!userSpreadsheet) throw new Error('userSpreadsheet tidak ditemukan');
    const sheet = userSpreadsheet.getSheetByName(SHEET_ACTION);
    sheet.getRange(row, col).setValue(status);
    SpreadsheetApp.flush();
    return { success: true, value: sheet.getRange(row, col).getDisplayValue() };
  } catch (e) { return { success: false, error: e.message }; }
}

function getActionPlan_(userSpreadsheet) {
  if (!userSpreadsheet) return { blocks: [], statusOptions: AP_STATUS_OPTIONS };
  const ss = userSpreadsheet;
  const s = ss.getSheetByName(SHEET_ACTION);
  if (!s) return { blocks: [], statusOptions: AP_STATUS_OPTIONS };

  const lastRow = Math.min(AP_MAX_ROW, s.getLastRow());
  const lastCol = 16;
  if (lastRow < AP_BLOCK_START_ROW) return { blocks: [], statusOptions: AP_STATUS_OPTIONS };
  const grid = s.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const formulaGrid = s.getRange(1, 1, lastRow, lastCol).getFormulas();
  const cell = (r, c) => (grid[r - 1] && grid[r - 1][c - 1] !== undefined) ? grid[r - 1][c - 1] : '';
  const cellF = (r, c) => (formulaGrid[r - 1] && formulaGrid[r - 1][c - 1] !== undefined) ? formulaGrid[r - 1][c - 1] : '';

  const blocks = [];
  let idx = 0;
  for (let start = AP_BLOCK_START_ROW; start <= lastRow; start += AP_BLOCK_SIZE) {
    AP_GROUPS.forEach(grp => {
      const headerRow = start + AP_HEADER_OFFSET;
      if (headerRow > lastRow) return;
      const headerText = cell(headerRow, grp.colGoal);
      if (!headerText || headerText === AP_PLACEHOLDER || isUnresolved_(headerText, cellF(headerRow, grp.colGoal)) || looksLikeArtifact_(headerText)) return;

      const dataStart = start + AP_DATA_OFFSET;
      const tactics = [];
      for (let j = 0; j < AP_DATA_ROWS; j++) {
        const r = dataStart + j;
        if (r > lastRow) break;
        const taktikVal = cell(r, grp.colTaktik);
        if (!taktikVal || isUnresolved_(taktikVal, cellF(r, grp.colTaktik)) || looksLikeArtifact_(taktikVal)) continue;

        const frekValRaw = cell(r, grp.colFrek);
        const frekBad = isUnresolved_(frekValRaw, cellF(r, grp.colFrek)) || looksLikeArtifact_(frekValRaw);
        const frekuensi = frekBad ? '' : frekValRaw;
        const statusValRaw = cell(r, grp.colStatus);
        const statusBad = isUnresolved_(statusValRaw, cellF(r, grp.colStatus)) || looksLikeArtifact_(statusValRaw);
        const status = statusBad ? 'Belum Dimulai' : (statusValRaw || 'Belum Dimulai');
        const meta = frekuensi ? parseFrequencyMeta_(frekuensi) : { target: 1, periodWeeks: 1 };

        tactics.push({
          no: cell(r, grp.colNo), taktik: taktikVal, frekuensi, status,
          target: meta.target, periodWeeks: meta.periodWeeks,
          row: r, statusCol: grp.colStatus,
        });
      }
      if (tactics.length === 0) return;
      blocks.push({ idx: idx++, quarter: grp.key, quarterLabel: grp.label, goal: headerText, tactics });
    });
  }
  const statusOptions = getValidationOptions_(s, AP_BLOCK_START_ROW + AP_DATA_OFFSET, AP_GROUPS[0].colStatus, AP_STATUS_OPTIONS);
  return { blocks, statusOptions };
}

  const lastRow = Math.min(AP_MAX_ROW, s.getLastRow());
  const lastCol = 16;
  if (lastRow < AP_BLOCK_START_ROW) return { blocks: [], statusOptions: AP_STATUS_OPTIONS };
  const grid = s.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const formulaGrid = s.getRange(1, 1, lastRow, lastCol).getFormulas();
  const cell = (r, c) => (grid[r - 1] && grid[r - 1][c - 1] !== undefined) ? grid[r - 1][c - 1] : '';
  const cellF = (r, c) => (formulaGrid[r - 1] && formulaGrid[r - 1][c - 1] !== undefined) ? formulaGrid[r - 1][c - 1] : '';

  const blocks = [];
  let idx = 0;
  for (let start = AP_BLOCK_START_ROW; start <= lastRow; start += AP_BLOCK_SIZE) {
    AP_GROUPS.forEach(grp => {
      const headerRow = start + AP_HEADER_OFFSET;
      if (headerRow > lastRow) return;
      const headerText = cell(headerRow, grp.colGoal);
      if (!headerText || headerText === AP_PLACEHOLDER || isUnresolved_(headerText, cellF(headerRow, grp.colGoal)) || looksLikeArtifact_(headerText)) return;

      const dataStart = start + AP_DATA_OFFSET;
      const tactics = [];
      for (let j = 0; j < AP_DATA_ROWS; j++) {
        const r = dataStart + j;
        if (r > lastRow) break;
        const taktikVal = cell(r, grp.colTaktik);
        if (!taktikVal || isUnresolved_(taktikVal, cellF(r, grp.colTaktik)) || looksLikeArtifact_(taktikVal)) continue;

        const frekValRaw = cell(r, grp.colFrek);
        const frekBad = isUnresolved_(frekValRaw, cellF(r, grp.colFrek)) || looksLikeArtifact_(frekValRaw);
        const frekuensi = frekBad ? '' : frekValRaw;
        const statusValRaw = cell(r, grp.colStatus);
        const statusBad = isUnresolved_(statusValRaw, cellF(r, grp.colStatus)) || looksLikeArtifact_(statusValRaw);
        const status = statusBad ? 'Belum Dimulai' : (statusValRaw || 'Belum Dimulai');
        const meta = frekuensi ? parseFrequencyMeta_(frekuensi) : { target: 1, periodWeeks: 1 };

        tactics.push({
          no: cell(r, grp.colNo), taktik: taktikVal, frekuensi, status,
          target: meta.target, periodWeeks: meta.periodWeeks,
          row: r, statusCol: grp.colStatus,
        });
      }
      if (tactics.length === 0) return;
      blocks.push({ idx: idx++, quarter: grp.key, quarterLabel: grp.label, goal: headerText, tactics });
    });
  }
  const statusOptions = getValidationOptions_(s, AP_BLOCK_START_ROW + AP_DATA_OFFSET, AP_GROUPS[0].colStatus, AP_STATUS_OPTIONS);
  return { blocks, statusOptions };
}
const ID_NUM_WORDS_ = { satu:1, dua:2, tiga:3, empat:4, lima:5, enam:6, tujuh:7, delapan:8 };
function wordOrDigitToNum_(s) {
  if (!s) return null;
  if (/^\d+$/.test(s)) return Number(s);
  return ID_NUM_WORDS_[s] || null;
}
function parseFrequencyMeta_(freq) {
  const t = String(freq || '').toLowerCase();

  let periodWeeks = 1;
  const pm = t.match(/setiap\s+(dua|tiga|empat|lima|enam|\d+)\s*minggu/) ||
             t.match(/(dua|tiga|empat|lima|enam|\d+)\s*minggu\s*sekali/) ||
             t.match(/(?:tiap|per)\s+(dua|tiga|empat|lima|enam|\d+)\s*minggu/);
  if (pm) {
    const n = wordOrDigitToNum_(pm[1]);
    if (n) periodWeeks = n;
  } else if (/bulan\s*sekali|per\s*bulan|\bbulanan\b|\bmonthly\b|tiap\s*bulan|\bsebulan\b/.test(t)) {
    periodWeeks = 4;
  }

  let target = 1;
  if (/setiap\s*hari|everyday|\bdaily\b|\bharian\b/.test(t)) {
    target = 7; periodWeeks = 1;
  } else {
    const xm = t.match(/(\d+)\s*x/);
    if (xm) target = Math.max(1, Number(xm[1]));
    else if (periodWeeks === 1) {
      const m2 = t.match(/(\d+)/);
      if (m2) target = Math.max(1, Number(m2[1]));
    }
  }
  return { target, periodWeeks };
}

function ensureTrackerState_(userSpreadsheet) {
  if (!userSpreadsheet) throw new Error('userSpreadsheet tidak ditemukan');
  const ss = userSpreadsheet;
  let s = ss.getSheetByName(SHEET_TRACKER_STATE);
  if (!s) {
    s = ss.insertSheet(SHEET_TRACKER_STATE);
    s.getRange(1, 1, 1, 8).setValues([['Key', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']])
      .setBackground('#143358').setFontColor('#fff').setFontWeight('bold');
    s.setFrozenRows(1);
    s.hideSheet();
  }
  return s;
}
function periodLabelForWeek_(week, periodWeeks) {
  if (!periodWeeks || periodWeeks <= 1) return 'W' + week;
  const periodStartWeek = Math.floor((week - 1) / periodWeeks) * periodWeeks + 1;
  return 'P' + periodStartWeek;
}
function trackerKey_(quarter, goalIdx, tacticNo, weekLabel) { return quarter + '|' + goalIdx + '|' + tacticNo + '|' + weekLabel; }
function readTrackerStateMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    data.forEach((row, i) => { map[row[0]] = { row: i + 2, days: row.slice(1).map(v => v === true) }; });
  }
  return map;
}
function trackerRow_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) return i + 2;
  }
  const newRow = lastRow + 1;
  sheet.getRange(newRow, 1, 1, 8).setValues([[key, false, false, false, false, false, false, false]]);
  return newRow;
}

function getStartDate_() {
  const prop = PropertiesService.getScriptProperties().getProperty('ikigai_tracker_start');
  if (prop) return new Date(prop);
  const now = new Date();
  const day = now.getDay();
  const diff = (day === 0 ? -6 : 1 - day);
  const monday = new Date(now);
  monday.setDate(now.getDate() + diff);
  monday.setHours(0, 0, 0, 0);
  return monday;
}
function setTrackerStartDate(dateStr) {
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return { success: false, error: 'Tanggal tidak valid' };
    PropertiesService.getScriptProperties().setProperty('ikigai_tracker_start', d.toISOString());
    return { success: true, startDate: d.toISOString() };
  } catch (e) { return { success: false, error: e.message }; }
}

function getTrackerView(userSpreadsheet, quarter, cachedBlocks) {
  const blocksAll = Array.isArray(cachedBlocks) ? cachedBlocks : getActionPlan_(userSpreadsheet).blocks;
  const start = getStartDate_();
  const stateSheet = ensureTrackerState_(userSpreadsheet);
  const stateMap = readTrackerStateMap_(stateSheet);
  const blocks = blocksAll.filter(b => b.quarter === quarter);

  const weeks = [];
  for (let w = 1; w <= TRACKER_WEEKS_PER_QUARTER; w++) {
    const weekStart = new Date(start);
    weekStart.setDate(start.getDate() + (w - 1) * 7);
    const days = DAY_NAMES.map((nm, di) => {
      const dt = new Date(weekStart);
      dt.setDate(weekStart.getDate() + di);
      return { name: nm, date: Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM') };
    });
    weeks.push({ w, label: 'Minggu ' + w, rangeLabel: days[0].date + ' – ' + days[6].date, days });
  }

  const newKeys = [];
  const goals = blocks.map(b => {
    const tactics = b.tactics.map(t => {
      const rows = [];
      for (let w = 1; w <= TRACKER_WEEKS_PER_QUARTER; w++) {
        const label = periodLabelForWeek_(w, t.periodWeeks);
        const key = trackerKey_(quarter, b.idx, t.no, label);
        let entry = stateMap[key];
        if (!entry) { entry = { days: [false, false, false, false, false, false, false] }; stateMap[key] = entry; newKeys.push(key); }
        rows.push({ week: w, periodStartWeek: Number(label.replace(/^[A-Z]/, '')) || w, days: entry.days, checkedCount: entry.days.filter(Boolean).length });
      }
      return { no: t.no, taktik: t.taktik, frekuensi: t.frekuensi, target: t.target, periodWeeks: t.periodWeeks || 1, weeks: rows };
    });
    return { idx: b.idx, goal: b.goal, tactics };
  });

  if (newKeys.length > 0) {
    const startRow = stateSheet.getLastRow() + 1;
    const values = newKeys.map(k => [k, false, false, false, false, false, false, false]);
    stateSheet.getRange(startRow, 1, values.length, 8).setValues(values);
  }
  return { quarter, weeks, goals, startDate: start.toISOString() };
}

function toggleTrackerCheck2(userSpreadsheet, quarter, goalIdx, tacticNo, week, dayIndex, value, target, periodWeeks) {
  const s = ensureTrackerState_(userSpreadsheet);
  const label = periodLabelForWeek_(week, periodWeeks);
  const key = trackerKey_(quarter, goalIdx, tacticNo, label);
  const row = trackerRow_(s, key);
  const vals = s.getRange(row, 2, 1, 7).getValues()[0];
  const checkedCount = vals.filter(v => v === true).length;
  if (value === true && vals[dayIndex] !== true && checkedCount >= target) {
    return { success: false, locked: true, error: 'Target minggu ini udah tercapai (' + target + 'x). Uncheck salah satu dulu kalau mau ganti.' };
  }
  s.getRange(row, 2 + dayIndex).setValue(value === true);
  SpreadsheetApp.flush();
  return { success: true };
}

function checkAllTrackerWeek(userSpreadsheet, quarter, goalIdx, tacticNo, week, target, periodWeeks) {
  try {
    const s = ensureTrackerState_(userSpreadsheet);
    const label = periodLabelForWeek_(week, periodWeeks);
    const key = trackerKey_(quarter, goalIdx, tacticNo, label);
    const row = trackerRow_(s, key);
    const vals = [];
    for (let i = 0; i < 7; i++) vals.push(i < target);
    s.getRange(row, 2, 1, 7).setValues([vals]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function getGamification_(userSpreadsheet, dream, college) {
  dream = dream || getDreamPlan_(userSpreadsheet);
  college = college || getCollegePlan_(userSpreadsheet);
  const dreamScore = dream.summary.tercapai * SCORE_DREAM_TERCAPAI + dream.summary.onProgress * SCORE_DREAM_PROGRESS;
  const collegeTercapai = college.missions.filter(m => String(m.progress).toLowerCase().indexOf('tercapai') >= 0).length;
  const collegeScore = collegeTercapai * SCORE_COLLEGE_TERCAPAI;

  let trackerScore = 0;
  const stateSheet = ensureTrackerState_(userSpreadsheet);
  const lastRow = stateSheet.getLastRow();
  if (lastRow >= 2) {
    const vals = stateSheet.getRange(2, 2, lastRow - 1, 7).getValues();
    vals.forEach(row => row.forEach(v => { if (v === true) trackerScore += SCORE_TRACKER_CHECK; }));
  }

  const total = dreamScore + collegeScore + trackerScore;
  let badge = BADGES[0], nextBadge = null;
  for (let i = 0; i < BADGES.length; i++) {
    if (total >= BADGES[i].min) badge = BADGES[i];
    else { nextBadge = BADGES[i]; break; }
  }

  const specialBadges = [];
  if (dream.summary.tercapai >= 1) specialBadges.push({ icon: '🥇', name: 'First Win', desc: 'Dream goal pertama tercapai' });
  if (dream.summary.tercapai >= 3) specialBadges.push({ icon: '🎯', name: 'Goal Getter', desc: '3 dream goals tercapai' });
  if (collegeTercapai >= 3) specialBadges.push({ icon: '📚', name: 'Akademis Solid', desc: '3 misi kuliah tercapai' });
  if (trackerScore >= 28) specialBadges.push({ icon: '💪', name: 'Grinder', desc: 'Setara 4 minggu checklist penuh' });
  if (trackerScore >= 84) specialBadges.push({ icon: '🚀', name: 'Full Quarter', desc: 'Setara 12 minggu checklist penuh' });

  return { total, badge, nextBadge, specialBadges, breakdown: { dream: dreamScore, college: collegeScore, tracker: trackerScore } };
}
function getGamificationOnly(userSpreadsheet) { return getGamification_(userSpreadsheet); }

function getReportData_(userSpreadsheet) {
  const personal = getPersonal_(userSpreadsheet);
  const dream = getDreamPlan_(userSpreadsheet);
  const college = getCollegePlan_(userSpreadsheet);
  const gamification = getGamification_(userSpreadsheet, dream, college);
  const curNum = personal ? personal.semesterNum : 999;

  const dreamEval = dream.goals.filter(g => semesterNum_(g.targetSem) <= curNum)
    .map(g => Object.assign({}, g, { overdue: semesterNum_(g.targetSem) < curNum && String(g.status).toLowerCase().indexOf('tercapai') < 0 }));
  const collegeEval = college.missions.filter(m => semesterNum_(m.semester) <= curNum)
    .map(m => Object.assign({}, m, { overdue: semesterNum_(m.semester) < curNum && String(m.progress).toLowerCase().indexOf('tercapai') < 0 }));

  return {
    personal, dreamEval, collegeEval, gamification,
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy, HH:mm"),
    dreamDoneCount: dreamEval.filter(g => String(g.status).toLowerCase().indexOf('tercapai') >= 0).length,
    collegeDoneCount: collegeEval.filter(m => String(m.progress).toLowerCase().indexOf('tercapai') >= 0).length,
  };
}
function getReportBundle(userSpreadsheet) { return getReportData_(userSpreadsheet); }

// ============================================================
// PERSONAL IDENTITY REPORT (penjabaran MBTI dsb)
// ============================================================
const MBTI_INFO = {
  INTJ: { title: 'The Architect', strengths: ['Strategis & visioner', 'Mandiri, kerja efisien', 'Analitis, suka sistem'], weaknesses: ['Kadang terlalu perfeksionis', 'Kurang ekspresif secara emosi', 'Bisa terkesan dingin/jaga jarak'], career: 'Cocok di peran riset, strategi bisnis, engineering, atau apapun yang butuh perencanaan jangka panjang & problem solving kompleks.' },
  INTP: { title: 'The Logician', strengths: ['Pemikir logis & kreatif', 'Haus ilmu baru', 'Objektif, ga gampang bias'], weaknesses: ['Suka overthinking', 'Kurang sabar sama hal teknis-administratif', 'Sering nunda eksekusi demi mikir lebih matang'], career: 'Cocok di riset, data science, software development, atau bidang yang kasih ruang eksplorasi ide.' },
  ENTJ: { title: 'The Commander', strengths: ['Leader natural', 'Tegas & efisien ambil keputusan', 'Ambisius, orientasi hasil'], weaknesses: ['Bisa terlalu dominan', 'Kurang sabar sama proses lambat', 'Kadang ngabaikan sisi emosional tim'], career: 'Cocok di manajemen, kewirausahaan, konsultan bisnis, atau posisi kepemimpinan proyek.' },
  ENTP: { title: 'The Debater', strengths: ['Ide cepat & inovatif', 'Jago diskusi/debat', 'Adaptif ke perubahan'], weaknesses: ['Gampang bosan sama rutinitas', 'Suka nunda hal detail', 'Kadang argumentatif berlebihan'], career: 'Cocok di marketing kreatif, entrepreneurship, konsultan, atau bidang yang butuh inovasi konstan.' },
  INFJ: { title: 'The Advocate', strengths: ['Empatik & insightful', 'Idealis tapi terorganisir', 'Peka sama makna & tujuan'], weaknesses: ['Gampang burnout krn terlalu banyak mikirin orang lain', 'Perfeksionis', 'Susah nolak permintaan orang'], career: 'Cocok di bidang konseling, pendidikan, HR, social impact, atau penulisan.' },
  INFP: { title: 'The Mediator', strengths: ['Idealis & penuh empati', 'Kreatif, kuat di nilai personal', 'Fleksibel & terbuka'], weaknesses: ['Kadang terlalu sensitif kritik', 'Sulit ambil keputusan cepat', 'Gampang teralihkan dari target'], career: 'Cocok di penulisan, seni, psikologi, non-profit, atau bidang yang selaras values pribadi.' },
  ENFJ: { title: 'The Protagonist', strengths: ['Karismatik & memotivasi orang', 'Empatik, jago komunikasi', 'Terorganisir dalam mengejar goal'], weaknesses: ['Kadang terlalu ngurusin kebutuhan orang lain', 'Sensitif kritik', 'Bisa overcommit'], career: 'Cocok di pendidikan, HR, public speaking, community building, atau leadership sosial.' },
  ENFP: { title: 'The Campaigner', strengths: ['Enerjik & penuh ide', 'People person, gampang connect', 'Adaptif & antusias'], weaknesses: ['Gampang teralih fokus', 'Kurang suka rutinitas detail', 'Overthinking soal kemungkinan'], career: 'Cocok di marketing, content creation, event, entrepreneurship, atau bidang kreatif-sosial.' },
  ISTJ: { title: 'The Logistician', strengths: ['Disiplin & bisa diandalkan', 'Terorganisir, detail-oriented', 'Konsisten & bertanggung jawab'], weaknesses: ['Kaku sama perubahan mendadak', 'Kurang ekspresif', 'Bisa terlalu terpaku aturan'], career: 'Cocok di akuntansi, administrasi, hukum, project management, atau bidang yang butuh presisi.' },
  ISFJ: { title: 'The Defender', strengths: ['Teliti & suportif', 'Loyal & bertanggung jawab', 'Sabar & telaten'], weaknesses: ['Susah nolak permintaan', 'Kurang nyaman jadi sorotan', 'Gampang overwhelmed kalau banyak tuntutan'], career: 'Cocok di kesehatan, pendidikan, administrasi, customer service, atau social work.' },
  ESTJ: { title: 'The Executive', strengths: ['Terorganisir & tegas', 'Praktis, orientasi hasil', 'Bertanggung jawab & disiplin'], weaknesses: ['Kadang terlalu kaku sama aturan', 'Kurang sabar sama ide abstrak', 'Bisa dominan dalam diskusi'], career: 'Cocok di manajemen operasional, project lead, hukum, atau bidang yang butuh struktur jelas.' },
  ESFJ: { title: 'The Consul', strengths: ['Suportif & perhatian ke orang lain', 'Terorganisir & kooperatif', 'Jago menjaga harmoni tim'], weaknesses: ['Terlalu peduli opini orang', 'Kurang nyaman sama konflik', 'Bisa overcommit demi orang lain'], career: 'Cocok di HR, event organizing, pendidikan, customer relations, atau community management.' },
  ISTP: { title: 'The Virtuoso', strengths: ['Praktis & jago solve masalah teknis', 'Tenang di situasi krisis', 'Mandiri & adaptif'], weaknesses: ['Kurang suka komitmen jangka panjang', 'Kadang kurang ekspresif emosi', 'Bisa impulsif'], career: 'Cocok di engineering, IT, teknisi, olahraga, atau bidang hands-on/problem solving langsung.' },
  ISFP: { title: 'The Adventurer', strengths: ['Kreatif & estetik', 'Fleksibel & rendah hati', 'Peka sama detail & suasana'], weaknesses: ['Kurang suka konflik/konfrontasi', 'Sulit rencana jangka panjang', 'Gampang stres kalau dikritik'], career: 'Cocok di desain, seni, fotografi, kuliner, atau bidang kreatif yang fleksibel.' },
  ESTP: { title: 'The Entrepreneur', strengths: ['Energik & berani ambil aksi', 'Jago improvisasi', 'Realistis & praktis'], weaknesses: ['Kurang sabar sama teori/rencana panjang', 'Bisa impulsif', 'Gampang bosan sama rutinitas'], career: 'Cocok di sales, entrepreneurship, event, olahraga, atau bidang yang dinamis & fast-paced.' },
  ESFP: { title: 'The Entertainer', strengths: ['Ceria & mudah bergaul', 'Spontan & energik', 'Jago baca suasana sosial'], weaknesses: ['Kurang suka planning detail', 'Gampang teralihkan', 'Sensitif sama kritik'], career: 'Cocok di entertainment, marketing, event, hospitality, atau bidang yang banyak interaksi orang.' },
};
function normKey_(s) { return String(s || '').toLowerCase().replace(/[^a-z]/g, ''); }
const VIA_INFO = {
  creativity: 'Kamu jago mikirin cara baru/orisinal buat nyelesain masalah atau bikin sesuatu — modal kuat buat kerja kreatif & problem-solving non-standar.',
  curiosity: 'Kamu punya dorongan alami buat eksplorasi hal baru & nanya "kenapa" — bagus buat riset, belajar mandiri, dan nemuin peluang yang orang lain lewatin.',
  judgment: 'Kamu mikir kritis, nimbang bukti sebelum ambil keputusan — cocok buat analisis, riset, atau peran yang butuh objektivitas.',
  loveoflearning: 'Kamu seneng nambah ilmu/skill baru demi belajar itu sendiri, bukan cuma buat nilai — modal kuat buat terus upgrade diri sepanjang karir.',
  perspective: 'Kamu bisa ngasih sudut pandang yang matang & masuk akal ke orang lain, sering jadi tempat curhat/minta saran — modal bagus buat mentoring atau peran penasihat.',
  bravery: 'Kamu berani ambil sikap atau tindakan meski ada risiko/tekanan — modal penting buat situasi yang butuh keputusan cepat atau speak up.',
  perseverance: 'Kamu tetap jalan meski capek atau ada hambatan, ga gampang nyerah di tengah jalan — modal kuat buat ngejar goal jangka panjang.',
  honesty: 'Kamu jujur & konsisten antara ucapan & tindakan — bikin orang lain gampang percaya sama kamu, modal penting buat kerja tim & leadership.',
  zest: 'Kamu jalanin hidup/kerja dengan energi & antusiasme, bukan sekadar ngejalanin rutinitas — bikin kamu lebih tahan lama dari burnout.',
  love: 'Kamu ngasih ruang buat hubungan yang deket & saling peduli sama orang lain — modal penting buat kerja yang butuh kolaborasi erat.',
  kindness: 'Kamu terbiasa bantu orang lain tanpa itung-itungan balasan — bikin kamu gampang dipercaya & disenengin di lingkungan kerja/tim.',
  socialintelligence: 'Kamu peka sama perasaan & motivasi orang lain, tahu gimana harus bersikap di situasi sosial yang beda-beda — modal kuat buat kerja yang banyak interaksi orang.',
  teamwork: 'Kamu bisa kerja bareng orang lain demi tujuan bersama, bukan cuma mikirin diri sendiri — modal penting buat kerja tim/proyek grup.',
  fairness: 'Kamu perlakuin orang lain secara adil, ga pilih kasih berdasarkan perasaan pribadi — modal bagus buat peran yang butuh objektivitas & kepercayaan tim.',
  leadership: 'Kamu natural buat ngajak/ngarahin orang lain ke tujuan bersama — modal kuat buat peran yang butuh koordinasi tim atau project lead.',
  forgiveness: 'Kamu bisa ngasih maaf & ga nyimpen dendam ke orang yang salah — bikin kamu lebih tenang & gampang lanjut kerja bareng orang lagi.',
  humility: 'Kamu ga suka nyombong atau nunjukin diri lebih dari orang lain, biarin karya kamu yang ngomong — bikin orang lain nyaman kerja bareng kamu.',
  prudence: 'Kamu hati-hati dalam ambil keputusan, mikirin konsekuensi jangka panjang sebelum bertindak — modal bagus buat peran yang butuh manajemen risiko.',
  selfregulation: 'Kamu bisa ngatur emosi/dorongan diri sendiri dengan baik — penting buat kerja di bawah tekanan/deadline.',
  appreciationofbeauty: 'Kamu peka & terinspirasi sama keindahan/excellence di sekitar kamu — modal bagus buat desain, seni, atau bidang kreatif.',
  gratitude: 'Kamu terbiasa sadar & berterima kasih atas hal baik yang kamu punya — bikin kamu lebih resilient hadapin tekanan.',
  hope: 'Kamu optimis & yakin masa depan bisa lebih baik, plus kerja buat nyampe ke sana — modal penting buat ngejalanin rencana jangka panjang.',
  humor: 'Kamu jago bikin suasana lebih ringan lewat humor — modal sosial yang kuat, bikin orang nyaman deket kamu.',
  spirituality: 'Kamu punya rasa purpose/makna yang kuat tentang kenapa kamu ngelakuin sesuatu — bikin kamu lebih tahan banting ngejar tujuan jangka panjang.',
};
const VIA_ALIASES = {
  kreativitas:'creativity', kreatif:'creativity',
  rasaingintahu:'curiosity', keingintahuan:'curiosity',
  penilaian:'judgment', berpikirkritis:'judgment', criticalthinking:'judgment', pemikirankritis:'judgment',
  cintabelajar:'loveoflearning',
  perspektif:'perspective', kebijaksanaan:'perspective', wisdom:'perspective',
  keberanian:'bravery',
  ketekunan:'perseverance', kegigihan:'perseverance',
  kejujuran:'honesty', autentik:'honesty', authenticity:'honesty',
  semangat:'zest', antusiasme:'zest',
  cinta:'love',
  kebaikan:'kindness', kebaikanhati:'kindness',
  kecerdasansosial:'socialintelligence',
  kerjasama:'teamwork', kewargaan:'teamwork', citizenship:'teamwork',
  keadilan:'fairness',
  kepemimpinan:'leadership',
  pemaaf:'forgiveness', pengampunan:'forgiveness',
  kerendahanhati:'humility', modesty:'humility',
  kehatihatian:'prudence',
  pengaturandiri:'selfregulation', selfregulation:'selfregulation',
  apresiasikeindahan:'appreciationofbeauty', appreciationofbeauty:'appreciationofbeauty',
  rasasyukur:'gratitude', bersyukur:'gratitude',
  harapan:'hope', optimisme:'hope', optimism:'hope',
  humor:'humor', seleraumor:'humor',
  spiritualitas:'spirituality', maknahidup:'spirituality', purpose:'spirituality',
};
function lookupVia_(name) {
  const norm = normKey_(name);
  const key = VIA_INFO[norm] ? norm : VIA_ALIASES[norm];
  return key && VIA_INFO[key] ? VIA_INFO[key] : null;
}
function genericCareerNote_(career) {
  return 'Kalau kamu condong ke arah "' + career + '", biasanya itu nunjukkin kecocokan alami sama dunia kerja yang fokus di bidang itu. Mulai validasi lewat magang, proyek kecil, ikut komunitas terkait, atau ngobrol langsung sama orang yang udah kerja di sana — biar tahu beneran cocok sebelum komit penuh.';
}
function genericSkillNote_(skill, type) {
  return type === 'hard'
    ? 'Skill teknis "' + skill + '" ini biasanya jadi pembeda di CV/portofolio kamu. Mulai diasah lewat kursus online, project pribadi, atau tugas kuliah yang relevan.'
    : 'Soft skill "' + skill + '" ini penting buat kerja sama tim & pertumbuhan karir jangka panjang. Latih lewat organisasi, project kelompok, atau pengalaman kepanitiaan.';
}
const TIPS_PERSONAL = [
  'Coba 1 aktivitas kecil tiap minggu yang manfaatin kekuatan utama kamu (dari VIA character) — biar makin kebiasa dipakai secara sadar.',
  'Kalau ada weakness yang keulang terus, jangan buru-buru "diperbaiki" sendirian — cari cara kerja/tim yang nutupin gap itu (misal partneran sama orang yang komplementer).',
  'Update ulang hasil MBTI/VIA/Ikigai ini tiap semester — karakter & minat kamu bisa berkembang seiring pengalaman baru.',
];
const TIPS_COLLEGE = [
  'Pilih organisasi/UKM/proyek kuliah yang nyambung sama sweetspot kamu, bukan asal ikut rame-rame — biar portofolionya konsisten.',
  'Manfaatin tugas/skripsi buat latihan hardskill yang kamu butuhin, jangan cuma ngejar nilai doang.',
  'Cari mentor atau senior yang udah di jalur karir yang kamu incar, minimal buat sekali ngobrol tiap semester.',
];

function getIdentityReport_(userSpreadsheet) {
  const p = getPersonal_(userSpreadsheet);
  const match = String(p ? p.mbti : '').toUpperCase().match(/[EI][NS][FT][JP]/);
  const code = match ? match[0] : '';
  const info = MBTI_INFO[code] || null;
  const viaDetailed = (p ? p.via : []).map(name => ({ name, desc: lookupVia_(name) }));
  const hardDetailed = (p ? p.hardskills : []).map(name => ({ name, note: genericSkillNote_(name, 'hard') }));
  const softDetailed = (p ? p.softskills : []).map(name => ({ name, note: genericSkillNote_(name, 'soft') }));
  return {
    personal: p, mbtiCode: code, mbtiInfo: info,
    viaDetailed, hardDetailed, softDetailed,
    tipsPersonal: TIPS_PERSONAL, tipsCollege: TIPS_COLLEGE,
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy"),
  };
}
function getIdentityBundle(userSpreadsheet) { return getIdentityReport_(userSpreadsheet); }

// ============================================================
// SETUP — wizard pengisian Ikigai & SWOT
// Prompt AI tetap dibaca dari formula cell (A17/A31/SWOT!A17) — TIDAK
// dipindah kemanapun, sesuai keputusan terbaru.
// ============================================================

const SETUP_FIELDS = {
  ctx_name:        { sheet: SHEET_IKIGAI, cell: 'D4' },
  ctx_major:       { sheet: SHEET_IKIGAI, cell: 'D5' },
  ctx_semester:    { sheet: SHEET_IKIGAI, cell: 'D6' },
  ctx_linear:      { sheet: SHEET_IKIGAI, cell: 'D7' },
  ik_mbti:         { sheet: SHEET_IKIGAI, cell: 'C13' },
  ik_via1:         { sheet: SHEET_IKIGAI, cell: 'D14' },
  ik_via2:         { sheet: SHEET_IKIGAI, cell: 'E14' },
  ik_via3:         { sheet: SHEET_IKIGAI, cell: 'F14' },
  ik_career1:      { sheet: SHEET_IKIGAI, cell: 'G14' },
  ik_career2:      { sheet: SHEET_IKIGAI, cell: 'H14' },
  ik_career3:      { sheet: SHEET_IKIGAI, cell: 'I14' },
  ik_spot:         { sheet: SHEET_IKIGAI, cell: 'C23' },
  ik_slice:        { sheet: SHEET_IKIGAI, cell: 'C27' },
  ik_sweetspot:    { sheet: SHEET_IKIGAI, cell: 'C37' },
  ik_hard1:        { sheet: SHEET_IKIGAI, cell: 'D38' },
  ik_hard2:        { sheet: SHEET_IKIGAI, cell: 'E38' },
  ik_hard3:        { sheet: SHEET_IKIGAI, cell: 'F38' },
  ik_soft1:        { sheet: SHEET_IKIGAI, cell: 'G38' },
  ik_soft2:        { sheet: SHEET_IKIGAI, cell: 'H38' },
  ik_soft3:        { sheet: SHEET_IKIGAI, cell: 'I38' },
  sw_strength1:    { sheet: SHEET_SWOT, cell: 'D23' },
  sw_strength2:    { sheet: SHEET_SWOT, cell: 'D24' },
  sw_strength3:    { sheet: SHEET_SWOT, cell: 'D25' },
  sw_weakness1:    { sheet: SHEET_SWOT, cell: 'F23' },
  sw_weakness2:    { sheet: SHEET_SWOT, cell: 'F24' },
  sw_weakness3:    { sheet: SHEET_SWOT, cell: 'F25' },
  sw_opportunity1: { sheet: SHEET_SWOT, cell: 'H23' },
  sw_opportunity2: { sheet: SHEET_SWOT, cell: 'H24' },
  sw_opportunity3: { sheet: SHEET_SWOT, cell: 'H25' },
  sw_threat1:      { sheet: SHEET_SWOT, cell: 'J23' },
  sw_threat2:      { sheet: SHEET_SWOT, cell: 'J24' },
  sw_threat3:      { sheet: SHEET_SWOT, cell: 'J25' },
  ai_choice:       { sheet: SHEET_WELCOME, cell: 'C12' },
};

const AI_OPTIONS = [
  { key: 'ChatGPT', label: 'ChatGPT', base: 'https://chatgpt.com/?q=' },
  { key: 'Claude',  label: 'Claude',  base: 'https://claude.ai/new?q=' },
  { key: 'Gemini',  label: 'Gemini (AI Studio)', base: 'https://aistudio.google.com/app/prompts/new_chat?prompt=' },
  { key: 'Manus',   label: 'Manus',   base: 'https://manus.im/?q=' },
];

const MBTI_OPTIONS = Object.keys(MBTI_INFO);
const VIA_LABEL_ID = {
  creativity:'Kreativitas', curiosity:'Rasa Ingin Tahu', judgment:'Berpikir Kritis',
  loveoflearning:'Cinta Belajar', perspective:'Perspektif (Kebijaksanaan)', bravery:'Keberanian',
  perseverance:'Ketekunan', honesty:'Kejujuran', zest:'Semangat', love:'Cinta',
  kindness:'Kebaikan Hati', socialintelligence:'Kecerdasan Sosial', teamwork:'Kerja Sama Tim',
  fairness:'Keadilan', leadership:'Kepemimpinan', forgiveness:'Pemaaf', humility:'Kerendahan Hati',
  prudence:'Kehati-hatian', selfregulation:'Pengaturan Diri', appreciationofbeauty:'Apresiasi Keindahan',
  gratitude:'Rasa Syukur', hope:'Harapan', humor:'Humor', spirituality:'Spiritualitas',
};
const VIA_OPTIONS = Object.keys(VIA_INFO).map(k => ({ key: k, label: VIA_LABEL_ID[k] || k }));

function getSetupData_(userSpreadsheet) {
  if (!userSpreadsheet) return null;
  const ss = userSpreadsheet;
  const ik = ss.getSheetByName(SHEET_IKIGAI);
  const sw = ss.getSheetByName(SHEET_SWOT);
  const wc = ss.getSheetByName(SHEET_WELCOME);
  if (!ik) return null;

  const ikGrid = ik.getRange(1, 1, ik.getLastRow(), ik.getLastColumn()).getDisplayValues();
  const g = a1 => cellFromGrid_(ikGrid, a1);

  let swot = { prompt: '', strength: ['', '', ''], weakness: ['', '', ''], opportunity: ['', '', ''], threat: ['', '', ''] };
  if (sw) {
    const swGrid = sw.getRange(1, 1, sw.getLastRow(), sw.getLastColumn()).getDisplayValues();
    const gs = a1 => cellFromGrid_(swGrid, a1);
    swot = {
      prompt: gs('A17'),
      strength: [gs('D23'), gs('D24'), gs('D25')],
      weakness: [gs('F23'), gs('F24'), gs('F25')],
      opportunity: [gs('H23'), gs('H24'), gs('H25')],
      threat: [gs('J23'), gs('J24'), gs('J25')],
    };
  }

  const aiChoice = wc ? String(wc.getRange('C12').getDisplayValue() || '') : '';

  return {
    context: { name: g('D4'), major: g('D5'), semester: g('D6'), linear: g('D7') },
    semesterOptions: getValidationOptions_(ik, 6, 4, ['Semester 1','Semester 2','Semester 3','Semester 4','Semester 5','Semester 6','Semester 7','Semester 8']),
    linearOptions: getValidationOptions_(ik, 7, 4, ['YA','TIDAK']),
    ikigai: {
      mbti: g('C13'),
      via: [g('D14'), g('E14'), g('F14')],
      career: [g('G14'), g('H14'), g('I14')],
      prompt1: g('A17'),
      spot: g('C23'),
      slice: g('C27'),
      prompt2: g('A31'),
      sweetspot: g('C37'),
      hardskills: [g('D38'), g('E38'), g('F38')],
      softskills: [g('G38'), g('H38'), g('I38')],
    },
    swot: swot,
    aiChoice: aiChoice,
    aiOptions: AI_OPTIONS,
    mbtiOptions: MBTI_OPTIONS,
    viaOptions: VIA_OPTIONS,
  };
}
function getSetupBundle(userSpreadsheet) { return getSetupData_(userSpreadsheet); }

function saveSetupField(userSpreadsheet, fieldKey, value) {
  try {
    if (!userSpreadsheet) throw new Error('userSpreadsheet tidak ditemukan');
    const def = SETUP_FIELDS[fieldKey];
    if (!def) throw new Error('Field Setup tidak dikenal: ' + fieldKey);
    const sheet = userSpreadsheet.getSheetByName(def.sheet);
    if (!sheet) throw new Error('Sheet "' + def.sheet + '" tidak ditemukan.');
    sheet.getRange(def.cell).setValue(value);
    SpreadsheetApp.flush();
    return { success: true, setup: getSetupData_(userSpreadsheet) };
  } catch (e) {
    return { success: false, error: e.message };
  }
}