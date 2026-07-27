/**
 * Menangani permintaan masuk dari website Vercel via HTTP POST
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var token = data.token || data.code; // Token aktivasi unik user
    var payload = data.payload || {};
    
    // 0. Penanganan khusus untuk validasi lisensi dari Rumah B / Dashboard
    if (action === "validate") {
      var validationResult = handleValidateAction_(data);
      return ContentService.createTextOutput(JSON.stringify(validationResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 1. Validasi Token dan Ambil Spreadsheet ID milik user dari database 'activations'
    var userInfo = getUserInfoByToken_(token);
    if (!userInfo) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, 
        message: "Token aktivasi tidak valid atau tidak ditemukan."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var userSpreadsheetId = userInfo.spreadsheetId;
    if (!userSpreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, 
        message: "Spreadsheet ID pengguna belum terdaftar di kolom database (Kolom I)."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var userSpreadsheet = SpreadsheetApp.openById(userSpreadsheetId);
    
    // 2. Routing aksi berdasarkan permintaan dari Vercel
    var result = {};
    switch (action) {
      case "getBundle":
        result = getBundle(userSpreadsheet);
        break;
      case "getInitialBundle":
        result = getInitialBundle(userSpreadsheet);
        break;
      case "getActionPlanOnly":
        result = getActionPlanOnly(userSpreadsheet);
        break;
      case "getDreamPlan":
        result = getDreamPlan_(userSpreadsheet);
        break;
      case "updateDreamStatus":
        result = updateDreamStatus(userSpreadsheet, payload.row, payload.status);
        break;
      case "getCollegePlan":
        result = getCollegePlan_(userSpreadsheet);
        break;
      case "updateCollegeProgress":
        result = updateCollegeProgress(userSpreadsheet, payload.row, payload.status);
        break;
      case "updateActionStatus":
        result = updateActionStatus(userSpreadsheet, payload.row, payload.col, payload.status);
        break;
      case "getActionPlan":
        result = getActionPlan_(userSpreadsheet);
        break;
      case "getTrackerView":
        result = getTrackerView(userSpreadsheet, payload.quarter, payload.cachedBlocks);
        break;
      case "toggleTrackerCheck2":
        result = toggleTrackerCheck2(userSpreadsheet, payload.quarter, payload.goalIdx, payload.tacticNo, payload.week, payload.dayIndex, payload.value, payload.target, payload.periodWeeks);
        break;
      case "checkAllTrackerWeek":
        result = checkAllTrackerWeek(userSpreadsheet, payload.quarter, payload.goalIdx, payload.tacticNo, payload.week, payload.target, payload.periodWeeks);
        break;
      case "setTrackerStartDate":
        result = setTrackerStartDate(payload.dateStr);
        break;
      case "getGamification":
        result = getGamificationOnly(userSpreadsheet);
        break;
      case "getReportBundle":
        result = getReportBundle(userSpreadsheet);
        break;
      case "getIdentityBundle":
        result = getIdentityBundle(userSpreadsheet);
        break;
      case "getSetupBundle":
        result = getSetupBundle(userSpreadsheet);
        break;
      case "saveSetupField":
        result = saveSetupField(userSpreadsheet, payload.fieldKey, payload.value);
        break;
      case "verifyWebAppUrl":
        result = verifyWebAppUrl();
        break;
      case "activateWithToken":
        result = activateWithToken(userSpreadsheet, payload.inputToken);
        break;
      case "getActivationStatus":
        result = getActivationStatus(userSpreadsheet);
        break;
      // Legacy/Custom Handlers
      case "getData":
        result = handleGetData_(userSpreadsheet);
        break;
      case "updateData":
        result = handleUpdateData_(userSpreadsheet, payload);
        break;
      default:
        result = { success: false, message: "Aksi '" + action + "' tidak dikenal." };
        break;
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false, 
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper: Penanganan aksi 'validate' lisensi dari Rumah B / Dashboard
 */
function handleValidateAction_(data) {
  var secret = data.secret;
  var email = data.email;
  var code = data.code || data.token;

  var expectedSecret = PropertiesService.getScriptProperties().getProperty('VALIDATE_SECRET') || "ikg_valid_z4Tn9wRfB7cJ";
  if (secret !== expectedSecret) {
    return { valid: false, reason: "Secret key tidak cocok." };
  }

  var activationsSpreadsheetId = "MASUKKAN_ACTIVATIONS_SPREADSHEET_ID_DI_SINI";
  var sheet = SpreadsheetApp.openById(activationsSpreadsheetId).getSheetByName("Activations");
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    var rowCode = String(rows[i][6] || '').trim(); // Indeks 6 (Kolom G - Code)
    var rowEmail = String(rows[i][3] || '').trim(); // Indeks 3 (Kolom D - Email)
    var rowStatus = String(rows[i][7] || '').trim().toLowerCase(); // Indeks 7 (Kolom H - Status)

    if (rowCode === code) {
      if (rowStatus !== 'aktif') {
        return { valid: false, reason: 'Lisensi/token ini sudah tidak aktif.' };
      }
      if (rowEmail.toLowerCase() !== email.toLowerCase()) {
        return { valid: false, reason: 'Token ini tidak terdaftar untuk email Google ' + email };
      }
      return { valid: true, produk: rows[i][5] || '' };
    }
  }
  return { valid: false, reason: 'Token aktivasi tidak ditemukan.' };
}

/**
 * Helper: Cari data user di sheet 'activations' berdasarkan token
 */
function getUserInfoByToken_(token) {
  var activationsSpreadsheetId = "MASUKKAN_ACTIVATIONS_SPREADSHEET_ID_DI_SINI";
  var sheet = SpreadsheetApp.openById(activationsSpreadsheetId).getSheetByName("Activations");
  var rows = sheet.getDataRange().getValues();
  
  // Asumsi kolom: [0] No, [1] Timestamp, [2] Nama, [3] Email, [4] Whatsapp, [5] Produk, [6] Code, [7] Status
  for (var i = 1; i < rows.length; i++) {
    var rowCode = String(rows[i][6] || '').trim(); // Kolom Code (indeks 6 / Kolom G)
    if (rowCode === token && String(rows[i][7] || '').trim().toLowerCase() === "aktif") {
      return {
        email: rows[i][3],
        spreadsheetId: rows[i][8] // Kolom tempat Anda menyimpan Spreadsheet ID user (Indeks 8 / Kolom I)
      };
    }
  }
  return null;
}

/**
 * Helper: Ambil data untuk dikirim ke Sidebar Vercel
 */
function handleGetData_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("DREAM PLAN");
  var data = sheet.getDataRange().getValues();
  return { success: true, data: data };
}

/**
 * Helper: Update data dari Sidebar Vercel ke Spreadsheet User
 */
function handleUpdateData_(spreadsheet, payload) {
  var sheet = spreadsheet.getSheetByName(payload.sheetName);
  if (!sheet) {
    return { success: false, message: "Sheet target tidak ditemukan." };
  }
  
  sheet.getRange(payload.range).setValue(payload.value);
  return { success: true, message: "Data berhasil diperbarui." };
}