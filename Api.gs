/**
 * Menangani permintaan masuk dari website Vercel via HTTP POST
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var token = data.token; // Token aktivasi unik user
    var payload = data.payload || {};
    
    // 1. Validasi Token dan Ambil Spreadsheet ID milik user dari database 'activations'
    var userInfo = getUserInfoByToken_(token);
    if (!userInfo) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, 
        message: "Token aktivasi tidak valid atau tidak ditemukan."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var userSpreadsheetId = userInfo.spreadsheetId;
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
 * Helper: Cari data user di sheet 'activations' berdasarkan token
 */
function getUserInfoByToken_(token) {
  var activationsSpreadsheetId = "MASUKKAN_ACTIVATIONS_SPREADSHEET_ID_DI_SINI";
  var sheet = SpreadsheetApp.openById(activationsSpreadsheetId).getSheetByName("Activations");
  var rows = sheet.getDataRange().getValues();
  
  // Asumsi kolom: [1] No, [2] Timestamp, [3] Nama, [4] Email, [5] Whatsapp, [6] Produk, [7] Code, [8] Status
  // Catatan: Pastikan Spreadsheet ID disimpan di kolom tersembunyi atau kita petakan dengan benar.
  // Sebagai contoh sederhana, mari kita asumsikan Spreadsheet ID disimpan di kolom ke-9 (kolom I).
  for (var i = 1; i < rows.length; i++) {
    var rowCode = rows[i][6]; // Kolom Code (indeks 6)
    if (rowCode === token && rows[i][7].toLowerCase() === "aktif") {
      return {
        email: rows[i][3],
        spreadsheetId: rows[i][8] // Kolom tempat Anda menyimpan Spreadsheet ID user
      };
    }
  }
  return null;
}

/**
 * Helper: Ambil data untuk dikirim ke Sidebar Vercel
 */
function handleGetData_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("DREAM PLAN"); // Contoh sheet
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
  
  // Contoh menulis data ke cell tertentu
  sheet.getRange(payload.range).setValue(payload.value);
  return { success: true, message: "Data berhasil diperbarui." };
}