/**
 * Menangani permintaan masuk dari website Vercel via HTTP POST
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var token = data.token; // Token aktivasi unik user
    
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
    if (action === "getData") {
      result = handleGetData_(userSpreadsheet);
    } else if (action === "updateData") {
      result = handleUpdateData_(userSpreadsheet, data.payload);
    } else {
      result = { success: false, message: "Aksi tidak dikenal." };
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