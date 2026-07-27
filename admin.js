/**
 * Fungsi utama untuk mendaftarkan dan membuatkan file unik untuk user baru.
 * Dijalankan langsung dari editor Apps Script (sebagai admin).
 * 
 * @param {string} nama - Nama lengkap pembeli
 * @param {string} email - Email Google pembeli
 * @param {string} whatsapp - Nomor WhatsApp pembeli
 * @param {string} produk - Nama produk yang dibeli
 */
function createNewUser(nama, email, whatsapp, produk) {
  try {
    // 1. Masukkan ID Master Sheet dan ID Folder "users_sheets" Anda di sini
    var MASTER_SPREADSHEET_ID = "1ZKC4uK-CptizDSDZc2zOkbHmKag3psf-d8_HgAy-TvE";
    var USERS_FOLDER_ID = "1w8RvkKlOLEwJdhyGxYo7pzqGe_OpNHzv";
    
    var masterFile = DriveApp.getFileById(MASTER_SPREADSHEET_ID);
    var usersFolder = DriveApp.getFolderById(USERS_FOLDER_ID);
    
    // 2. Buat nama file unik untuk user baru di Google Drive
    var newFileName = "ElevatEd Student Kit - " + nama;
    
    // 3. Salin (Make a copy) dari master sheet ke folder users_sheets
    var copiedFile = masterFile.makeCopy(newFileName, usersFolder);
    var newUserSpreadsheetId = copiedFile.getId();
    
    // 4. Berikan akses (share) otomatis ke email user tersebut sebagai Editor
    copiedFile.addEditor(email);
    
    // 5. Generate Kode Aktivasi Unik secara otomatis (Format: ESDK-XXXX-XXXX)
    var activationCode = generateUniqueActivationCode_();
    
    // 6. Catat data lengkap ke database activations (sesuai struktur kolom Anda)
    saveToActivationsDatabase_(nama, email, whatsapp, produk || "ElevatEd Student Kit", newUserSpreadsheetId, activationCode);
    
    Logger.log("=== BERHASIL MEMBUAT USER BARU ===");
    Logger.log("Nama: " + nama);
    Logger.log("Email: " + email);
    Logger.log("Spreadsheet URL: " + copiedFile.getUrl());
    Logger.log("Kode Aktivasi: " + activationCode);
    
    return {
      success: true,
      nama: nama,
      email: email,
      url: copiedFile.getUrl(),
      code: activationCode
    };
    
  } catch (error) {
    Logger.log("Error saat membuat user: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Helper untuk membuat kode aktivasi acak yang unik
 */
function generateUniqueActivationCode_() {
  var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  var code = "ESDK";
  for (var i = 0; i < 2; i++) {
    code += "-";
    for (var j = 0; j < 4; j++) {
      code += chars.charAt(Math.floor(Math.random() * chars.length));
    }
  }
  return code; 
}

/**
 * Helper untuk menyimpan data ke database activations pusat 
 * Berdasarkan kolom: No, Timestamp, Nama, Email, Whatsapp, Produk, Code, Status
 */
function saveToActivationsDatabase_(nama, email, whatsapp, produk, spreadsheetId, code) {
  // Masukkan ID Spreadsheet activations Anda di sini
  var activationsSpreadsheetId = "1OkbFWqAdQmO_R2aoyH2i9cQ6QOf3t4HmcEnis021BWc";
  var sheet = SpreadsheetApp.openById(activationsSpreadsheetId).getSheetByName("Activations");
  
  if (!sheet) {
    throw new Error("Sheet dengan nama 'Activations' tidak ditemukan di database pusat.");
  }
  
  // Hitung nomor urut otomatis berdasarkan baris terakhir yang terisi
  var lastRow = sheet.getLastRow();
  var nextNo = lastRow === 0 ? 1 : lastRow; // Jika baris 1 kosong/header, disesuaikan
  
  var timestamp = new Date();
  var status = "aktif"; // Default status awal saat dibuat
  
  // Urutan kolom: 
  // [1] No, [2] Timestamp, [3] Nama, [4] Email, [5] Whatsapp, [6] Produk, [7] Code, [8] Status, [9] Spreadsheet ID
  sheet.appendRow([
    nextNo,
    timestamp,
    nama,
    email,
    whatsapp,
    produk,
    code,
    status,
    spreadsheetId
  ]);
}

/**
 * Fungsi uji coba (Test Run) untuk dijalankan langsung dari editor malam ini.
 * Cukup isi data dummy di bawah lalu klik Run.
 */
function testCreateUser() {
  var hasil = createNewUser(
    "Budi Santoso", 
    "budi.contoh@gmail.com", 
    "081234567890", 
    "ElevatEd Student Kit"
  );
  Logger.log(hasil);
}