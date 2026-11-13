// ==========================================
// BACKEND LOGIC: SIPAT-BTI (Google Apps Script)
// ==========================================

// 1. Fungsi Utama untuk Menyajikan WebApp
function doGet(e) {
  var html = HtmlService.createTemplateFromFile('index');
  return html.evaluate()
      .setTitle('SIPAT-BTI | Bengkel TI')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 2. FUNGSI PENGATURAN DATABASE (WAJIB DIJALANKAN PERTAMA KALI)
// Fungsi ini akan otomatis membuat Sheet dan Header yang dibutuhkan
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Setup Sheet: Users
  let userSheet = ss.getSheetByName("Users");
  if (!userSheet) {
    userSheet = ss.insertSheet("Users");
    userSheet.appendRow(["user_id", "nama", "username", "password", "role", "status"]);
    userSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#d9ead3");
    // Insert default admin
    userSheet.appendRow(["USR-001", "Administrator", "admin", "edudigital", "Admin PLP", "Aktif"]);
  }

  // Setup Sheet: Items (Inventaris)
  let itemSheet = ss.getSheetByName("Items");
  if (!itemSheet) {
    itemSheet = ss.insertSheet("Items");
    itemSheet.appendRow(["item_id", "kode_inventaris", "nama_alat", "kategori", "total_qty", "qty_tersedia", "kondisi"]);
    itemSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#c9daf8");
    // Insert dummy item
    itemSheet.appendRow(["INV-001", "AL-001", "Multimeter Digital Fluke", "Alat Ukur", 5, 3, "Tersedia"]);
  }

  // Setup Sheet: Transactions (Transaksi)
  let trxSheet = ss.getSheetByName("Transactions");
  if (!trxSheet) {
    trxSheet = ss.insertSheet("Transactions");
    trxSheet.appendRow(["trx_id", "user_id", "item_id", "jumlah_pinjam", "tgl_pinjam", "rencana_kembali", "tgl_kembali_aktual", "status"]);
    trxSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#fce5cd");
  }

  return "Setup Database Berhasil! Sheet Users, Items, dan Transactions telah dibuat.";
}

// 3. Fungsi Verifikasi Login dari Database
function verifyLoginServer(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName("Users");
  
  if(!userSheet) return { success: false, message: "Database belum disetup!" };
  
  const data = userSheet.getDataRange().getValues();
  
  // Looping data (mulai dari baris ke-2 / index 1 karena baris 1 adalah header)
  for (let i = 1; i < data.length; i++) {
    let dbUser = data[i][2]; // Kolom Username
    let dbPass = data[i][3]; // Kolom Password
    
    if (dbUser === username && dbPass === password) {
      return {
        success: true,
        user: {
          id: data[i][0],
          nama: data[i][1],
          role: data[i][4]
        }
      };
    }
  }
  
  return { success: false, message: "Username atau password salah!" };
}