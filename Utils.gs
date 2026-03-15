/**
 * UTILS (Utility / Helper Functions)
 * File ini khusus menampung fungsi-fungsi bantuan agar Code.gs tetap bersih.
 */

const Utils = {
  // Format tanggal ke dd/MM/yyyy
  formatDate: function(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
  },
  
  // Format tanggal dan jam ke dd/MM/yyyy HH:mm
  formatDateTime: function(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  },
  
  // Generate ID Surat Masuk (Contoh: SM-0001)
  generateIdMasuk: function(sheet) {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const idUrut = lastRow.toString().padStart(4, '0');
    return "SM-" + idUrut;
  },
  
  // Generate Nomor Surat Keluar sesuai format instansi
  generateNomorKeluar: function(sheet, settings, kodeBidang) {
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const noUrutStr = lastRow.toString().padStart(3, '0');
    return `${settings.kodeSurat}/${noUrutStr}/${kodeBidang}/${settings.bulanRomawi}/${settings.tahun}`;
  },
  
  // Mengambil data pengaturan dari Sheet 'Pengaturan'
  getPengaturan: function() {
    // Memanggil active spreadsheet secara independen agar lebih aman
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengaturan');
    if (!sheet) throw new Error("Sheet Pengaturan tidak ditemukan!");
    
    const data = sheet.getDataRange().getValues();
    let config = { kodeSurat: "000", bulanRomawi: "I", tahun: new Date().getFullYear().toString() };
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'Kode Klasifikasi Surat') config.kodeSurat = data[i][1];
      if (data[i][0] === 'Bulan Berjalan') config.bulanRomawi = data[i][1];
      if (data[i][0] === 'Tahun Berjalan') config.tahun = data[i][1];
    }
    return config;
  }
};
