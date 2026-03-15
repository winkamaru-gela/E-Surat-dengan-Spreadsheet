/**
 * SISTEM E-SURAT & JDIH ASN - MAIN CONTROLLER
 * Menangani Routing, UI, Auth, dan API Backend
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

// ==========================================
// 1. ROUTING & UI INITIATION
// ==========================================

function doGet(e) {
  checkAndSetupDatabase(); 
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Sistem E-Surat & JDIH')
      .setFaviconUrl('https://img.icons8.com/color/48/000000/document.png')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Menu E-Surat')
    .addItem('💻 Buka Aplikasi E-Surat', 'bukaAplikasi')
    .addToUi();
}

function bukaAplikasi() {
  checkAndSetupDatabase(); 
  const html = HtmlService.createTemplateFromFile('Index').evaluate()
      .setWidth(1000)
      .setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aplikasi E-Surat Terpadu');
}

// ==========================================
// 2. AUTHENTICATION & USERS API
// ==========================================

function apiLogin(username, password) {
  const sheet = SS.getSheetByName('Pengguna');
  if (!sheet) return { success: false, message: "Sistem sedang disiapkan, harap muat ulang aplikasi." };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password) {
      return { success: true, role: data[i][2], nama: data[i][3] }; 
    }
  }
  return { success: false, message: "Username atau password salah!" };
}

function apiGetDaftarUser() {
  const sheet = SS.getSheetByName('Pengguna');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  
  let result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({ username: data[i][0], role: data[i][2], nama: data[i][3] });
  }
  return result;
}

function apiTambahUser(data) {
  try {
    const sheet = SS.getSheetByName('Pengguna');
    const existing = sheet.getDataRange().getValues();
    
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][0] === data.username) {
        return { success: false, message: "Username sudah digunakan!" };
      }
    }
    
    sheet.appendRow([data.username, data.password, data.role, data.nama]);
    return { success: true, message: "Pengguna baru berhasil ditambahkan." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function apiUbahPassword(username, newPassword) {
  try {
    const sheet = SS.getSheetByName('Pengguna');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username) {
        sheet.getRange(i + 1, 2).setValue(newPassword);
        return { success: true, message: "Password berhasil diperbarui." };
      }
    }
    return { success: false, message: "Username tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 3. E-SURAT API (MASUK & KELUAR)
// ==========================================

function apiSimpanSuratMasuk(data) {
  try {
    const sheet = SS.getSheetByName('SuratMasuk');
    const idSurat = Utils.generateIdMasuk(sheet);
    const tglTerima = Utils.formatDateTime(new Date());
    
    sheet.appendRow([
      idSurat, tglTerima, data.noSurat, data.asalSurat, 
      data.perihal, data.disposisi, data.linkDoc
    ]);
    return { success: true, message: "Berhasil disimpan!", id: idSurat };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function apiSimpanSuratKeluar(data) {
  try {
    const sheet = SS.getSheetByName('SuratKeluar');
    const tglKeluar = Utils.formatDate(new Date());
    const noUrutStr = Math.max(sheet.getLastRow(), 1).toString().padStart(3, '0');
    
    sheet.appendRow([
      noUrutStr, tglKeluar, data.noSurat, data.tujuan, 
      data.perihal, data.ttd, data.linkDoc
    ]);
    return { success: true, message: "Berhasil disimpan!", nomor: data.noSurat };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function apiUpdateSurat(data) {
  try {
    const sheetName = data.jenis === 'Masuk' ? 'SuratMasuk' : 'SuratKeluar';
    const sheet = SS.getSheetByName(sheetName);
    const barisTarget = parseInt(data.rowId);

    sheet.getRange(barisTarget, 3).setValue(data.noSurat);
    sheet.getRange(barisTarget, 5).setValue(data.perihal);
    sheet.getRange(barisTarget, 7).setValue(data.linkDoc);
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function apiGetDaftarSuratMasuk() {
  const sheet = SS.getSheetByName('SuratMasuk');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  
  let result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      rowId: i + 1, idSurat: data[i][0], tglTerima: data[i][1], noSurat: data[i][2],
      asalSurat: data[i][3], perihal: data[i][4], disposisi: data[i][5], linkDoc: data[i][6]
    });
  }
  return result.reverse();
}

function apiGetDaftarSuratKeluar() {
  const sheet = SS.getSheetByName('SuratKeluar');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  
  let result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      rowId: i + 1, noUrut: data[i][0], tglSurat: data[i][1], noSuratKeluar: data[i][2],
      tujuan: data[i][3], perihal: data[i][4], penandatangan: data[i][5], linkDoc: data[i][6]
    });
  }
  return result.reverse(); 
}

function apiGetDashboardStatistik() {
  const sheetSM = SS.getSheetByName('SuratMasuk');
  const sheetSK = SS.getSheetByName('SuratKeluar');
  return {
    jmlMasuk: sheetSM ? Math.max(0, sheetSM.getLastRow() - 1) : 0,
    jmlKeluar: sheetSK ? Math.max(0, sheetSK.getLastRow() - 1) : 0
  };
}

// ==========================================
// 4. JDIH / PERATURAN API
// ==========================================

function apiSimpanPeraturan(data) {
  try {
    const sheet = SS.getSheetByName('Peraturan');
    if (!sheet) throw new Error("Sheet 'Peraturan' belum dibuat.");
    
    const lastRow = Math.max(sheet.getLastRow(), 1);
    const idPeraturan = "PRT-" + lastRow.toString().padStart(3, '0');
    
    sheet.appendRow([
      idPeraturan, data.kategori, data.nomor, data.tahun, 
      data.tentang, data.linkPdf
    ]);
    
    return { success: true, message: "Peraturan berhasil dipublikasikan!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function apiGetDaftarPeraturan() {
  const sheet = SS.getSheetByName('Peraturan');
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  
  let result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      rowId: i + 1,
      id: data[i][0], kategori: data[i][1], nomor: data[i][2],
      tahun: data[i][3], tentang: data[i][4], linkPdf: data[i][5]
    });
  }
  return result.reverse();
}

function apiUpdatePeraturan(data) {
  try {
    const sheet = SS.getSheetByName('Peraturan');
    if (!sheet) throw new Error("Sheet tidak ditemukan.");
    
    const barisTarget = parseInt(data.rowId);

    // Update kolom B (Kategori), C (Nomor), D (Tahun), E (Tentang), F (Link PDF)
    sheet.getRange(barisTarget, 2).setValue(data.kategori);
    sheet.getRange(barisTarget, 3).setValue(data.nomor);
    sheet.getRange(barisTarget, 4).setValue(data.tahun);
    sheet.getRange(barisTarget, 5).setValue(data.tentang);
    sheet.getRange(barisTarget, 6).setValue(data.linkPdf);
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 5. API HAPUS DATA (Dinamis)
// ==========================================

function apiHapusData(jenis, rowId) {
  try {
    let sheetName = '';
    if (jenis === 'Masuk') sheetName = 'SuratMasuk';
    else if (jenis === 'Keluar') sheetName = 'SuratKeluar';
    else if (jenis === 'Peraturan') sheetName = 'Peraturan';
    
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) throw new Error("Sheet tidak ditemukan.");
    
    sheet.deleteRow(parseInt(rowId));
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ==========================================
// 6. DATABASE SETUP (OTOMATIS)
// ==========================================

function checkAndSetupDatabase() {
  if (!SS.getSheetByName('SuratMasuk')) buatSheet('SuratMasuk', ['ID Surat', 'Tanggal Terima', 'Nomor Surat Asal', 'Asal Surat', 'Perihal', 'Disposisi Ke', 'Link Dokumen/Arsip']);
  if (!SS.getSheetByName('SuratKeluar')) buatSheet('SuratKeluar', ['Nomor Urut', 'Tanggal Surat', 'Nomor Surat Keluar', 'Tujuan', 'Perihal', 'Penandatangan', 'Link Dokumen/Draft']);
  if (!SS.getSheetByName('Peraturan')) buatSheet('Peraturan', ['ID', 'Kategori', 'Nomor', 'Tahun', 'Tentang', 'Link PDF']);
  
  if (!SS.getSheetByName('Pengguna')) {
    buatSheet('Pengguna', ['Username', 'Password', 'Role', 'Nama Lengkap']);
    SS.getSheetByName('Pengguna').appendRow(['admin', 'asn123', 'Admin', 'Administrator Sistem']);
  }

  if (!SS.getSheetByName('Pengaturan')) {
    buatSheet('Pengaturan', ['Parameter', 'Nilai', 'Keterangan']);
    const sheetPengaturan = SS.getSheetByName('Pengaturan');
    if (sheetPengaturan.getLastRow() <= 1) {
      sheetPengaturan.appendRow(['Kode Klasifikasi Surat', '005', 'Contoh: 005 untuk Undangan']);
      sheetPengaturan.appendRow(['Bulan Berjalan', 'III', 'Contoh: III (Romawi)']);
      sheetPengaturan.appendRow(['Tahun Berjalan', new Date().getFullYear().toString(), 'Tahun berjalan saat ini']);
    }
  }
}

function buatSheet(namaSheet, headerArray) {
  let sheet = SS.insertSheet(namaSheet);
  const range = sheet.getRange(1, 1, 1, headerArray.length);
  range.setValues([headerArray]);
  range.setFontWeight('bold').setBackground('#d9ead3');
  sheet.setFrozenRows(1);
  for (let i = 1; i <= headerArray.length; i++) sheet.autoResizeColumn(i);
}
