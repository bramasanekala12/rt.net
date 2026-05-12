function pancingIzin() {
  DriveApp.getRootFolder();
}

function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
        .setTitle('Sistem RT: Login & Daftar')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function registerUser(userData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Warga');
    if (!sheet) {
      sheet = ss.insertSheet('Warga');
      sheet.appendRow(['ID', 'Username', 'No_KK', 'Alamat', 'Role', 'Password']);
    }
    const data = sheet.getDataRange().getValues();
    if (!userData.username || !userData.password) return "Nama dan Password wajib diisi!";
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === userData.username) return "Username sudah digunakan!";
    }
    sheet.appendRow([data.length, userData.username, userData.nokk, userData.alamat, 'Warga', userData.password]);
    return "Berhasil";
  } catch (e) {
    return "Error Server: " + e.toString();
  }
}

function checkLogin(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Warga');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] == username && data[i][5] == password) {
        return { status: 'success', role: data[i][4], nama: data[i][1] };
      }
    }
    return { status: 'fail' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKeuangan = ss.getSheetByName('Keuangan');
    const data = sheetKeuangan.getDataRange().getValues();
    let totalSaldo = 0;
    for (let i = 1; i < data.length; i++) {
      let tipe = data[i][2];
      let jumlah = Number(data[i][4] || 0);
      if (tipe === "Masuk") totalSaldo += jumlah;
      else if (tipe === "Keluar") totalSaldo -= jumlah;
    }
    return { saldo: "Rp " + totalSaldo.toLocaleString('id-ID'), statusIuran: "Update Terkini" };
  } catch (e) {
    return { saldo: "Rp 0", statusIuran: "Error: " + e.message };
  }
}

// ======================================================
// HELPER: parse rawTgl → { tglObj, tglDisplay }
// Semua tanggal baru disimpan sebagai Date object.
// Data lama (string dd/MM/yyyy) juga tetap dihandle.
// ======================================================
function parseTanggal_(raw) {
  if (!raw) return { tglObj: null, tglDisplay: '-' };

  // Date object → format untuk display
  if (raw instanceof Date) {
    return {
      tglObj: raw,
      tglDisplay: Utilities.formatDate(raw, "GMT+7", "dd/MM/yyyy")
    };
  }

  const str = raw.toString().trim();

  // String "dd/MM/yyyy" — hati-hati: Sheets mungkin sudah salah parse data lama
  // Kita anggap format ini: bagian[0]=dd, bagian[1]=MM, bagian[2]=yyyy
  const s1 = str.split('/');
  if (s1.length === 3 && s1[2].length === 4) {
    return {
      tglObj: new Date(parseInt(s1[2]), parseInt(s1[1]) - 1, parseInt(s1[0])),
      tglDisplay: str
    };
  }

  return { tglObj: null, tglDisplay: str };
}

// ======================================================
// FUNGSI DEBUG
// ======================================================
function debugLaporan() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Keuangan');
  const data  = sheet.getDataRange().getValues();
  const now   = new Date();
  Logger.log('Bulan sekarang (1-based): ' + (now.getMonth() + 1) + ' | Tahun: ' + now.getFullYear());
  for (let i = 1; i < data.length; i++) {
    const raw = data[i][1];
    const { tglObj, tglDisplay } = parseTanggal_(raw);
    Logger.log('Baris ' + i
      + ' | raw=' + raw
      + ' | isDate=' + (raw instanceof Date)
      + ' | parsed=' + tglDisplay
      + ' | bulan=' + (tglObj ? tglObj.getMonth() + 1 : 'null')
      + ' | tahun=' + (tglObj ? tglObj.getFullYear() : 'null'));
  }
  const hasil = getLaporan('bulan_ini');
  Logger.log('Rows ditemukan: ' + hasil.rows.length);
  Logger.log('Masuk: ' + hasil.totalMasuk + ' | Keluar: ' + hasil.totalKeluar);
  if (hasil.error) Logger.log('ERROR: ' + hasil.error);
}

// ======================================================
// FUNGSI LAPORAN
// ======================================================
function getLaporan(filter) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Keuangan');
    const data  = sheet.getDataRange().getValues();

    let totalMasuk = 0, totalKeluar = 0, rows = [];

    const sekarang = new Date();
    const bulanIni = sekarang.getMonth() + 1; // 1-based
    const tahunIni = sekarang.getFullYear();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === '' && row[1] === '') continue;

      const { tglObj, tglDisplay } = parseTanggal_(row[1]);

      if (filter === 'bulan_ini') {
        if (!tglObj) continue;
        if ((tglObj.getMonth() + 1) !== bulanIni || tglObj.getFullYear() !== tahunIni) continue;
      } else if (filter === 'tahun_ini') {
        if (!tglObj) continue;
        if (tglObj.getFullYear() !== tahunIni) continue;
      }

      const tipe   = (row[2] || '').toString().trim();
      const jumlah = Number(row[4] || 0);

      if (tipe === 'Masuk')  totalMasuk  += jumlah;
      if (tipe === 'Keluar') totalKeluar += jumlah;

      rows.push({
        id        : row[0],
        tanggal   : tglDisplay,
        tipe      : tipe,
        kategori  : row[3] || '-',
        jumlah    : jumlah,
        keterangan: row[5] || '-',
        bukti     : row[6] || ''
      });
    }

    return { rows: rows, totalMasuk: totalMasuk, totalKeluar: totalKeluar, saldo: totalMasuk - totalKeluar };
  } catch (e) {
    return { rows: [], totalMasuk: 0, totalKeluar: 0, saldo: 0, error: e.toString() };
  }
}

// ======================================================
// FUNGSI KEUANGAN
// Tanggal disimpan sebagai Date object (bukan string)
// agar tidak salah parse oleh Sheets locale.
// Kolom tanggal di sheet harus diformat sebagai Date.
// ======================================================
function simpanKeuangan(form, linkFoto) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Keuangan');
    const data  = sheet.getDataRange().getValues();
    const newId = data.length;
    
    // Simpan sebagai Date object, bukan string → Sheets tidak salah baca
    const today = new Date();
    
    sheet.appendRow([newId, today, "Masuk", form.kategori, Number(form.jumlah), form.keterangan, linkFoto || ""]);
    
    // Format kolom tanggal (kolom B) baris baru sebagai dd/MM/yyyy
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat("dd/MM/yyyy");
    
    return "Berhasil Simpan!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function getChartData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Keuangan');
    if (!sheet) return { labels: [], masuk: [], keluar: [] };
    
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const result = {};
    
    // Inisialisasi 6 bulan terakhir
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      const key = Utilities.formatDate(d, "GMT+7", "MMM yyyy");
      result[key] = { masuk: 0, keluar: 0 };
    }
    
    const labels = Object.keys(result);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const { tglObj } = parseTanggal_(row[1]);
      if (!tglObj) continue;
      
      const key = Utilities.formatDate(tglObj, "GMT+7", "MMM yyyy");
      if (result[key]) {
        const tipe = (row[2] || '').toString().trim();
        const jumlah = Number(row[4] || 0);
        if (tipe === 'Masuk') result[key].masuk += jumlah;
        if (tipe === 'Keluar') result[key].keluar += jumlah;
      }
    }
    
    return {
      labels: labels,
      masuk: labels.map(l => result[l].masuk),
      keluar: labels.map(l => result[l].keluar)
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

function simpanPengeluaran(form, linkFoto) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Keuangan');
    const data  = sheet.getDataRange().getValues();
    const newId = data.length;

    const today = new Date();

    sheet.appendRow([newId, today, "Keluar", form.kategori, Number(form.jumlah), form.keterangan, linkFoto || ""]);

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 2).setNumberFormat("dd/MM/yyyy");

    return "Pengeluaran Berhasil Dicatat!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

// Upload helper generik
function uploadKeDrive_(folderId, data, fileName) {
  const folder      = DriveApp.getFolderById(folderId);
  const contentType = data.substring(5, data.indexOf(';'));
  const bytes       = Utilities.base64Decode(data.split(',')[1]);
  const blob        = Utilities.newBlob(bytes, contentType, fileName);
  const file        = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Bukti pemasukan / iuran
function uploadBuktIuran(data, fileName) {
  try { return uploadKeDrive_("1txfgqg20RJBbokYKHN6DcZ7eqlghnPFO", data, fileName); }
  catch (e) { return "Error Upload Iuran: " + e.toString(); }
}

// Bukti pengeluaran
function uploadBuktiPengeluaran(data, fileName) {
  try { return uploadKeDrive_("1Hl9H8jta4FtEON5T09oSJ_Goy-CZK1mE", data, fileName); }
  catch (e) { return "Error Upload Pengeluaran: " + e.toString(); }
}

// Foto inventaris
function uploadFotoInventaris(data, fileName) {
  try { return uploadKeDrive_("1bl4Md-u_nJrkHI4FUOjWOGltwuqs107J", data, fileName); }
  catch (e) { return "Error Upload Inventaris: " + e.toString(); }
}

// Tetap ada untuk kompatibilitas lama (tidak dipakai lagi tapi jaga-jaga)
function uploadFileKeDrive(data, fileName) {
  return uploadBuktIuran(data, fileName);
}

// ======================================================
// FUNGSI PERBAIKI DATA LAMA
// Jalankan SEKALI di editor untuk fix baris yang sudah terlanjur salah
// ======================================================
function perbaikiDataLama() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Keuangan');
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const raw = data[i][1];
    // Hanya proses yang sudah jadi Date object (salah parse oleh Sheets)
    if (raw instanceof Date) {
      // Sheets baca "11/05/2026" sebagai 5 Nov → kita balik: ambil hari & bulan, tukar
      const hari  = raw.getDate();    // ini sebenarnya bulan (karena terbalik)
      const bulan = raw.getMonth();   // ini sebenarnya hari - 1
      // Buat Date yang benar: hari=bulan, bulan=hari-1
      const tglBenar = new Date(raw.getFullYear(), hari - 1, bulan + 1);
      sheet.getRange(i + 1, 2).setValue(tglBenar).setNumberFormat("dd/MM/yyyy");
      Logger.log('Baris ' + (i+1) + ': ' + raw + ' → diperbaiki → ' + tglBenar);
    }
  }
  Logger.log('Selesai perbaiki data lama.');
}

// ======================================================
// FUNGSI INVENTARIS
// Sheet: Inventaris | Kolom: ID, Nama, Jumlah, Satuan, Kondisi, Lokasi, Keterangan, TglMasuk
// ======================================================

function initSheetInventaris_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Inventaris');
  if (!sheet) {
    sheet = ss.insertSheet('Inventaris');
    sheet.appendRow(['ID', 'Nama Barang', 'Jumlah', 'Satuan', 'Kondisi', 'Lokasi', 'Keterangan', 'Tgl Masuk', 'Foto']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f0f0f0');
  }
  return sheet;
}

function getInventaris() {
  try {
    const sheet = initSheetInventaris_();
    const data  = sheet.getDataRange().getValues();
    const rows  = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === '' && row[1] === '') continue;
      const rawTgl = row[7];
      let tglDisplay = '-';
      if (rawTgl instanceof Date) {
        tglDisplay = Utilities.formatDate(rawTgl, "GMT+7", "dd/MM/yyyy");
      } else if (rawTgl) {
        tglDisplay = rawTgl.toString();
      }
      rows.push({
        rowIndex  : i + 1,
        id        : row[0],
        nama      : row[1] || '-',
        jumlah    : Number(row[2] || 0),
        satuan    : row[3] || '-',
        kondisi   : row[4] || '-',
        lokasi    : row[5] || '-',
        keterangan: row[6] || '-',
        tglMasuk  : tglDisplay,
        foto      : row[8] || ''
      });
    }
    return { rows: rows };
  } catch (e) {
    return { rows: [], error: e.toString() };
  }
}

function tambahInventaris(form, linkFoto) {
  try {
    const sheet = initSheetInventaris_();
    const data  = sheet.getDataRange().getValues();
    const newId = data.length;
    const today = new Date();
    sheet.appendRow([
      newId, form.nama, Number(form.jumlah), form.satuan,
      form.kondisi, form.lokasi, form.keterangan, today, linkFoto || ""
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 8).setNumberFormat("dd/MM/yyyy");
    return "Barang berhasil ditambahkan!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function editInventaris(rowIndex, form) {
  try {
    const sheet = initSheetInventaris_();
    sheet.getRange(rowIndex, 2).setValue(form.nama);
    sheet.getRange(rowIndex, 3).setValue(Number(form.jumlah));
    sheet.getRange(rowIndex, 4).setValue(form.satuan);
    sheet.getRange(rowIndex, 5).setValue(form.kondisi);
    sheet.getRange(rowIndex, 6).setValue(form.lokasi);
    sheet.getRange(rowIndex, 7).setValue(form.keterangan);
    return "Barang berhasil diupdate!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function hapusInventaris(rowIndex) {
  try {
    const sheet = initSheetInventaris_();
    sheet.deleteRow(rowIndex);
    return "Barang berhasil dihapus!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function catatKeluarInventaris(rowIndex, jumlahKeluar, keteranganKeluar) {
  try {
    const sheet  = initSheetInventaris_();
    const stokLama = Number(sheet.getRange(rowIndex, 3).getValue());
    if (jumlahKeluar > stokLama) return "Stok tidak cukup! Stok saat ini: " + stokLama;
    sheet.getRange(rowIndex, 3).setValue(stokLama - jumlahKeluar);
    // Tambah catatan di keterangan
    const ketLama = sheet.getRange(rowIndex, 7).getValue();
    const today = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
    sheet.getRange(rowIndex, 7).setValue(ketLama + " | Keluar " + jumlahKeluar + " (" + today + "): " + keteranganKeluar);
    return "Barang keluar berhasil dicatat!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

// ======================================================
// PROFILE / WARGA
// Sheet: Warga | Kolom: ID, Username, No_KK, Alamat, Role, Password, Foto, Email, Telepon, AnggotaKeluarga
// ======================================================

function initSheetWarga_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Warga');
  if (!sheet) {
    sheet = ss.insertSheet('Warga');
    sheet.appendRow(['ID', 'Username', 'No_KK', 'Alamat', 'Role', 'Password', 'Foto', 'Email', 'Telepon', 'AnggotaKeluarga']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#f0f0f0');
  }
  return sheet;
}

function getProfileUser(username) {
  try {
    const sheet = initSheetWarga_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        return {
          username: data[i][1],
          nokk: data[i][2],
          alamat: data[i][3],
          role: data[i][4],
          foto: data[i][6] || "",
          email: data[i][7] || "",
          telepon: data[i][8] || "",
          anggota: data[i][9] ? JSON.parse(data[i][9]) : [],
          rowIndex: i + 1
        };
      }
    }
    return null;
  } catch (e) {
    return { error: e.toString() };
  }
}

function updateProfile(username, form) {
  try {
    const sheet = initSheetWarga_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        const rowIndex = i + 1;
        if (form.email) sheet.getRange(rowIndex, 8).setValue(form.email);
        if (form.telepon) sheet.getRange(rowIndex, 9).setValue(form.telepon);
        if (form.alamat) sheet.getRange(rowIndex, 4).setValue(form.alamat);
        if (form.anggota) sheet.getRange(rowIndex, 10).setValue(JSON.stringify(form.anggota));
        return "Profil berhasil diupdate!";
      }
    }
    return "User tidak ditemukan!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function updatePassword(username, oldPass, newPass) {
  try {
    const sheet = initSheetWarga_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        if (data[i][5] !== oldPass) return "Password lama salah!";
        sheet.getRange(i + 1, 6).setValue(newPass);
        return "Password berhasil diubah!";
      }
    }
    return "User tidak ditemukan!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}

function uploadFotoProfile(data, fileName) {
  try {
    return uploadKeDrive_("1txfgqg20RJBbokYKHN6DcZ7eqlghnPFO", data, fileName);
  } catch (e) {
    return "Error Upload Foto: " + e.toString();
  }
}

function updateFotoProfile(username, fotoLink) {
  try {
    const sheet = initSheetWarga_();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        sheet.getRange(i + 1, 7).setValue(fotoLink);
        return "Foto profil berhasil diupdate!";
      }
    }
    return "User tidak ditemukan!";
  } catch (e) {
    return "Gagal: " + e.toString();
  }
}
