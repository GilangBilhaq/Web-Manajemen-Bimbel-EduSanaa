// ============================================================
// EduSanaa Bimbel - Google Apps Script Backend
// ============================================================

var SS_ID = ''; // Kosongkan jika menggunakan Bound Script (dari Spreadsheet langsung)
var ADMIN_PASSCODE = '123456'; // Ganti dengan passcode Anda

function getSpreadsheet() {
  if (SS_ID) return SpreadsheetApp.openById(SS_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================================
// ROUTING
// ============================================================
function doGet(e) {
  var view = e.parameter.view || 'tutor';
  var template;
  if (view === 'admin') {
    template = HtmlService.createTemplateFromFile('Admin');
  } else {
    template = HtmlService.createTemplateFromFile('Tutor');
  }
  template.scriptUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle(view === 'admin' ? 'EduSanaa - Admin Portal' : 'EduSanaa - Portal Tutor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// DATABASE SETUP
// ============================================================
function setupDatabase() {
  var ss = getSpreadsheet();
  var sheetsConfig = {
    'Siswa': ['ID_Siswa', 'Nama_Siswa', 'NISN', 'Program', 'Kelas', 'Nama_Ortu', 'No_WA_Ortu', 'Alamat', 'Status', 'Tgl_Daftar'],
    'Tutor': ['ID_Tutor', 'Nama_Tutor', 'PIN', 'No_WA', 'Mapel_Utama', 'Tarif_Per_Sesi', 'Status', 'Tgl_Bergabung'],
    'Jadwal': ['ID_Jadwal', 'Hari', 'Jam_Mulai', 'Jam_Selesai', 'ID_Siswa', 'ID_Tutor', 'Mata_Pelajaran', 'Ruang', 'Status', 'Tanggal'],
    'Absensi': ['ID_Absen', 'Timestamp', 'ID_Jadwal', 'ID_Tutor', 'Nama_Tutor', 'Nama_Siswa', 'Mata_Pelajaran', 'Laporan', 'URL_Foto', 'Status_Verifikasi', 'Tanggal_Sesi'],
    'Keuangan': ['ID_Trx', 'Timestamp', 'Tipe', 'Kategori', 'Deskripsi', 'Nominal', 'Metode', 'ID_Siswa', 'ID_Tutor', 'Status'],
    'Pengaturan': ['Key', 'Value']
  };

  for (var sheetName in sheetsConfig) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = sheetsConfig[sheetName];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#F7C8D3')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
  }

  // Migrasi: pastikan kolom Tanggal_Sesi ada di sheet Absensi lama
  ensureColumn('Absensi', 'Tanggal_Sesi');
  // Migrasi: kolom Tanggal di sheet Jadwal (jadwal per tanggal, bukan hanya hari)
  ensureColumn('Jadwal', 'Tanggal');

  // Seed Pengaturan jika kosong
  var setPengaturan = ss.getSheetByName('Pengaturan');
  if (setPengaturan && setPengaturan.getLastRow() <= 1) {
    setPengaturan.appendRow(['nama_lembaga', 'EduSanaa']);
    setPengaturan.appendRow(['alamat', 'Jl. Contoh No. 1, Jakarta']);
    setPengaturan.appendRow(['no_wa_admin', '628123456789']);
    setPengaturan.appendRow(['admin_passcode', ADMIN_PASSCODE]);
    setPengaturan.appendRow(['rekening', 'Bank BCA - 1234567890 a.n. EduSanaa']);
  }

  return { success: true, message: 'Database Setup Selesai! Semua sheet telah dibuat.' };
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================
function getSheetData(sheetName) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(header, index) {
      var val = row[index];
      if (val instanceof Date) {
        var yr = val.getFullYear();
        // Jika tahun 1899 atau 1900 -> nilai murni waktu (time-only) dari Sheets
        if (yr <= 1900) {
          var hh = String(val.getHours()).padStart(2, '0');
          var mm = String(val.getMinutes()).padStart(2, '0');
          obj[header] = hh + ':' + mm;
        } else {
          obj[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
        }
      } else {
        obj[header] = val;
      }
    });
    return obj;
  });
}

function appendToSheet(sheetName, rowData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.appendRow(rowData);
    return true;
  }
  return false;
}

function updateSheetRow(sheetName, idColumn, idValue, updateData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idColIndex = headers.indexOf(idColumn);
  if (idColIndex === -1) return false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex]) === String(idValue)) {
      for (var key in updateData) {
        var colIndex = headers.indexOf(key);
        if (colIndex !== -1) {
          sheet.getRange(i + 1, colIndex + 1).setValue(updateData[key]);
        }
      }
      return true;
    }
  }
  return false;
}

function generateId(prefix) {
  return prefix + '_' + new Date().getTime();
}

// Tambahkan kolom di akhir header jika belum ada (migrasi aman untuk sheet lama)
function ensureColumn(sheetName, colName) {
  var sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return;
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf(colName) === -1) {
    var cell = sheet.getRange(1, lastCol + 1);
    cell.setValue(colName).setFontWeight('bold').setBackground('#F7C8D3').setFontColor('#FFFFFF');
  }
}

// Tulis baris berdasarkan nama header, format TEXT agar tanggal tidak di-parse ulang oleh Sheets
function appendRowByHeaders(sheetName, obj) {
  var sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return false;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function(h) { return obj.hasOwnProperty(h) ? obj[h] : ''; });
  var range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length);
  range.setNumberFormat('@');
  range.setValues([row]);
  return true;
}

function pad2(n) { return ('0' + n).slice(-2); }

// Normalisasi berbagai format tanggal ke 'yyyy-MM-dd'
function toDateKey(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(val).trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + pad2(m[2]) + '-' + pad2(m[1]);
  return '';
}

// Kunci tanggal sesi suatu absensi (pakai Tanggal_Sesi, fallback ke Timestamp untuk data lama)
function getSesiKey(a) {
  return toDateKey(a.Tanggal_Sesi) || toDateKey(a.Timestamp);
}

function getOrCreateFolder(name) {
  var folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function getTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

// ============================================================
// AUTH
// ============================================================
function verifyAdminLogin(passcode) {
  var ss = getSpreadsheet();
  var pengaturan = ss.getSheetByName('Pengaturan');
  var storedPasscode = ADMIN_PASSCODE;
  if (pengaturan) {
    var data = pengaturan.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === 'admin_passcode') {
        storedPasscode = String(data[i][1]);
        break;
      }
    }
  }
  if (String(passcode) === storedPasscode) {
    return { success: true };
  }
  return { success: false, message: 'Passcode salah. Coba lagi.' };
}

// ============================================================
// DASHBOARD ADMIN
// ============================================================
function getDashboardData() {
  return {
    siswa: getSheetData('Siswa'),
    tutor: getSheetData('Tutor'),
    jadwal: getSheetData('Jadwal').map(function(j) { j.Tanggal = toDateKey(j.Tanggal); return j; }),
    absensi: getSheetData('Absensi').map(function(a) { a.Tanggal_Sesi = getSesiKey(a); return a; }),
    keuangan: getSheetData('Keuangan'),
    pengaturan: getSheetData('Pengaturan')
  };
}

// ============================================================
// SISWA CRUD
// ============================================================
function tambahSiswa(data) {
  var id = generateId('SSW');
  appendToSheet('Siswa', [
    id, data.nama, data.nisn || '', data.program || '', data.kelas || '',
    data.nama_ortu, data.no_wa, data.alamat || '', 'Aktif', getTimestamp()
  ]);
  return { success: true, message: 'Data siswa berhasil disimpan.', id: id };
}

function editSiswa(data) {
  var result = updateSheetRow('Siswa', 'ID_Siswa', data.id, {
    'Nama_Siswa': data.nama,
    'NISN': data.nisn || '',
    'Program': data.program || '',
    'Kelas': data.kelas || '',
    'Nama_Ortu': data.nama_ortu,
    'No_WA_Ortu': data.no_wa,
    'Alamat': data.alamat || '',
    'Status': data.status || 'Aktif'
  });
  return { success: result, message: result ? 'Data siswa diperbarui.' : 'Siswa tidak ditemukan.' };
}

function hapusSiswa(idSiswa) {
  return hapusBaris('Siswa', 'ID_Siswa', idSiswa);
}

// ============================================================
// TUTOR CRUD
// ============================================================
function tambahTutor(data) {
  var id = generateId('TUT');
  appendToSheet('Tutor', [
    id, data.nama, data.pin, data.no_wa,
    data.mapel || '', data.tarif || 0, 'Aktif', getTimestamp()
  ]);
  return { success: true, message: 'Data tutor berhasil disimpan.', id: id };
}

function editTutor(data) {
  var result = updateSheetRow('Tutor', 'ID_Tutor', data.id, {
    'Nama_Tutor': data.nama,
    'PIN': data.pin,
    'No_WA': data.no_wa,
    'Mapel_Utama': data.mapel || '',
    'Tarif_Per_Sesi': data.tarif || 0,
    'Status': data.status || 'Aktif'
  });
  return { success: result, message: result ? 'Data tutor diperbarui.' : 'Tutor tidak ditemukan.' };
}

function hapusTutor(idTutor) {
  return hapusBaris('Tutor', 'ID_Tutor', idTutor);
}

// ============================================================
// JADWAL CRUD
// ============================================================
function tambahJadwal(data) {
  var id = generateId('JDW');
  ensureColumn('Jadwal', 'Tanggal');
  appendRowByHeaders('Jadwal', {
    'ID_Jadwal': id,
    'Hari': data.hari,
    'Jam_Mulai': data.jam_mulai,
    'Jam_Selesai': data.jam_selesai,
    'ID_Siswa': data.id_siswa,
    'ID_Tutor': data.id_tutor,
    'Mata_Pelajaran': data.mapel,
    'Ruang': data.ruang || '',
    'Status': 'Aktif',
    'Tanggal': toDateKey(data.tanggal) // 'yyyy-MM-dd'
  });
  return { success: true, message: 'Jadwal berhasil disimpan.', id: id };
}

function hapusJadwal(idJadwal) {
  return hapusBaris('Jadwal', 'ID_Jadwal', idJadwal);
}

// ============================================================
// KEUANGAN CRUD
// ============================================================
function tambahTransaksi(data) {
  var id = generateId('TRX');
  appendToSheet('Keuangan', [
    id, getTimestamp(), data.tipe, data.kategori,
    data.deskripsi, data.nominal, data.metode || 'Kas Tunai',
    data.id_siswa || '', data.id_tutor || '', 'Sukses'
  ]);
  return { success: true, message: 'Transaksi berhasil dicatat.', id: id };
}

function getRingkasanKeuangan() {
  var keuangan = getSheetData('Keuangan');
  var pemasukan = 0;
  var pengeluaran = 0;
  keuangan.forEach(function(k) {
    var nominal = parseFloat(k.Nominal) || 0;
    if (k.Tipe === 'Masuk') pemasukan += nominal;
    else if (k.Tipe === 'Keluar') pengeluaran += nominal;
  });
  return {
    pemasukan: pemasukan,
    pengeluaran: pengeluaran,
    saldo: pemasukan - pengeluaran,
    total_transaksi: keuangan.length
  };
}

// ============================================================
// ABSENSI & JADWAL FOR TUTOR PORTAL
// ============================================================

// Mengembalikan SEMUA jadwal aktif beserta nama tutor & siswa
// Sehingga portal tutor bisa filter per hari di frontend
function getAllJadwalWithNames() {
  var jadwal = getSheetData('Jadwal');
  var tutor = getSheetData('Tutor');
  var siswa = getSheetData('Siswa');

  return jadwal
    .filter(function(j) {
      return j.Status === 'Aktif';
    })
    .map(function(j) {
      var t = tutor.find(function(x) { return x.ID_Tutor === j.ID_Tutor; }) || {};
      var s = siswa.find(function(x) { return x.ID_Siswa === j.ID_Siswa; }) || {};
      return {
        ID_Jadwal: j.ID_Jadwal,
        Hari: j.Hari,
        Tanggal: toDateKey(j.Tanggal),
        Jam_Mulai: j.Jam_Mulai,
        Jam_Selesai: j.Jam_Selesai,
        Nama_Tutor: t.Nama_Tutor || '-',
        Nama_Siswa: s.Nama_Siswa || '-',
        Kelas: s.Kelas || '-',
        Mata_Pelajaran: j.Mata_Pelajaran,
        Ruang: j.Ruang || '-',
        ID_Tutor: j.ID_Tutor,
        ID_Siswa: j.ID_Siswa
      };
    });
}

// LEGACY: tetap dipertahankan agar tidak break existing deployment
function getJadwalHariIni() {
  return getAllJadwalWithNames();
}

// Data ringan untuk portal tutor: status absensi per jadwal+tanggal (tanpa data sensitif)
function getStatusAbsensiTutor() {
  return getSheetData('Absensi').map(function(a) {
    return {
      ID_Jadwal: a.ID_Jadwal,
      Tanggal_Sesi: getSesiKey(a),
      Status_Verifikasi: a.Status_Verifikasi || 'Menunggu'
    };
  });
}

function submitAbsensi(payload) {
  var idJadwal = payload.idJadwal;
  var idTutor = payload.idTutor;
  var pinTutor = payload.pinTutor;
  var laporan = payload.laporan;
  var base64Image = payload.base64Image;
  var tanggalSesi = toDateKey(payload.tanggalSesi) || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // Tolak tanggal di masa depan (toleransi 1 hari untuk perbedaan zona waktu)
  var besok = new Date();
  besok.setDate(besok.getDate() + 1);
  if (tanggalSesi > Utilities.formatDate(besok, Session.getScriptTimeZone(), 'yyyy-MM-dd')) {
    throw new Error('Tidak bisa mengisi absen untuk tanggal yang belum tiba.');
  }

  // Validasi PIN
  var tutors = getSheetData('Tutor');
  var tutor = null;
  for (var i = 0; i < tutors.length; i++) {
    if (tutors[i].ID_Tutor === idTutor && String(tutors[i].PIN) === String(pinTutor)) {
      tutor = tutors[i];
      break;
    }
  }
  if (!tutor) throw new Error('PIN Tutor salah atau data tidak ditemukan.');

  // Ambil info jadwal
  var jadwals = getSheetData('Jadwal');
  var jadwal = null;
  for (var j = 0; j < jadwals.length; j++) {
    if (jadwals[j].ID_Jadwal === idJadwal) { jadwal = jadwals[j]; break; }
  }

  var siswaList = getSheetData('Siswa');
  var siswa = jadwal ? siswaList.find(function(s) { return s.ID_Siswa === jadwal.ID_Siswa; }) : null;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Cegah absen ganda: 1 jadwal hanya boleh 1x absen per tanggal sesi
    // (kecuali absen sebelumnya DITOLAK admin, maka boleh absen ulang)
    ensureColumn('Absensi', 'Tanggal_Sesi');
    var sudahAda = getSheetData('Absensi').some(function(a) {
      return String(a.ID_Jadwal) === String(idJadwal) &&
             getSesiKey(a) === tanggalSesi &&
             a.Status_Verifikasi !== 'Ditolak';
    });
    if (sudahAda) {
      throw new Error('Sesi ini sudah diabsen untuk tanggal ' + tanggalSesi.split('-').reverse().join('/') + '. Absen tidak bisa diisi dua kali.');
    }

    var imageUrl = '';
    if (base64Image) {
      try {
        var folderName = 'Absensi_EduSanaa';
        var folder = getOrCreateFolder(folderName);
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var contentType = base64Image.substring(5, base64Image.indexOf(';'));
        var bytes = Utilities.base64Decode(base64Image.split(',')[1]);
        var blob = Utilities.newBlob(bytes, contentType, 'Absensi_' + idJadwal + '_' + new Date().getTime() + '.jpg');
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        // Gunakan URL thumbnail export agar gambar bisa di-embed langsung (bukan halaman preview)
        imageUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      } catch (e) {
        throw new Error('Gagal mengunggah foto: ' + e.message);
      }
    }

    var idAbsen = generateId('ABS');
    appendRowByHeaders('Absensi', {
      'ID_Absen': idAbsen,
      'Timestamp': getTimestamp(),
      'ID_Jadwal': idJadwal,
      'ID_Tutor': idTutor,
      'Nama_Tutor': tutor.Nama_Tutor,
      'Nama_Siswa': siswa ? siswa.Nama_Siswa : '-',
      'Mata_Pelajaran': jadwal ? jadwal.Mata_Pelajaran : '-',
      'Laporan': laporan,
      'URL_Foto': imageUrl,
      'Status_Verifikasi': 'Menunggu',
      'Tanggal_Sesi': tanggalSesi
    });
    return { success: true, message: 'Absensi berhasil disimpan!', idAbsen: idAbsen, tanggalSesi: tanggalSesi };
  } catch (e) {
    throw new Error(e.message);
  } finally {
    lock.releaseLock();
  }
}

function verifikasiAbsensi(idAbsen, status) {
  var result = updateSheetRow('Absensi', 'ID_Absen', idAbsen, { 'Status_Verifikasi': status });
  return { success: result, message: result ? 'Status absensi diperbarui.' : 'Absensi tidak ditemukan.' };
}

// ============================================================
// PENGATURAN
// ============================================================
function getPengaturan() {
  var data = getSheetData('Pengaturan');
  var result = {};
  data.forEach(function(row) { result[row.Key] = row.Value; });
  return result;
}

function simpanPengaturan(data) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Pengaturan');
  if (!sheet) return { success: false };
  var existing = sheet.getDataRange().getValues();
  for (var key in data) {
    var found = false;
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(data[key]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, data[key]]);
  }
  return { success: true, message: 'Pengaturan berhasil disimpan.' };
}

// ============================================================
// HELPER: Hapus Baris
// ============================================================
function hapusBaris(sheetName, idColumn, idValue) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idColIndex = headers.indexOf(idColumn);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idColIndex]) === String(idValue)) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Data berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan.' };
}

// ============================================================
// PENAGIHAN WHATSAPP (+ Rekening dari Pengaturan)
// ============================================================
function getUrlTagihanWA(idSiswa, nominal, keterangan) {
  var siswaList = getSheetData('Siswa');
  var siswa = null;
  for (var i = 0; i < siswaList.length; i++) {
    if (siswaList[i].ID_Siswa === idSiswa) { siswa = siswaList[i]; break; }
  }
  if (!siswa) return { success: false, message: 'Siswa tidak ditemukan.' };

  // Ambil info rekening & template dari pengaturan
  var pengaturan = getPengaturan();
  var rekening = pengaturan['rekening'] || '-';
  var templateCustom = pengaturan['template_tagihan'] || '';

  var noWa = String(siswa.No_WA_Ortu).replace(/\D/g, '');
  var keteranganText = keterangan || 'SPP Bimbingan Belajar';
  var nominalFormatted = 'Rp ' + parseInt(nominal).toLocaleString('id-ID');

  var pesan;
  if (templateCustom.trim()) {
    // Gunakan template kustom, ganti variabel
    pesan = templateCustom
      .replace(/\{nama_ortu\}/g, siswa.Nama_Ortu)
      .replace(/\{nama_siswa\}/g, siswa.Nama_Siswa)
      .replace(/\{nominal\}/g, nominalFormatted)
      .replace(/\{keterangan\}/g, keteranganText)
      .replace(/\{rekening\}/g, rekening);
  } else {
    // Pesan default
    pesan = 'Yth. Bapak/Ibu ' + siswa.Nama_Ortu + ',\n\n' +
      'Kami dari *EduSanaa* ingin menginformasikan tagihan bimbingan belajar:\n\n' +
      '\u{1F464} *Siswa:* ' + siswa.Nama_Siswa + '\n' +
      '\u{1F4B0} *Nominal:* ' + nominalFormatted + '\n' +
      '\u{1F4DD} *Keterangan:* ' + keteranganText + '\n\n' +
      '\u{1F3E6} *Rekening Pembayaran:*\n' + rekening + '\n\n' +
      'Harap segera melakukan pembayaran. Terima kasih atas kepercayaan Anda kepada EduSanaa. \u{1F64F}';
  }

  // Buat Google Calendar Event sebagai reminder tagihan
  var calendarEventId = '';
  try {
    var cal = CalendarApp.getDefaultCalendar();
    var jatuhTempo = new Date();
    jatuhTempo.setDate(jatuhTempo.getDate() + 7);
    jatuhTempo.setHours(9, 0, 0, 0);
    var event = cal.createEvent(
      'Reminder Tagihan - ' + siswa.Nama_Siswa,
      jatuhTempo,
      new Date(jatuhTempo.getTime() + 30 * 60 * 1000),
      {
        description: 'Tagihan ' + keteranganText + '\nNominal: ' + nominalFormatted + '\nNo WA Ortu: ' + noWa
      }
    );
    calendarEventId = event.getId();
  } catch (e) {
    Logger.log('Calendar error: ' + e.message);
  }

  return {
    success: true,
    waUrl: 'https://wa.me/' + noWa + '?text=' + encodeURIComponent(pesan),
    noWa: noWa,
    namaSiswa: siswa.Nama_Siswa,
    calendarEventId: calendarEventId
  };
}

// ============================================================
// NOTIFIKASI GAJI TUTOR VIA WHATSAPP
// ============================================================
function getUrlNotifGajiTutor(idTutor, nominal, keterangan) {
  var tutors = getSheetData('Tutor');
  var tutor = null;
  for (var i = 0; i < tutors.length; i++) {
    if (tutors[i].ID_Tutor === idTutor) { tutor = tutors[i]; break; }
  }
  if (!tutor) return { success: false, message: 'Tutor tidak ditemukan.' };

  var noWa = String(tutor.No_WA || '').replace(/\D/g, '');
  if (!noWa) return { success: false, message: 'No WA tutor belum diisi.' };

  var nominalFormatted = 'Rp ' + parseInt(nominal).toLocaleString('id-ID');
  var ketText = keterangan || 'Honor mengajar bulan ini';

  // Ambil template kustom dari pengaturan
  var pengaturan = getPengaturan();
  var templateCustom = pengaturan['template_gaji'] || '';

  var pesan;
  if (templateCustom.trim()) {
    pesan = templateCustom
      .replace(/\{nama_tutor\}/g, tutor.Nama_Tutor)
      .replace(/\{nominal\}/g, nominalFormatted)
      .replace(/\{keterangan\}/g, ketText);
  } else {
    pesan = 'Halo Kak ' + tutor.Nama_Tutor + ',\n\n' +
      'Kami informasikan bahwa *EduSanaa* telah melakukan pembayaran honor Anda:\n\n' +
      '\u{1F4B0} *Nominal:* ' + nominalFormatted + '\n' +
      '\u{1F4DD} *Keterangan:* ' + ketText + '\n\n' +
      'Mohon dicek ke rekening Anda. Terima kasih atas dedikasi mengajarnya! \u{1F64F}\u{2728}';
  }

  return {
    success: true,
    waUrl: 'https://wa.me/' + noWa + '?text=' + encodeURIComponent(pesan),
    noWa: noWa,
    namaTutor: tutor.Nama_Tutor
  };
}

// ============================================================
// TUTOR - Riwayat Sesi & Honorarium
// ============================================================
function getRiwayatTutor(idTutor) {
  var absensi = getSheetData('Absensi');
  var jadwal = getSheetData('Jadwal');
  var siswaList = getSheetData('Siswa');
  var tutors = getSheetData('Tutor');
  var tutor = tutors.find(function(t) { return t.ID_Tutor === idTutor; });

  var riwayat = absensi.filter(function(a) { return a.ID_Tutor === idTutor; })
    .map(function(a) { 
      a.Tanggal_Sesi = getSesiKey(a); 
      var j = jadwal.find(function(x) { return String(x.ID_Jadwal) === String(a.ID_Jadwal); });
      var s = j ? siswaList.find(function(x) { return String(x.ID_Siswa) === String(j.ID_Siswa); }) : null;
      a.Kelas = s ? s.Kelas : '-';
      return a; 
    });
  var tarifPerSesi = tutor ? (parseFloat(tutor.Tarif_Per_Sesi) || 0) : 0;
  var totalHonor = riwayat.filter(function(a) {
    return a.Status_Verifikasi === 'Terverifikasi';
  }).length * tarifPerSesi;

  return {
    tutor: tutor || {},
    riwayat: riwayat,
    totalSesi: riwayat.length,
    sesiVerifikasi: riwayat.filter(function(a) { return a.Status_Verifikasi === 'Terverifikasi'; }).length,
    totalHonor: totalHonor,
    tarifPerSesi: tarifPerSesi
  };
}

// ============================================================
// CEK KONEKSI GOOGLE DRIVE (untuk verifikasi dari admin)
// ============================================================
function cekKoneksiDrive() {
  try {
    var folderName = 'Absensi_EduSanaa';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : null;
    if (!folder) {
      return { success: true, connected: true, message: 'Drive terhubung, tapi folder Absensi_EduSanaa belum ada. Akan dibuat otomatis saat absensi pertama disubmit.', fileCount: 0, folderId: '' };
    }
    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext()) { files.next(); count++; }
    return {
      success: true,
      connected: true,
      message: 'Terhubung ke Google Drive. Folder Absensi_EduSanaa ditemukan.',
      fileCount: count,
      folderId: folder.getId(),
      folderUrl: folder.getUrl()
    };
  } catch (e) {
    return { success: false, connected: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// FOTO ABSENSI - fallback ambil langsung dari Drive (jika embed gagal)
// ============================================================
function getFotoAbsensiDataUrl(url) {
  try {
    var s = String(url || '');
    var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/) || s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (!m) return { success: false, message: 'ID file tidak dikenali.' };
    var blob = DriveApp.getFileById(m[1]).getBlob();
    return { success: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  } catch (e) {
    return { success: false, message: 'Gagal mengambil foto: ' + e.message };
  }
}

// ============================================================
// QRIS - upload, ambil, hapus (file disimpan di Drive folder EduSanaa_Assets)
// ============================================================
function uploadQris(base64Image) {
  if (!base64Image || base64Image.indexOf('base64,') === -1) throw new Error('File gambar tidak valid.');
  var contentType = base64Image.substring(5, base64Image.indexOf(';'));
  if (contentType.indexOf('image/') !== 0) throw new Error('File harus berupa gambar (PNG/JPG).');
  var bytes = Utilities.base64Decode(base64Image.split(',')[1]);
  var ext = contentType.indexOf('png') > -1 ? '.png' : '.jpg';

  // Hapus QRIS lama (pindah ke sampah)
  var p = getPengaturan();
  if (p.qris_file_id) {
    try { DriveApp.getFileById(p.qris_file_id).setTrashed(true); } catch (e) {}
  }

  var folder = getOrCreateFolder('EduSanaa_Assets');
  var file = folder.createFile(Utilities.newBlob(bytes, contentType, 'QRIS_EduSanaa' + ext));
  simpanPengaturan({ qris_file_id: file.getId() });
  return { success: true, message: 'QRIS berhasil diunggah.', fileId: file.getId() };
}

function getQrisDataUrl() {
  var p = getPengaturan();
  if (!p.qris_file_id) return { success: true, dataUrl: '' };
  try {
    var blob = DriveApp.getFileById(p.qris_file_id).getBlob();
    return { success: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  } catch (e) {
    return { success: false, dataUrl: '', message: 'File QRIS tidak ditemukan di Drive. Silakan upload ulang.' };
  }
}

function hapusQris() {
  var p = getPengaturan();
  if (p.qris_file_id) {
    try { DriveApp.getFileById(p.qris_file_id).setTrashed(true); } catch (e) {}
  }
  simpanPengaturan({ qris_file_id: '' });
  return { success: true, message: 'QRIS dihapus.' };
}

// ============================================================
// LOGO LEMBAGA - upload, ambil, hapus (file di Drive folder EduSanaa_Assets)
// ============================================================
function uploadLogo(base64Image) {
  if (!base64Image || base64Image.indexOf('base64,') === -1) throw new Error('File gambar tidak valid.');
  var contentType = base64Image.substring(5, base64Image.indexOf(';'));
  if (contentType.indexOf('image/') !== 0) throw new Error('File harus berupa gambar (PNG/JPG).');
  var bytes = Utilities.base64Decode(base64Image.split(',')[1]);
  var ext = contentType.indexOf('png') > -1 ? '.png' : '.jpg';

  // Hapus logo lama (pindah ke sampah)
  var p = getPengaturan();
  if (p.logo_file_id) {
    try { DriveApp.getFileById(p.logo_file_id).setTrashed(true); } catch (e) {}
  }

  var folder = getOrCreateFolder('EduSanaa_Assets');
  var file = folder.createFile(Utilities.newBlob(bytes, contentType, 'Logo_EduSanaa' + ext));
  simpanPengaturan({ logo_file_id: file.getId() });
  return { success: true, message: 'Logo berhasil diunggah.', fileId: file.getId() };
}

function getLogoDataUrl() {
  var p = getPengaturan();
  if (!p.logo_file_id) return { success: true, dataUrl: '' };
  try {
    var blob = DriveApp.getFileById(p.logo_file_id).getBlob();
    return { success: true, dataUrl: 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes()) };
  } catch (e) {
    return { success: false, dataUrl: '', message: 'File logo tidak ditemukan di Drive. Silakan upload ulang.' };
  }
}

function hapusLogo() {
  var p = getPengaturan();
  if (p.logo_file_id) {
    try { DriveApp.getFileById(p.logo_file_id).setTrashed(true); } catch (e) {}
  }
  simpanPengaturan({ logo_file_id: '' });
  return { success: true, message: 'Logo lembaga dihapus.' };
}

// ============================================================
// HAPUS ABSENSI LAMA (> 1 bulan) - batch delete berdasarkan array ID_Absen
// ============================================================
function hapusAbsensiLama(ids) {
  if (!ids || !ids.length) return { success: false, message: 'Tidak ada ID absensi yang diberikan.' };
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Absensi');
  if (!sheet) return { success: false, message: 'Sheet Absensi tidak ditemukan.' };

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idColIndex = headers.indexOf('ID_Absen');
  if (idColIndex === -1) return { success: false, message: 'Kolom ID_Absen tidak ditemukan.' };

  // Buat Set ID yang akan dihapus agar lookup O(1)
  var idSet = {};
  ids.forEach(function(id) { idSet[String(id)] = true; });

  // Kumpulkan row index yang akan dihapus (dari bawah ke atas agar index tidak bergeser)
  var toDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (idSet[String(data[i][idColIndex])]) {
      toDelete.push(i + 1); // 1-indexed row number in sheet
    }
  }

  if (toDelete.length === 0) return { success: false, message: 'Data tidak ditemukan di spreadsheet.' };

  // Hapus dari bawah ke atas
  toDelete.reverse().forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });

  return { success: true, message: toDelete.length + ' data absensi lama berhasil dihapus.' };
}
