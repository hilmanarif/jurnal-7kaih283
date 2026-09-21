/**
 * ============================================
 *  Code.gs — Pintu Masuk Aplikasi
 *  Jurnal 7 Kebiasaan Anak Indonesia Hebat
 *  © Develop by R●Gunawan | Powered by Google Apps Script 2026
 * ============================================
 */

function doGet(e) {
  try {
    setupDatabase();
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Jurnal 7 Kebiasaan Anak Indonesia Hebat')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput("<h3>Error Sistem</h3><p>" + err.message + "</p>");
  }
}

/** Include file HTML (untuk template) */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Setup database: buat sheet jika belum ada, migrasi jika perlu */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;

  const H = { bg: "#00A896", color: "white" };

  /* ====== SHEET: Students ====== */
  if (!ss.getSheetByName("Students")) {
    const sh = ss.insertSheet("Students");
    sh.appendRow(["Username", "Password", "Nama", "Kelas", "Cita-cita", "Agama"]);
    sh.getRange("A1:F1").setFontWeight("bold").setBackground(H.bg).setFontColor(H.color);
    const sample = [
      ["andi",  "123", "Andi Pratama",    "7-A", "Dokter",     "Islam"],
      ["siti",  "123", "Siti Nurhaliza",   "7-A", "Guru",       "Islam"],
      ["budi",  "123", "Budi Setiawan",    "7-B", "Insinyur",   "Kristen"],
      ["maria", "123", "Maria Anggraeni",  "7-B", "Perawat",    "Katolik"],
      ["kadek", "123", "Kadek Ayu Dewi",   "7-C", "Desainer",   "Hindu"],
      ["wayan", "123", "I Wayan Dharma",   "7-C", "Pilot",      "Hindu"],
      ["suci",  "123", "Suci Rahayu",      "8-A", "Penyanyi",   "Buddha"],
      ["kevin", "123", "Kevin Wijaya",     "8-A", "Programmer", "Konghucu"],
      ["dewi",  "123", "Dewi Lestari",     "8-B", "Penulis",    "Kepercayaan"],
      ["rizki", "123", "Rizki Ramadhan",   "8-B", "TNI",        "Islam"]
    ];
    sh.getRange(2, 1, sample.length, 6).setValues(sample);
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, 6); } catch(e) {}
  } else {
    const sh = ss.getSheetByName("Students");
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (headers.indexOf("Agama") === -1) {
      const c = sh.getLastColumn() + 1;
      sh.getRange(1, c).setValue("Agama").setFontWeight("bold").setBackground(H.bg).setFontColor(H.color);
    }
  }

  /* ====== SHEET: Settings ====== */
  if (!ss.getSheetByName("Settings")) {
    const sh = ss.insertSheet("Settings");
    sh.appendRow(["Key", "Value"]);
    sh.getRange("A1:B1").setFontWeight("bold").setBackground(H.bg).setFontColor(H.color);
    const def = [
      ["pemerintah",      "PEMERINTAH KABUPATEN KONOHA"],
      ["nama_sekolah",    "SMP NEGERI 1 KONOHA"],
      ["alamat",          "Jl. Tengah, Kab, Konoha"],
      ["nama_guru",       "Budi , S.Pd"],
      ["nip_guru",        "198001012005011001"],
      ["nama_kepsek",     "Juysuf, M.Pd"],
      ["nip_kepsek",      "197502022000032002"],
      ["tempat_ttd",      "Konoha"],
      ["logo_daerah",     ""],
      ["deskripsi_login", "Anak Indonesia Hebat 🇮🇩"]
    ];
    sh.getRange(2, 1, def.length, 2).setValues(def);
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, 2); } catch(e) {}
  } else {
    // Migrasi: tambah deskripsi_login jika belum ada
    const sh = ss.getSheetByName("Settings");
    const data = sh.getDataRange().getValues();
    let hasDeskripsi = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === "deskripsi_login") { hasDeskripsi = true; break; }
    }
    if (!hasDeskripsi) sh.appendRow(["deskripsi_login", "Anak Indonesia Hebat 🇮🇩"]);
  }

  /* ====== SHEET: Habits ====== */
  if (!ss.getSheetByName("Habits")) {
    const sh = ss.insertSheet("Habits");
    sh.appendRow(["Timestamp", "Username", "Type", "Tanggal", "Detail1", "Detail2", "Latitude", "Longitude", "IP_Address"]);
    sh.getRange("A1:I1").setFontWeight("bold").setBackground(H.bg).setFontColor(H.color);
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, 9); } catch(e) {}
  } else {
    const sh = ss.getSheetByName("Habits");
    const lc = sh.getLastColumn();
    if (lc > 0 && lc < 9) {
      sh.getRange(1, 7).setValue("Latitude");
      sh.getRange(1, 8).setValue("Longitude");
      sh.getRange(1, 9).setValue("IP_Address");
    }
  }

  /* ====== SHEET: Admins ====== */
  if (!ss.getSheetByName("Admins")) {
    const sh = ss.insertSheet("Admins");
    sh.appendRow(["Username", "Password", "Nama"]);
    sh.getRange("A1:C1").setFontWeight("bold").setBackground(H.bg).setFontColor(H.color);
    sh.appendRow(["admin", "admin123", "Guru Kelas Utama"]);
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, 3); } catch(e) {}
  }
}

/** Reset semua data — jalankan manual dari Script Editor */
function resetDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ["Students", "Settings", "Habits", "Admins"].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      if (ss.getSheets().length <= 1) ss.insertSheet("_temp");
      ss.deleteSheet(sh);
    }
  });
  setupDatabase();
  const temp = ss.getSheetByName("_temp");
  if (temp && ss.getSheets().length > 1) ss.deleteSheet(temp);
  SpreadsheetApp.getUi().alert("✅ Database berhasil di-reset!");
}

/** Ambil data awal untuk client (settings saja, tanpa daftar siswa) */
function getInitialData() {
  try {
    setupDatabase();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("Spreadsheet tidak terhubung!");

    const settingSheet = ss.getSheetByName("Settings");
    let settings = {};
    if (settingSheet) {
      const data = settingSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) settings[String(data[i][0])] = String(data[i][1]);
      }
    }
    return { error: false, settings: settings };
  } catch (e) {
    return { error: true, message: e.toString() };
  }
}