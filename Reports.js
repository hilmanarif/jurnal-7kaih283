/**
 * ============================================
 *  Reports.gs — Laporan & Cetak
 *  (Tahap 2: per-kelas, simpan sheet, kesimpulan)
 * ============================================
 */

/** Ambil data laporan 1 siswa (untuk cetak per-siswa) */
function getReportForAdmin(username) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = ss.getSheetByName("Students");
    const data = studentSheet.getDataRange().getValues();

    let user = null;
    const target = String(username).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === target) {
        user = {
          username: String(data[i][0]),
          nama: String(data[i][2]),
          kelas: String(data[i][3]),
          cita_cita: String(data[i][4]),
          agama: String(data[i][5] || "")
        };
        break;
      }
    }
    if (!user) throw new Error("Siswa tidak ditemukan.");

    const habitSheet = ss.getSheetByName("Habits");
    const habitData = habitSheet.getDataRange().getValues();
    let habits = { bangun: [], ibadah: [], olahraga: [], makan: [], belajar: [], masyarakat: [], tidur: [] };

    for (let i = 1; i < habitData.length; i++) {
      if (String(habitData[i][1]).trim().toLowerCase() === target) {
        const type = String(habitData[i][2]).trim().toLowerCase();
        if (habits[type]) {
          let tglVal = habitData[i][3];
          if (tglVal instanceof Date) {
            tglVal = Utilities.formatDate(tglVal, "Asia/Jakarta", "yyyy-MM-dd");
          } else {
            tglVal = String(tglVal).trim().substring(0, 10);
          }

          let d1Val = habitData[i][4];
          if (d1Val instanceof Date) {
            d1Val = Utilities.formatDate(d1Val, "Asia/Jakarta", "HH:mm");
          } else {
            d1Val = String(d1Val || "");
          }

          habits[type].push({ tgl: tglVal, detail1: d1Val, detail2: String(habitData[i][5] || "") });
        }
      }
    }
    return { user: user, habits: habits };
  } catch (err) {
    throw new Error(err.message);
  }
}

/** Ambil daftar kelas unik dari sheet Students */
function getClassList() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  let classes = {};
  for (let i = 1; i < data.length; i++) {
    const kelas = String(data[i][3] || "").trim();
    if (kelas) classes[kelas] = true;
  }
  return Object.keys(classes).sort();
}

/** Ambil data laporan per kelas — Tahap 2 */
function getClassReport(kelas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("Students");
  const sData = studentSheet.getDataRange().getValues();
  const habitSheet = ss.getSheetByName("Habits");
  const hData = habitSheet.getDataRange().getValues();

  let students = [];
  for (let i = 1; i < sData.length; i++) {
    if (String(sData[i][3]).trim() === String(kelas).trim()) {
      students.push({
        username: String(sData[i][0]),
        nama: String(sData[i][2]),
        kelas: String(sData[i][3]),
        agama: String(sData[i][5] || "")
      });
    }
  }

  // Ambil habits semua siswa di kelas
  let habitsMap = {};
  students.forEach(s => {
    habitsMap[s.username.toLowerCase()] = { bangun: [], ibadah: [], olahraga: [], makan: [], belajar: [], masyarakat: [], tidur: [] };
  });

  for (let i = 1; i < hData.length; i++) {
    const uname = String(hData[i][1]).trim().toLowerCase();
    if (habitsMap[uname]) {
      const type = String(hData[i][2]).trim().toLowerCase();
      if (habitsMap[uname][type]) {
        let tglVal = hData[i][3];
        if (tglVal instanceof Date) {
          tglVal = Utilities.formatDate(tglVal, "Asia/Jakarta", "yyyy-MM-dd");
        } else {
          tglVal = String(tglVal).trim().substring(0, 10);
        }
        let d1Val = hData[i][4];
        if (d1Val instanceof Date) {
          d1Val = Utilities.formatDate(d1Val, "Asia/Jakarta", "HH:mm");
        } else {
          d1Val = String(d1Val || "");
        }
        habitsMap[uname][type].push({ tgl: tglVal, detail1: d1Val, detail2: String(hData[i][5] || "") });
      }
    }
  }

  let result = students.map(s => ({
    user: s,
    habits: habitsMap[s.username.toLowerCase()]
  }));

  return { kelas: kelas, students: result };
}

/** Simpan laporan ke sheet baru — Tahap 2 */
function saveReportToSheet(sheetName, reportRows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hapus sheet lama jika nama sama
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const sh = ss.insertSheet(sheetName);
  if (reportRows && reportRows.length > 0) {
    sh.getRange(1, 1, reportRows.length, reportRows[0].length).setValues(reportRows);
    // Format header
    sh.getRange("1:1").setFontWeight("bold").setBackground("#00A896").setFontColor("white");
    sh.setFrozenRows(1);
    try { sh.autoResizeColumns(1, reportRows[0].length); } catch(e) {}
  }
  return { success: true, sheetName: sheetName };
}
