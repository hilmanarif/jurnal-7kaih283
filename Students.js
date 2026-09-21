/**
 * ============================================
 *  Students.gs — Manajemen Data Siswa
 * ============================================
 */

/** Ambil semua data siswa untuk admin */
function getAdminStudents() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  let res = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      res.push({
        username: String(data[i][0]),
        nama: String(data[i][2]),
        kelas: String(data[i][3]),
        cita_cita: String(data[i][4]),
        agama: String(data[i][5] || "")
      });
    }
  }
  return res;
}

/** Tambah 1 siswa baru */
function addStudent(nama, username, pass, kelas, citaCita, agama) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  const uname = String(username).trim().toLowerCase();

  if (!uname || !nama) return { success: false, message: "Nama dan Username wajib diisi!" };

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === uname) {
      return { success: false, message: "Username '" + uname + "' sudah digunakan!" };
    }
  }
  sh.appendRow([uname, String(pass || "123"), String(nama), String(kelas || ""), String(citaCita || ""), String(agama || "")]);
  return { success: true };
}

/** Admin edit data siswa */
function editStudent(username, updateData) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  const target = String(username).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) {
      if (updateData.nama) sh.getRange(i + 1, 3).setValue(String(updateData.nama));
      if (updateData.kelas !== undefined) sh.getRange(i + 1, 4).setValue(String(updateData.kelas));
      if (updateData.cita_cita !== undefined) sh.getRange(i + 1, 5).setValue(String(updateData.cita_cita));
      if (updateData.agama !== undefined) sh.getRange(i + 1, 6).setValue(String(updateData.agama));
      return { success: true, message: "Data siswa berhasil diperbarui!" };
    }
  }
  return { success: false, message: "Siswa tidak ditemukan!" };
}

/** Hapus siswa + semua data kebiasaannya */
function deleteStudent(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      sh.deleteRow(i + 1); break;
    }
  }
  const shH = ss.getSheetByName("Habits");
  const hData = shH.getDataRange().getValues();
  for (let i = hData.length - 1; i >= 1; i--) {
    if (String(hData[i][1]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      shH.deleteRow(i + 1);
    }
  }
  return true;
}

/** Import batch siswa dari CSV (array of arrays) */
function bulkAddStudents(studentsArray) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();

  // Kumpulkan username yang sudah ada
  let existing = {};
  for (let i = 1; i < data.length; i++) {
    existing[String(data[i][0]).trim().toLowerCase()] = true;
  }

  let added = 0, skipped = 0, errors = [];
  for (let j = 0; j < studentsArray.length; j++) {
    const row = studentsArray[j];
    const uname = String(row[0] || "").trim().toLowerCase();
    const pass = String(row[1] || "123");
    const nama = String(row[2] || "").trim();
    const kelas = String(row[3] || "");
    const cita = String(row[4] || "");
    const agama = String(row[5] || "");

    if (!uname || !nama) {
      errors.push("Baris " + (j + 1) + ": Username/Nama kosong");
      continue;
    }
    if (existing[uname]) {
      skipped++;
      continue;
    }
    sh.appendRow([uname, pass, nama, kelas, cita, agama]);
    existing[uname] = true;
    added++;
  }

  return { added: added, skipped: skipped, errors: errors };
}
