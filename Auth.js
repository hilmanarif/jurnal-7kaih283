/**
 * ============================================
 *  Auth.gs — Autentikasi & Password
 * ============================================
 */

/** Login siswa: cari username+password di sheet Students */
function loginSiswa(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Students");
  const data = sh.getDataRange().getValues();

  let user = null;
  const target = String(username).trim().toLowerCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target && String(data[i][1]) === String(password)) {
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

  if (!user) return { success: false, message: "Username atau sandi salah!" };

  // Ambil data habits siswa
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
  return { success: true, user: user, habits: habits };
}

/** Login admin */
function loginAdmin(user, pass) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admins");
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(user) && String(data[i][1]) === String(pass)) {
      return { success: true, nama: String(data[i][2]), username: String(data[i][0]) };
    }
  }
  return { success: false, message: "Username atau sandi admin salah!" };
}

/** Siswa ganti password sendiri (perlu verifikasi password lama) */
function changePasswordSiswa(username, oldPass, newPass) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  const target = String(username).trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) {
      if (String(data[i][1]) !== String(oldPass)) {
        return { success: false, message: "Password lama tidak cocok!" };
      }
      if (!newPass || String(newPass).length < 3) {
        return { success: false, message: "Password baru minimal 3 karakter!" };
      }
      sh.getRange(i + 1, 2).setValue(String(newPass));
      return { success: true, message: "Password berhasil diubah!" };
    }
  }
  return { success: false, message: "Username tidak ditemukan!" };
}

/** Admin ganti password sendiri */
function changePasswordAdmin(username, oldPass, newPass) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Admins");
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(username)) {
      if (String(data[i][1]) !== String(oldPass)) {
        return { success: false, message: "Password lama tidak cocok!" };
      }
      if (!newPass || String(newPass).length < 3) {
        return { success: false, message: "Password baru minimal 3 karakter!" };
      }
      sh.getRange(i + 1, 2).setValue(String(newPass));
      return { success: true, message: "Password admin berhasil diubah!" };
    }
  }
  return { success: false, message: "Username admin tidak ditemukan!" };
}

/** Admin reset password siswa (tanpa perlu tahu password lama) */
function resetPasswordSiswa(username, newPass) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
  const data = sh.getDataRange().getValues();
  const target = String(username).trim().toLowerCase();

  if (!newPass || String(newPass).length < 3) {
    return { success: false, message: "Password baru minimal 3 karakter!" };
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === target) {
      sh.getRange(i + 1, 2).setValue(String(newPass));
      return { success: true, message: "Password siswa berhasil di-reset!" };
    }
  }
  return { success: false, message: "Username siswa tidak ditemukan!" };
}
