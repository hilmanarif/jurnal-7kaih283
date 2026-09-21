function myFunction() {
  
}
/**
 * ============================================
 *  Admin.gs — Settings & Tracking Harian
 * ============================================
 */

/** Simpan pengaturan sekolah */
function saveAdminSettings(settingsObj) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  const data = sh.getDataRange().getValues();
  for (let key in settingsObj) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sh.getRange(i + 1, 2).setValue(settingsObj[key]);
        found = true;
        break;
      }
    }
    if (!found) sh.appendRow([key, settingsObj[key]]);
  }
  return true;
}

/** Tracking harian admin: siapa mengisi apa pada tanggal tertentu */
function getDailyTrackingAdmin(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const habitSheet = ss.getSheetByName("Habits");
  const hData = habitSheet.getDataRange().getValues();
  const studentSheet = ss.getSheetByName("Students");
  const sData = studentSheet.getDataRange().getValues();

  let userMap = {};
  for (let i = 1; i < sData.length; i++) { userMap[sData[i][0]] = sData[i][2]; }

  let tracking = [];
  for (let i = 1; i < hData.length; i++) {
    let tglVal = hData[i][3];
    if (tglVal instanceof Date) {
      tglVal = Utilities.formatDate(tglVal, "Asia/Jakarta", "yyyy-MM-dd");
    } else {
      tglVal = String(tglVal).trim().substring(0, 10);
    }

    if (tglVal === dateStr) {
      let ts = hData[i][0];
      if (ts instanceof Date) {
        ts = Utilities.formatDate(ts, "Asia/Jakarta", "HH:mm:ss");
      }
      tracking.push({
        time: String(ts),
        nama: userMap[hData[i][1]] || String(hData[i][1]),
        type: String(hData[i][2]),
        ip: String(hData[i][8] || "-"),
        lat: String(hData[i][6] || "-"),
        lng: String(hData[i][7] || "-")
      });
    }
  }
  tracking.reverse();
  return tracking;
}
