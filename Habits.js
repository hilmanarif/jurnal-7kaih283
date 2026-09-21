/**
 * ============================================
 *  Habits.gs — Data Kebiasaan Harian
 * ============================================
 */

/** Simpan 1 entri kebiasaan */
function saveHabit(username, type, tgl, detail1, detail2, lat, lng, ip) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Habits");
  let detailAman = detail1;
  if (type === 'bangun' || type === 'tidur') { detailAman = "'" + detail1; }

  const ts = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
  sh.appendRow([ts, username, type, "'" + tgl, detailAman || "", detail2 || "", lat || "-", lng || "-", ip || "-"]);
  return true;
}

/** Dashboard admin: status pengisian per siswa untuk tanggal tertentu */
function getAdminDashboardData(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName("Students");
  const habitSheet = ss.getSheetByName("Habits");
  const sData = studentSheet.getDataRange().getValues();
  const hData = habitSheet.getDataRange().getValues();

  const habitTypes = ['bangun', 'ibadah', 'olahraga', 'makan', 'belajar', 'masyarakat', 'tidur'];

  // Hitung kebiasaan yang sudah diisi per siswa
  let filledMap = {}; // { username: Set(['bangun','ibadah',...]) }
  for (let i = 1; i < hData.length; i++) {
    let tglVal = hData[i][3];
    if (tglVal instanceof Date) {
      tglVal = Utilities.formatDate(tglVal, "Asia/Jakarta", "yyyy-MM-dd");
    } else {
      tglVal = String(tglVal).trim().substring(0, 10);
    }

    if (tglVal === dateStr) {
      const uname = String(hData[i][1]).trim().toLowerCase();
      const type = String(hData[i][2]).trim().toLowerCase();
      if (!filledMap[uname]) filledMap[uname] = {};
      filledMap[uname][type] = true;
    }
  }

  // Bangun result per siswa
  let result = [];
  for (let i = 1; i < sData.length; i++) {
    if (!sData[i][0]) continue;
    const uname = String(sData[i][0]).trim().toLowerCase();
    const filled = filledMap[uname] || {};
    const count = habitTypes.filter(t => filled[t]).length;
    result.push({
      username: String(sData[i][0]),
      nama: String(sData[i][2]),
      kelas: String(sData[i][3]),
      filled: count,
      total: 7
    });
  }

  return result;
}
