/**
 * Google Apps Script — lưu & duyệt lời chúc cho thiệp cưới.
 * Gắn vào Google Sheet: Tiện ích mở rộng > Apps Script > dán toàn bộ file này.
 *
 * Cột trong sheet "LoiChuc":
 *   A: Ngày giờ chúc | B: Họ tên | C: Lời chúc | D: Duyệt (tick ô vuông để hiển thị lên web)
 */
const SHEET_NAME = "LoiChuc";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Ngày giờ chúc", "Họ tên", "Lời chúc", "Duyệt"]);
    sh.getRange("A1:D1").setFontWeight("bold").setBackground("#d8e3b5");
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 1, 140);
    sh.setColumnWidths(2, 1, 180);
    sh.setColumnWidths(3, 1, 480);
    sh.setColumnWidths(4, 1, 80);
  }
  return sh;
}

// Chống chèn công thức vào Google Sheet
function safe_(v) {
  v = String(v == null ? "" : v).trim();
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

// Khách gửi lời chúc -> thêm dòng mới, mặc định CHƯA duyệt
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const name = safe_(d.name).slice(0, 60);
    const message = safe_(d.message).slice(0, 500);
    const time = safe_(d.time).slice(0, 30);
    if (!name || !message) throw new Error("Thiếu thông tin");

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    const sh = getSheet_();
    sh.appendRow([time, name, message, false]);
    const row = sh.getLastRow();
    sh.getRange(row, 4).insertCheckboxes();      // ô tick duyệt
    lock.releaseLock();
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// Web đọc danh sách lời chúc ĐÃ DUYỆT (cột D = TRUE)
function doGet(e) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues().slice(1);
  const wishes = values
    .filter(r => r[3] === true || String(r[3]).toLowerCase() === "true")
    .map(r => ({ time: String(r[0]), name: String(r[1]), message: String(r[2]) }))
    .reverse();                                   // mới nhất lên đầu
  return json_({ ok: true, wishes: wishes });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
