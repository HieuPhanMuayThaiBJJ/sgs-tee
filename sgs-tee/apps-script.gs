// Google Apps Script - nhận đơn từ form và ghi vào Google Sheet
// Mỗi dòng áo trong đơn = 1 dòng trong Sheet (cùng Mã đơn). Tương thích với bản form cũ.

const SHEET_NAME = "Đăng ký áo";
const HEADERS = ["Thời gian","Tên","SĐT","Size","Đã chuyển khoản","Số tiền","Nội dung CK","Màu","Địa chỉ","Số lượng","Mã đơn","Tổng đơn"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.setFrozenRows(1); }
    for (let i = 0; i < HEADERS.length; i++) {
      if (sh.getRange(1, i + 1).getValue() !== HEADERS[i]) sh.getRange(1, i + 1).setValue(HEADERS[i]);
    }
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0b1f5c").setFontColor("#ffffff");
    sh.getRange("C:C").setNumberFormat("@");

    const d = JSON.parse(e.postData.contents);
    const time = d.submittedAt || new Date();
    const unit = d.unitPrice || 159000;
    const items = (d.items && d.items.length) ? d.items : [{ color: d.color, size: d.size, qty: 1 }];
    const total = d.amount || items.reduce((a, i) => a + (i.qty || 1) * unit, 0);
    const rows = items.map(i => [
      time, d.name, "'" + d.phone, i.size, d.paid, (i.qty || 1) * unit, d.transferContent,
      i.color || "", d.address || "", i.qty || 1, d.orderId || "", total
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet() { return ContentService.createTextOutput("SGS form endpoint OK"); }
