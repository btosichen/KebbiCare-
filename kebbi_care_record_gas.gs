// Kebbi Care+｜照護紀錄專用 Google Apps Script
// 請部署為第一個 Web App；此份只寫入照護事件，不寫入環境感測資料。
const SPREADSHEET_ID = "1mh6qGH9pg8mH6Uc9TX5NfvhJh6MUZYQyNaGy6hBh2MI";
const SHEET_NAME = "CareRecord";
const HEADERS = ["日期", "時間", "睡眠", "早餐", "心情", "吃藥", "AI摘要", "資料來源", "老人編號"];

function doGet() {
  return ContentService.createTextOutput("Kebbi Care 照護紀錄 GAS 運作中");
}

function initSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#E2F0D9");
  sheet.setFrozenRows(1);
  return sheet;
}

function saveRecord(data) {
  try {
    const now = new Date();
    initSheet().appendRow([
      Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd"),
      Utilities.formatDate(now, "Asia/Taipei", "HH:mm:ss"),
      Number(data.sleep ?? 0),
      Number(data.breakfast ?? 0),
      Number(data.mood ?? 0),
      Number(data.medicine ?? 0),
      String(data.summary ?? "無摘要資料"),
      String(data.source ?? "Web App"),
      String(data.elderId ?? "尚未辨識")
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "照護紀錄已寫入" };
  } catch (error) {
    console.error(error.stack);
    return { success: false, message: error.message };
  }
}

function doPost(e) {
  try {
    if (!e?.postData?.contents) throw new Error("沒有收到 POST 資料");
    const data = JSON.parse(e.postData.contents);
    const result = saveRecord(data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getHistoryRecords() {
  try {
    const sheet = initSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const startRow = Math.max(2, lastRow - 6);
    return sheet.getRange(startRow, 1, lastRow - startRow + 1, HEADERS.length).getValues().map(row => ({
      date: row[0], time: row[1], sleep: row[2], breakfast: row[3], mood: row[4], medicine: row[5],
      summary: row[6], source: row[7], elderId: row[8] || "尚未辨識"
    }));
  } catch (error) {
    console.error(error.stack);
    return [];
  }
}

function testWrite() {
  console.log(JSON.stringify(saveRecord({
    sleep: 5, breakfast: 1, mood: 5, medicine: 1,
    summary: "照護紀錄 GAS 測試成功", source: "testWrite", elderId: "A-001"
  })));
}
