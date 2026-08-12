// Kebbi Care+｜環境感測專用 Google Apps Script
// 先建立「新的 Google 試算表」，再將它的 ID 貼到下方雙引號中並部署為第二個 Web App。
const SPREADSHEET_ID = "請貼上新的環境感測試算表 ID";
const SHEET_NAME = "EnvironmentRecord";
const HEADERS = ["日期", "時間", "室內溫度(°C)", "室內濕度(%)", "室內亮度(%)", "資料來源"];

function doGet() {
  return ContentService.createTextOutput("Kebbi Care 環境感測 GAS 運作中");
}

function initSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#DDEBF7");
  sheet.setFrozenRows(1);
  return sheet;
}

function optionalNumber(value) {
  if (value === undefined || value === null || value === "") return "";
  const numberValue = Number(value);
  return Number.isFinite(numberValue) ? numberValue : "";
}

function saveEnvironmentRecord(data) {
  try {
    const temperature = optionalNumber(data.indoorTemperature ?? data.temperature);
    const humidity = optionalNumber(data.humidity);
    const brightness = optionalNumber(data.brightness ?? data.lux);
    if (temperature === "" && humidity === "" && brightness === "") {
      throw new Error("沒有可寫入的溫度、濕度或亮度資料");
    }

    const now = new Date();
    initSheet().appendRow([
      Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd"),
      Utilities.formatDate(now, "Asia/Taipei", "HH:mm:ss"),
      temperature, humidity, brightness,
      String(data.source ?? "ESP32 環境感測器")
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: "環境感測紀錄已寫入" };
  } catch (error) {
    console.error(error.stack);
    return { success: false, message: error.message };
  }
}

function doPost(e) {
  try {
    if (!e?.postData?.contents) throw new Error("沒有收到 POST 資料");
    const result = saveEnvironmentRecord(JSON.parse(e.postData.contents));
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function testWrite() {
  console.log(JSON.stringify(saveEnvironmentRecord({
    indoorTemperature: 26.5, humidity: 62, brightness: 42.5, source: "testWrite"
  })));
}
