/**
 * 【一鍵全自動修正時間格式與看板】
 * 
 * 使用方式：
 * 回到 Google Apps Script，把內容覆蓋貼上，點「執行 (Run)」即可！
 * 會自動將「最後回報時間」格式化為清晰漂亮的「HH:mm:ss」或「MM/dd HH:mm」格式！
 */

function fixTimeFormatAndDashboard() {
  Logger.log("🔧 正在自動格式化時間與看板欄位...");

  const ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GF3Y9oR_KidYlfdnQwkSF_xCRgElAlAC1PwJG8FJSrU/edit");
  const sheets = ss.getSheets();
  
  let formSheet = null;
  let dashboardSheet = null;

  for (let s of sheets) {
    if (s.getName().includes("即時戰情看板")) {
      dashboardSheet = s;
    } else {
      formSheet = s;
    }
  }

  const formSheetName = formSheet.getName();

  // 1. 設定 E 欄（最後回報時間）格式為「時:分:秒」 (例如 08:30:15)
  dashboardSheet.getRange("E5:E54").setNumberFormat("HH:mm:ss");
  
  // 2. 也將 C 欄（實到人數）設為整數置中
  dashboardSheet.getRange("C5:C54").setNumberFormat("0").setHorizontalAlignment("center");
  dashboardSheet.getRange("D5:D54").setHorizontalAlignment("center");
  dashboardSheet.getRange("E5:E54").setHorizontalAlignment("center");

  // 3. 自動更新 50 台車最新數據抓取
  const rows = [];
  for (let i = 1; i <= 50; i++) {
    const busName = `${String(i).padStart(2, '0')}車`;
    const rowIdx = 4 + i;
    const formulaActual = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!D:D, 0, 0, -1), 0)`;
    const formulaStatus = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!E:E, "⚪ 尚未回報", 0, -1), "⚪ 尚未回報")`;
    const formulaTime   = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!A:A, "-", 0, -1), "-")`;
    const formulaNote   = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!F:F, "-", 0, -1), "-")`;
    rows.push([busName, 40, formulaActual, formulaStatus, formulaTime, formulaNote]);
  }

  dashboardSheet.getRange("A5:F54").setValues(rows);
  dashboardSheet.autoResizeColumns(1, 6);

  Logger.log("🎉【時間格式已修正完成！】現在顯示為標準 時:分:秒 (HH:mm:ss)");
}
