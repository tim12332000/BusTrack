/**
 * 【一鍵全自動修復看板公式】
 * 
 * 使用方式：
 * 回到剛才的 Google Apps Script 網頁，把內容全選覆蓋貼上，點「執行 (Run)」即可！
 * 會自動抓取正確的表單工作表名稱，並將 50 台車的即時公式 100% 修正完成！
 */

function autoFixDashboard() {
  Logger.log("🔧 開始自動診斷與修復看板公式...");

  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GF3Y9oR_KidYlfdnQwkSF_xCRgElAlAC1PwJG8FJSrU/edit");
  const sheets = ss.getSheets();
  
  // 1. 自動尋找哪一個是表單回應工作表
  let formSheet = null;
  let dashboardSheet = null;

  for (let s of sheets) {
    const name = s.getName();
    if (name.includes("即時戰情看板")) {
      dashboardSheet = s;
    } else {
      // 另一個就是存放資料的工作表
      formSheet = s;
    }
  }

  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet("即時戰情看板", 0);
  }
  if (!formSheet) {
    Logger.log("❌ 找不到表單回應工作表");
    return;
  }

  const formSheetName = formSheet.getName();
  Logger.log(`✅ 找到您的表單資料分頁名稱為：'${formSheetName}'`);

  // 2. 重新設定頂部 KPI 統計公式
  dashboardSheet.getRange("A1:E1").setValues([["總車數", "已出發", "全員到齊", "上車中", "尚未回報"]]).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
  dashboardSheet.getRange("A2").setValue(50);
  dashboardSheet.getRange("B2").setFormula(`=COUNTIF(D5:D54, "*已出發*")`);
  dashboardSheet.getRange("C2").setFormula(`=COUNTIF(D5:D54, "*全員到齊*")`);
  dashboardSheet.getRange("D2").setFormula(`=COUNTIF(D5:D54, "*上車中*")`);
  dashboardSheet.getRange("E2").setFormula(`=COUNTIF(D5:D54, "*尚未回報*")`);

  // 3. 標題列
  dashboardSheet.getRange("A4:F4").setValues([["車號", "應到人數", "最新實到人數", "最新狀態", "最後回報時間", "備註"]]).setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");

  // 4. 一鍵寫入 50 台車保證正確的 XLOOKUP 自動更新公式
  const rows = [];
  for (let i = 1; i <= 50; i++) {
    const busName = `${String(i).padStart(2, '0')}車`;
    const rowIdx = 4 + i;
    
    // 使用 XLOOKUP(目標, 車號欄, 資料欄, 預設值, 0, -1 由後往前搜最新)
    const formulaActual = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!D:D, 0, 0, -1), 0)`;
    const formulaStatus = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!E:E, "⚪ 尚未回報", 0, -1), "⚪ 尚未回報")`;
    const formulaTime   = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!A:A, "-", 0, -1), "-")`;
    const formulaNote   = `=IFNA(XLOOKUP(A${rowIdx}, '${formSheetName}'!B:B, '${formSheetName}'!F:F, "-", 0, -1), "-")`;

    rows.push([busName, 40, formulaActual, formulaStatus, formulaTime, formulaNote]);
  }

  dashboardSheet.getRange("A5:F54").setValues(rows);
  dashboardSheet.autoResizeColumns(1, 6);

  Logger.log("🎉【修復成功！】即時戰情看板公式已全數重新校正並生效！");
}
