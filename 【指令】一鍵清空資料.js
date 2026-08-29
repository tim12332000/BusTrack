/**
 * 🗑️【指令：一鍵清空所有回報資料 (全場歸零)】
 * 
 * 使用方式：
 * 1. 複製這段程式碼。
 * 2. 貼到 Google Apps Script 覆蓋。
 * 3. 點「執行 (Run)」。
 * 👉 試算表所有數據瞬間歸零，隨時準備正式上線！
 */

function runClearAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let formSheet = null;
  for (let s of ss.getSheets()) {
    if (s.getName().includes("表單回應") || s.getName().includes("Form Responses")) {
      formSheet = s;
      break;
    }
  }

  if (!formSheet) {
    Logger.log("❌ 找不到表單回應工作表！");
    return;
  }

  const lastRow = formSheet.getLastRow();
  if (lastRow > 1) {
    formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).clearContent();
    Logger.log("🗑️ 已成功清空所有回報資料！戰情看板已全數歸零。");
  } else {
    Logger.log("ℹ️ 資料表本來就是空的，無需清除。");
  }
}
