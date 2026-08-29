/**
 * 🎲 一鍵產生「國防知性之旅-成功嶺營區開放」即時測試資料 (進場 / 離場)
 */
function runGenerateTestData() {
  function getTargetSpreadsheet() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
    const files = DriveApp.getFilesByName("「國防知性之旅-成功嶺營區開放」即時戰情中心");
    if (files.hasNext()) return SpreadsheetApp.open(files.next());
    const fallback = DriveApp.searchFiles('title contains "即時戰情中心" and mimeType = "application/vnd.google-apps.spreadsheet"');
    if (fallback.hasNext()) return SpreadsheetApp.open(fallback.next());
    return null;
  }

  const ss = getTargetSpreadsheet();
  if (!ss) {
    Logger.log("❌ 找不到戰情中心試算表！請先執行一次 `createMultiStationBusSystem` 建立系統。");
    return;
  }

  let formSheet = null;
  for (let s of ss.getSheets()) {
    if (s.getName().includes("表單回應") || s.getName().includes("Form Responses")) {
      formSheet = s;
      break;
    }
  }

  if (!formSheet) {
    Logger.log("❌ 找不到表單回應工作表！請確認試算表已綁定表單。");
    return;
  }

  const testRows = [];
  const now = new Date();
  const datePrefix = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd");

  // 1. 接駁車測試數據 (進場 / 離場)
  const stList = [
    { name: "🚌 成功車站", count: 20 },
    { name: "🚌 新烏日台鐵站", count: 50 },
    { name: "🚌 經貿六停車場", count: 40 },
    { name: "🚌 水湳轉運站", count: 20 }
  ];

  stList.forEach(st => {
    for (let i = 1; i <= st.count; i++) {
      const busName = `${i} 號車`;
      const goPax = Math.floor(Math.random() * 15) + 25; // 25~40人
      testRows.push([`${datePrefix} 08:${String(10 + (i%40)).padStart(2,'0')}:15`, "進場", st.name, busName, goPax, ""]);
      
      if (Math.random() > 0.15) {
        const backPax = Math.floor(goPax * (0.8 + Math.random() * 0.2));
        testRows.push([`${datePrefix} 17:${String(20 + (i%35)).padStart(2,'0')}:30`, "離場", st.name, busName, backPax, ""]);
      }
    }
  });

  // 2. 步行通道測試數據 (進場 / 離場)
  const gates = ["🚶 1號門", "🚶 3號門", "🚶 4號門"];
  gates.forEach(gate => {
    for (let p = 1; p <= 4; p++) {
      const walkIn = Math.floor(Math.random() * 80) + 120;
      testRows.push([`${datePrefix} 09:${String(10 + p*15).padStart(2,'0')}:00`, "進場", gate, "🚶 步行通道", walkIn, ""]);
    }
    for (let p = 1; p <= 4; p++) {
      const walkOut = Math.floor(Math.random() * 70) + 100;
      testRows.push([`${datePrefix} 18:${String(5 + p*15).padStart(2,'0')}:00`, "離場", gate, "🚶 步行通道", walkOut, ""]);
    }
  });

  formSheet.getRange(formSheet.getLastRow() + 1, 1, testRows.length, 6).setValues(testRows);
  Logger.log("🎉 已成功產生 " + testRows.length + " 筆測試數據！請回到「總即時戰情看板」查看數據！");
}
