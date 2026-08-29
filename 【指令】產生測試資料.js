/**
 * 🎲【指令：一鍵產生 4 大車站 + 3 大步行門 測試資料】
 * 
 * 使用方式：
 * 1. 複製這段程式碼。
 * 2. 貼到 Google Apps Script 覆蓋。
 * 3. 點「執行 (Run)」。
 * 👉 試算表就會立刻填滿去程、返程、步行門數據與動態進度條！
 */

function runGenerateTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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

  // 1. 車站測試數據 (成功 20、新烏日 50、經貿六 40、水湳 20)
  const stList = [
    { name: "成功車站（綠線）", count: 20 },
    { name: "新烏日台鐵站（藍線）", count: 50 },
    { name: "經貿六停車場（黃線）", count: 40 },
    { name: "水湳轉運站（黃線）", count: 20 }
  ];

  stList.forEach(st => {
    for (let i = 1; i <= st.count; i++) {
      const busName = `${i} 號車`;
      const goPax = Math.floor(Math.random() * 15) + 25; // 25~40人
      testRows.push([`${datePrefix} 08:${String(10 + (i%40)).padStart(2,'0')}:15`, "👉 去程 (入場)", st.name, busName, goPax, ""]);
      
      // 90% 機率已返程
      if (Math.random() > 0.1) {
        const backPax = Math.floor(goPax * (0.85 + Math.random() * 0.15));
        testRows.push([`${datePrefix} 17:${String(20 + (i%35)).padStart(2,'0')}:30`, "👈 返程 (離場)", st.name, busName, backPax, ""]);
      }
    }
  });

  // 2. 步行通道測試數據 (1號門、3號門、4號門)
  const gates = ["🚶 步行：1號門", "🚶 步行：3號門", "🚶 步行：4號門"];
  gates.forEach(gate => {
    for (let p = 1; p <= 3; p++) {
      const walkIn = Math.floor(Math.random() * 80) + 120;
      testRows.push([`${datePrefix} 09:${String(15 + p*15).padStart(2,'0')}:00`, "👉 去程 (入場)", gate, "🚶 步行通道 (無車號)", walkIn, "步行批次回報"]);
    }
    for (let p = 1; p <= 3; p++) {
      const walkOut = Math.floor(Math.random() * 70) + 110;
      testRows.push([`${datePrefix} 18:${String(10 + p*15).padStart(2,'0')}:00`, "👈 返程 (離場)", gate, "🚶 步行通道 (無車號)", walkOut, "步行離場回報"]);
    }
  });

  formSheet.getRange(formSheet.getLastRow() + 1, 1, testRows.length, 6).setValues(testRows);
  Logger.log("🎉 已成功產生 " + testRows.length + " 筆測試數據！請回到試算表查看戰情看板。");
}
