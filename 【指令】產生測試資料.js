/**
 * 🎲【指令：一鍵產生 車 ABCD + 人 ABC 測試資料】
 * 
 * 使用方式：
 * 1. 複製這段程式碼。
 * 2. 貼到 Google Apps Script 執行 `runGenerateTestData`。
 * 👉 試算表就會立刻填滿 4 站接駁車 + 3 大步行門去/返程數據與動態條！
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

  // 1. 🚌 接駁車測試數據 (車 ABCD：成功 20、新烏日 50、經貿六 40、水湳 20)
  const stList = [
    { name: "🚌 成功車站（綠線）", count: 20 },
    { name: "🚌 新烏日台鐵站（藍線）", count: 50 },
    { name: "🚌 經貿六停車場（黃線）", count: 40 },
    { name: "🚌 水湳轉運站（黃線）", count: 20 }
  ];

  stList.forEach(st => {
    for (let i = 1; i <= st.count; i++) {
      const busName = `${i} 號車`;
      const goPax = Math.floor(Math.random() * 15) + 25; // 25~40人
      testRows.push([`${datePrefix} 08:${String(10 + (i%40)).padStart(2,'0')}:15`, "👉 去程 (入場)", st.name, busName, goPax, ""]);
      
      // 85% 機率已返程
      if (Math.random() > 0.15) {
        const backPax = Math.floor(goPax * (0.8 + Math.random() * 0.2));
        testRows.push([`${datePrefix} 17:${String(20 + (i%35)).padStart(2,'0')}:30`, "👈 返程 (離場)", st.name, busName, backPax, ""]);
      }
    }
  });

  // 2. 🚶 步行通道測試數據 (人 ABC：1號門綠、3號門藍、4號門黃)
  const gates = ["🚶 1號門（綠線）", "🚶 3號門（藍線）", "🚶 4號門（黃線）"];
  gates.forEach(gate => {
    // 模擬 4 批入場
    for (let p = 1; p <= 4; p++) {
      const walkIn = Math.floor(Math.random() * 80) + 120; // 120~200人
      testRows.push([`${datePrefix} 09:${String(10 + p*15).padStart(2,'0')}:00`, "👉 去程 (入場)", gate, "🚶 步行通道 (無車號)", walkIn, "入場批次"]);
    }
    // 模擬 4 批離場
    for (let p = 1; p <= 4; p++) {
      const walkOut = Math.floor(Math.random() * 70) + 100; // 100~170人
      testRows.push([`${datePrefix} 18:${String(5 + p*15).padStart(2,'0')}:00`, "👈 返程 (離場)", gate, "🚶 步行通道 (無車號)", walkOut, "離場批次"]);
    }
  });

  formSheet.getRange(formSheet.getLastRow() + 1, 1, testRows.length, 6).setValues(testRows);
  Logger.log("🎉 已成功產生 " + testRows.length + " 筆測試數據！請回到「總即時戰情看板」查看數據！");
}
