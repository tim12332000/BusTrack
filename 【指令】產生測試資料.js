/**
 * 🎲 一鍵產生「國防知性之旅-成功嶺營區開放」即時測試資料 (雙表單整合版)
 * 包含：接駁車進離場、步行通道進離場、以及七處停車場剩餘車位回報
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

  const sheets = ss.getSheets();
  let peopleSheet = null;
  let parkingSheet = null;

  for (let s of sheets) {
    const name = s.getName();
    if (name.includes("表單回應 1") || (name.includes("表單回應") && !peopleSheet)) {
      peopleSheet = s;
    }
    if (name.includes("表單回應 2") || name.includes("停車場車位回報")) {
      parkingSheet = s;
    }
  }

  const now = new Date();
  const datePrefix = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd");
  const timePrefix = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");

  // ==========================================
  // 1. 接駁車與步行通道測試數據 (寫入表單回應 1)
  // ==========================================
  if (peopleSheet) {
    const peopleRows = [];
    const stList = [
      { name: "🚌 成功車站", count: 20 },
      { name: "🚌 新烏日台鐵站", count: 50 },
      { name: "🚌 經貿六停車場", count: 40 },
      { name: "🚌 水湳轉運站", count: 20 }
    ];

    stList.forEach(st => {
      for (let i = 1; i <= Math.min(st.count, 5); i++) {
        const busName = `${i} 號車`;
        const goPax = Math.floor(Math.random() * 15) + 30; // 30~45人
        peopleRows.push([`${datePrefix} 08:${String(10 + (i%40)).padStart(2,'0')}:15`, "進場", st.name, busName, goPax, "測試資料"]);
        
        const backPax = Math.floor(goPax * (0.6 + Math.random() * 0.3));
        peopleRows.push([`${datePrefix} 17:${String(20 + (i%35)).padStart(2,'0')}:30`, "離場", st.name, busName, backPax, "測試資料"]);
      }
    });

    const gates = ["🚶 1號門", "🚶 3號門", "🚶 4號門"];
    gates.forEach(gate => {
      for (let p = 1; p <= 3; p++) {
        const walkIn = Math.floor(Math.random() * 80) + 150;
        peopleRows.push([`${datePrefix} 09:${String(10 + p*15).padStart(2,'0')}:00`, "進場", gate, "🚶 步行通道", walkIn, "測試資料"]);
        const walkOut = Math.floor(Math.random() * 60) + 100;
        peopleRows.push([`${datePrefix} 18:${String(5 + p*15).padStart(2,'0')}:00`, "離場", gate, "🚶 步行通道", walkOut, "測試資料"]);
      }
    });

    peopleSheet.getRange(peopleSheet.getLastRow() + 1, 1, peopleRows.length, 6).setValues(peopleRows);
    Logger.log(`🎉 [人員疏運] 已成功產生 ${peopleRows.length} 筆測試數據！`);
  }

  // ==========================================
  // 2. 七處停車場即時剩餘車位測試數據 (寫入表單回應 2)
  // ==========================================
  if (parkingSheet) {
    const parkingLots = [
      { name: "五都日出", cars: 650, motos: 80 },
      { name: "新烏日", cars: 140, motos: 210 },
      { name: "*嶺東科大", cars: 35, motos: 520 },
      { name: "*台中科大", cars: 0, motos: 110 },
      { name: "水湳轉運站", cars: 380, motos: 860 },
      { name: "*經貿六", cars: 310, motos: 0 },
      { name: "*經貿八", cars: 290, motos: 0 }
    ];

    const parkingRows = parkingLots.map((lot, idx) => {
      const timeStr = `${datePrefix} 14:${String(10 + idx * 5).padStart(2, '0')}:00`;
      return [timeStr, lot.name, lot.cars, lot.motos, "定時巡檢回報"];
    });

    parkingSheet.getRange(parkingSheet.getLastRow() + 1, 1, parkingRows.length, 5).setValues(parkingRows);
    Logger.log(`🎉 [停車場] 已成功產生 7 處停車場即時剩餘車位數據！`);
  }

  Logger.log("✨ 全部測試資料已成功寫入，請回到「總即時戰情看板」查看數據！");
}
