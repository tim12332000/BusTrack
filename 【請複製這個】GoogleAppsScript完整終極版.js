/**
 * =========================================================================
 * 🚌 接駁車管理系統 - 【現代極簡戰情看板 + 一鍵測試/清空按鈕版】
 * =========================================================================
 * 
 * ✨ 本版新功能：
 * 1. 【上方選單 🛠️ 戰情控制台】：
 *    - 點選上方「🛠️ 戰情控制台」➔「🎲 產生模擬測試資料」或「🗑️ 清空所有回報資料 (歸零)」
 * 2. 【右下角實體按鈕引導區】：
 *    - 右下角設有按鈕引導區，方便主控人員隨時一鍵測試與上線前清空歸零！
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行 (Run)」即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🎨 開始建立【現代極簡高階戰情系統】...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🚌 接駁車【去程 / 返程】搭乘人數即時回報");
  form.setDescription("請車長於各趟次發車時回報搭乘人數。\n回報流程：1. 去/返程 ➔ 2. 站點與車號 ➔ 3. 搭乘人數");

  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向")
    .setChoiceValues(["👉 去程", "👈 返程"])
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle("2. 選擇接駁站點")
    .setChoiceValues([
      "成功車站（綠線）",
      "新烏日台鐵站（藍線）",
      "經貿六停車場（黃線）",
      "水湳轉運站（黃線）"
    ])
    .setRequired(true);

  const busChoices = [];
  for (let i = 1; i <= 50; i++) {
    busChoices.push(`${i} 號車`);
  }
  form.addListItem()
    .setTitle("3. 選擇車號")
    .setChoiceValues(busChoices)
    .setRequired(true);

  const textValidation = FormApp.createTextValidation()
    .setHelpText("請輸入大於等於 0 的數字")
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem()
    .setTitle("4. 搭乘人數")
    .setValidation(textValidation)
    .setRequired(true);

  form.addTextItem()
    .setTitle("5. 備註 (若有特殊狀況填寫)")
    .setRequired(false);

  // 2. 建立 Google 試算表並綁定
  const ss = SpreadsheetApp.create("🚌 各站接駁車【去/返程】即時戰情中心");
  const ssId = ss.getId();
  const ssUrl = ss.getUrl();

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  Utilities.sleep(2000);

  // 3. 建立 2 個工作表：首頁【總即時戰情看板】 + 第二頁【各車即時明細】
  let dashboardSheet = ss.getSheets()[0];
  dashboardSheet.setName("總即時戰情看板");

  let detailSheet = ss.insertSheet("各車即時明細");

  let formSheetName = "表單回應 1";
  for (let s of ss.getSheets()) {
    if (s.getName() !== "總即時戰情看板" && s.getName() !== "各車即時明細") {
      formSheetName = s.getName();
      break;
    }
  }

  // ==========================================
  // 【A. 第二頁：各車即時明細 (底層計算清單)】
  // ==========================================
  const stations = [
    { name: "成功車站（綠線）", formKeyword: "成功車站", startCol: 1, busCount: 20 },
    { name: "新烏日台鐵站（藍線）", formKeyword: "新烏日", startCol: 6, busCount: 50 },
    { name: "經貿六停車場（黃線）", formKeyword: "經貿六", startCol: 11, busCount: 40 },
    { name: "水湳轉運站（黃線）", formKeyword: "水湳", startCol: 16, busCount: 20 }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    detailSheet.getRange(1, col, 1, 5).merge()
      .setValue(`📍 ${st.name}`)
      .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

    detailSheet.getRange(2, col, 1, 5)
      .setValues([["車號", "👉去程人數", "去程時間", "👈返程人數", "返程時間"]])
      .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setHorizontalAlignment("center");

    const busRows = [];
    for (let b = 1; b <= st.busCount; b++) {
      const busName = `${b} 號車`;
      const fGoActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), '${formSheetName}'!B:B="👉 去程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      const fGoTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), '${formSheetName}'!B:B="👉 去程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      const fBackActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), '${formSheetName}'!B:B="👈 返程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      const fBackTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), '${formSheetName}'!B:B="👈 返程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      busRows.push([busName, fGoActual, fGoTime, fBackActual, fBackTime]);
    }
    detailSheet.getRange(3, col, st.busCount, 5).setValues(busRows);
    detailSheet.getRange(2, col, st.busCount + 1, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // ==========================================
  // 【B. 第一頁：總即時戰情看板 (簡約高階設計)】
  // ==========================================
  try {
    dashboardSheet.getRange(1, 1, 40, 20).breakApart();
  } catch (e) {}

  // 1. 頂部全場總大盤
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總運輸人次").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#0F172A").setFontColor("#38BDF8").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 全場去程總人數").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=B8+E8+H8+K8").setBackground("#0F172A").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 全場返程總人數").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=C8+F8+I8+L8").setBackground("#0F172A").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("J1:L1").merge().setValue("📈 全場返程疏運率").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("J2:L3").merge().setFormula(`=IF(D2>0, TEXT(G2/D2, "0.0%"), "0.0%")`).setBackground("#0F172A").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 全場疏運進度條 (第 4 列)
  dashboardSheet.getRange("A4:C4").merge().setValue("⚡ 總疏運進度").setBackground("#1E293B").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D4:L4").merge().setFormula(`=IF(D2>0, SPARKLINE(G2, {"charttype", "bar"; "max", D2; "color1", "#38BDF8"}), "")`).setBackground("#1E293B");

  // 2. 4 大站點獨立現代白底卡片 (第 6~10 列)
  const cardStations = [
    { name: "📍 成功車站（綠線）", startCol: 1, detailGoCol: "B", detailBackCol: "D", detailEndRow: 22, tagColor: "#059669" },
    { name: "📍 新烏日台鐵站（藍線）", startCol: 4, detailGoCol: "G", detailBackCol: "I", detailEndRow: 52, tagColor: "#2563EB" },
    { name: "📍 經貿六停車場（黃線）", startCol: 7, detailGoCol: "L", detailBackCol: "N", detailEndRow: 42, tagColor: "#D97706" },
    { name: "📍 水湳轉運站（黃線）", startCol: 10, detailGoCol: "Q", detailBackCol: "S", detailEndRow: 22, tagColor: "#D97706" }
  ];

  cardStations.forEach(cs => {
    const col = cs.startCol;
    
    // 站名大標題
    dashboardSheet.getRange(6, col, 1, 3).merge()
      .setValue(cs.name)
      .setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");

    // 項目標籤
    dashboardSheet.getRange(7, col).setValue("👉 去程人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 1).setValue("👈 返程人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 2).setValue("🏆 累計總量").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列
    const colLetterGo = String.fromCharCode(64 + col);
    const colLetterBack = String.fromCharCode(64 + col + 1);

    dashboardSheet.getRange(8, col).setFormula(`=SUM('各車即時明細'!${cs.detailGoCol}3:${cs.detailGoCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 1).setFormula(`=SUM('各車即時明細'!${cs.detailBackCol}3:${cs.detailBackCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 2).setFormula(`=${colLetterGo}8+${colLetterBack}8`).setBackground("#FFFFFF").setFontColor("#2563EB").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");

    // 疏運率標題
    dashboardSheet.getRange(9, col, 1, 2).merge()
      .setFormula(`=IF(${colLetterGo}8>0, "📈 返程疏運率: " & TEXT(${colLetterBack}8/${colLetterGo}8, "0.0%"), "📈 返程疏運率: 0.0%")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    dashboardSheet.getRange(9, col + 2).setFormula(`=IF(${colLetterGo}8>0, "尚餘 " & MAX(0, ${colLetterGo}8 - ${colLetterBack}8) & " 人", "已完成")`)
      .setBackground("#F8FAFC").setFontColor("#EF4444").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條
    dashboardSheet.getRange(10, col, 1, 3).merge()
      .setFormula(`=IF(${colLetterGo}8>0, SPARKLINE(${colLetterBack}8, {"charttype", "bar"; "max", ${colLetterGo}8; "color1", "#2563EB"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(6, col, 5, 3).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    
    dashboardSheet.setColumnWidth(col, 115);
    dashboardSheet.setColumnWidth(col + 1, 115);
    dashboardSheet.setColumnWidth(col + 2, 115);
  });

  // 3. 🌟 右下角操作按鈕卡片 (第 12~14 列)
  dashboardSheet.getRange("J12:L12").merge()
    .setValue("🎲 產生模擬測試資料 (上方選單執行)")
    .setBackground("#2563EB").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

  dashboardSheet.getRange("J13:L13").merge()
    .setValue("🗑️ 一鍵清空所有資料 (上方選單執行)")
    .setBackground("#DC2626").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

  dashboardSheet.getRange("J14:L14").merge()
    .setValue("💡 提示：點上方選單「🛠️ 戰情控制台」即可一鍵操作！")
    .setFontColor("#64748B").setFontSize(10).setHorizontalAlignment("center");

  // 列高設定
  dashboardSheet.setRowHeight(1, 26);
  dashboardSheet.setRowHeight(2, 28);
  dashboardSheet.setRowHeight(3, 28);
  dashboardSheet.setRowHeight(4, 24);
  dashboardSheet.setRowHeight(5, 16);
  dashboardSheet.setRowHeight(6, 36);
  dashboardSheet.setRowHeight(7, 28);
  dashboardSheet.setRowHeight(8, 42);
  dashboardSheet.setRowHeight(9, 28);
  dashboardSheet.setRowHeight(10, 20);
  dashboardSheet.setRowHeight(11, 14);
  dashboardSheet.setRowHeight(12, 32);
  dashboardSheet.setRowHeight(13, 32);
  dashboardSheet.setRowHeight(14, 24);

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新現代極簡戰情看板建立完成！】");
  Logger.log("\n📱【最新車長填寫表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

// ==========================================
// 🛠️ 選單與按鈕動作函式
// ==========================================

// 自動在試算表上方增加自訂選單
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛠️ 戰情控制台")
    .addItem("🎲 產生模擬測試資料", "generateTestData")
    .addItem("🗑️ 清空所有回報資料 (歸零)", "clearAllData")
    .addToUi();
}

// 🎲 產生 4 站隨機測試資料
function generateTestData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GF3Y9oR_KidYlfdnQwkSF_xCRgElAlAC1PwJG8FJSrU/edit");
  let formSheet = null;
  for (let s of ss.getSheets()) {
    if (s.getName().includes("表單回應") || s.getName().includes("Form Responses")) {
      formSheet = s;
      break;
    }
  }

  if (!formSheet) {
    SpreadsheetApp.getUi().alert("❌ 找不到表單回應工作表！");
    return;
  }

  const testRows = [];
  const now = new Date();
  const datePrefix = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd");

  const stList = [
    { name: "成功車站（綠線）", count: 20 },
    { name: "新烏日台鐵站（藍線）", count: 50 },
    { name: "經貿六停車場（黃線）", count: 40 },
    { name: "水湳轉運站（黃線）", count: 20 }
  ];

  stList.forEach(st => {
    for (let i = 1; i <= st.count; i++) {
      const busName = `${i} 號車`;
      // 去程
      const goPax = Math.floor(Math.random() * 15) + 25; // 25~40人
      testRows.push([`${datePrefix} 08:${String(10 + (i%40)).padStart(2,'0')}:15`, "👉 去程", st.name, busName, goPax, ""]);
      
      // 返程 (約 85%~95% 車輛已返程)
      if (Math.random() > 0.1) {
        const backPax = Math.floor(goPax * (0.85 + Math.random() * 0.15));
        testRows.push([`${datePrefix} 17:${String(20 + (i%35)).padStart(2,'0')}:30`, "👈 返程", st.name, busName, backPax, ""]);
      }
    }
  });

  formSheet.getRange(formSheet.getLastRow() + 1, 1, testRows.length, 6).setValues(testRows);
  
  try {
    SpreadsheetApp.getUi().alert("🎉 已成功隨機產生 4 站共 " + testRows.length + " 筆去/返程測試數據！");
  } catch(e) {
    Logger.log("✅ 已成功產生測試數據 " + testRows.length + " 筆！");
  }
}

// 🗑️ 清空所有流水帳回報資料
function clearAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1GF3Y9oR_KidYlfdnQwkSF_xCRgElAlAC1PwJG8FJSrU/edit");
  let formSheet = null;
  for (let s of ss.getSheets()) {
    if (s.getName().includes("表單回應") || s.getName().includes("Form Responses")) {
      formSheet = s;
      break;
    }
  }

  if (!formSheet) return;

  const lastRow = formSheet.getLastRow();
  if (lastRow > 1) {
    formSheet.getRange(2, 1, lastRow - 1, formSheet.getLastColumn()).clearContent();
  }

  try {
    SpreadsheetApp.getUi().alert("🗑️ 所有回報資料已全數清空歸零，戰情看板已重置完成！");
  } catch(e) {
    Logger.log("✅ 所有回報資料已清空！");
  }
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
