/**
 * =========================================================================
 * 🚌 接駁車管理系統 - 【路線色名稱客製版】
 * =========================================================================
 * 
 * 📍 4 大路線接駁站點：
 * 1. 成功車站（綠線）      (1 ~ 20 號車)
 * 2. 新烏日台鐵站（藍線）   (1 ~ 50 號車)
 * 3. 經貿六停車場（黃線）   (1 ~ 40 號車)
 * 4. 水湳轉運站（黃線）     (1 ~ 20 號車)
 * 
 * 👉 使用方式：
 * 1. 按 Ctrl + A 全選複製。
 * 2. 到 Google Apps Script 編輯器全選貼上覆蓋。
 * 3. 點擊上方的「執行 (Run)」按鈕即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🚀 開始建立【路線色名稱版 接駁戰情系統】...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🚌 接駁車【去程 / 返程】搭乘人數即時回報");
  form.setDescription("請車長於各趟次發車時回報搭乘人數。\n回報流程：1. 去/返程 ➔ 2. 站點與車號 ➔ 3. 搭乘人數");

  // 題目 1: 行程方向
  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向")
    .setChoiceValues(["👉 去程", "👈 返程"])
    .setRequired(true);

  // 題目 2: 接駁站點 (含綠線、藍線、黃線)
  form.addMultipleChoiceItem()
    .setTitle("2. 選擇接駁站點")
    .setChoiceValues([
      "成功車站（綠線）",
      "新烏日台鐵站（藍線）",
      "經貿六停車場（黃線）",
      "水湳轉運站（黃線）"
    ])
    .setRequired(true);

  // 題目 3: 車號
  const busChoices = [];
  for (let i = 1; i <= 50; i++) {
    busChoices.push(`${i} 號車`);
  }
  form.addListItem()
    .setTitle("3. 選擇車號")
    .setChoiceValues(busChoices)
    .setRequired(true);

  // 題目 4: 搭乘人數
  const textValidation = FormApp.createTextValidation()
    .setHelpText("請輸入大於等於 0 的數字")
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem()
    .setTitle("4. 搭乘人數")
    .setValidation(textValidation)
    .setRequired(true);

  // 題目 5: 備註
  form.addTextItem()
    .setTitle("5. 備註 (若有特殊狀況填寫)")
    .setRequired(false);

  // 2. 建立 Google 試算表並綁定
  const ss = SpreadsheetApp.create("🚌 各站接駁車【去/返程】即時戰情中心");
  const ssId = ss.getId();
  const ssUrl = ss.getUrl();

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  Utilities.sleep(2000);

  // 3. 設定戰情看板
  let dashboardSheet = ss.getSheets()[0];
  dashboardSheet.setName("總即時戰情看板");

  let formSheetName = "表單回應 1";
  for (let s of ss.getSheets()) {
    if (s.getName() !== "總即時戰情看板") {
      formSheetName = s.getName();
      break;
    }
  }

  try {
    dashboardSheet.getRange(1, 1, 70, 25).breakApart();
  } catch (e) {}

  // --- A. 頂部三大超醒目 KPI 戰情大卡片 ---
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總運輸人次").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#020617").setFontColor("#38BDF8").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 全場去程總人數").setBackground("#1E3A8A").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=B6+G6+L6+Q6").setBackground("#172554").setFontColor("#60A5FA").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 全場返程總人數").setBackground("#064E3B").setFontColor("#6EE7B7").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=D6+I6+N6+S6").setBackground("#022C22").setFontColor("#34D399").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // --- B. 4 大站點戰情明細 (成功-綠線 / 新烏日-藍線 / 經貿六-黃線 / 水湳-黃線) ---
  const stations = [
    { name: "成功車站（綠線）", formKeyword: "成功車站", startCol: 1, busCount: 20, bg: "#059669" },
    { name: "新烏日台鐵站（藍線）", formKeyword: "新烏日", startCol: 6, busCount: 50, bg: "#2563EB" },
    { name: "經貿六停車場（黃線）", formKeyword: "經貿六", startCol: 11, busCount: 40, bg: "#D97706" },
    { name: "水湳轉運站（黃線）", formKeyword: "水湳", startCol: 16, busCount: 20, bg: "#D97706" }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    const startRow = 8;
    const endRow = startRow + st.busCount - 1;
    
    // 1. 路線站名大標題 (第 5 列)
    dashboardSheet.getRange(5, col, 1, 5).merge()
      .setValue(`📍 ${st.name}`)
      .setBackground(st.bg).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    // 2. 各站獨立小計 (第 6 列)
    const colGo = String.fromCharCode(64 + col + 1);
    const colBack = String.fromCharCode(64 + col + 3);

    dashboardSheet.getRange(6, col).setValue("📊 本站小計").setBackground("#1E293B").setFontColor("#F8FAFC").setFontWeight("bold").setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 1).setFormula(`=SUM(${colGo}${startRow}:${colGo}${endRow})`).setBackground("#172554").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 2).setValue("-").setBackground("#1E293B").setFontColor("#64748B").setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 3).setFormula(`=SUM(${colBack}${startRow}:${colBack}${endRow})`).setBackground("#022C22").setFontColor("#6EE7B7").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 4).setFormula(`=${colGo}6+${colBack}6`).setBackground("#312E81").setFontColor("#FDE047").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 3. 子表頭 (第 7 列)
    dashboardSheet.getRange(7, col, 1, 5)
      .setValues([["車號", "👉去程人數", "去程時間", "👈返程人數", "返程時間"]])
      .setBackground("#334155").setFontColor("#E2E8F0").setFontWeight("bold").setHorizontalAlignment("center");

    // 4. 每台車公式 (第 8 列開始)
    const busRows = [];
    for (let b = 1; b <= st.busCount; b++) {
      const busName = `${b} 號車`;
      
      // 去程人數
      const fGoActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), '${formSheetName}'!B:B="👉 去程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      // 去程時間
      const fGoTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), '${formSheetName}'!B:B="👉 去程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      
      // 返程人數
      const fBackActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), '${formSheetName}'!B:B="👈 返程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      // 返程時間
      const fBackTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), '${formSheetName}'!B:B="👈 返程", ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;

      busRows.push([busName, fGoActual, fGoTime, fBackActual, fBackTime]);
    }
    dashboardSheet.getRange(startRow, col, st.busCount, 5).setValues(busRows);
    dashboardSheet.getRange(6, col, st.busCount + 2, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);
    
    dashboardSheet.setColumnWidth(col, 75);
    dashboardSheet.setColumnWidth(col + 1, 85);
    dashboardSheet.setColumnWidth(col + 2, 85);
    dashboardSheet.setColumnWidth(col + 3, 85);
    dashboardSheet.setColumnWidth(col + 4, 90);
  });

  dashboardSheet.setRowHeight(1, 28);
  dashboardSheet.setRowHeight(2, 28);
  dashboardSheet.setRowHeight(3, 28);
  dashboardSheet.setRowHeight(5, 32);
  dashboardSheet.setRowHeight(6, 30);
  dashboardSheet.setRowHeight(7, 26);
  for (let r = 8; r <= 57; r++) { dashboardSheet.setRowHeight(r, 24); }

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新路線色版 戰情看板建立完成！】");
  Logger.log("\n📱【最新車長填寫表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
