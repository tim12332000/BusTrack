/**
 * =========================================================================
 * 🚌 接駁車管理系統 - 【動態比例進度條 (Sparkline) 戰情版】
 * =========================================================================
 * 
 * ✨ 本版視覺與數據升級：
 * 1. 【各站獨立進度條 (Sparkline)】：
 *    - 👉 去程發車比例條：自動計算「已發車輛數 / 總車輛數」，即時動態長條圖與百分比！
 *    - 👈 返程發車比例條：自動計算「已回程車數 / 總車輛數」，即時動態長條圖與百分比！
 * 2. 【各站小計與頂部全場總大盤】：去程、返程、總人次一目了然！
 * 
 * 👉 使用方式：
 * 1. 按 Ctrl + A 全選複製。
 * 2. 到 Google Apps Script 編輯器全選貼上覆蓋。
 * 3. 點擊上方的「執行 (Run)」按鈕即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🚀 開始建立【動態比例進度條 戰情看板】...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🚌 接駁車【去程 / 返程】搭乘人數即時回報");
  form.setDescription("請車長於各趟次發車時回報搭乘人數。\n回報流程：1. 去/返程 ➔ 2. 站點與車號 ➔ 3. 搭乘人數");

  // 題目 1: 行程方向
  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向")
    .setChoiceValues(["👉 去程", "👈 返程"])
    .setRequired(true);

  // 題目 2: 接駁站點
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
    dashboardSheet.getRange(1, 1, 80, 25).breakApart();
  } catch (e) {}

  // --- A. 頂部三大超醒目 KPI 戰情大卡片 ---
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總運輸人次").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#020617").setFontColor("#38BDF8").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 全場去程總人數").setBackground("#1E3A8A").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=B6+G6+L6+Q6").setBackground("#172554").setFontColor("#60A5FA").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 全場返程總人數").setBackground("#064E3B").setFontColor("#6EE7B7").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=D6+I6+N6+S6").setBackground("#022C22").setFontColor("#34D399").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 全場總發車進度條 (第 4 列)
  dashboardSheet.getRange("A4:C4").merge().setValue("⚡ 全場車輛進度: 130 輛").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D4:F4").merge().setFormula(`=SPARKLINE(COUNTIF(B9:B28, ">0")+COUNTIF(G9:G58, ">0")+COUNTIF(L9:L48, ">0")+COUNTIF(Q9:Q28, ">0"), {"charttype", "bar"; "max", 130; "color1", "#3B82F6"})`).setBackground("#0F172A");
  dashboardSheet.getRange("G4:I4").merge().setFormula(`=SPARKLINE(COUNTIF(D9:D28, ">0")+COUNTIF(I9:I58, ">0")+COUNTIF(N9:N48, ">0")+COUNTIF(S9:S28, ">0"), {"charttype", "bar"; "max", 130; "color1", "#10B981"})`).setBackground("#0F172A");

  // --- B. 4 大站點戰情明細 (含每站獨立小計 + 比例進度條) ---
  const stations = [
    { name: "成功車站（綠線）", formKeyword: "成功車站", startCol: 1, busCount: 20, bg: "#059669" },
    { name: "新烏日台鐵站（藍線）", formKeyword: "新烏日", startCol: 6, busCount: 50, bg: "#2563EB" },
    { name: "經貿六停車場（黃線）", formKeyword: "經貿六", startCol: 11, busCount: 40, bg: "#D97706" },
    { name: "水湳轉運站（黃線）", formKeyword: "水湳", startCol: 16, busCount: 20, bg: "#D97706" }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    const startRow = 9;
    const endRow = startRow + st.busCount - 1;
    
    // 1. 路線站名大標題 (第 5 列)
    dashboardSheet.getRange(5, col, 1, 5).merge()
      .setValue(`📍 ${st.name}`)
      .setBackground(st.bg).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    // 2. 各站獨立小計人數 (第 6 列)
    const colGo = String.fromCharCode(64 + col + 1);
    const colBack = String.fromCharCode(64 + col + 3);

    dashboardSheet.getRange(6, col).setValue("📊 本站小計").setBackground("#1E293B").setFontColor("#F8FAFC").setFontWeight("bold").setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 1).setFormula(`=SUM(${colGo}${startRow}:${colGo}${endRow})`).setBackground("#172554").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 2).setValue("-").setBackground("#1E293B").setFontColor("#64748B").setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 3).setFormula(`=SUM(${colBack}${startRow}:${colBack}${endRow})`).setBackground("#022C22").setFontColor("#6EE7B7").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(6, col + 4).setFormula(`=${colGo}6+${colBack}6`).setBackground("#312E81").setFontColor("#FDE047").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 3. 🌟 各站專屬【去程 / 返程 比例進度條】(第 7 列)
    dashboardSheet.getRange(7, col).setValue("📈 發車進度").setBackground("#1E293B").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    
    // 👉 去程比例與 Sparkline 進度條
    dashboardSheet.getRange(7, col + 1).setFormula(`=TEXT(COUNTIF(${colGo}${startRow}:${colGo}${endRow}, ">0")/${st.busCount}, "0.0%")`).setBackground("#172554").setFontColor("#60A5FA").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 2).setFormula(`=SPARKLINE(COUNTIF(${colGo}${startRow}:${colGo}${endRow}, ">0"), {"charttype", "bar"; "max", ${st.busCount}; "color1", "#3B82F6"})`).setBackground("#0F172A");

    // 👈 返程比例與 Sparkline 進度條
    dashboardSheet.getRange(7, col + 3).setFormula(`=TEXT(COUNTIF(${colBack}${startRow}:${colBack}${endRow}, ">0")/${st.busCount}, "0.0%")`).setBackground("#022C22").setFontColor("#34D399").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 4).setFormula(`=SPARKLINE(COUNTIF(${colBack}${startRow}:${colBack}${endRow}, ">0"), {"charttype", "bar"; "max", ${st.busCount}; "color1", "#10B981"})`).setBackground("#0F172A");

    // 4. 子表頭 (第 8 列)
    dashboardSheet.getRange(8, col, 1, 5)
      .setValues([["車號", "👉去程人數", "去程時間", "👈返程人數", "返程時間"]])
      .setBackground("#334155").setFontColor("#E2E8F0").setFontWeight("bold").setHorizontalAlignment("center");

    // 5. 每台車公式 (第 9 列開始)
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
    dashboardSheet.getRange(6, col, st.busCount + 3, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);
    
    dashboardSheet.setColumnWidth(col, 75);
    dashboardSheet.setColumnWidth(col + 1, 85);
    dashboardSheet.setColumnWidth(col + 2, 85);
    dashboardSheet.setColumnWidth(col + 3, 85);
    dashboardSheet.setColumnWidth(col + 4, 90);
  });

  dashboardSheet.setRowHeight(1, 28);
  dashboardSheet.setRowHeight(2, 28);
  dashboardSheet.setRowHeight(3, 28);
  dashboardSheet.setRowHeight(4, 22);
  dashboardSheet.setRowHeight(5, 32);
  dashboardSheet.setRowHeight(6, 30);
  dashboardSheet.setRowHeight(7, 24);
  dashboardSheet.setRowHeight(8, 26);
  for (let r = 9; r <= 58; r++) { dashboardSheet.setRowHeight(r, 24); }

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新動態進度條版 戰情看板建立完成！】");
  Logger.log("\n📱【最新車長填寫表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
