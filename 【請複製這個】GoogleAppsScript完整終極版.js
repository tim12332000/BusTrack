/**
 * =========================================================================
 * 🚌 接駁車管理系統 - 【依最新回報流程客製版】
 * =========================================================================
 * 
 * 📋 最新回報流程順序：
 * 1. 【行程方向】：👉 去程 / 👈 返程
 * 2. 【接駁站點】：成功車站(20台) / 新烏日(50台) / 經貿六(40台) / 水湳(20台)
 * 3. 【車號】：1 ~ 50 號車 (下拉選單防呆)
 * 4. 【搭乘人數】：數字 (大於等於 0)
 * 5. 【備註說明】：選填
 * 
 * 👉 使用方式：
 * 1. 按 Ctrl + A 全選複製。
 * 2. 到 Google Apps Script 編輯器全選貼上覆蓋。
 * 3. 點擊「執行 (Run)」。
 * 4. 下方紀錄會直接產生【全新表單網址】與【試算表看板網址】！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🚀 開始依最新回報流程建立系統...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🚌 接駁車【去程 / 返程】搭乘人數即時回報");
  form.setDescription("請車長於各趟次發車時回報搭乘人數。\n回報流程：1. 去/返程 ➔ 2. 站點與車號 ➔ 3. 搭乘人數");

  // 【步驟 1】：行程方向 (去程 / 返程)
  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向")
    .setChoiceValues(["👉 去程", "👈 返程"])
    .setRequired(true);

  // 【步驟 2-1】：接駁站點
  form.addMultipleChoiceItem()
    .setTitle("2. 選擇接駁站點")
    .setChoiceValues([
      "成功車站 (1~20號車)",
      "新烏日台鐵站 (1~50號車)",
      "經貿六停車場 (1~40號車)",
      "水湳轉運站 (1~20號車)"
    ])
    .setRequired(true);

  // 【步驟 2-2】：車號 (1 ~ 50 號車)
  const busChoices = [];
  for (let i = 1; i <= 50; i++) {
    busChoices.push(`${i} 號車`);
  }
  form.addListItem()
    .setTitle("3. 選擇車號")
    .setChoiceValues(busChoices)
    .setRequired(true);

  // 【步驟 3】：搭乘人數
  const textValidation = FormApp.createTextValidation()
    .setHelpText("請輸入大於等於 0 的數字")
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem()
    .setTitle("4. 搭乘人數")
    .setValidation(textValidation)
    .setRequired(true);

  // 【步驟 4】：備註
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

  // 拆散舊儲存格防衝突
  try {
    dashboardSheet.getRange(1, 1, 60, 25).breakApart();
  } catch (e) {}

  // --- A. 頂部三大 KPI 卡片 ---
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總運輸人次").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#020617").setFontColor("#38BDF8").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 去程總人數").setBackground("#1E3A8A").setFontColor("#93C5FD").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=SUM(B7:B26) + SUM(G7:G56) + SUM(L7:L46) + SUM(Q7:Q26)").setBackground("#172554").setFontColor("#60A5FA").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 返程總人數").setBackground("#064E3B").setFontColor("#6EE7B7").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=SUM(D7:D26) + SUM(I7:I56) + SUM(N7:N46) + SUM(S7:S26)").setBackground("#022C22").setFontColor("#34D399").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");

  // --- B. 4 大站點戰情明細 ---
  const stations = [
    { name: "成功車站", formKeyword: "成功車站", startCol: 1, busCount: 20, bg: "#3B82F6" },
    { name: "新烏日台鐵站", formKeyword: "新烏日台鐵站", startCol: 6, busCount: 50, bg: "#8B5CF6" },
    { name: "經貿六停車場", formKeyword: "經貿六停車場", startCol: 11, busCount: 40, bg: "#EC4899" },
    { name: "水湳轉運站", formKeyword: "水湳轉運站", startCol: 16, busCount: 20, bg: "#10B981" }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    
    // 站名標題 (第 5 列)
    dashboardSheet.getRange(5, col, 1, 5).merge()
      .setValue(`📍 ${st.name} (共 ${st.busCount} 輛)`)
      .setBackground(st.bg).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    // 子表頭 (第 6 列)
    dashboardSheet.getRange(6, col, 1, 5)
      .setValues([["車號", "👉去程人數", "去程時間", "👈返程人數", "返程時間"]])
      .setBackground("#1E293B").setFontColor("#E2E8F0").setFontWeight("bold").setHorizontalAlignment("center");

    // 每台車公式 (B欄:行程方向, C欄:站點, D欄:車號, E欄:人數)
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
    dashboardSheet.getRange(7, col, st.busCount, 5).setValues(busRows);
    dashboardSheet.getRange(7, col, st.busCount, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);
    
    dashboardSheet.setColumnWidth(col, 75);
    dashboardSheet.setColumnWidth(col + 1, 85);
    dashboardSheet.setColumnWidth(col + 2, 85);
    dashboardSheet.setColumnWidth(col + 3, 85);
    dashboardSheet.setColumnWidth(col + 4, 85);
  });

  dashboardSheet.setRowHeight(5, 30);
  dashboardSheet.setRowHeight(6, 26);
  for (let r = 7; r <= 56; r++) { dashboardSheet.setRowHeight(r, 24); }

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新客製表單與戰情看板已建立完成！】");
  Logger.log("\n📱【最新車長填寫表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
