/**
 * =========================================================================
 * 🎖️「國防知性之旅-成功嶺營區開放」進場人數統計表 - 【正式官方戰情版】
 * =========================================================================
 * 
 * ✨ 規格與特色：
 * 1. 【正式活動標題】：
 *    - 最上方加入大氣莊嚴官方大標題：🎖️「國防知性之旅-成功嶺營區開放」進場人數統計表
 * 2. 【24 格黃金對稱網格 (100% 絕不切字)】：
 *    - 🚌 車總統計 ➔ 4 大接駁車站 (成功綠、新烏日藍、經貿六黃、水湳黃)
 *    - 🚶 人總統計 ➔ 3 大步行通道 (1號門綠、3號門藍、4號門黃)
 * 3. 【三合一整合】：
 *    - 內建系統建立、🎲 產生測試資料 (`runGenerateTestData`)、🗑️ 一鍵清空歸零 (`clearAllData`)
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行」即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🎨 開始建立【「國防知性之旅-成功嶺營區開放」正式戰情系統】...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🎖️「國防知性之旅-成功嶺營區開放」接駁車與步行人數回報表");
  form.setDescription("請各站車長或門口工作人員即時回報人數。\n回報流程：1. 方向 ➔ 2. 站點/門號 ➔ 3. 車號/通道 ➔ 4. 人數");

  // 題目 1: 行程方向
  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向 / 進出方向")
    .setChoiceValues(["👉 去程 (入場)", "👈 返程 (離場)"])
    .setRequired(true);

  // 題目 2: 選擇接駁站點或步行大門
  form.addMultipleChoiceItem()
    .setTitle("2. 選擇接駁站點 / 步行大門")
    .setChoiceValues([
      "🚌 成功車站（綠線）",
      "🚌 新烏日台鐵站（藍線）",
      "🚌 經貿六停車場（黃線）",
      "🚌 水湳轉運站（黃線）",
      "🚶 1號門（綠線）",
      "🚶 3號門（藍線）",
      "🚶 4號門（黃線）"
    ])
    .setRequired(true);

  // 題目 3: 選擇車號 (或步行不分車號)
  const busChoices = ["🚶 步行通道 (無車號)"];
  for (let i = 1; i <= 50; i++) {
    busChoices.push(`${i} 號車`);
  }
  form.addListItem()
    .setTitle("3. 選擇車號 (若為步行進出請選第一項)")
    .setChoiceValues(busChoices)
    .setRequired(true);

  // 題目 4: 搭乘 / 入場人數
  const textValidation = FormApp.createTextValidation()
    .setHelpText("請輸入大於等於 0 的數字")
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem()
    .setTitle("4. 搭乘 / 進出人數")
    .setValidation(textValidation)
    .setRequired(true);

  // 題目 5: 備註
  form.addTextItem()
    .setTitle("5. 備註 (選填)")
    .setRequired(false);

  // 2. 建立 Google 試算表並綁定
  const ss = SpreadsheetApp.create("🎖️「國防知性之旅-成功嶺營區開放」即時戰情中心");
  const ssId = ss.getId();
  const ssUrl = ss.getUrl();

  try {
    ss.setSharing(SpreadsheetApp.Access.ANYONE_WITH_LINK, SpreadsheetApp.Permission.EDIT);
  } catch(e) {}

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
  // 【A. 第二頁：各車即時明細 (4 大車站車輛明細)】
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
      const fGoActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), ISNUMBER(SEARCH("去程", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      const fGoTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), ISNUMBER(SEARCH("去程", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      const fBackActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), ISNUMBER(SEARCH("返程", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      const fBackTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), ISNUMBER(SEARCH("返程", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      busRows.push([busName, fGoActual, fGoTime, fBackActual, fBackTime]);
    }
    detailSheet.getRange(3, col, st.busCount, 5).setValues(busRows);
    detailSheet.getRange(2, col, st.busCount + 1, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);
  });

  // ==========================================
  // 【B. 第一頁：總即時戰情看板 (24 格黃金對稱)】
  // ==========================================
  try {
    dashboardSheet.getRange(1, 1, 45, 26).breakApart();
  } catch (e) {}

  // 設定 24 欄標準欄寬 (每欄 48px，總寬 1152px 完美滿版)
  for (let c = 1; c <= 24; c++) {
    dashboardSheet.setColumnWidth(c, 48);
  }

  const THEME_MAIN_BANNER = "#090D16"; // 頂部大標題深軍黑
  const THEME_HEADER_BG   = "#0F172A"; // 大盤深石墨黑
  const THEME_CARD_BAR    = "#2563EB"; // 統一極簡商務藍跑條
  const THEME_TOP_BAR     = "#38BDF8"; // 頂部高亮青藍

  // -------------------------------------------------------------
  // 🎖️【第 1 列：全場官方總大標題】
  // -------------------------------------------------------------
  dashboardSheet.getRange(1, 1, 1, 24).merge()
    .setValue("🎖️「國防知性之旅-成功嶺營區開放」進場人數統計表")
    .setBackground(THEME_MAIN_BANNER).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center");

  // -------------------------------------------------------------
  // 🚌【段落一：車總統計大盤 (第 2~5 列)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(2, 1, 1, 24).merge()
    .setValue("🚌 接駁車輛全場總統計（共 4 條路線 / 130 輛車）")
    .setBackground(THEME_HEADER_BG).setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  // 👉 去程總人數 (A3:H4, 8欄)
  dashboardSheet.getRange(3, 1, 1, 8).merge().setValue("👉 全場去程總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 1, 1, 8).merge().setFormula("=A10+G10+M10+S10").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 👈 返程總人數 (I3:P4, 8欄)
  dashboardSheet.getRange(3, 9, 1, 8).merge().setValue("👈 全場返程總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 9, 1, 8).merge().setFormula("=D10+J10+P10+V10").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 📈 車輛疏運完成率 (Q3:X4, 8欄)
  dashboardSheet.getRange(3, 17, 1, 8).merge().setValue("📈 車輛疏運完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 17, 1, 8).merge().setFormula(`=IF(A4>0, TEXT(I4/A4, "0.0%"), "0.0%")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 接駁車進度條 (第 5 列)
  dashboardSheet.getRange(5, 1, 1, 6).merge().setValue("⚡ 車輛疏運進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(5, 7, 1, 18).merge().setFormula(`=IF(A4>0, SPARKLINE(I4, {"charttype", "bar"; "max", A4; "color1", "${THEME_TOP_BAR}"}), "")`).setBackground(THEME_HEADER_BG);

  dashboardSheet.getRange(2, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------
  // 🚌【段落一：4 大接駁車站 (第 7~12 列，每站 6 欄)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(7, 1, 1, 24).merge()
    .setValue("🚌 各接駁車站點即時疏運明細")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const busStations = [
    { name: "📍 成功車站（綠線）", startCol: 1, detailGoCol: "B", detailBackCol: "D", detailEndRow: 22, tagColor: "#059669" },
    { name: "📍 新烏日台鐵站（藍線）", startCol: 7, detailGoCol: "G", detailBackCol: "I", detailEndRow: 52, tagColor: "#2563EB" },
    { name: "📍 經貿六停車場（黃線）", startCol: 13, detailGoCol: "L", detailBackCol: "N", detailEndRow: 42, tagColor: "#D97706" },
    { name: "📍 水湳轉運站（黃線）", startCol: 19, detailGoCol: "Q", detailBackCol: "S", detailEndRow: 22, tagColor: "#D97706" }
  ];

  busStations.forEach(cs => {
    const col = cs.startCol;
    
    // 車站大標題 (第 8 列，合併 6 欄)
    dashboardSheet.getRange(8, col, 1, 6).merge()
      .setValue(cs.name)
      .setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 項目標籤 (第 9 列：去程 3 欄 ｜ 返程 3 欄)
    dashboardSheet.getRange(9, col, 1, 3).merge().setValue("👉 去程人數").setBackground("#F8FAFC").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(9, col + 3, 1, 3).merge().setValue("👈 返程人數").setBackground("#F8FAFC").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 10 列：去程 3 欄 ｜ 返程 3 欄)
    const colLetterGo = String.fromCharCode(64 + col);
    const colLetterBack = String.fromCharCode(64 + col + 3);

    dashboardSheet.getRange(10, col, 1, 3).merge()
      .setFormula(`=SUM('各車即時明細'!${cs.detailGoCol}3:${cs.detailGoCol}${cs.detailEndRow})`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(22).setHorizontalAlignment("center");
    dashboardSheet.getRange(10, col + 3, 1, 3).merge()
      .setFormula(`=SUM('各車即時明細'!${cs.detailBackCol}3:${cs.detailBackCol}${cs.detailEndRow})`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(22).setHorizontalAlignment("center");

    // 疏運率與尚餘 (第 11 列，跨滿 6 欄)
    dashboardSheet.getRange(11, col, 1, 6).merge()
      .setFormula(`=IF(${colLetterGo}10>0, "📈 返程疏運率: " & TEXT(${colLetterBack}10/${colLetterGo}10, "0.0%") & "   (尚餘 " & MAX(0, ${colLetterGo}10 - ${colLetterBack}10) & " 人)", "📈 返程疏運率: 0.0% (已完成)")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條 (第 12 列，跨滿 6 欄)
    dashboardSheet.getRange(12, col, 1, 6).merge()
      .setFormula(`=IF(${colLetterGo}10>0, SPARKLINE(${colLetterBack}10, {"charttype", "bar"; "max", ${colLetterGo}10; "color1", "${THEME_CARD_BAR}"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(8, col, 5, 6).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // -------------------------------------------------------------
  // 🚶【段落二：人總統計大盤 (第 14~17 列，跨 A~X 24欄)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(14, 1, 1, 24).merge()
    .setValue("🚶 步行通道全場總統計（1號門綠線 / 3號門藍線 / 4號門黃線）")
    .setBackground(THEME_HEADER_BG).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  // 👉 入場總人數 (A15:H16, 8欄)
  dashboardSheet.getRange(15, 1, 1, 8).merge().setValue("👉 全場入場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 1, 1, 8).merge().setFormula("=A22+I22+Q22").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 👈 離場總人數 (I15:P16, 8欄)
  dashboardSheet.getRange(15, 9, 1, 8).merge().setValue("👈 全場離場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 9, 1, 8).merge().setFormula("=E22+M22+U22").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 📈 步行離場完成率 (Q15:X16, 8欄)
  dashboardSheet.getRange(15, 17, 1, 8).merge().setValue("📈 步行離場完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 17, 1, 8).merge().setFormula(`=IF(A16>0, TEXT(I16/A16, "0.0%"), "0.0%")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center");

  // 步行進度條 (第 17 列)
  dashboardSheet.getRange(17, 1, 1, 6).merge().setValue("⚡ 步行疏運進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(17, 7, 1, 18).merge().setFormula(`=IF(A16>0, SPARKLINE(I16, {"charttype", "bar"; "max", A16; "color1", "${THEME_TOP_BAR}"}), "")`).setBackground(THEME_HEADER_BG);

  dashboardSheet.getRange(14, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------
  // 🚶【段落二：3 大步行門 (第 19~24 列，每門 8 欄)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(19, 1, 1, 24).merge()
    .setValue("🚶 各步行通道進出即時明細")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const walkGates = [
    { name: "🚶 1號門（綠線）", keyword: "1號門", startCol: 1, tagColor: "#059669" },
    { name: "🚶 3號門（藍線）", keyword: "3號門", startCol: 9, tagColor: "#2563EB" },
    { name: "🚶 4號門（黃線）", keyword: "4號門", startCol: 17, tagColor: "#D97706" }
  ];

  walkGates.forEach(wg => {
    const col = wg.startCol;
    
    // 門名大標題 (第 20 列，合併 8 欄)
    dashboardSheet.getRange(20, col, 1, 8).merge()
      .setValue(wg.name)
      .setBackground(wg.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 項目標籤 (第 21 列：入場 4 欄 ｜ 離場 4 欄)
    dashboardSheet.getRange(21, col, 1, 4).merge().setValue("👉 入場人數").setBackground("#F8FAFC").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(21, col + 4, 1, 4).merge().setValue("👈 離場人數").setBackground("#F8FAFC").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 22 列：入場 4 欄 ｜ 離場 4 欄)
    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 4);

    dashboardSheet.getRange(22, col, 1, 4).merge()
      .setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*去程*", '${formSheetName}'!C:C, "*${wg.keyword}*")`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(22).setHorizontalAlignment("center");

    dashboardSheet.getRange(22, col + 4, 1, 4).merge()
      .setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*返程*", '${formSheetName}'!C:C, "*${wg.keyword}*")`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(22).setHorizontalAlignment("center");

    // 疏運率與尚餘 (第 23 列，跨滿 8 欄)
    dashboardSheet.getRange(23, col, 1, 8).merge()
      .setFormula(`=IF(${colLetterIn}22>0, "📈 離場疏運率: " & TEXT(${colLetterOut}22/${colLetterIn}22, "0.0%") & "   (尚餘 " & MAX(0, ${colLetterIn}22 - ${colLetterOut}22) & " 人)", "📈 離場疏運率: 0.0% (已完成)")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條 (第 24 列，跨滿 8 欄)
    dashboardSheet.getRange(24, col, 1, 8).merge()
      .setFormula(`=IF(${colLetterIn}22>0, SPARKLINE(${colLetterOut}22, {"charttype", "bar"; "max", ${colLetterIn}22; "color1", "${THEME_CARD_BAR}"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(20, col, 5, 8).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 列高設定
  dashboardSheet.setRowHeight(1, 38); // 🎖️ 大標題
  dashboardSheet.setRowHeight(2, 26);
  dashboardSheet.setRowHeight(3, 22);
  dashboardSheet.setRowHeight(4, 38);
  dashboardSheet.setRowHeight(5, 24);
  dashboardSheet.setRowHeight(6, 14);
  dashboardSheet.setRowHeight(7, 26);
  dashboardSheet.setRowHeight(8, 34);
  dashboardSheet.setRowHeight(9, 26);
  dashboardSheet.setRowHeight(10, 38);
  dashboardSheet.setRowHeight(11, 28);
  dashboardSheet.setRowHeight(12, 18);
  dashboardSheet.setRowHeight(13, 18);
  dashboardSheet.setRowHeight(14, 26);
  dashboardSheet.setRowHeight(15, 22);
  dashboardSheet.setRowHeight(16, 38);
  dashboardSheet.setRowHeight(17, 24);
  dashboardSheet.setRowHeight(18, 14);
  dashboardSheet.setRowHeight(19, 26);
  dashboardSheet.setRowHeight(20, 34);
  dashboardSheet.setRowHeight(21, 26);
  dashboardSheet.setRowHeight(22, 38);
  dashboardSheet.setRowHeight(23, 28);
  dashboardSheet.setRowHeight(24, 18);

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【🎖️「國防知性之旅-成功嶺營區開放」進場人數統計表 建立完成！】");
  Logger.log("\n📱【最新回報表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

// =========================================================================
// 🔍【輔助函式：全自動智慧獲取試算表 (防 null 崩潰)】
// =========================================================================
function getTargetSpreadsheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const files = DriveApp.getFilesByName("🎖️「國防知性之旅-成功嶺營區開放」即時戰情中心");
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }

  const fallbackFiles = DriveApp.searchFiles('title contains "即時戰情中心" and mimeType = "application/vnd.google-apps.spreadsheet"');
  if (fallbackFiles.hasNext()) {
    return SpreadsheetApp.open(fallbackFiles.next());
  }

  return null;
}

// =========================================================================
// 🎲【指令：一鍵產生 測試資料】
// =========================================================================
function runGenerateTestData() {
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

  // 1. 🚌 接駁車測試數據 (成功 20、新烏日 50、經貿六 40、水湳 20)
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

  // 2. 🚶 步行通道測試數據 (1號門綠、3號門藍、4號門黃)
  const gates = ["🚶 1號門（綠線）", "🚶 3號門（藍線）", "🚶 4號門（黃線）"];
  gates.forEach(gate => {
    for (let p = 1; p <= 4; p++) {
      const walkIn = Math.floor(Math.random() * 80) + 120; // 120~200人
      testRows.push([`${datePrefix} 09:${String(10 + p*15).padStart(2,'0')}:00`, "👉 去程 (入場)", gate, "🚶 步行通道 (無車號)", walkIn, "入場批次"]);
    }
    for (let p = 1; p <= 4; p++) {
      const walkOut = Math.floor(Math.random() * 70) + 100; // 100~170人
      testRows.push([`${datePrefix} 18:${String(5 + p*15).padStart(2,'0')}:00`, "👈 返程 (離場)", gate, "🚶 步行通道 (無車號)", walkOut, "離場批次"]);
    }
  });

  formSheet.getRange(formSheet.getLastRow() + 1, 1, testRows.length, 6).setValues(testRows);
  Logger.log("🎉 已成功產生 " + testRows.length + " 筆測試數據！請回到「總即時戰情看板」查看數據！");
}

// =========================================================================
// 🗑️【指令：一鍵清空全場數據歸零】
// =========================================================================
function clearAllData() {
  const ss = getTargetSpreadsheet();
  if (!ss) {
    Logger.log("❌ 找不到戰情中心試算表！");
    return;
  }

  let formSheet = null;
  for (let s of ss.getSheets()) {
    if (s.getName().includes("表單回應") || s.getName().includes("Form Responses")) {
      formSheet = s;
      break;
    }
  }

  if (formSheet && formSheet.getLastRow() > 1) {
    formSheet.getRange(2, 1, formSheet.getLastRow() - 1, formSheet.getLastColumn()).clearContent();
    Logger.log("🧹 表單流水帳已全部清空！");
  }

  const detailSheet = ss.getSheetByName("各車即時明細");
  if (detailSheet) {
    Logger.log("🧹 各車即時明細已透過公式自動歸零！");
  }

  Logger.log("✨ 全場數據已一秒清空歸零完畢！");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
