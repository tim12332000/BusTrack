/**
 * =========================================================================
 * 🎖️「國防知性之旅-成功嶺營區開放」進場人數統計表 - 【鎖定既有網址專用版】
 * =========================================================================
 * 
 * 📌 鎖定官方指定永久網址：
 * - 戰情看板：https://docs.google.com/spreadsheets/d/1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w/edit
 * - 回報表單：https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform
 * 
 * 👉 每次點「執行」，100% 在您原本這份試算表與表單上原地更新，網址永遠不變！
 * =========================================================================
 */

const TARGET_SPREADSHEET_ID = "1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w";

function createMultiStationBusSystem() {
  Logger.log("🎨 開始對指定試算表 [" + TARGET_SPREADSHEET_ID + "] 進行原地極簡重繪...");

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  
  // 1. 取得並更新既有表單題目 (網址保持不變)
  let form = null;
  const formUrl = ss.getFormUrl();
  if (formUrl) {
    try {
      form = FormApp.openByUrl(formUrl);
    } catch (e) {}
  }

  if (form) {
    form.setTitle("「國防知性之旅-成功嶺營區開放」人數回報");
    const items = form.getItems();
    for (let i = items.length - 1; i >= 0; i--) {
      form.deleteItem(items[i]);
    }

    // 題目 1: 方向 (進場 / 離場)
    form.addMultipleChoiceItem()
      .setTitle("1. 方向")
      .setChoiceValues(["進場", "離場"])
      .setRequired(true);

    // 題目 2: 站點 / 門號
    form.addMultipleChoiceItem()
      .setTitle("2. 站點 / 門號")
      .setChoiceValues([
        "🚌 成功車站",
        "🚌 新烏日台鐵站",
        "🚌 經貿六停車場",
        "🚌 水湳轉運站",
        "🚶 1號門",
        "🚶 3號門",
        "🚶 4號門"
      ])
      .setRequired(true);

    // 題目 3: 車號
    const busChoices = ["🚶 步行通道"];
    for (let i = 1; i <= 50; i++) {
      busChoices.push(`${i} 號車`);
    }
    form.addListItem()
      .setTitle("3. 車號")
      .setChoiceValues(busChoices)
      .setRequired(true);

    // 題目 4: 人數
    const textValidation = FormApp.createTextValidation()
      .setHelpText("請輸入數字")
      .requireNumberGreaterThanOrEqualTo(0)
      .build();
    form.addTextItem()
      .setTitle("4. 人數")
      .setValidation(textValidation)
      .setRequired(true);

    // 題目 5: 備註
    form.addTextItem()
      .setTitle("5. 備註")
      .setRequired(false);
  }

  // 2. 取得/建立工作表
  let dashboardSheet = ss.getSheetByName("總即時戰情看板");
  if (!dashboardSheet) {
    dashboardSheet = ss.getSheets()[0];
    dashboardSheet.setName("總即時戰情看板");
  }

  let detailSheet = ss.getSheetByName("各車即時明細");
  if (!detailSheet) {
    detailSheet = ss.insertSheet("各車即時明細");
  }

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
    { name: "🚌 成功車站", formKeyword: "成功車站", startCol: 1, busCount: 20 },
    { name: "🚌 新烏日台鐵站", formKeyword: "新烏日", startCol: 6, busCount: 50 },
    { name: "🚌 經貿六停車場", formKeyword: "經貿六", startCol: 11, busCount: 40 },
    { name: "🚌 水湳轉運站", formKeyword: "水湳", startCol: 16, busCount: 20 }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    detailSheet.getRange(1, col, 1, 5).merge()
      .setValue(st.name)
      .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

    detailSheet.getRange(2, col, 1, 5)
      .setValues([["車號", "進場人數", "進場時間", "離場人數", "離場時間"]])
      .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setHorizontalAlignment("center");

    const busRows = [];
    for (let b = 1; b <= st.busCount; b++) {
      const busName = `${b} 號車`;
      const fInActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), ISNUMBER(SEARCH("進場", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      const fInTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), ISNUMBER(SEARCH("進場", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      const fOutActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), ISNUMBER(SEARCH("離場", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), 0)`;
      const fOutTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), ISNUMBER(SEARCH("離場", '${formSheetName}'!B:B)), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!C:C)), '${formSheetName}'!D:D="${busName}"))), "hh:mm:ss"), "-")`;
      busRows.push([busName, fInActual, fInTime, fOutActual, fOutTime]);
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

  // 24 欄標準欄寬 (每欄 48px，總寬 1152px 完美滿版)
  for (let c = 1; c <= 24; c++) {
    dashboardSheet.setColumnWidth(c, 48);
  }

  const THEME_MAIN_BANNER = "#090D16"; // 官方標題深軍黑
  const THEME_HEADER_BG   = "#0F172A"; // 大盤深石墨黑
  const THEME_CARD_BAR    = "#2563EB"; // 統一極簡商務藍跑條
  const THEME_TOP_BAR     = "#38BDF8"; // 頂部高亮青藍

  // -------------------------------------------------------------
  // 🎖️【第 1 列：活動總標題】
  // -------------------------------------------------------------
  dashboardSheet.getRange(1, 1, 1, 24).merge()
    .setValue("「國防知性之旅-成功嶺營區開放」進場人數統計表")
    .setBackground(THEME_MAIN_BANNER).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center");

  // -------------------------------------------------------------
  // 🚌【段落一：接駁車大盤 (第 2~5 列)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(2, 1, 1, 24).merge()
    .setValue("🚌 接駁車疏運概況")
    .setBackground(THEME_HEADER_BG).setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  // 👉 進場總人數 (A3:H4, 8欄)
  dashboardSheet.getRange(3, 1, 1, 8).merge().setValue("👉 進場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 1, 1, 8).merge().setFormula("=A10+G10+M10+S10").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 👈 離場總人數 (I3:P4, 8欄)
  dashboardSheet.getRange(3, 9, 1, 8).merge().setValue("👈 離場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 9, 1, 8).merge().setFormula("=D10+J10+P10+V10").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 📈 離場完成率 (Q3:X4, 8欄)
  dashboardSheet.getRange(3, 17, 1, 8).merge().setValue("📈 離場完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 17, 1, 8).merge().setFormula(`=IF(A4>0, TEXT(I4/A4, "0.0%"), "0.0%")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 接駁車進度條 (第 5 列)
  dashboardSheet.getRange(5, 1, 1, 6).merge().setValue("離場進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(5, 7, 1, 18).merge().setFormula(`=IF(A4>0, SPARKLINE(I4, {"charttype", "bar"; "max", A4; "color1", "${THEME_TOP_BAR}"}), "")`).setBackground(THEME_HEADER_BG);

  dashboardSheet.getRange(2, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------
  // 🚌【段落一：4 大接駁車站 (第 7~12 列，每站 6 欄)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(7, 1, 1, 24).merge()
    .setValue("接駁車站即時數據")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const busStations = [
    { name: "🚌 成功車站", startCol: 1, detailInCol: "B", detailOutCol: "D", detailEndRow: 22, tagColor: "#059669" },
    { name: "🚌 新烏日台鐵站", startCol: 7, detailInCol: "G", detailOutCol: "I", detailEndRow: 52, tagColor: "#2563EB" },
    { name: "🚌 經貿六停車場", startCol: 13, detailInCol: "L", detailOutCol: "N", detailEndRow: 42, tagColor: "#D97706" },
    { name: "🚌 水湳轉運站", startCol: 19, detailInCol: "Q", detailOutCol: "S", detailEndRow: 22, tagColor: "#D97706" }
  ];

  busStations.forEach(cs => {
    const col = cs.startCol;
    
    // 站名大標題 (第 8 列，合併 6 欄)
    dashboardSheet.getRange(8, col, 1, 6).merge()
      .setValue(cs.name)
      .setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 項目標籤 (第 9 列：進場 3 欄 ｜ 離場 3 欄)
    dashboardSheet.getRange(9, col, 1, 3).merge().setValue("進場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(9, col + 3, 1, 3).merge().setValue("離場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 10 列：24pt 特大粗體黑字)
    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 3);

    dashboardSheet.getRange(10, col, 1, 3).merge()
      .setFormula(`=SUM('各車即時明細'!${cs.detailInCol}3:${cs.detailInCol}${cs.detailEndRow})`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(10, col + 3, 1, 3).merge()
      .setFormula(`=SUM('各車即時明細'!${cs.detailOutCol}3:${cs.detailOutCol}${cs.detailEndRow})`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");

    // 狀態列 (第 11 列，極致洗鍊：70.9% · 尚餘 183 人)
    dashboardSheet.getRange(11, col, 1, 6).merge()
      .setFormula(`=IF(${colLetterIn}10>0, TEXT(${colLetterOut}10/${colLetterIn}10, "0.0%") & "  ·  尚餘 " & MAX(0, ${colLetterIn}10 - ${colLetterOut}10) & " 人", "0.0%  ·  已完成")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 進度條 (第 12 列，跨滿 6 欄)
    dashboardSheet.getRange(12, col, 1, 6).merge()
      .setFormula(`=IF(${colLetterIn}10>0, SPARKLINE(${colLetterOut}10, {"charttype", "bar"; "max", ${colLetterIn}10; "color1", "${THEME_CARD_BAR}"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(8, col, 5, 6).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // -------------------------------------------------------------
  // 🚶【段落二：步行大盤 (第 14~17 列)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(14, 1, 1, 24).merge()
    .setValue("🚶 步行通道疏運概況")
    .setBackground(THEME_HEADER_BG).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  // 👉 進場總人數 (A15:H16, 8欄)
  dashboardSheet.getRange(15, 1, 1, 8).merge().setValue("👉 進場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 1, 1, 8).merge().setFormula("=A22+I22+Q22").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 👈 離場總人數 (I15:P16, 8欄)
  dashboardSheet.getRange(15, 9, 1, 8).merge().setValue("👈 離場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 9, 1, 8).merge().setFormula("=E22+M22+U22").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 📈 離場完成率 (Q15:X16, 8欄)
  dashboardSheet.getRange(15, 17, 1, 8).merge().setValue("📈 離場完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 17, 1, 8).merge().setFormula(`=IF(A16>0, TEXT(I16/A16, "0.0%"), "0.0%")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 離場進度條 (第 17 列)
  dashboardSheet.getRange(17, 1, 1, 6).merge().setValue("離場進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(17, 7, 1, 18).merge().setFormula(`=IF(A16>0, SPARKLINE(I16, {"charttype", "bar"; "max", A16; "color1", "${THEME_TOP_BAR}"}), "")`).setBackground(THEME_HEADER_BG);

  dashboardSheet.getRange(14, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------
  // 🚶【段落二：3 大步行門 (第 19~24 列，每門 8 欄)】
  // -------------------------------------------------------------
  dashboardSheet.getRange(19, 1, 1, 24).merge()
    .setValue("步行通道即時數據")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const walkGates = [
    { name: "🚶 1號門", keyword: "1號門", startCol: 1, tagColor: "#059669" },
    { name: "🚶 3號門", keyword: "3號門", startCol: 9, tagColor: "#2563EB" },
    { name: "🚶 4號門", keyword: "4號門", startCol: 17, tagColor: "#D97706" }
  ];

  walkGates.forEach(wg => {
    const col = wg.startCol;
    
    // 門名大標題 (第 20 列，合併 8 欄)
    dashboardSheet.getRange(20, col, 1, 8).merge()
      .setValue(wg.name)
      .setBackground(wg.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 項目標籤 (第 21 列：進場 4 欄 ｜ 離場 4 欄)
    dashboardSheet.getRange(21, col, 1, 4).merge().setValue("進場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(21, col + 4, 1, 4).merge().setValue("離場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 22 列：24pt 特大粗體黑字)
    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 4);

    dashboardSheet.getRange(22, col, 1, 4).merge()
      .setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*進場*", '${formSheetName}'!C:C, "*${wg.keyword}*")`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");

    dashboardSheet.getRange(22, col + 4, 1, 4).merge()
      .setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*離場*", '${formSheetName}'!C:C, "*${wg.keyword}*")`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");

    // 狀態列 (第 23 列，極致洗鍊：70.4% · 尚餘 202 人)
    dashboardSheet.getRange(23, col, 1, 8).merge()
      .setFormula(`=IF(${colLetterIn}22>0, TEXT(${colLetterOut}22/${colLetterIn}22, "0.0%") & "  ·  尚餘 " & MAX(0, ${colLetterIn}22 - ${colLetterOut}22) & " 人", "0.0%  ·  已完成")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 進度條 (第 24 列，跨滿 8 欄)
    dashboardSheet.getRange(24, col, 1, 8).merge()
      .setFormula(`=IF(${colLetterIn}22>0, SPARKLINE(${colLetterOut}22, {"charttype", "bar"; "max", ${colLetterIn}22; "color1", "${THEME_CARD_BAR}"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(20, col, 5, 8).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 列高設定
  dashboardSheet.setRowHeight(1, 38);
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

  Logger.log("\n=======================================================");
  Logger.log("🎉【鎖定既有網址・原地更新完成！】");
  Logger.log("📊【戰情看板網址 (完全不變)】: " + ss.getUrl());
  if (form) {
    Logger.log("📱【回報表單網址 (完全不變)】: " + form.getPublishedUrl());
  }
  Logger.log("=======================================================\n");
}

// =========================================================================
// 🔍【獲取鎖定試算表】
// =========================================================================
function getTargetSpreadsheet() {
  try {
    return SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  } catch (e) {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
    return null;
  }
}

// =========================================================================
// 🎲【指令：一鍵產生 測試資料】
// =========================================================================
function runGenerateTestData() {
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

  if (!formSheet) {
    Logger.log("❌ 找不到表單回應工作表！請確認試算表已綁定表單。");
    return;
  }

  const testRows = [];
  const now = new Date();
  const datePrefix = Utilities.formatDate(now, "Asia/Taipei", "yyyy/MM/dd");

  // 1. 🚌 接駁車測試數據 (進場 / 離場)
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

  // 2. 🚶 步行通道測試數據 (進場 / 離場)
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
