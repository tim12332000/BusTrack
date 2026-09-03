/**
 * =========================================================================
 * 🎖️「國防知性之旅-成功嶺營區開放」進場人數統計表 - 【表單欄位自動校準修復版】
 * =========================================================================
 * 
 * 🔧 本版重大修復：
 * 1. 【自動校準表單回應欄位】：
 *    - 自動清理表單回應中的多餘舊欄位，將用戶填寫的最新資料精確歸位到 B~F 欄！
 *    - 用戶剛才回報的數據（例如成功車站 10 號車）立刻生效跳動！
 * 2. 【防欄位漂移機制】：
 *    - 若表單題目已存在，直接保留不再重複刪建，徹底防止 Google Form 往右側多開欄位！
 * 3. 【純淨看板與極速同步】：
 *    - 鎖定試算表 ID：1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行」即可！
 * =========================================================================
 */

const TARGET_SPREADSHEET_ID = "1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w";

function createMultiStationBusSystem() {
  Logger.log("🎨 開始對指定試算表 [" + TARGET_SPREADSHEET_ID + "] 進行欄位校準與看板重繪...");

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  
  // 1. 取得並更新既有表單題目 (智慧維護，不重複刪建防漂移)
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
    // 只有在題目數量不對時才重建，避免 Google 試算表往右多開欄位
    if (items.length !== 5) {
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
  }

  // 2. 取得工作表
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

  // 🛠️【核心修復：清理「表單回應 1」中歷史冗餘空欄，將最新數據對齊回 B~F 欄】
  let formSheet = ss.getSheetByName(formSheetName);
  if (formSheet) {
    const lastCol = formSheet.getLastColumn();
    // 如果欄位數超過 6 欄 (代表前面有重複刪題留下的歷史空欄)，將中間的空欄刪除
    if (lastCol > 6) {
      const deleteCount = lastCol - 6;
      formSheet.deleteColumns(2, deleteCount);
      Logger.log("✨ 成功校準表單回應欄位！最新數據已自動移回 B、C、D、E 欄！");
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
  // 【B. 第一頁：總即時戰情看板 (深度清洗與純淨排版)】
  // ==========================================
  try {
    const maxRows = Math.max(dashboardSheet.getMaxRows(), 50);
    const maxCols = Math.max(dashboardSheet.getMaxColumns(), 30);
    const cleanRange = dashboardSheet.getRange(1, 1, maxRows, maxCols);
    cleanRange.breakApart();
    cleanRange.clearDataValidations();
    dashboardSheet.clearConditionalFormatRules();
    cleanRange.clear();
  } catch (e) {}

  dashboardSheet.setHiddenGridlines(true);

  for (let c = 1; c <= 24; c++) {
    dashboardSheet.setColumnWidth(c, 48);
  }

  const THEME_MAIN_BANNER = "#090D16";
  const THEME_HEADER_BG   = "#0F172A";
  const THEME_CARD_BAR    = "#2563EB";
  const THEME_TOP_BAR     = "#38BDF8";

  // 第 1 列：活動總標題
  dashboardSheet.getRange(1, 1, 1, 24).merge()
    .setValue("「國防知性之旅-成功嶺營區開放」進場人數統計表")
    .setBackground(THEME_MAIN_BANNER).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center");

  // 段落一：接駁車大盤 (第 2~5 列)
  dashboardSheet.getRange(2, 1, 1, 24).merge()
    .setValue("🚌 接駁車疏運概況")
    .setBackground(THEME_HEADER_BG).setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  dashboardSheet.getRange(3, 1, 1, 8).merge().setValue("👉 進場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 1, 1, 8).merge().setFormula("=A10+G10+M10+S10").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(3, 9, 1, 8).merge().setValue("👈 離場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 9, 1, 8).merge().setFormula("=D10+J10+P10+V10").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(3, 17, 1, 8).merge().setValue("📈 離場完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 17, 1, 8).merge().setFormula(`=IF(A4>0, TEXT(I4/A4, "0.0%"), "0.0%")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(5, 1, 1, 6).merge().setValue("離場進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(5, 7, 1, 18).merge().setFormula(`=IF(A4>0, SPARKLINE(I4, {"charttype", "bar"; "max", A4; "color1", "${THEME_TOP_BAR}"}), "")`).setBackground(THEME_HEADER_BG);

  dashboardSheet.getRange(2, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // 4 大接駁車站 (第 7~12 列)
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
    dashboardSheet.getRange(8, col, 1, 6).merge().setValue(cs.name).setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(9, col, 1, 3).merge().setValue("進場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(9, col + 3, 1, 3).merge().setValue("離場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 3);

    dashboardSheet.getRange(10, col, 1, 3).merge().setFormula(`=SUM('各車即時明細'!${cs.detailInCol}3:${cs.detailInCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(10, col + 3, 1, 3).merge().setFormula(`=SUM('各車即時明細'!${cs.detailOutCol}3:${cs.detailOutCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(11, col, 1, 6).merge().setFormula(`=IF(${colLetterIn}10>0, TEXT(${colLetterOut}10/${colLetterIn}10, "0.0%") & "  ·  尚餘 " & MAX(0, ${colLetterIn}10 - ${colLetterOut}10) & " 人", "0.0%  ·  已完成")`).setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(12, col, 1, 6).merge().setFormula(`=IF(${colLetterIn}10>0, SPARKLINE(${colLetterOut}10, {"charttype", "bar"; "max", ${colLetterIn}10; "color1", "${THEME_CARD_BAR}"}), "")`).setBackground("#F1F5F9");
    dashboardSheet.getRange(8, col, 5, 6).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 段落二：步行大盤 (第 14~17 列)
  dashboardSheet.getRange(14, 1, 1, 24).merge().setValue("🚶 步行通道疏運概況").setBackground(THEME_HEADER_BG).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange(15, 1, 1, 8).merge().setValue("👉 進場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 1, 1, 8).merge().setFormula("=A22+I22+Q22").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(15, 9, 1, 8).merge().setValue("👈 離場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 9, 1, 8).merge().setFormula("=E22+M22+U22").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(15, 17, 1, 8).merge().setValue("📈 離場完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(16, 17, 1, 8).merge().setFormula(`=IF(A16>0, TEXT(I16/A16, "0.0%"), "0.0%")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(17, 1, 1, 6).merge().setValue("離場進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(17, 7, 1, 18).merge().setFormula(`=IF(A16>0, SPARKLINE(I16, {"charttype", "bar"; "max", A16; "color1", "${THEME_TOP_BAR}"}), "")`).setBackground(THEME_HEADER_BG);
  dashboardSheet.getRange(14, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // 3 大步行門 (第 19~24 列)
  dashboardSheet.getRange(19, 1, 1, 24).merge().setValue("步行通道即時數據").setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const walkGates = [
    { name: "🚶 1號門", keyword: "1號門", startCol: 1, tagColor: "#059669" },
    { name: "🚶 3號門", keyword: "3號門", startCol: 9, tagColor: "#2563EB" },
    { name: "🚶 4號門", keyword: "4號門", startCol: 17, tagColor: "#D97706" }
  ];

  walkGates.forEach(wg => {
    const col = wg.startCol;
    dashboardSheet.getRange(20, col, 1, 8).merge().setValue(wg.name).setBackground(wg.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(21, col, 1, 4).merge().setValue("進場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(21, col + 4, 1, 4).merge().setValue("離場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 4);

    dashboardSheet.getRange(22, col, 1, 4).merge().setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*進場*", '${formSheetName}'!C:C, "*${wg.keyword}*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(22, col + 4, 1, 4).merge().setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*離場*", '${formSheetName}'!C:C, "*${wg.keyword}*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(23, col, 1, 8).merge().setFormula(`=IF(${colLetterIn}22>0, TEXT(${colLetterOut}22/${colLetterIn}22, "0.0%") & "  ·  尚餘 " & MAX(0, ${colLetterIn}22 - ${colLetterOut}22) & " 人", "0.0%  ·  已完成")`).setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(24, col, 1, 8).merge().setFormula(`=IF(${colLetterIn}22>0, SPARKLINE(${colLetterOut}22, {"charttype", "bar"; "max", ${colLetterIn}22; "color1", "${THEME_CARD_BAR}"}), "")`).setBackground("#F1F5F9");
    dashboardSheet.getRange(20, col, 5, 8).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 列高設定
  dashboardSheet.setRowHeight(1, 38);
  dashboardSheet.setRowHeight(2, 26);
  dashboardSheet.setRowHeight(3, 22);
  dashboardSheet.setRowHeight(4, 38);
  dashboardSheet.setRowHeight(5, 24);
  dashboardSheet.setRowHeight(6, 12);
  dashboardSheet.setRowHeight(7, 26);
  dashboardSheet.setRowHeight(8, 34);
  dashboardSheet.setRowHeight(9, 26);
  dashboardSheet.setRowHeight(10, 38);
  dashboardSheet.setRowHeight(11, 28);
  dashboardSheet.setRowHeight(12, 18);
  dashboardSheet.setRowHeight(13, 14);
  dashboardSheet.setRowHeight(14, 26);
  dashboardSheet.setRowHeight(15, 22);
  dashboardSheet.setRowHeight(16, 38);
  dashboardSheet.setRowHeight(17, 24);
  dashboardSheet.setRowHeight(18, 12);
  dashboardSheet.setRowHeight(19, 26);
  dashboardSheet.setRowHeight(20, 34);
  dashboardSheet.setRowHeight(21, 26);
  dashboardSheet.setRowHeight(22, 38);
  dashboardSheet.setRowHeight(23, 28);
  dashboardSheet.setRowHeight(24, 18);

  Logger.log("\n=======================================================");
  Logger.log("🎉【表單欄位自動校準完成！數據已即時生效！】");
  Logger.log("=======================================================\n");
}

// 快速修復欄位指令 (可單獨執行)
function fixFormColumnsAlignment() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  let formSheetName = "表單回應 1";
  for (let s of ss.getSheets()) {
    if (s.getName() !== "總即時戰情看板" && s.getName() !== "各車即時明細") {
      formSheetName = s.getName();
      break;
    }
  }
  let formSheet = ss.getSheetByName(formSheetName);
  if (formSheet) {
    const lastCol = formSheet.getLastColumn();
    if (lastCol > 6) {
      formSheet.deleteColumns(2, lastCol - 6);
      Logger.log("✨ 已成功刪除多餘舊欄位，數據已精確移回 B~F 欄！");
    } else {
      Logger.log("✨ 欄位結構正常，無須清理！");
    }
  }
}
