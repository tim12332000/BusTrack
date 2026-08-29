/**
 * =========================================================================
 * 🚌 接駁車與步行通道管理系統 - 【純淨專業戰情版 (修復步行計算精準公式)】
 * =========================================================================
 * 
 * ✨ 本版修復：
 * - 步行通道改用精準 SUMIFS 條件加總（鎖定 1/3/4 號門關鍵字），徹底告別負數與誤差！
 * - 4 大接駁車站 (成功、新烏日、經貿六、水湳) + 3 大步行門 (1/3/4號門) 完美運算！
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行 (Run)」即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🎨 開始建立【純淨專業戰情系統】...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🚌 接駁車與步行進出【去程(入場) / 返程(離場)】人數回報");
  form.setDescription("請車長或門口工作人員即時回報人數。\n回報流程：1. 方向 ➔ 2. 站點/門號 ➔ 3. 車號/批次 ➔ 4. 人數");

  // 題目 1: 行程方向
  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向 / 進出方向")
    .setChoiceValues(["👉 去程 (入場)", "👈 返程 (離場)"])
    .setRequired(true);

  // 題目 2: 選擇接駁站點或步行大門
  form.addMultipleChoiceItem()
    .setTitle("2. 選擇接駁站點 / 步行大門")
    .setChoiceValues([
      "成功車站（綠線）",
      "新烏日台鐵站（藍線）",
      "經貿六停車場（黃線）",
      "水湳轉運站（黃線）",
      "🚶 步行：1號門",
      "🚶 步行：3號門",
      "🚶 步行：4號門"
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
  const ss = SpreadsheetApp.create("🚌 各站接駁車與步行通道【去/返程】即時戰情中心");
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
    detailSheet.getRange(2, col, st.busCount + 1, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // ==========================================
  // 【B. 第一頁：總即時戰情看板 (純淨 5 大卡片)】
  // ==========================================
  try {
    dashboardSheet.getRange(1, 1, 40, 20).breakApart();
  } catch (e) {}

  // 1. 頂部全場總大盤 (含 4 站接駁車 + 步行進出)
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總入場人次 (車輛+步行)").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#0F172A").setFontColor("#38BDF8").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 全場入場/去程總人數").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=B8+E8+H8+K8+M8").setBackground("#0F172A").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 全場離場/返程總人數").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=C8+F8+I8+L8+N8").setBackground("#0F172A").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("J1:O1").merge().setValue("📈 全場總疏運/離場完成率").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("J2:O3").merge().setFormula(`=IF(D2>0, TEXT(G2/D2, "0.0%"), "0.0%")`).setBackground("#0F172A").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 全場疏運進度條 (第 4 列)
  dashboardSheet.getRange("A4:C4").merge().setValue("⚡ 總疏運進度").setBackground("#1E293B").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D4:O4").merge().setFormula(`=IF(D2>0, SPARKLINE(G2, {"charttype", "bar"; "max", D2; "color1", "#38BDF8"}), "")`).setBackground("#1E293B");

  // 2. 前 4 大接駁車站卡片 (第 6~10 列)
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
      .setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

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
    
    dashboardSheet.setColumnWidth(col, 105);
    dashboardSheet.setColumnWidth(col + 1, 105);
    dashboardSheet.setColumnWidth(col + 2, 105);
  });

  // 3. 🌟 第 5 大卡片：【🚶 步行進出通道卡片 (1/3/4號門)】 (M6:O10)
  const walkCol = 13; // M 欄
  dashboardSheet.getRange(6, walkCol, 1, 3).merge()
    .setValue("🚶 步行進出 (1/3/4號門)")
    .setBackground("#64748B").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

  dashboardSheet.getRange(7, walkCol).setValue("👉 入場人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(7, walkCol + 1).setValue("👈 離場人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(7, walkCol + 2).setValue("🏆 步行總量").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

  // 🌟【修復】：精準鎖定步行大門關鍵字做 SUMIFS，絕無負數！
  dashboardSheet.getRange(8, walkCol).setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*去程*", '${formSheetName}'!C:C, "*門*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
  dashboardSheet.getRange(8, walkCol + 1).setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*返程*", '${formSheetName}'!C:C, "*門*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
  dashboardSheet.getRange(8, walkCol + 2).setFormula(`=M8+N8`).setBackground("#FFFFFF").setFontColor("#7C3AED").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");

  // 步行疏運率
  dashboardSheet.getRange(9, walkCol, 1, 2).merge()
    .setFormula(`=IF(M8>0, "📈 離場疏運率: " & TEXT(N8/M8, "0.0%"), "📈 離場疏運率: 0.0%")`)
    .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

  dashboardSheet.getRange(9, walkCol + 2).setFormula(`=IF(M8>0, "尚餘 " & MAX(0, M8 - N8) & " 人", "已完成")`)
    .setBackground("#F8FAFC").setFontColor("#EF4444").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

  // 步行疏運進度條
  dashboardSheet.getRange(10, walkCol, 1, 3).merge()
    .setFormula(`=IF(M8>0, SPARKLINE(N8, {"charttype", "bar"; "max", M8; "color1", "#8B5CF6"}), "")`)
    .setBackground("#F1F5F9");

  dashboardSheet.getRange(6, walkCol, 5, 3).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.setColumnWidth(walkCol, 105);
  dashboardSheet.setColumnWidth(walkCol + 1, 105);
  dashboardSheet.setColumnWidth(walkCol + 2, 105);

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

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新純淨專業戰情看板建立完成！】");
  Logger.log("\n📱【最新回報表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
