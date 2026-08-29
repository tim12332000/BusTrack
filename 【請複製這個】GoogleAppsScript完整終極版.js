/**
 * =========================================================================
 * 🚌 接駁車與步行通道管理系統 - 【純淨 3 門等寬・人車徹底分開 戰情版】
 * =========================================================================
 * 
 * ✨ 本版優化：
 * 1. 【移除多餘重複的步行小計】：
 *    - 頂部右側已經是「步行全場總大盤」，下方不再重複放小計卡片！
 * 2. 【3 大步行門均勻等寬展開】：
 *    - 1號門(綠線)、3號門(藍線)、4號門(黃線) 3 張大卡片均分螢幕寬度，排版更寬敞大氣！
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行 (Run)」即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🎨 開始建立【純淨 3 門・人車分區 現代戰情系統】...");

  // 1. 建立全新 Google 表單
  const form = FormApp.create("🚌 接駁車與步行進出【去程(入場) / 返程(離場)】人數回報");
  form.setDescription("請車長或門口工作人員即時回報人數。\n回報流程：1. 方向 ➔ 2. 站點/門號 ➔ 3. 車號/通道 ➔ 4. 人數");

  // 題目 1: 行程方向
  form.addMultipleChoiceItem()
    .setTitle("1. 行程方向 / 進出方向")
    .setChoiceValues(["👉 去程 (入場)", "👈 返程 (離場)"])
    .setRequired(true);

  // 題目 2: 選擇接駁站點或步行大門 (人車分明、色系對應)
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
  const ss = SpreadsheetApp.create("🚌 各站接駁車與步行通道【去/返程】即時戰情中心");
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
    detailSheet.getRange(2, col, st.busCount + 1, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // ==========================================
  // 【B. 第一頁：總即時戰情看板】
  // ==========================================
  try {
    dashboardSheet.getRange(1, 1, 40, 20).breakApart();
  } catch (e) {}

  // 1. 🌟【頂部左半邊：🚌 接駁車輛全場總大盤 (A1:F4)】
  dashboardSheet.getRange("A1:F1").merge().setValue("🚌 接駁車輛全場大盤（共 4 條路線 / 130 輛車）").setBackground("#0F172A").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  
  dashboardSheet.getRange("A2:B2").merge().setValue("👉 去程總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("A3:B3").merge().setFormula("=A8+D8+G8+J8").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("C2:D2").merge().setValue("👈 返程總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("C3:D3").merge().setFormula("=B8+E8+H8+K8").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("E2:F2").merge().setValue("🏆 車運總量").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("E3:F3").merge().setFormula("=A3+C3").setBackground("#1E293B").setFontColor("#38BDF8").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("A4:B4").merge().setFormula(`=IF(A3>0, "📈 疏運率: " & TEXT(C3/A3, "0.0%"), "📈 疏運率: 0.0%")`).setBackground("#0F172A").setFontColor("#34D399").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("C4:F4").merge().setFormula(`=IF(A3>0, SPARKLINE(C3, {"charttype", "bar"; "max", A3; "color1", "#38BDF8"}), "")`).setBackground("#0F172A");

  // 2. 🌟【頂部右半邊：🚶 步行通道全場總大盤 (G1:L4)】
  dashboardSheet.getRange("G1:L1").merge().setValue("🚶 步行通道全場大盤（1號門綠 / 3號門藍 / 4號門黃）").setBackground("#312E81").setFontColor("#C7D2FE").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  
  // 步行入場 (G3) = A14 + E14 + I14 (3 門入場加總)
  // 步行離場 (I3) = C14 + G14 + K14 (3 門離場加總)
  dashboardSheet.getRange("G2:H2").merge().setValue("👉 入場總人數").setBackground("#1E1B4B").setFontColor("#94A3B8").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("G3:H3").merge().setFormula("=A14+E14+I14").setBackground("#1E1B4B").setFontColor("#FFFFFF").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("I2:J2").merge().setValue("👈 離場總人數").setBackground("#1E1B4B").setFontColor("#94A3B8").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("I3:J3").merge().setFormula("=C14+G14+K14").setBackground("#1E1B4B").setFontColor("#FFFFFF").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("K2:L2").merge().setValue("🏆 步行總量").setBackground("#1E1B4B").setFontColor("#94A3B8").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("K3:L3").merge().setFormula("=G3+I3").setBackground("#1E1B4B").setFontColor("#A78BFA").setFontSize(20).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G4:H4").merge().setFormula(`=IF(G3>0, "📈 離場率: " & TEXT(I3/G3, "0.0%"), "📈 離場率: 0.0%")`).setBackground("#312E81").setFontColor("#FDE047").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("I4:L4").merge().setFormula(`=IF(G3>0, SPARKLINE(I3, {"charttype", "bar"; "max", G3; "color1", "#8B5CF6"}), "")`).setBackground("#312E81");

  dashboardSheet.getRange("A1:F4").setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);
  dashboardSheet.getRange("G1:L4").setBorder(true, true, true, true, true, true, "#4338CA", SpreadsheetApp.BorderStyle.SOLID);

  // 3. 🚌【中間：接駁車 4 大站點專區 (第 5~10 列)】
  dashboardSheet.getRange("A5:L5").merge()
    .setValue("🚌 各接駁車站點即時疏運明細（每站佔 3 欄）")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("left");

  const busStations = [
    { name: "📍 成功車站（綠線）", startCol: 1, detailGoCol: "B", detailBackCol: "D", detailEndRow: 22, tagColor: "#059669" },
    { name: "📍 新烏日台鐵站（藍線）", startCol: 4, detailGoCol: "G", detailBackCol: "I", detailEndRow: 52, tagColor: "#2563EB" },
    { name: "📍 經貿六停車場（黃線）", startCol: 7, detailGoCol: "L", detailBackCol: "N", detailEndRow: 42, tagColor: "#D97706" },
    { name: "📍 水湳轉運站（黃線）", startCol: 10, detailGoCol: "Q", detailBackCol: "S", detailEndRow: 22, tagColor: "#D97706" }
  ];

  busStations.forEach(cs => {
    const col = cs.startCol;
    
    dashboardSheet.getRange(6, col, 1, 3).merge()
      .setValue(cs.name)
      .setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    dashboardSheet.getRange(7, col).setValue("👉 去程人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 1).setValue("👈 返程人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 2).setValue("🏆 累計總量").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    const colLetterGo = String.fromCharCode(64 + col);
    const colLetterBack = String.fromCharCode(64 + col + 1);

    dashboardSheet.getRange(8, col).setFormula(`=SUM('各車即時明細'!${cs.detailGoCol}3:${cs.detailGoCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 1).setFormula(`=SUM('各車即時明細'!${cs.detailBackCol}3:${cs.detailBackCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 2).setFormula(`=${colLetterGo}8+${colLetterBack}8`).setBackground("#FFFFFF").setFontColor("#2563EB").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");

    dashboardSheet.getRange(9, col, 1, 2).merge()
      .setFormula(`=IF(${colLetterGo}8>0, "📈 返程疏運率: " & TEXT(${colLetterBack}8/${colLetterGo}8, "0.0%"), "📈 返程疏運率: 0.0%")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    dashboardSheet.getRange(9, col + 2).setFormula(`=IF(${colLetterGo}8>0, "尚餘 " & MAX(0, ${colLetterGo}8 - ${colLetterBack}8) & " 人", "已完成")`)
      .setBackground("#F8FAFC").setFontColor("#EF4444").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    dashboardSheet.getRange(10, col, 1, 3).merge()
      .setFormula(`=IF(${colLetterGo}8>0, SPARKLINE(${colLetterBack}8, {"charttype", "bar"; "max", ${colLetterGo}8; "color1", "#2563EB"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(6, col, 5, 3).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    
    dashboardSheet.setColumnWidth(col, 110);
    dashboardSheet.setColumnWidth(col + 1, 110);
    dashboardSheet.setColumnWidth(col + 2, 110);
  });

  // 4. 🚶【下方：純淨 3 大步行門區 (每門佔 4 欄，均勻平分 12 欄)】(第 11~16 列)
  dashboardSheet.getRange("A11:L11").merge()
    .setValue("🚶 各步行通道進出即時明細（1號門綠線 / 3號門藍線 / 4號門黃線）")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("left");

  const walkGates = [
    { name: "🚶 1號門（綠線）", keyword: "1號門", startCol: 1, tagColor: "#059669", barColor: "#059669" },
    { name: "🚶 3號門（藍線）", keyword: "3號門", startCol: 5, tagColor: "#2563EB", barColor: "#2563EB" },
    { name: "🚶 4號門（黃線）", keyword: "4號門", startCol: 9, tagColor: "#D97706", barColor: "#D97706" }
  ];

  walkGates.forEach(wg => {
    const col = wg.startCol;
    
    // 門名大標題 (第 12 列，合併 4 欄)
    dashboardSheet.getRange(12, col, 1, 4).merge()
      .setValue(wg.name)
      .setBackground(wg.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    // 項目標籤 (第 13 列)
    dashboardSheet.getRange(13, col, 1, 2).merge().setValue("👉 入場人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(13, col + 2, 1, 2).merge().setValue("👈 離場人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 14 列)
    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 2);

    dashboardSheet.getRange(14, col, 1, 2).merge()
      .setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*去程*", '${formSheetName}'!C:C, "*${wg.keyword}*")`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(20).setHorizontalAlignment("center");

    dashboardSheet.getRange(14, col + 2, 1, 2).merge()
      .setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*返程*", '${formSheetName}'!C:C, "*${wg.keyword}*")`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(20).setHorizontalAlignment("center");

    // 疏運率標題 (第 15 列)
    dashboardSheet.getRange(15, col, 1, 2).merge()
      .setFormula(`=IF(${colLetterIn}14>0, "📈 離場疏運率: " & TEXT(${colLetterOut}14/${colLetterIn}14, "0.0%"), "📈 離場疏運率: 0.0%")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    dashboardSheet.getRange(15, col + 2, 1, 2).merge()
      .setFormula(`=IF(${colLetterIn}14>0, "尚餘 " & MAX(0, ${colLetterIn}14 - ${colLetterOut}14) & " 人", "已完成")`)
      .setBackground("#F8FAFC").setFontColor("#EF4444").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條 (第 16 列)
    dashboardSheet.getRange(16, col, 1, 4).merge()
      .setFormula(`=IF(${colLetterIn}14>0, SPARKLINE(${colLetterOut}14, {"charttype", "bar"; "max", ${colLetterIn}14; "color1", "${wg.barColor}"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(12, col, 5, 4).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 列高設定
  dashboardSheet.setRowHeight(1, 26);
  dashboardSheet.setRowHeight(2, 22);
  dashboardSheet.setRowHeight(3, 36);
  dashboardSheet.setRowHeight(4, 24);
  dashboardSheet.setRowHeight(5, 26);
  dashboardSheet.setRowHeight(6, 34);
  dashboardSheet.setRowHeight(7, 26);
  dashboardSheet.setRowHeight(8, 38);
  dashboardSheet.setRowHeight(9, 26);
  dashboardSheet.setRowHeight(10, 18);
  dashboardSheet.setRowHeight(11, 26);
  dashboardSheet.setRowHeight(12, 34);
  dashboardSheet.setRowHeight(13, 26);
  dashboardSheet.setRowHeight(14, 38);
  dashboardSheet.setRowHeight(15, 26);
  dashboardSheet.setRowHeight(16, 18);

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新純淨 3 門・人車徹底分開 戰情看板建立完成！】");
  Logger.log("\n📱【最新回報表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
