/**
 * =========================================================================
 * 🚌 接駁車與步行通道管理系統 - 【人車分區・路線色對應 戰情版】
 * =========================================================================
 * 
 * ✨ 業務與視覺架構：
 * 1. 【人車完全分開】：
 *    - 上半部：【🚌 接駁車專區 (4大站點，共 130 輛)】
 *    - 下半部：【🚶 步行進出專區 (1/3/4號門)】
 * 2. 【門號與路線色精準對應】：
 *    - 1 號門（綠線）➔ 翡翠綠 (#059669)
 *    - 3 號門（藍線）➔ 皇家藍 (#2563EB)
 *    - 4 號門（黃線）➔ 琥珀黃 (#D97706)
 * 3. 【頂部全場總大盤】：
 *    - 全場總進場人次、全場入場(車+人)、全場離場(車+人)、全場總疏運完成率
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行 (Run)」即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🎨 開始建立【人車分區・路線色對應 現代戰情系統】...");

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

  // 1. 🏆 頂部全場總大盤 (第 1~4 列)
  // 接駁車去程: A8+D8+G8+J8, 返程: B8+E8+H8+K8
  // 步行門去程: A14+D14+G14, 返程: B14+E14+H14
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總進場人次 (車輛+步行)").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#0F172A").setFontColor("#38BDF8").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 👉 全場去程總人數 (接駁車 4 站 + 步行 3 門)
  dashboardSheet.getRange("D1:F1").merge().setValue("👉 全場入場/去程總人數").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=(A8+D8+G8+J8) + (A14+D14+G14)").setBackground("#0F172A").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 👈 全場返程總人數 (接駁車 4 站 + 步行 3 門)
  dashboardSheet.getRange("G1:I1").merge().setValue("👈 全場離場/返程總人數").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=(B8+E8+H8+K8) + (B14+E14+H14)").setBackground("#0F172A").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 📈 全場總疏運率
  dashboardSheet.getRange("J1:L1").merge().setValue("📈 全場總疏運/離場完成率").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
  dashboardSheet.getRange("J2:L3").merge().setFormula(`=IF(D2>0, TEXT(G2/D2, "0.0%"), "0.0%")`).setBackground("#0F172A").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 全場疏運進度條 (第 4 列)
  dashboardSheet.getRange("A4:C4").merge().setValue("⚡ 總疏運進度").setBackground("#1E293B").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D4:L4").merge().setFormula(`=IF(D2>0, SPARKLINE(G2, {"charttype", "bar"; "max", D2; "color1", "#38BDF8"}), "")`).setBackground("#1E293B");

  // 2. 🚌【上半部：接駁車輛疏運專區 (4大站點)】(第 5~10 列)
  dashboardSheet.getRange("A5:L5").merge()
    .setValue("🚌 接駁車輛疏運專區（共 4 條路線 / 130 輛車）")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("left");

  const busStations = [
    { name: "📍 成功車站（綠線）", startCol: 1, detailGoCol: "B", detailBackCol: "D", detailEndRow: 22, tagColor: "#059669" },
    { name: "📍 新烏日台鐵站（藍線）", startCol: 4, detailGoCol: "G", detailBackCol: "I", detailEndRow: 52, tagColor: "#2563EB" },
    { name: "📍 經貿六停車場（黃線）", startCol: 7, detailGoCol: "L", detailBackCol: "N", detailEndRow: 42, tagColor: "#D97706" },
    { name: "📍 水湳轉運站（黃線）", startCol: 10, detailGoCol: "Q", detailBackCol: "S", detailEndRow: 22, tagColor: "#D97706" }
  ];

  busStations.forEach(cs => {
    const col = cs.startCol;
    
    // 站名大標題 (第 6 列)
    dashboardSheet.getRange(6, col, 1, 3).merge()
      .setValue(cs.name)
      .setBackground(cs.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    // 項目標籤 (第 7 列)
    dashboardSheet.getRange(7, col).setValue("👉 去程人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 1).setValue("👈 返程人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 2).setValue("🏆 累計總量").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 8 列)
    const colLetterGo = String.fromCharCode(64 + col);
    const colLetterBack = String.fromCharCode(64 + col + 1);

    dashboardSheet.getRange(8, col).setFormula(`=SUM('各車即時明細'!${cs.detailGoCol}3:${cs.detailGoCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 1).setFormula(`=SUM('各車即時明細'!${cs.detailBackCol}3:${cs.detailBackCol}${cs.detailEndRow})`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 2).setFormula(`=${colLetterGo}8+${colLetterBack}8`).setBackground("#FFFFFF").setFontColor("#2563EB").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");

    // 疏運率標題 (第 9 列)
    dashboardSheet.getRange(9, col, 1, 2).merge()
      .setFormula(`=IF(${colLetterGo}8>0, "📈 返程疏運率: " & TEXT(${colLetterBack}8/${colLetterGo}8, "0.0%"), "📈 返程疏運率: 0.0%")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    dashboardSheet.getRange(9, col + 2).setFormula(`=IF(${colLetterGo}8>0, "尚餘 " & MAX(0, ${colLetterGo}8 - ${colLetterBack}8) & " 人", "已完成")`)
      .setBackground("#F8FAFC").setFontColor("#EF4444").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條 (第 10 列)
    dashboardSheet.getRange(10, col, 1, 3).merge()
      .setFormula(`=IF(${colLetterGo}8>0, SPARKLINE(${colLetterBack}8, {"charttype", "bar"; "max", ${colLetterGo}8; "color1", "#2563EB"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(6, col, 5, 3).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
    
    dashboardSheet.setColumnWidth(col, 110);
    dashboardSheet.setColumnWidth(col + 1, 110);
    dashboardSheet.setColumnWidth(col + 2, 110);
  });

  // 3. 🚶【下半部：步行進出通道專區 (1/3/4號門，色系對應)】(第 11~16 列)
  dashboardSheet.getRange("A11:L11").merge()
    .setValue("🚶 步行進出通道專區（1號門綠線 / 3號門藍線 / 4號門黃線）")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("left");

  const walkGates = [
    { name: "🚶 1號門（綠線）", keyword: "1號門", startCol: 1, tagColor: "#059669", barColor: "#059669" },
    { name: "🚶 3號門（藍線）", keyword: "3號門", startCol: 4, tagColor: "#2563EB", barColor: "#2563EB" },
    { name: "🚶 4號門（黃線）", keyword: "4號門", startCol: 7, tagColor: "#D97706", barColor: "#D97706" },
    { name: "🚶 步行通道小計", startCol: 10, isSummary: true, tagColor: "#475569", barColor: "#8B5CF6" }
  ];

  walkGates.forEach(wg => {
    const col = wg.startCol;
    
    // 門名大標題 (第 12 列)
    dashboardSheet.getRange(12, col, 1, 3).merge()
      .setValue(wg.name)
      .setBackground(wg.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    // 項目標籤 (第 13 列)
    dashboardSheet.getRange(13, col).setValue("👉 入場人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(13, col + 1).setValue("👈 離場人數").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(13, col + 2).setValue("🏆 門區總量").setBackground("#F1F5F9").setFontColor("#475569").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 數據列 (第 14 列)
    const colLetterGo = String.fromCharCode(64 + col);
    const colLetterBack = String.fromCharCode(64 + col + 1);

    if (!wg.isSummary) {
      dashboardSheet.getRange(14, col).setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*去程*", '${formSheetName}'!C:C, "*${wg.keyword}*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
      dashboardSheet.getRange(14, col + 1).setFormula(`=SUMIFS('${formSheetName}'!E:E, '${formSheetName}'!B:B, "*返程*", '${formSheetName}'!C:C, "*${wg.keyword}*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
      dashboardSheet.getRange(14, col + 2).setFormula(`=${colLetterGo}14+${colLetterBack}14`).setBackground("#FFFFFF").setFontColor("#2563EB").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    } else {
      // 步行小計 = 1號門 + 3號門 + 4號門
      dashboardSheet.getRange(14, col).setFormula(`=A14+D14+G14`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
      dashboardSheet.getRange(14, col + 1).setFormula(`=B14+E14+H14`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
      dashboardSheet.getRange(14, col + 2).setFormula(`=C14+F14+I14`).setBackground("#FFFFFF").setFontColor("#7C3AED").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    }

    // 疏運率標題 (第 15 列)
    dashboardSheet.getRange(15, col, 1, 2).merge()
      .setFormula(`=IF(${colLetterGo}14>0, "📈 離場疏運率: " & TEXT(${colLetterBack}14/${colLetterGo}14, "0.0%"), "📈 離場疏運率: 0.0%")`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    dashboardSheet.getRange(15, col + 2).setFormula(`=IF(${colLetterGo}14>0, "尚餘 " & MAX(0, ${colLetterGo}14 - ${colLetterBack}14) & " 人", "已完成")`)
      .setBackground("#F8FAFC").setFontColor("#EF4444").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條 (第 16 列)
    dashboardSheet.getRange(16, col, 1, 3).merge()
      .setFormula(`=IF(${colLetterGo}14>0, SPARKLINE(${colLetterBack}14, {"charttype", "bar"; "max", ${colLetterGo}14; "color1", "${wg.barColor}"}), "")`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(12, col, 5, 3).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 列高設定
  dashboardSheet.setRowHeight(1, 26);
  dashboardSheet.setRowHeight(2, 28);
  dashboardSheet.setRowHeight(3, 28);
  dashboardSheet.setRowHeight(4, 24);
  dashboardSheet.setRowHeight(5, 26); // 接駁車專區分隔列
  dashboardSheet.setRowHeight(6, 34); // 車站名
  dashboardSheet.setRowHeight(7, 26); // 去/返標籤
  dashboardSheet.setRowHeight(8, 38); // 大數字
  dashboardSheet.setRowHeight(9, 26); // 疏運率
  dashboardSheet.setRowHeight(10, 18); // 進度條
  dashboardSheet.setRowHeight(11, 26); // 步行通道專區分隔列
  dashboardSheet.setRowHeight(12, 34); // 門名
  dashboardSheet.setRowHeight(13, 26); // 入/離標籤
  dashboardSheet.setRowHeight(14, 38); // 大數字
  dashboardSheet.setRowHeight(15, 26); // 疏運率
  dashboardSheet.setRowHeight(16, 18); // 進度條

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新人車分區・路線色對應 戰情看板建立完成！】");
  Logger.log("\n📱【最新回報表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
