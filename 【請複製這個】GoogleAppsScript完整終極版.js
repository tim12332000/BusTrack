/**
 * =========================================================================
 * 🚌 接駁車管理系統 - 【純淨高階戰情大螢幕版 (進度條避開綠色混淆)】
 * =========================================================================
 * 
 * ✨ 顏色優化：
 * - 進度條改為【電光亮青色 (#38BDF8)】或【活力橙 (#F97316)】，徹底避開綠色，不與綠線/返程標籤混淆！
 * - 路線色（成功-綠 / 新烏日-藍 / 經貿六-黃 / 水湳-黃）清晰分明！
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行 (Run)」即可！
 * =========================================================================
 */

function createMultiStationBusSystem() {
  Logger.log("🚀 開始建立【純淨高階戰情大螢幕系統】...");

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
  // 【A. 第二頁：各車即時明細 (計算底層數據)】
  // ==========================================
  const stations = [
    { name: "成功車站（綠線）", formKeyword: "成功車站", startCol: 1, busCount: 20, bg: "#059669" },
    { name: "新烏日台鐵站（藍線）", formKeyword: "新烏日", startCol: 6, busCount: 50, bg: "#2563EB" },
    { name: "經貿六停車場（黃線）", formKeyword: "經貿六", startCol: 11, busCount: 40, bg: "#D97706" },
    { name: "水湳轉運站（黃線）", formKeyword: "水湳", startCol: 16, busCount: 20, bg: "#D97706" }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    detailSheet.getRange(1, col, 1, 5).merge()
      .setValue(`📍 ${st.name}`)
      .setBackground(st.bg).setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");

    detailSheet.getRange(2, col, 1, 5)
      .setValues([["車號", "👉去程人數", "去程時間", "👈返程人數", "返程時間"]])
      .setBackground("#334155").setFontColor("#E2E8F0").setFontWeight("bold").setHorizontalAlignment("center");

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
    detailSheet.getRange(2, col, st.busCount + 1, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);
  });

  // ==========================================
  // 【B. 第一頁：總即時戰情看板 (大氣純淨無明細)】
  // ==========================================
  try {
    dashboardSheet.getRange(1, 1, 40, 20).breakApart();
  } catch (e) {}

  // 1. 頂部三大超醒目 KPI 戰情大卡片
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總運輸人次").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#020617").setFontColor("#38BDF8").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 全場去程總人數").setBackground("#1E3A8A").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=B8+E8+H8+K8").setBackground("#172554").setFontColor("#60A5FA").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 全場返程總人數").setBackground("#064E3B").setFontColor("#6EE7B7").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=C8+F8+I8+L8").setBackground("#022C22").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("J1:L1").merge().setValue("📈 全場返程疏運完成率").setBackground("#312E81").setFontColor("#C7D2FE").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");
  dashboardSheet.getRange("J2:L3").merge().setFormula(`=IF(D2>0, TEXT(G2/D2, "0.0%"), "0.0%")`).setBackground("#1E1B4B").setFontColor("#FDE047").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  // 全場總疏運進度條 (第 4 列：改為電光亮青色 #38BDF8，徹底避開綠色)
  dashboardSheet.getRange("A4:C4").merge().setValue("⚡ 全場疏運進度").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D4:L4").merge().setFormula(`=IF(D2>0, SPARKLINE(G2, {"charttype", "bar"; "max", D2; "color1", "#38BDF8"}), "")`).setBackground("#0F172A");

  // 2. 4 大站點獨立戰情大卡片
  const cardStations = [
    { name: "📍 成功車站（綠線）", startCol: 1, detailGoCol: "B", detailBackCol: "D", detailEndRow: 22, bg: "#059669" },
    { name: "📍 新烏日台鐵站（藍線）", startCol: 4, detailGoCol: "G", detailBackCol: "I", detailEndRow: 52, bg: "#2563EB" },
    { name: "📍 經貿六停車場（黃線）", startCol: 7, detailGoCol: "L", detailBackCol: "N", detailEndRow: 42, bg: "#D97706" },
    { name: "📍 水湳轉運站（黃線）", startCol: 10, detailGoCol: "Q", detailBackCol: "S", detailEndRow: 22, bg: "#D97706" }
  ];

  cardStations.forEach(cs => {
    const col = cs.startCol;
    
    // 站名大標題 (第 6 列)
    dashboardSheet.getRange(6, col, 1, 3).merge()
      .setValue(cs.name)
      .setBackground(cs.bg).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(15).setHorizontalAlignment("center");

    // 標題說明列 (第 7 列)
    dashboardSheet.getRange(7, col).setValue("👉 去程人數").setBackground("#1E3A8A").setFontColor("#93C5FD").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 1).setValue("👈 返程人數").setBackground("#064E3B").setFontColor("#6EE7B7").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(7, col + 2).setValue("🏆 本站總量").setBackground("#312E81").setFontColor("#FDE047").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 數據列 (第 8 列)
    const colLetterGo = String.fromCharCode(64 + col);
    const colLetterBack = String.fromCharCode(64 + col + 1);

    dashboardSheet.getRange(8, col).setFormula(`=SUM('各車即時明細'!${cs.detailGoCol}3:${cs.detailGoCol}${cs.detailEndRow})`).setBackground("#172554").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 1).setFormula(`=SUM('各車即時明細'!${cs.detailBackCol}3:${cs.detailBackCol}${cs.detailEndRow})`).setBackground("#022C22").setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");
    dashboardSheet.getRange(8, col + 2).setFormula(`=${colLetterGo}8+${colLetterBack}8`).setBackground("#1E1B4B").setFontColor("#FDE047").setFontWeight("bold").setFontSize(18).setHorizontalAlignment("center");

    // 疏運率標題 (第 9 列：返程 / 去程)
    dashboardSheet.getRange(9, col, 1, 2).merge()
      .setFormula(`=IF(${colLetterGo}8>0, "📈 返程疏運率: " & TEXT(${colLetterBack}8/${colLetterGo}8, "0.0%"), "📈 返程疏運率: 0.0%")`)
      .setBackground("#1E293B").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    dashboardSheet.getRange(9, col + 2).setFormula(`=IF(${colLetterGo}8>0, "尚餘 " & MAX(0, ${colLetterGo}8 - ${colLetterBack}8) & " 人", "已完成")`)
      .setBackground("#1E293B").setFontColor("#F87171").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    // 疏運進度條 (第 10 列：改用高對比電光亮青色 #38BDF8，絕不與綠/藍/黃撞色)
    dashboardSheet.getRange(10, col, 1, 3).merge()
      .setFormula(`=IF(${colLetterGo}8>0, SPARKLINE(${colLetterBack}8, {"charttype", "bar"; "max", ${colLetterGo}8; "color1", "#38BDF8"}), "")`)
      .setBackground("#0F172A");

    dashboardSheet.getRange(6, col, 5, 3).setBorder(true, true, true, true, true, true, "#475569", SpreadsheetApp.BorderStyle.SOLID);
    dashboardSheet.setColumnWidth(col, 110);
    dashboardSheet.setColumnWidth(col + 1, 110);
    dashboardSheet.setColumnWidth(col + 2, 110);
  });

  dashboardSheet.setRowHeight(1, 28);
  dashboardSheet.setRowHeight(2, 28);
  dashboardSheet.setRowHeight(3, 28);
  dashboardSheet.setRowHeight(4, 26);
  dashboardSheet.setRowHeight(5, 16);
  dashboardSheet.setRowHeight(6, 38);
  dashboardSheet.setRowHeight(7, 30);
  dashboardSheet.setRowHeight(8, 42);
  dashboardSheet.setRowHeight(9, 32);
  dashboardSheet.setRowHeight(10, 26);

  const newFormUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉【全新進度條高對比配色 戰情看板建立完成！】");
  Logger.log("\n📱【最新車長填寫表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function fixAndFormatEverything() { createMultiStationBusSystem(); }
function autoFixDashboard() { createMultiStationBusSystem(); }
function forceFormatTimeAsText() { createMultiStationBusSystem(); }
