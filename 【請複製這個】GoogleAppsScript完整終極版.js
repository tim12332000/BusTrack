/**
 * =========================================================================
 * 🚌 一鍵建立全新【4大站點 去/回 表單 + 試算表看板】
 * =========================================================================
 * 
 * 👉 使用方式：
 * 1. 按 Ctrl + A 全選複製。
 * 2. 回到 Google Apps Script 全選貼上。
 * 3. 點擊「執行 (Run)」。
 * 4. 下方會直接產生【全新 4 站點 去/回 車長填寫表單網址】！
 * =========================================================================
 */

function fixAndFormatEverything() {
  Logger.log("🚀 開始為您建立全新 4 站點【去/回】表單與戰情看板...");

  // 1. 建立全新 Google 表單 (4 站點 + 去/回 + 人數)
  const form = FormApp.create("🚌 接駁車【去/回】搭乘人數即時回報");
  form.setDescription("請車長於各趟次發車時回報搭乘人數。\n只需選擇【站點】、【車號】、【去/回程】及【人數】。");

  // 題目 1: 接駁站點
  form.addMultipleChoiceItem()
    .setTitle("1. 選擇接駁站點")
    .setChoiceValues(["成功車站", "新烏日台鐵站", "經貿六停車場", "水湳轉運站"])
    .setRequired(true);

  // 題目 2: 車號 (1 ~ 50 號車)
  const busChoices = [];
  for (let i = 1; i <= 50; i++) {
    busChoices.push(`${i} 號車`);
  }
  form.addListItem()
    .setTitle("2. 選擇車號")
    .setChoiceValues(busChoices)
    .setRequired(true);

  // 題目 3: 行程方向 (去程 / 回程)
  form.addMultipleChoiceItem()
    .setTitle("3. 行程方向")
    .setChoiceValues(["👉 去程", "👈 回程"])
    .setRequired(true);

  // 題目 4: 實到搭乘人數
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

  // 2. 建立/開啟 試算表並綁定
  const ss = SpreadsheetApp.create("🚌 各站接駁車【去/回】即時戰情中心");
  const ssId = ss.getId();
  const ssUrl = ss.getUrl();

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  Utilities.sleep(2000); // 等待 Google 表單回應分頁生成

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

  // 頂部 KPI 卡片
  dashboardSheet.getRange("A1:C1").merge().setValue("🏆 全場總運輸人次").setBackground("#0F172A").setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("A2:C3").merge().setFormula("=D2+G2").setBackground("#020617").setFontColor("#38BDF8").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("D1:F1").merge().setValue("👉 去程總人數").setBackground("#1E3A8A").setFontColor("#93C5FD").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("D2:F3").merge().setFormula("=SUM(B7:B26) + SUM(G7:G56) + SUM(L7:L46) + SUM(Q7:Q26)").setBackground("#172554").setFontColor("#60A5FA").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange("G1:I1").merge().setValue("👈 回程總人數").setBackground("#064E3B").setFontColor("#6EE7B7").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("G2:I3").merge().setFormula("=SUM(D7:D26) + SUM(I7:I56) + SUM(N7:N46) + SUM(S7:S26)").setBackground("#022C22").setFontColor("#34D399").setFontSize(22).setFontWeight("bold").setHorizontalAlignment("center");

  // 4 大站點
  const stations = [
    { name: "成功車站", startCol: 1, busCount: 20, bg: "#3B82F6" },
    { name: "新烏日台鐵站", startCol: 6, busCount: 50, bg: "#8B5CF6" },
    { name: "經貿六停車場", startCol: 11, busCount: 40, bg: "#EC4899" },
    { name: "水湳轉運站", startCol: 16, busCount: 20, bg: "#10B981" }
  ];

  stations.forEach(st => {
    const col = st.startCol;
    dashboardSheet.getRange(5, col, 1, 5).merge()
      .setValue(`📍 ${st.name} (共 ${st.busCount} 輛)`)
      .setBackground(st.bg).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center");

    dashboardSheet.getRange(6, col, 1, 5)
      .setValues([["車號", "👉去程人數", "去程時間", "👈回程人數", "回程時間"]])
      .setBackground("#1E293B").setFontColor("#E2E8F0").setFontWeight("bold").setHorizontalAlignment("center");

    const busRows = [];
    for (let b = 1; b <= st.busCount; b++) {
      const busName = `${b} 號車`;
      const fGoActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), '${formSheetName}'!B:B="${st.name}", '${formSheetName}'!C:C="${busName}", '${formSheetName}'!D:D="👉 去程"))), 0)`;
      const fGoTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), '${formSheetName}'!B:B="${st.name}", '${formSheetName}'!C:C="${busName}", '${formSheetName}'!D:D="👉 去程"))), "hh:mm:ss"), "-")`;
      
      const fBackActual = `=IFERROR(INDEX('${formSheetName}'!E:E, MAX(FILTER(ROW('${formSheetName}'!E:E), '${formSheetName}'!B:B="${st.name}", '${formSheetName}'!C:C="${busName}", '${formSheetName}'!D:D="👈 回程"))), 0)`;
      const fBackTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!A:A, MAX(FILTER(ROW('${formSheetName}'!A:A), '${formSheetName}'!B:B="${st.name}", '${formSheetName}'!C:C="${busName}", '${formSheetName}'!D:D="👈 回程"))), "hh:mm:ss"), "-")`;

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
  Logger.log("🎉【全新 4 站點 去/回 表單與看板建立完成！】");
  Logger.log("\n📱【最新車長填寫表單網址】:\n" + newFormUrl);
  Logger.log("\n📊【最新主控即時戰情看板網址】:\n" + ssUrl);
  Logger.log("=======================================================\n");
}

function createMultiStationBusSystem() { fixAndFormatEverything(); }
function autoFixDashboard() { fixAndFormatEverything(); }
function forceFormatTimeAsText() { fixAndFormatEverything(); }
