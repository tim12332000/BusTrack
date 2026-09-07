/**
 * =========================================================================
 * 🎖️「國防知性之旅-成功嶺營區開放」進場人數與六大停車場統計表 - 【雙表單終極旗艦版】
 * =========================================================================
 * 
 * 🔧 本版核心升級：
 * 1. 【雙表單整合聯防】：
 *    - 表單 1：人員疏運（4大接駁站 + 3大步行門）
 *    - 表單 2：六大停車場車位回報（A、B、C、D、E、F 區，各 500 輛，共 3,000 席）
 *    - 兩套表單自動綁定在同一個試算表（分頁：表單回應 1、表單回應 2），互不干擾！
 * 2. 【防欄位漂移機制】：
 *    - 題目存在時不重複刪建，永久鎖定欄位結構！
 * 3. 【三合一綜合戰情看板】：
 *    - 人員疏運 + 停車場在席車位全部彙整在同一個工作表「總即時戰情看板」（gid=0）！
 *    - 前端大螢幕 HTML 一次 JSONP 抓取即可全部同步更新！
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行」即可！
 * =========================================================================
 */

const TARGET_SPREADSHEET_ID = "1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w";

function createMultiStationBusSystem() {
  Logger.log("🎨 開始對指定試算表 [" + TARGET_SPREADSHEET_ID + "] 進行雙表單建置與看板重繪...");

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  
  // =========================================================================
  // 1. 維護【表單 1：人員疏運回報】(依專屬名稱鎖定獨立表單，絕不與表單2混淆)
  // =========================================================================
  let form1 = null;
  const form1Files = DriveApp.getFilesByName("「國防知性之旅-成功嶺營區開放」人數回報");
  if (form1Files.hasNext()) {
    try {
      form1 = FormApp.openById(form1Files.next().getId());
    } catch (e) {}
  }

  if (form1) {
    form1.setTitle("「國防知性之旅-成功嶺營區開放」人數回報");
    const items1 = form1.getItems();
    if (items1.length !== 5) {
      for (let i = items1.length - 1; i >= 0; i--) {
        form1.deleteItem(items1[i]);
      }

      // 題目 1: 方向 (進場 / 離場)
      form1.addMultipleChoiceItem()
        .setTitle("1. 方向")
        .setChoiceValues(["進場", "離場"])
        .setRequired(true);

      // 題目 2: 站點 / 門號
      form1.addMultipleChoiceItem()
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
      form1.addListItem()
        .setTitle("3. 車號")
        .setChoiceValues(busChoices)
        .setRequired(true);

      // 題目 4: 人數
      const textValidation = FormApp.createTextValidation()
        .setHelpText("請輸入數字")
        .requireNumberGreaterThanOrEqualTo(0)
        .build();
      form1.addTextItem()
        .setTitle("4. 人數")
        .setValidation(textValidation)
        .setRequired(true);

      // 題目 5: 備註
      form1.addTextItem()
        .setTitle("5. 備註")
        .setRequired(false);
    }
  }

  // =========================================================================
  // 2. 建立或維護【表單 2：六大停車場車位回報】(A~F 各 500 席，總共 3,000 席)
  // =========================================================================
  const props = PropertiesService.getScriptProperties();
  let parkingFormId = props.getProperty("PARKING_FORM_ID");
  let parkingForm = null;

  if (parkingFormId) {
    try {
      parkingForm = FormApp.openById(parkingFormId);
    } catch (e) {
      parkingForm = null;
    }
  }

  if (!parkingForm) {
    // 檢查雲端是否已有同名表單
    const formFiles = DriveApp.getFilesByName("「國防知性之旅-成功嶺營區開放」六大停車場車位回報");
    if (formFiles.hasNext()) {
      const file = formFiles.next();
      parkingForm = FormApp.openById(file.getId());
      props.setProperty("PARKING_FORM_ID", file.getId());
    }
  }

  if (!parkingForm) {
    // 建立新停車場表單並綁定至試算表
    parkingForm = FormApp.create("「國防知性之旅-成功嶺營區開放」六大停車場車位回報");
    parkingForm.setDescription("現場交通哨專用：各停車場容量均為 500 輛。請於車流進出時即時回報車輛數。");
    parkingForm.setDestination(FormApp.DestinationType.SPREADSHEET, TARGET_SPREADSHEET_ID);
    props.setProperty("PARKING_FORM_ID", parkingForm.getId());
  }

  if (parkingForm) {
    parkingForm.setTitle("「國防知性之旅-成功嶺營區開放」六大停車場車位回報");
    const items2 = parkingForm.getItems();
    if (items2.length !== 4) {
      for (let i = items2.length - 1; i >= 0; i--) {
        parkingForm.deleteItem(items2[i]);
      }

      // 題目 1: 停車場區域
      parkingForm.addMultipleChoiceItem()
        .setTitle("1. 停車場區域")
        .setChoiceValues([
          "🅿️ A 區停車場 (上限 500 輛)",
          "🅿️ B 區停車場 (上限 500 輛)",
          "🅿️ C 區停車場 (上限 500 輛)",
          "🅿️ D 區停車場 (上限 500 輛)",
          "🅿️ E 區停車場 (上限 500 輛)",
          "🅿️ F 區停車場 (上限 500 輛)"
        ])
        .setRequired(true);

      // 題目 2: 回報項目 / 動作
      parkingForm.addMultipleChoiceItem()
        .setTitle("2. 回報項目 / 動作")
        .setChoiceValues([
          "🚗 汽車進場",
          "🚙 汽車離場"
        ])
        .setRequired(true);

      // 題目 3: 車輛數量
      const carValidation = FormApp.createTextValidation()
        .setHelpText("請輸入車輛數量（正整數）")
        .requireNumberGreaterThanOrEqualTo(1)
        .build();
      parkingForm.addTextItem()
        .setTitle("3. 車輛數量 (輛)")
        .setValidation(carValidation)
        .setRequired(true);

      // 題目 4: 備註
      parkingForm.addTextItem()
        .setTitle("4. 備註")
        .setRequired(false);
    }
  }

  // =========================================================================
  // 3. 識別與整理各工作表
  // =========================================================================
  let dashboardSheet = ss.getSheetByName("總即時戰情看板");
  if (!dashboardSheet) {
    dashboardSheet = ss.getSheets()[0];
    dashboardSheet.setName("總即時戰情看板");
  }

  let detailSheet = ss.getSheetByName("各車即時明細");
  if (!detailSheet) {
    detailSheet = ss.insertSheet("各車即時明細");
  }

  // 識別人員回報表單分頁 (表單回應 1)
  let formSheetName = "表單回應 1";
  for (let s of ss.getSheets()) {
    const sName = s.getName();
    if (sName !== "總即時戰情看板" && sName !== "各車即時明細" && !sName.includes("停車場") && !sName.includes("回應 2")) {
      formSheetName = sName;
      break;
    }
  }

  // 識別停車場回報表單分頁 (尋找標題包含「停車場」或「車輛數量」的真實回報分頁)
  let parkingSheetName = "表單回應 3";
  let maxParkingRows = -1;

  for (let s of ss.getSheets()) {
    const sName = s.getName();
    if (sName === "總即時戰情看板" || sName === "各車即時明細" || sName === formSheetName) continue;

    try {
      const lastCol = s.getLastColumn();
      if (lastCol >= 3) {
        const headerValues = s.getRange(1, 1, 1, Math.min(lastCol, 10)).getValues()[0].join(" ");
        if (headerValues.includes("停車場") || headerValues.includes("車輛數量")) {
          const rowCount = s.getLastRow();
          if (rowCount > maxParkingRows) {
            maxParkingRows = rowCount;
            parkingSheetName = sName;
          }
        }
      }
    } catch (e) {}
  }
  let pSheet = ss.getSheetByName(parkingSheetName);
  if (!pSheet) {
    pSheet = ss.insertSheet(parkingSheetName);
    pSheet.getRange(1, 1, 1, 5).setValues([["時間戳記", "1. 停車場區域", "2. 回報項目 / 動作", "3. 車輛數量 (輛)", "4. 備註"]]);
    pSheet.getRange(1, 1, 1, 5).setBackground("#334155").setFontColor("#FFFFFF").setFontWeight("bold");
  }

  // 🎯 動態精確抓取「區域」、「動作」、「數量」所在欄位字母 (防 Google 表單欄位向右偏移到 F/G/H 欄)
  let colAreaLetter = "B";
  let colActionLetter = "C";
  let colQtyLetter = "D";

  if (pSheet) {
    const lastCol = Math.max(pSheet.getLastColumn(), 10);
    const headerRow = pSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    headerRow.forEach((h, idx) => {
      const str = String(h || "");
      const letter = String.fromCharCode(65 + idx);
      if (str.includes("停車場區域")) colAreaLetter = letter;
      if (str.includes("回報項目") || str.includes("動作")) colActionLetter = letter;
      if (str.includes("車輛數量") || str.includes("數量")) colQtyLetter = letter;
    });
  }
  Logger.log(`🎯 停車場真實欄位鎖定：分頁=[${parkingSheetName}], 區域=[${colAreaLetter}欄], 動作=[${colActionLetter}欄], 數量=[${colQtyLetter}欄]`);

  // 🛠️ 清理人員表單回應欄位漂移
  let formSheet = ss.getSheetByName(formSheetName);
  if (formSheet) {
    const lastCol = formSheet.getLastColumn();
    if (lastCol > 6) {
      formSheet.deleteColumns(2, lastCol - 6);
      Logger.log("✨ 成功校準人員表單回應欄位！最新數據已自動移回 B、C、D、E 欄！");
    }
  }

  // =========================================================================
  // 【A. 第二頁：各車即時明細 (4 大車站車輛明細)】
  // =========================================================================
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

  // =========================================================================
  // 【B. 第一頁：總即時戰情看板 (人員疏運 + 六大停車場全合一)】
  // =========================================================================
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
    .setValue("「國防知性之旅-成功嶺營區開放」進場人數與停車場即時戰情表")
    .setBackground(THEME_MAIN_BANNER).setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(16).setHorizontalAlignment("center");

  // -------------------------------------------------------------------------
  // 區塊一：接駁車大盤 (第 2~5 列)
  // -------------------------------------------------------------------------
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

  // -------------------------------------------------------------------------
  // 區塊二：步行通道大盤 (第 14~17 列)
  // -------------------------------------------------------------------------
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

  // -------------------------------------------------------------------------
  // 區塊三：🅿️ 六大停車場車位大盤 (第 26~29 列) - 總容量 3,000 席
  // -------------------------------------------------------------------------
  dashboardSheet.getRange(26, 1, 1, 24).merge()
    .setValue("🅿️ 汽車停車場車位即時監控 (總容量 3,000 席)")
    .setBackground(THEME_HEADER_BG).setFontColor("#FBBF24").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

  dashboardSheet.getRange(27, 1, 1, 8).merge().setValue("👉 全區已停汽車").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  // A34(A區)+E34(B區)+I34(C區)+M34(D區)+Q34(E區)+U34(F區)
  dashboardSheet.getRange(28, 1, 1, 8).merge().setFormula("=A34+E34+I34+M34+Q34+U34").setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(27, 9, 1, 8).merge().setValue("👈 全區剩餘車位").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(28, 9, 1, 8).merge().setFormula("=MAX(0, 3000-A28)").setBackground("#1E293B").setFontColor("#38BDF8").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(27, 17, 1, 8).merge().setValue("📊 全區車位佔用率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(28, 17, 1, 8).merge().setFormula(`=TEXT(A28/3000, "0.0%")`).setBackground("#1E293B").setFontColor("#FBBF24").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(29, 1, 1, 6).merge().setValue("全區車位使用率").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(29, 7, 1, 18).merge().setFormula(`=SPARKLINE(A28, {"charttype", "bar"; "max", 3000; "color1", "#F59E0B"})`).setBackground(THEME_HEADER_BG);
  dashboardSheet.getRange(26, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------------------
  // 區塊四：六大停車場獨立卡片 (第 31~36 列) - 每區 500 席，共 24 欄 (每區 4 欄)
  // -------------------------------------------------------------------------
  dashboardSheet.getRange(31, 1, 1, 24).merge()
    .setValue("各區停車場即時車位 (每區容量 500 席)")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const parkingLots = [
    { name: "🅿️ A 區 (500席)", keyword: "A 區", startCol: 1, tagColor: "#2563EB" },
    { name: "🅿️ B 區 (500席)", keyword: "B 區", startCol: 5, tagColor: "#059669" },
    { name: "🅿️ C 區 (500席)", keyword: "C 區", startCol: 9, tagColor: "#D97706" },
    { name: "🅿️ D 區 (500席)", keyword: "D 區", startCol: 13, tagColor: "#7C3AED" },
    { name: "🅿️ E 區 (500席)", keyword: "E 區", startCol: 17, tagColor: "#DB2777" },
    { name: "🅿️ F 區 (500席)", keyword: "F 區", startCol: 21, tagColor: "#0D9488" }
  ];

  parkingLots.forEach(lot => {
    const col = lot.startCol;
    // 標題 (Row 32)
    dashboardSheet.getRange(32, col, 1, 4).merge()
      .setValue(lot.name)
      .setBackground(lot.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");

    // 標籤 (Row 33)
    dashboardSheet.getRange(33, col, 1, 2).merge().setValue("已停").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(33, col + 2, 1, 2).merge().setValue("剩餘").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    const colLetterPark = String.fromCharCode(64 + col);
    const colLetterRemain = String.fromCharCode(64 + col + 2);

    // 數值 (Row 34)
    dashboardSheet.getRange(34, col, 1, 2).merge()
      .setFormula(`=MAX(0, SUMIFS('${parkingSheetName}'!${colQtyLetter}:${colQtyLetter}, '${parkingSheetName}'!${colAreaLetter}:${colAreaLetter}, "*${lot.keyword}*", '${parkingSheetName}'!${colActionLetter}:${colActionLetter}, "*進場*") - SUMIFS('${parkingSheetName}'!${colQtyLetter}:${colQtyLetter}, '${parkingSheetName}'!${colAreaLetter}:${colAreaLetter}, "*${lot.keyword}*", '${parkingSheetName}'!${colActionLetter}:${colActionLetter}, "*離場*"))`)
      .setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(22).setHorizontalAlignment("center");

    dashboardSheet.getRange(34, col + 2, 1, 2).merge()
      .setFormula(`=MAX(0, 500 - ${colLetterPark}34)`)
      .setBackground("#FFFFFF").setFontColor("#0284C7").setFontWeight("bold").setFontSize(22).setHorizontalAlignment("center");

    // 佔用率與狀態 (Row 35)
    dashboardSheet.getRange(35, col, 1, 4).merge()
      .setFormula(`=TEXT(${colLetterPark}34/500, "0.0%") & "  ·  尚餘 " & ${colLetterRemain}34 & " 席"`)
      .setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");

    // 迷你進度條 (Row 36)
    dashboardSheet.getRange(36, col, 1, 4).merge()
      .setFormula(`=SPARKLINE(${colLetterPark}34, {"charttype", "bar"; "max", 500; "color1", "${THEME_CARD_BAR}"})`)
      .setBackground("#F1F5F9");

    dashboardSheet.getRange(32, col, 5, 4).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 設定列高
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
  dashboardSheet.setRowHeight(25, 14);
  dashboardSheet.setRowHeight(26, 26);
  dashboardSheet.setRowHeight(27, 22);
  dashboardSheet.setRowHeight(28, 38);
  dashboardSheet.setRowHeight(29, 24);
  dashboardSheet.setRowHeight(30, 12);
  dashboardSheet.setRowHeight(31, 26);
  dashboardSheet.setRowHeight(32, 34);
  dashboardSheet.setRowHeight(33, 26);
  dashboardSheet.setRowHeight(34, 38);
  dashboardSheet.setRowHeight(35, 28);
  dashboardSheet.setRowHeight(36, 18);

  Logger.log("\n=======================================================");
  Logger.log("🎉【雙表單整合戰情室建置完成！】");
  const form1UrlFinal = (form1 && form1.getPublishedUrl().includes("1FAIpQLSe")) ? form1.getPublishedUrl() : "https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform";
  const form2UrlFinal = (parkingForm && parkingForm.getPublishedUrl().includes("1FAIpQLSd")) ? parkingForm.getPublishedUrl() : "https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform";
  Logger.log("📱【表單 1・人員疏運回報專用】: " + form1UrlFinal);
  Logger.log("🅿️【表單 2・六大停車場車位回報】: " + form2UrlFinal);
  Logger.log("📊【Google 試算表看板網址】: " + ss.getUrl());
  Logger.log("=======================================================\n");
}
