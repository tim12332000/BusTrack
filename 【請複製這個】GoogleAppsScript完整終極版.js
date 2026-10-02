/**
 * =========================================================================
 * 🎖️「國防知性之旅-成功嶺營區開放」進場人數與六大停車場統計表 - 【雙表單終極旗艦版】
 * =========================================================================
 * 
 * 🔧 本版核心升級：
 * 1. 【雙表單整合聯防】：
 *    - 表單 1：人員疏運（5大接駁站 + 2大步行門）
 *    - 表單 2：七處停車場剩餘汽車位、機車位（每場以最新回報為準）
 *    - 兩套表單自動綁定在同一個試算表（分頁：表單回應 1、表單回應 2），互不干擾！
 * 2. 【防欄位漂移機制】：
 *    - 題目存在時不重複刪建，永久鎖定欄位結構！
 * 3. 【三合一綜合戰情看板】：
 *    - 人員疏運 + 停車場在席車位全部彙整在同一個工作表「總即時戰情看板」（依分頁名稱讀取）！
 *    - 前端大螢幕 HTML 一次 JSONP 抓取即可全部同步更新！
 * 
 * 👉 使用方式：全選複製貼到 Google Apps Script 覆蓋，點「執行」即可！
 * =========================================================================
 */

const TARGET_SPREADSHEET_ID = "1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w";

// 日常修復請執行此函式：只修復公式與看板，不修改表單題目或回應資料。
function repairDashboard() {
  createMultiStationBusSystem(true);
}

// 多個同類回應分頁無法唯一辨識時，在此填入目前使用的分頁名稱。
const RESPONSE_SHEET_NAMES = { people: "", parking: "" };

const PEOPLE_STATION_CHOICES = [
  "🚌 成功車站", "🚌 新烏日車站", "🚌 水湳轉運站", "🚌 經貿六停車場",
  "🚌 嶺東科大-寶文校區", "🚶 1號門", "🚶 3號門", "♿ 復康巴士"
];

// 一次性更新：只改原人員表單的站點選項，不重建題目、不動歷史回應或停車表單。
function updatePeopleStationChoices() {
  const files = DriveApp.getFilesByName("「國防知性之旅-成功嶺營區開放」人數回報");
  const candidates = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!file || typeof file.getId !== "function") continue;
    const form = FormApp.openById(file.getId());
    if (form.getPublishedUrl() === "https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform" &&
        form.getDestinationId() === TARGET_SPREADSHEET_ID) candidates.push(form);
  }
  if (candidates.length !== 1) throw new Error("找不到唯一符合既有網址與目的試算表的人員表單，未修改表單。");
  const form = candidates[0];
  const items = form.getItems().filter(item => /站點.*門號/.test(item.getTitle()));
  if (items.length !== 1 || items[0].getType() !== FormApp.ItemType.MULTIPLE_CHOICE) {
    throw new Error("站點題目缺少、重複或型別不符，未修改表單。");
  }
  items[0].asMultipleChoiceItem().setChoiceValues(PEOPLE_STATION_CHOICES);
  Logger.log("已更新原人員表單的八個站點選項，其他題目及歷史回應保留。");
}

// Append vehicle choices to the original people form without rebuilding questions.
function updatePeopleVehicleChoices() {
  const files = DriveApp.getFilesByName('「國防知性之旅-成功嶺營區開放」人數回報');
  const candidates = [];
  while (files.hasNext()) {
    const file = files.next();
    if (!file || typeof file.getId !== 'function') continue;
    const form = FormApp.openById(file.getId());
    if (form.getPublishedUrl() === 'https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform' &&
        form.getDestinationId() === TARGET_SPREADSHEET_ID) candidates.push(form);
  }
  if (candidates.length !== 1) throw new Error('找不到唯一原人員表單，未修改車號。');
  const items = candidates[0].getItems().filter(item => /車號/.test(item.getTitle()));
  if (items.length !== 1 || ![FormApp.ItemType.LIST, FormApp.ItemType.MULTIPLE_CHOICE].includes(items[0].getType())) {
    throw new Error('車號題目缺少、重複或型別不符，未修改表單。');
  }
  const item = items[0].getType() === FormApp.ItemType.LIST ? items[0].asListItem() : items[0].asMultipleChoiceItem();
  const before = item.getChoices().map(choice => choice.getValue());
  const after = before.slice();
  for (const [prefix, count] of [['復康巴士', 6], ['預備車號', 16]]) {
    for (let i = 1; i <= count; i++) {
      const value = prefix + i;
      if (!after.includes(value)) after.push(value);
    }
  }
  if (after.length !== before.length) item.setChoiceValues(after);
  return { before, after };
}

// 沿用每車、每方向最後一列累計值；車號取自實際回報，不預設新站車輛數。
function lingdongTotalFormula(sheetName, columns, direction) {
  const range = key => `'${sheetName}'!${columns[key]}2:${columns[key]}`;
  const conditions = `ISNUMBER(SEARCH("嶺東科大-寶文校區",${range("station")})),ISNUMBER(SEARCH("${direction}",${range("direction")})),${range("bus")}<>"",${range("bus")}<>"🚶 步行通道"`;
  const reports = `COUNTIFS(${range("station")},"*嶺東科大-寶文校區*",${range("direction")},"*${direction}*",${range("bus")},"<>",${range("bus")},"<>🚶 步行通道")`;
  return `=IF(${reports}=0,"未回報",LET(rawBuses,FILTER(${range("bus")},${conditions}),rawCounts,FILTER(${range("quantity")},${conditions}),valid,ARRAYFORMULA(IFERROR(ISNUMBER(rawCounts)*(rawCounts>=0)*(MOD(rawCounts,1)=0),FALSE)),IF(SUM(valid)=0,"未回報",LET(buses,FILTER(rawBuses,valid),counts,FILTER(rawCounts,valid),SUM(MAP(UNIQUE(buses),LAMBDA(vehicle,XLOOKUP(vehicle,buses,counts,0,0,-1))))))))`;
}

// 使用者提供的總容量；不代表目前剩餘車位。照片空白欄位代表未提供該車種。
const PARKING_LOTS = [
  { group: "4號門", name: "CITYPARKING春安站", motorcycle: 29, car: 81 },
  { group: "4號門", name: "協弘停車場", motorcycle: 0, car: 59 },
  { group: "4號門", name: "嶺東科大", motorcycle: 845, car: 0 },
  { group: "新烏日站", name: "五都日出", motorcycle: 168, car: 1008 },
  { group: "新烏日站", name: "嘟嘟房高鐵臺中站", motorcycle: 736, car: 1537 },
  { group: "新烏日站", name: "城市車旅高鐵五路三站", motorcycle: 0, car: 184 },
  { group: "新烏日站", name: "城市車旅烏日高鐵一站", motorcycle: 0, car: 390 },
  { group: "新烏日站", name: "TIMES烏日高鐵五路第2P", motorcycle: 0, car: 92 },
  { group: "新烏日站", name: "城市車旅-高鐵五路站", motorcycle: 0, car: 108 },
  { group: "新烏日站", name: "烏日臻愛站", motorcycle: 191, car: 187 },
  { group: "新烏日站", name: "城市車旅烏日高鐵二站P", motorcycle: 0, car: 518 },
  { group: "新烏日站", name: "五都高鐵烏日二站P", motorcycle: 0, car: 101 },
  { group: "新烏日站", name: "TIMES烏日高鐵五路P", motorcycle: 0, car: 63 },
  { group: "新烏日站", name: "五都高鐵烏日一站P", motorcycle: 0, car: 75 },
  { group: "新烏日站", name: "五都新幹線", motorcycle: 0, car: 325 },
  { group: "新烏日站", name: "城市車旅烏日高鐵三站P", motorcycle: 0, car: 82 },
  { group: "新烏日站", name: "大車河新烏日地下停車場", motorcycle: 352, car: 266 },
  { group: "水湳轉運站", name: "水湳轉運站停車場", motorcycle: 1253, car: 614 },
  { group: "水湳轉運站", name: "經貿6停車場", motorcycle: 318, car: 600 },
  { group: "水湳轉運站", name: "經貿8停車場", motorcycle: 703, car: 580 },
  { group: "水湳轉運站", name: "中央公園北側停車場", motorcycle: 0, car: 401 },
  { group: "水湳轉運站", name: "臺中國際會展停車場", motorcycle: 712, car: 699 },
  { group: "水湳轉運站", name: "綠美圖停車場", motorcycle: 386, car: 335 }
];
const PARKING_FORM_TITLE = "「國防知性之旅-成功嶺營區開放」停車場剩餘車位回報";

// 一次性更新：只改既有停車場表單的場地選項，不建立表單、不刪除歷史回應。
function updateParkingLotChoices() {
  const id = PropertiesService.getScriptProperties().getProperty("PARKING_FORM_ID");
  if (!id) throw new Error("找不到既有停車場表單 ID，未修改表單。");
  const form = FormApp.openById(id);
  if (form.getPublishedUrl() !== "https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform" ||
      form.getDestinationId() !== TARGET_SPREADSHEET_ID) {
    throw new Error("停車場表單網址或目的試算表不符，未修改表單。");
  }
  const items = form.getItems().filter(item => /停車場區域/.test(item.getTitle()));
  if (items.length !== 1 || items[0].getType() !== FormApp.ItemType.MULTIPLE_CHOICE) {
    throw new Error("停車場區域題目缺少、重複或型別不符，未修改表單。");
  }
  items[0].asMultipleChoiceItem().setChoiceValues(PARKING_LOTS.map(lot => lot.name));
  Logger.log("已更新既有停車場表單的 23 處場地選項，歷史回應保留。");
}

function findParkingFormFile() {
  const found = [];
  for (const title of [PARKING_FORM_TITLE, "「國防知性之旅-成功嶺營區開放」六大停車場車位回報"]) {
    const files = DriveApp.getFilesByName(title);
    while (files.hasNext()) found.push(files.next());
  }
  if (found.length > 1) throw new Error("找到多份停車場表單，請設定 PARKING_FORM_ID。");
  return found[0] || null;
}

// 一次性升級入口：沿用原表單與網址，舊回應欄位保留於試算表。
function upgradeParkingRemaining() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("PARKING_FORM_ID");
  let form = id ? FormApp.openById(id) : null;
  if (!form) {
    const file = findParkingFormFile();
    if (!file) throw new Error("找不到原停車場表單，未建立新表單。請設定 PARKING_FORM_ID。");
    form = FormApp.openById(file.getId());
  }
  if (form.getPublishedUrl() !== "https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform") {
    throw new Error("停車場表單網址不符，停止升級以免修改其他表單。");
  }
  if (form.getDestinationId() !== TARGET_SPREADSHEET_ID) throw new Error("原表單未連結指定試算表，請先確認目的地。");
  configureParkingRemainingForm(form);
  props.setProperty("PARKING_FORM_ID", form.getId());
  SpreadsheetApp.flush();
  // Google Forms 更新回應標題可能稍有延遲；只重試讀取，不重建表單。
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  for (let attempt = 0; attempt < 5; attempt++) {
    try { findResponseSource(ss, "parking"); break; }
    catch (error) {
      if (attempt === 4) throw new Error("表單已更新，回應標題尚未同步或來源不明。稍後執行 repairDashboard。" + error.message);
      Utilities.sleep(1000);
    }
  }
  repairDashboard();
}

function configureParkingRemainingForm(form) {
  const fields = [
    { title: "1. 停車場區域", match: /停車場區域/, type: FormApp.ItemType.MULTIPLE_CHOICE },
    { title: "2. 目前剩餘汽車停車位", match: /目前剩餘汽車停車位/, type: FormApp.ItemType.TEXT },
    { title: "3. 目前剩餘機車停車位", match: /目前剩餘機車停車位/, type: FormApp.ItemType.TEXT },
    { title: "4. 備註", match: /備註/, type: FormApp.ItemType.TEXT }
  ];
  const original = form.getItems();
  // 先檢查再修改，避免部分更新後才發現題目重複或型別不符。
  const existing = fields.map(field => {
    const matches = original.filter(item => field.match.test(item.getTitle()));
    if (matches.length > 1 || (matches[0] && matches[0].getType() !== field.type)) {
      throw new Error("停車場題目重複或型別不符：" + field.title);
    }
    return matches[0];
  });
  const validation = FormApp.createTextValidation()
    .requireTextMatchesPattern("^[0-9]+$")
    .setHelpText("請填目前剩餘格數（0 或正整數）；沒有剩餘車位請填 0。")
    .build();
  const keep = fields.map((field, index) => {
    const item = existing[index]
      ? (index === 0 ? existing[index].asMultipleChoiceItem() : existing[index].asTextItem())
      : (index === 0 ? form.addMultipleChoiceItem() : form.addTextItem());
    item.setTitle(field.title).setRequired(index !== 3);
    if (index === 0) item.setChoiceValues(PARKING_LOTS.map(lot => lot.name));
    if (index === 1 || index === 2) item.setValidation(validation);
    return item;
  });
  // 不把舊「數量」改名成剩餘車位，防止舊增減數字被誤讀。
  original.filter(item => !keep.some(k => k.getId() === item.getId()))
    .forEach(item => form.deleteItem(item));
  keep.forEach((item, index) => form.moveItem(item.getIndex(), index));
  form.setTitle(PARKING_FORM_TITLE);
  form.setDescription("請選擇停車場，直接填目前剩餘汽車與機車停車位，不填進場或離場。已停滿請填 0；*台中科大無汽車位、*經貿六與*經貿八無機車位，該欄填 0（看板以 — 顯示）。各場總容量（汽車／機車）：" + PARKING_LOTS.map(lot => `${lot.name} ${lot.car}／${lot.motorcycle}`).join("；") + "。總容量不代表即時空位。");
}

function parkingRemainingFormula(sheetName, columns, field, name, capacity) {
  if (capacity === 0) return '="不提供"';
  const area = `'${sheetName}'!${columns.area}2:${columns.area}`;
  const value = `'${sheetName}'!${columns[field]}2:${columns[field]}`;
  // 精確比對以保留名稱中的 *；不可當成萬用字元，也不可把舊 A～F 對應至新場地。
  const latest = `XLOOKUP("${name}", ${area}, ARRAYFORMULA(IF(${value}="", "未回報", IFERROR(VALUE(${value}), "回報異常"))), "未回報", 0, -1)`;
  return `=LET(latest,${latest},IF(ISNUMBER(latest),IF(AND(latest>=0,latest<=${capacity},latest=INT(latest)),latest,"回報異常"),latest))`;
}

function parkingTotalFormula(field) {
  const refs = PARKING_LOTS.map((lot, i) => lot[field] > 0 ? `${responseColumnLetter(i * 4 + (field === "car" ? 0 : 2))}34` : null).filter(Boolean);
  return `=IF(COUNT(${refs.join(",")})=${refs.length},SUM(${refs.join(",")}),"未完整回報")`;
}

function parkingReportedCondition(lot, index) {
  const conditions = [];
  if (lot.car > 0) conditions.push(`ISNUMBER(${responseColumnLetter(index * 4)}34)`);
  if (lot.motorcycle > 0) conditions.push(`ISNUMBER(${responseColumnLetter(index * 4 + 2)}34)`);
  return `AND(${conditions.join(",")})`;
}

// 提供停車場獨立網址讀取的直式資料；資料由目前回應分頁中同名場地的最新回報取得。
function rebuildParkingStatusSheet(ss, sheetName, columns) {
  let statusSheet = ss.getSheetByName("停車場現況");
  if (!statusSheet) statusSheet = ss.insertSheet("停車場現況");
  const requiredRows = PARKING_LOTS.length + 1;
  if (statusSheet.getMaxRows() < requiredRows) statusSheet.insertRowsAfter(statusSheet.getMaxRows(), requiredRows - statusSheet.getMaxRows());
  if (statusSheet.getMaxColumns() < 6) statusSheet.insertColumnsAfter(statusSheet.getMaxColumns(), 6 - statusSheet.getMaxColumns());
  statusSheet.getRange(1, 1, statusSheet.getMaxRows(), 6).clear();
  statusSheet.getRange(1, 1, 1, 6).setValues([["區域", "停車場", "剩餘汽車位", "剩餘機車位", "汽車容量", "機車容量"]])
    .setBackground("#1E293B").setFontColor("#FFFFFF").setFontWeight("bold");
  PARKING_LOTS.forEach((lot, index) => {
    const row = index + 2;
    statusSheet.getRange(row, 1, 1, 2).setValues([[lot.group, lot.name]]);
    statusSheet.getRange(row, 3).setFormula(parkingRemainingFormula(sheetName, columns, "car", lot.name, lot.car));
    statusSheet.getRange(row, 4).setFormula(parkingRemainingFormula(sheetName, columns, "motorcycle", lot.name, lot.motorcycle));
    statusSheet.getRange(row, 5, 1, 2).setValues([[lot.car || "", lot.motorcycle || ""]]);
  });
  if (statusSheet.setFrozenRows) statusSheet.setFrozenRows(1);
  if (statusSheet.autoResizeColumns) statusSheet.autoResizeColumns(1, 6);
}

function responseColumnLetter(index) {
  let result = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    result = String.fromCharCode(65 + (n - 1) % 26) + result;
  }
  return result;
}

function findResponseSource(ss, kind) {
  const patterns = kind === "people"
    ? { time: /^(時間戳記|時間標記|Timestamp)$/i, direction: /方向/, station: /站點.*門號/, bus: /車號/, quantity: /^(?:4[.、．]\s*)?人數$/ }
    : { area: /停車場區域/, car: /目前剩餘汽車停車位/, motorcycle: /目前剩餘機車停車位/ };
  const candidates = [];
  ss.getSheets().forEach(sheet => {
    if (!sheet.getLastColumn()) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columns = {};
    for (const key of Object.keys(patterns)) {
      const matches = headers.map((h, i) => patterns[key].test(String(h).trim()) ? i : -1).filter(i => i >= 0);
      // 同名題目欄位不可任意選取，以免讀到已停用的舊欄位。
      if (matches.length !== 1) return;
      columns[key] = responseColumnLetter(matches[0]);
    }
    candidates.push({ sheet, columns });
  });
  const explicitName = RESPONSE_SHEET_NAMES[kind];
  let selected = explicitName ? candidates.filter(c => c.sheet.getName() === explicitName) : candidates;
  if (!explicitName && selected.length > 1) {
    const linked = selected.filter(c => c.sheet.getFormUrl());
    if (linked.length === 1) selected = linked;
  }
  if (selected.length !== 1) {
    throw new Error(`${kind} 回應分頁無法唯一辨識。候選：${candidates.map(c => c.sheet.getName()).join("、") || "無（請检查題目標題是否缺少或重複）"}。請設定 RESPONSE_SHEET_NAMES。未改動回應資料。`);
  }
  Logger.log(`${kind} 資料來源：${selected[0].sheet.getName()}，欄位：${JSON.stringify(selected[0].columns)}`);
  return selected[0];
}

function createMultiStationBusSystem(dashboardOnly) {
  Logger.log("🎨 開始對指定試算表 [" + TARGET_SPREADSHEET_ID + "] 進行雙表單建置與看板重繪...");

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  
  // =========================================================================
  // 1. 維護【表單 1：人員疏運回報】(依專屬名稱鎖定獨立表單，絕不與表單2混淆)
  // =========================================================================
  let form1 = null;
  let parkingForm = null;
  if (dashboardOnly !== true) {
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
        .setChoiceValues(PEOPLE_STATION_CHOICES)
        .setRequired(true);

      // 題目 3: 車號
      const busChoices = ["🚶 步行通道"];
      for (let i = 1; i <= 50; i++) {
        busChoices.push(`${i} 號車`);
      }
      for (let i = 1; i <= 6; i++) busChoices.push(`復康巴士${i}`);
      for (let i = 1; i <= 16; i++) busChoices.push(`預備車號${i}`);
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
  // 2. 建立或維護【表單 2：六區剩餘汽車、機車位回報】
  // =========================================================================
  const props = PropertiesService.getScriptProperties();
  let parkingFormId = props.getProperty("PARKING_FORM_ID");
  if (parkingFormId) {
    try {
      parkingForm = FormApp.openById(parkingFormId);
    } catch (e) {
      parkingForm = null;
    }
  }

  if (!parkingForm) {
    // 檢查雲端是否已有同名表單
    const file = findParkingFormFile();
    if (file) {
      parkingForm = FormApp.openById(file.getId());
      props.setProperty("PARKING_FORM_ID", file.getId());
    }
  }

  if (!parkingForm) {
    // 建立新停車場表單並綁定至試算表
    parkingForm = FormApp.create(PARKING_FORM_TITLE);
    parkingForm.setDescription("現場交通哨專用：請回報各區目前剩餘汽車、機車停車位。");
    parkingForm.setDestination(FormApp.DestinationType.SPREADSHEET, TARGET_SPREADSHEET_ID);
    props.setProperty("PARKING_FORM_ID", parkingForm.getId());
  }

  if (parkingForm) {
    configureParkingRemainingForm(parkingForm);
  }

  }

  // 在任何看板寫入之前檢查來源；不依分頁順序、語言或資料列數猜測。
  const people = findResponseSource(ss, "people");
  const parking = findResponseSource(ss, "parking");
  const formSheetName = people.sheet.getName().replace(/'/g, "''");
  const parkingSheetName = parking.sheet.getName().replace(/'/g, "''");
  const personCols = people.columns;
  let dashboardSheet = ss.getSheetByName("總即時戰情看板");
  if (!dashboardSheet) dashboardSheet = ss.insertSheet("總即時戰情看板");
  let detailSheet = ss.getSheetByName("各車即時明細");
  if (!detailSheet) detailSheet = ss.insertSheet("各車即時明細");
  if (dashboardSheet.getMaxRows() < 50) dashboardSheet.insertRowsAfter(dashboardSheet.getMaxRows(), 50 - dashboardSheet.getMaxRows());
  if (dashboardSheet.getMaxColumns() < 30) dashboardSheet.insertColumnsAfter(dashboardSheet.getMaxColumns(), 30 - dashboardSheet.getMaxColumns());
  if (detailSheet.getMaxRows() < 75) detailSheet.insertRowsAfter(detailSheet.getMaxRows(), 75 - detailSheet.getMaxRows());
  if (detailSheet.getMaxColumns() < 40) detailSheet.insertColumnsAfter(detailSheet.getMaxColumns(), 40 - detailSheet.getMaxColumns());

  // =========================================================================
  // 【A. 第二頁：各車即時明細 (4 大車站車輛明細)】
  // =========================================================================
  const additionalVehicles = [...Array.from({ length: 6 }, (_, i) => `復康巴士${i + 1}`),
    ...Array.from({ length: 16 }, (_, i) => `預備車號${i + 1}`)];
  const stations = [
    { name: "🚌 成功車站", formKeyword: "成功車站", startCol: 1, busCount: 50 },
    { name: "🚌 新烏日車站", formKeyword: "新烏日", startCol: 6, busCount: 50 },
    { name: "🚌 經貿六停車場", formKeyword: "經貿六", startCol: 11, busCount: 50 },
    { name: "🚌 水湳轉運站", formKeyword: "水湳", startCol: 16, busCount: 50 },
    { name: "🚌 嶺東科大-寶文校區", formKeyword: "嶺東科大-寶文校區", startCol: 21, busCount: 50 },
    { name: "♿ 復康巴士", formKeyword: "復康巴士", startCol: 26, busCount: 50, additive: true },
    { name: "🚶 1號門", formKeyword: "1號門", startCol: 31, busCount: 0, additive: true },
    { name: "🚶 3號門", formKeyword: "3號門", startCol: 36, busCount: 0, additive: true }
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
    const vehicleNames = st.busCount === 0 ? ["全部回報"] : [
      ...(st.additive ? ["🚶 步行通道"] : []),
      ...Array.from({ length: st.busCount }, (_, i) => `${i + 1} 號車`), ...additionalVehicles];
    for (const busName of vehicleNames) {
      const fInActual = `=IFERROR(INDEX('${formSheetName}'!${personCols.quantity}:${personCols.quantity}, MAX(FILTER(ROW('${formSheetName}'!${personCols.quantity}:${personCols.quantity}), ISNUMBER(SEARCH("進場", '${formSheetName}'!${personCols.direction}:${personCols.direction})), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!${personCols.station}:${personCols.station})), '${formSheetName}'!${personCols.bus}:${personCols.bus}="${busName}"))), 0)`;
      const fInTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!${personCols.time}:${personCols.time}, MAX(FILTER(ROW('${formSheetName}'!${personCols.time}:${personCols.time}), ISNUMBER(SEARCH("進場", '${formSheetName}'!${personCols.direction}:${personCols.direction})), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!${personCols.station}:${personCols.station})), '${formSheetName}'!${personCols.bus}:${personCols.bus}="${busName}"))), "hh:mm:ss"), "-")`;
      const fOutActual = `=IFERROR(INDEX('${formSheetName}'!${personCols.quantity}:${personCols.quantity}, MAX(FILTER(ROW('${formSheetName}'!${personCols.quantity}:${personCols.quantity}), ISNUMBER(SEARCH("離場", '${formSheetName}'!${personCols.direction}:${personCols.direction})), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!${personCols.station}:${personCols.station})), '${formSheetName}'!${personCols.bus}:${personCols.bus}="${busName}"))), 0)`;
      const fOutTime = `=IFERROR(TEXT(INDEX('${formSheetName}'!${personCols.time}:${personCols.time}, MAX(FILTER(ROW('${formSheetName}'!${personCols.time}:${personCols.time}), ISNUMBER(SEARCH("離場", '${formSheetName}'!${personCols.direction}:${personCols.direction})), ISNUMBER(SEARCH("${st.formKeyword}", '${formSheetName}'!${personCols.station}:${personCols.station})), '${formSheetName}'!${personCols.bus}:${personCols.bus}="${busName}"))), "hh:mm:ss"), "-")`;
      const quantityRange = `'${formSheetName}'!${personCols.quantity}:${personCols.quantity}`;
      const valid = `,IFERROR(ISNUMBER(${quantityRange})*(${quantityRange}>=0)*(MOD(${quantityRange},1)=0),FALSE)`;
      const latestValid = formula => formula.replace(`, '${formSheetName}'!${personCols.bus}:${personCols.bus}="${busName}"`,
        `, '${formSheetName}'!${personCols.bus}:${personCols.bus}="${busName}"${valid}`);
      const additive = direction => `=SUMIFS(${quantityRange},'${formSheetName}'!${personCols.direction}:${personCols.direction},"*${direction}*",'${formSheetName}'!${personCols.station}:${personCols.station},"*${st.formKeyword}*"${st.busCount ? `,'${formSheetName}'!${personCols.bus}:${personCols.bus},"${busName}"` : ''})`;
      const gateTime = direction => `=IFERROR(TEXT(MAX(FILTER('${formSheetName}'!${personCols.time}:${personCols.time},ISNUMBER(SEARCH("${direction}",'${formSheetName}'!${personCols.direction}:${personCols.direction})),ISNUMBER(SEARCH("${st.formKeyword}",'${formSheetName}'!${personCols.station}:${personCols.station})))),"hh:mm:ss"),"-")`;
      busRows.push([busName, st.additive ? additive("進場") : latestValid(fInActual),
        st.busCount === 0 ? gateTime("進場") : st.additive ? fInTime : latestValid(fInTime),
        st.additive ? additive("離場") : latestValid(fOutActual),
        st.busCount === 0 ? gateTime("離場") : st.additive ? fOutTime : latestValid(fOutTime)]);
    }
    detailSheet.getRange(3, col, vehicleNames.length, 5).setValues(busRows);
    detailSheet.getRange(2, col, vehicleNames.length + 1, 5).setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);
  });

  // =========================================================================
  // 【B. 第一頁：總即時戰情看板 (人員疏運 + 六大停車場全合一)】
  // =========================================================================
  try {
    const maxRows = dashboardSheet.getMaxRows();
    const maxCols = dashboardSheet.getMaxColumns();
    const cleanRange = dashboardSheet.getRange(1, 1, maxRows, maxCols);
    cleanRange.breakApart();
    cleanRange.clearDataValidations();
    dashboardSheet.clearConditionalFormatRules();
    cleanRange.clear();
  } catch (e) { throw new Error("看板重設失敗：" + e.message); }

  dashboardSheet.setHiddenGridlines(true);

  for (let c = 1; c <= 28; c++) {
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
  dashboardSheet.getRange(4, 1, 1, 8).merge().setFormula('=IF(ISNUMBER(A40),A10+G10+M10+S10+A40,"未完整回報")').setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(3, 9, 1, 8).merge().setValue("👈 離場總人數").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 9, 1, 8).merge().setFormula('=IF(ISNUMBER(B40),D10+J10+P10+V10+B40,"未完整回報")').setBackground("#1E293B").setFontColor("#FFFFFF").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(3, 17, 1, 8).merge().setValue("📈 離場完成率").setBackground("#1E293B").setFontColor("#94A3B8").setFontSize(11).setHorizontalAlignment("center");
  dashboardSheet.getRange(4, 17, 1, 8).merge().setFormula(`=IF(AND(ISNUMBER(A4),ISNUMBER(I4)),IF(A4>0,TEXT(I4/A4,"0.0%"),"0.0%"),"未完整回報")`).setBackground("#1E293B").setFontColor("#34D399").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("center");

  dashboardSheet.getRange(5, 1, 1, 6).merge().setValue("離場進度").setBackground(THEME_HEADER_BG).setFontColor("#94A3B8").setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(5, 7, 1, 18).merge().setFormula(`=IF(AND(ISNUMBER(A4),ISNUMBER(I4)),IF(A4>0,SPARKLINE(I4,{"charttype","bar";"max",A4;"color1","${THEME_TOP_BAR}"}),""),"未完整回報")`).setBackground(THEME_HEADER_BG);
  dashboardSheet.getRange(2, 1, 4, 24).setBorder(true, true, true, true, true, true, "#334155", SpreadsheetApp.BorderStyle.SOLID);

  // 4 大接駁車站 (第 7~12 列)
  dashboardSheet.getRange(7, 1, 1, 24).merge()
    .setValue("接駁車站即時數據")
    .setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const busStations = [
    { name: "🚌 成功車站", startCol: 1, detailInCol: "B", detailOutCol: "D", detailEndRow: 74, tagColor: "#059669" },
    { name: "🚌 新烏日車站", startCol: 7, detailInCol: "G", detailOutCol: "I", detailEndRow: 74, tagColor: "#2563EB" },
    { name: "🚌 經貿六停車場", startCol: 13, detailInCol: "L", detailOutCol: "N", detailEndRow: 74, tagColor: "#D97706" },
    { name: "🚌 水湳轉運站", startCol: 19, detailInCol: "Q", detailOutCol: "S", detailEndRow: 74, tagColor: "#D97706" }
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

  // 2 大步行門 (第 19~24 列；原 4 號門歷史回應保留於來源分頁)
  dashboardSheet.getRange(19, 1, 1, 24).merge().setValue("步行通道即時數據").setBackground("#334155").setFontColor("#F8FAFC").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("left");

  const walkGates = [
    { name: "🚶 1號門", keyword: "1號門", startCol: 1, tagColor: "#059669" },
    { name: "🚶 3號門", keyword: "3號門", startCol: 9, tagColor: "#2563EB" },
    { name: "♿ 復康巴士", keyword: "復康巴士", startCol: 17, tagColor: "#7C3AED" }
  ];

  walkGates.forEach(wg => {
    const col = wg.startCol;
    dashboardSheet.getRange(20, col, 1, 8).merge().setValue(wg.name).setBackground(wg.tagColor).setFontColor("#FFFFFF").setFontWeight("bold").setFontSize(12).setHorizontalAlignment("center");
    dashboardSheet.getRange(21, col, 1, 4).merge().setValue("進場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(21, col + 4, 1, 4).merge().setValue("離場").setBackground("#F8FAFC").setFontColor("#64748B").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");

    const colLetterIn = String.fromCharCode(64 + col);
    const colLetterOut = String.fromCharCode(64 + col + 4);

    dashboardSheet.getRange(22, col, 1, 4).merge().setFormula(`=SUMIFS('${formSheetName}'!${personCols.quantity}:${personCols.quantity}, '${formSheetName}'!${personCols.direction}:${personCols.direction}, "*進場*", '${formSheetName}'!${personCols.station}:${personCols.station}, "*${wg.keyword}*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(22, col + 4, 1, 4).merge().setFormula(`=SUMIFS('${formSheetName}'!${personCols.quantity}:${personCols.quantity}, '${formSheetName}'!${personCols.direction}:${personCols.direction}, "*離場*", '${formSheetName}'!${personCols.station}:${personCols.station}, "*${wg.keyword}*")`).setBackground("#FFFFFF").setFontColor("#0F172A").setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center");
    dashboardSheet.getRange(23, col, 1, 8).merge().setFormula(`=IF(${colLetterIn}22>0, TEXT(${colLetterOut}22/${colLetterIn}22, "0.0%") & "  ·  尚餘 " & MAX(0, ${colLetterIn}22 - ${colLetterOut}22) & " 人", "0.0%  ·  已完成")`).setBackground("#F8FAFC").setFontColor("#0F172A").setFontWeight("bold").setFontSize(11).setHorizontalAlignment("center");
    dashboardSheet.getRange(24, col, 1, 8).merge().setFormula(`=IF(${colLetterIn}22>0, SPARKLINE(${colLetterOut}22, {"charttype", "bar"; "max", ${colLetterIn}22; "color1", "${THEME_CARD_BAR}"}), "")`).setBackground("#F1F5F9");
    dashboardSheet.getRange(20, col, 5, 8).setBorder(true, true, true, true, true, true, "#CBD5E1", SpreadsheetApp.BorderStyle.SOLID);
  });

  // 停車場改由獨立的「停車場現況」分頁供停車場網址讀取；不再受舊七場橫向欄位限制。
  rebuildParkingStatusSheet(ss, parkingSheetName, parking.columns);

  // 在既有區塊後追加；前端依 A 欄標題辨識，下一列 A/B 為進場/離場。
  dashboardSheet.getRange(38, 1, 1, 2).setValues([["進場", "離場"]]);
  dashboardSheet.getRange(39, 1).setValue("嶺東科大-寶文校區接駁統計");
  dashboardSheet.getRange(40, 1).setFormula(lingdongTotalFormula(formSheetName, personCols, "進場"));
  dashboardSheet.getRange(40, 2).setFormula(lingdongTotalFormula(formSheetName, personCols, "離場"));

  dashboardSheet.getRange(8, 25, 1, 6).merge().setValue("🚌 嶺東科大-寶文校區").setBackground("#DB2777").setFontColor("#FFFFFF").setFontWeight("bold");
  dashboardSheet.getRange(9, 25, 1, 3).merge().setValue("進場");
  dashboardSheet.getRange(9, 28, 1, 3).merge().setValue("離場");
  dashboardSheet.getRange(10, 25, 1, 3).merge().setFormula("=A40").setFontSize(24).setFontWeight("bold");
  dashboardSheet.getRange(10, 28, 1, 3).merge().setFormula("=B40").setFontSize(24).setFontWeight("bold");
  dashboardSheet.getRange(11, 25, 1, 6).merge().setFormula('=IF(AND(ISNUMBER(A40),ISNUMBER(B40)),IF(A40>0,TEXT(B40/A40,"0.0%"),"0.0%"),"未完整回報")');
  dashboardSheet.getRange(12, 25, 1, 6).merge().setFormula(`=IF(AND(ISNUMBER(A40),ISNUMBER(B40)),IF(A40>0,SPARKLINE(B40,{"charttype","bar";"max",A40;"color1","#DB2777"}),""),"未完整回報")`);
  dashboardSheet.getRange(8, 25, 5, 6).setHorizontalAlignment("center").setBorder(true,true,true,true,true,true,"#CBD5E1",SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.flush();
  const valA = dashboardSheet.getRange("A34").getValue();
  const valC = dashboardSheet.getRange("I34").getValue();
  const valTotal = dashboardSheet.getRange("A28").getValue();
  Logger.log(`📊【現場驗證結果】五都日出剩餘汽車位: ${valA} 輛 | 嶺東科大剩餘汽車位: ${valC} 輛 | 全區剩餘汽車位: ${valTotal} 輛`);

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

// Read-only deployment verification; never returns individual response records.
function inspectPeopleDashboardSetup() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const people = findResponseSource(ss, 'people');
  const parking = findResponseSource(ss, 'parking');
  const board = ss.getSheetByName('總即時戰情看板');
  const files = DriveApp.getFilesByName('「國防知性之旅-成功嶺營區開放」人數回報');
  const forms = [];
  while (files.hasNext()) {
    const form = FormApp.openById(files.next().getId());
    if (form.getDestinationId() !== TARGET_SPREADSHEET_ID) continue;
    forms.push({ url: form.getPublishedUrl(), vehicles: form.getItems()
      .filter(item => /車號/.test(item.getTitle()))
      .map(item => (item.getType() === FormApp.ItemType.LIST ? item.asListItem() : item.asMultipleChoiceItem()).getChoices().map(choice => choice.getValue())), stations: form.getItems()
      .filter(item => /站點.*門號/.test(item.getTitle()))
      .map(item => item.asMultipleChoiceItem().getChoices().map(choice => choice.getValue())) });
  }
  return { forms, responseRows: { people: people.sheet.getLastRow(), parking: parking.sheet.getLastRow() },
    board: board.getRange(1, 1, 40, 30).getDisplayValues(),
    detailGid: ss.getSheetByName("各車即時明細").getSheetId(),
    boardGid: board.getSheetId(),
    detailHeaders: ss.getSheetByName("各車即時明細").getRange(1, 1, 1, 40).getDisplayValues()[0],
    detailValues: ss.getSheetByName("各車即時明細").getRange(3, 1, 73, 40).getDisplayValues(),
    walkingFormulas: [board.getRange(16, 1).getFormula(), board.getRange(16, 9).getFormula()],
    accessibleFormulas: [board.getRange(22, 17).getFormula(), board.getRange(22, 21).getFormula()] };
}
