/**
 * 【一鍵全自動建立 50 臺遊覽車上車進度系統】
 * 
 * 使用方式：
 * 1. 打開瀏覽器前往 https://script.new (會自動開啟 Google Apps Script)
 * 2. 把這整份程式碼複製貼上，覆蓋原有內容
 * 3. 點擊上方的「執行 (Run)」按鈕（第一次執行會跳出權限確認，點允許即可）
 * 4. 執行完畢後，下方的「執行紀錄」會直接印出：
 *    - 📱 車長填寫表單網址 (Google Form URL)
 *    - 📊 主控台即時看板網址 (Google Sheet URL)
 */

function createBusTrackingSystem() {
  Logger.log("🚀 開始自動建立系統中，請稍候約 5~10 秒...");

  // 1. 建立 Google 試算表
  const ss = SpreadsheetApp.create("🚌 50臺遊覽車上車進度戰情看板");
  const ssId = ss.getId();
  const ssUrl = ss.getUrl();

  // 2. 建立 Google 表單
  const form = FormApp.create("🚌 遊覽車上車進度回報系統");
  form.setDescription("請各車車長於上車/出發時即時回報最新人數與狀態。");

  // 加入車號（下拉選單 01車 ~ 50車）
  const busChoices = [];
  for (let i = 1; i <= 50; i++) {
    const num = String(i).padStart(2, '0');
    busChoices.push(`${num}車`);
  }
  form.addListItem()
    .setTitle("1. 選擇車號")
    .setChoiceValues(busChoices)
    .setRequired(true);

  // 加入車長姓名
  form.addTextItem()
    .setTitle("2. 車長姓名")
    .setRequired(false);

  // 加入實到人數
  const textValidation = FormApp.createTextValidation()
    .setHelpText("請輸入大於等於 0 的數字")
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem()
    .setTitle("3. 目前實到人數")
    .setValidation(textValidation)
    .setRequired(true);

  // 加入狀態
  form.addMultipleChoiceItem()
    .setTitle("4. 目前車輛狀態")
    .setChoiceValues(["🟡 上車中", "🟢 全員到齊", "🚀 已出發"])
    .setRequired(true);

  // 加入備註
  form.addParagraphTextItem()
    .setTitle("5. 備註 / 缺席名單說明")
    .setRequired(false);

  // 3. 將表單綁定到該試算表
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ssId);
  Utilities.sleep(1500); // 等待 Google 同步建立表單回應分頁

  // 4. 設定試算表看板
  const sheets = ss.getSheets();
  let dashboardSheet = sheets[0];
  dashboardSheet.setName("即時戰情看板");

  // 設置頂部 KPI 樣式與公式
  dashboardSheet.getRange("A1:E1").setValues([["總車數", "已出發", "全員到齊", "上車中", "尚未回報"]]);
  dashboardSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");

  dashboardSheet.getRange("A2").setValue(50);
  dashboardSheet.getRange("B2").setFormula('=COUNTIF(D5:D54, "🚀 已出發")');
  dashboardSheet.getRange("C2").setFormula('=COUNTIF(D5:D54, "🟢 全員到齊")');
  dashboardSheet.getRange("D2").setFormula('=COUNTIF(D5:D54, "🟡 上車中")');
  dashboardSheet.getRange("E2").setFormula('=COUNTIF(D5:D54, "⚪ 尚未回報")');
  dashboardSheet.getRange("A2:E2").setFontSize(14).setFontWeight("bold").setHorizontalAlignment("center");

  // 設置車輛明細標題
  const headers = [["車號", "應到人數", "最新實到人數", "最新狀態", "最後回報時間", "備註"]];
  dashboardSheet.getRange("A4:F4").setValues(headers);
  dashboardSheet.getRange("A4:F4").setFontWeight("bold").setBackground("#334155").setFontColor("#ffffff");

  // 填入 50 台車資料與動態抓取公式
  const rows = [];
  for (let i = 1; i <= 50; i++) {
    const num = String(i).padStart(2, '0');
    const busName = `${num}車`;
    const rowIdx = 4 + i;
    
    // 公式說明：由下往上抓最新一筆回報
    const formulaActual = `=IFERROR(INDEX('表單回應 1'!D:D, MAX(FILTER(ROW('表單回應 1'!B:B), '表單回應 1'!B:B=A${rowIdx}))), 0)`;
    const formulaStatus = `=IFERROR(INDEX('表單回應 1'!E:E, MAX(FILTER(ROW('表單回應 1'!B:B), '表單回應 1'!B:B=A${rowIdx}))), "⚪ 尚未回報")`;
    const formulaTime   = `=IFERROR(INDEX('表單回應 1'!A:A, MAX(FILTER(ROW('表單回應 1'!B:B), '表單回應 1'!B:B=A${rowIdx}))), "-")`;
    const formulaNote   = `=IFERROR(INDEX('表單回應 1'!F:F, MAX(FILTER(ROW('表單回應 1'!B:B), '表單回應 1'!B:B=A${rowIdx}))), "-")`;

    rows.push([busName, 40, formulaActual, formulaStatus, formulaTime, formulaNote]);
  }
  dashboardSheet.getRange("A5:F54").setValues(rows);

  // 自動調整欄寬
  dashboardSheet.autoResizeColumns(1, 6);

  const formPublishedUrl = form.getPublishedUrl();

  Logger.log("\n=======================================================");
  Logger.log("🎉 建立完成！請保存以下網址：");
  Logger.log("📱 車長手機填寫網址 (Google 表單): \n" + formPublishedUrl);
  Logger.log("\n📊 主控台即時戰情看板 (Google 試算表): \n" + ssUrl);
  Logger.log("=======================================================\n");
}
