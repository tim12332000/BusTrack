// 衛勤支援獨立資料：不寫入總即時戰情看板，也不使用人員回應分頁。
const MEDICAL_FORM_TITLE = '「國防知性之旅-成功嶺營區開放」衛勤支援回報';
const MEDICAL_STATS_SHEET = '衛勤支援統計';

function medicalLatestFormula(sheetName, column) {
  const source = "'" + sheetName.replace(/'/g, "''") + "'!";
  // Google Forms inserts rows and shifts A2 references; whole columns stay anchored.
  return '=IFERROR(LET(data,FILTER(' + source + 'A:C,ROW(' + source + 'A:A)>1,' + source + 'B:B<>"",' + source + 'C:C<>""),valid,FILTER(data,ISNUMBER(INDEX(data,,2)),ISNUMBER(INDEX(data,,3)),INDEX(data,,2)>=0,INDEX(data,,3)>=0,MOD(INDEX(data,,2),1)=0,MOD(INDEX(data,,3),1)=0),INDEX(valid,ROWS(valid),' + column + ')),"未回報")';
}

function setupMedicalSupport() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    let id = props.getProperty('MEDICAL_FORM_ID');
    let form = id ? FormApp.openById(id) : null;
    if (!form) {
      const files = DriveApp.getFilesByName(MEDICAL_FORM_TITLE);
      const matches = [];
      while (files.hasNext()) matches.push(FormApp.openById(files.next().getId()));
      if (matches.length > 1) throw new Error('找到多份衛勤表單，請確認 MEDICAL_FORM_ID。');
      form = matches[0] || FormApp.create(MEDICAL_FORM_TITLE);
      props.setProperty('MEDICAL_FORM_ID', form.getId());
    }
    let destination = null;
    try { destination = form.getDestinationId(); }
    catch (error) {
      if (!String(error.message).includes('no response destination')) throw error;
    }
    if (destination && destination !== TARGET_SPREADSHEET_ID) throw new Error('衛勤表單目的試算表不符。');
    const titles = ['醫療協處累計人數', '後送累計人數'];
    const items = form.getItems();
    if (items.length && (items.length !== 2 || items.some((item, i) => item.getTitle() !== titles[i] || item.getType() !== FormApp.ItemType.TEXT))) {
      throw new Error('衛勤表單題目不符，保留現有表單與回應。');
    }
    if (!items.length) {
      const validation = FormApp.createTextValidation().requireTextMatchesPattern('^[0-9]{1,9}$').setHelpText('請填 0 或正整數（最多九位數）。').build();
      titles.forEach(title => form.addTextItem().setTitle(title).setRequired(true).setValidation(validation));
      form.setDescription('請填目前累計人數，沒有請填 0。以最新一次完整回報更新衛勤卡片，不累加每次回報，也不納入進離場人數總計。請勿填入病患姓名或個人資料。');
      form.setConfirmationMessage('衛勤支援回報已收到，看板將顯示最新累計人數。');
    }
    if (!destination) form.setDestination(FormApp.DestinationType.SPREADSHEET, TARGET_SPREADSHEET_ID);
    if (form.supportsAdvancedResponderPermissions()) form.setPublished(true);
    form.setAcceptingResponses(true);
    SpreadsheetApp.flush();
    const sources = ss.getSheets().filter(sheet => {
      if (sheet.getLastColumn() < 3) return false;
      const header = sheet.getRange(1, 1, 1, 3).getValues()[0];
      return header[1] === titles[0] && header[2] === titles[1];
    });
    if (sources.length !== 1) throw new Error('無法唯一辨識衛勤回應分頁，未寫入統計；可重試 setupMedicalSupport。');
    const stats = ss.getSheetByName(MEDICAL_STATS_SHEET) || ss.insertSheet(MEDICAL_STATS_SHEET);
    stats.getRange('A1:D1').setValues([['醫療協處', '後送', '最新回報時間', '填報表單']]);
    stats.getRange('A2:C2').setFormulas([[medicalLatestFormula(sources[0].getName(), 2), medicalLatestFormula(sources[0].getName(), 3), medicalLatestFormula(sources[0].getName(), 1)]]);
    stats.getRange('C2').setNumberFormat('yyyy/MM/dd HH:mm:ss');
    stats.getRange('D2').setValue(form.getPublishedUrl());
    SpreadsheetApp.flush();
    return inspectMedicalSupport();
  } finally { lock.releaseLock(); }
}

function inspectMedicalSupport() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const id = PropertiesService.getScriptProperties().getProperty('MEDICAL_FORM_ID');
  const form = id ? FormApp.openById(id) : null;
  const stats = ss.getSheetByName(MEDICAL_STATS_SHEET);
  return {
    formUrl: form ? form.getPublishedUrl() : null,
    titles: form ? form.getItems().map(item => item.getTitle()) : [],
    accepting: form ? form.isAcceptingResponses() : false,
    stats: stats ? stats.getRange('A1:D2').getDisplayValues() : null,
    formulas: stats ? stats.getRange('A2:C2').getFormulas()[0] : [],
    responseRows: ss.getSheets().filter(sheet => /回應|Responses/.test(sheet.getName())).map(sheet => ({ name: sheet.getName(), rows: sheet.getLastRow() })),
    peopleBoard: ss.getSheetByName('總即時戰情看板').getRange('A1:AD40').getDisplayValues()
  };
}
