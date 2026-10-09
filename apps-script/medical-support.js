// 衛勤支援獨立資料：不寫入總即時戰情看板，也不使用人員回應分頁。
const MEDICAL_FORM_TITLE = '「國防知性之旅-成功嶺營區開放」衛勤支援回報';
const MEDICAL_STATS_SHEET = '衛勤支援統計';

function medicalLatestFormula(sheetName, column) {
  const source = "'" + sheetName.replace(/'/g, "''") + "'!";
  // Google Forms inserts rows and shifts A2 references; whole columns stay anchored.
  return '=IFERROR(LET(data,FILTER(' + source + 'A:C,ROW(' + source + 'A:A)>1,' + source + 'B:B<>"",' + source + 'C:C<>""),valid,FILTER(data,ISNUMBER(INDEX(data,,2)),ISNUMBER(INDEX(data,,3)),INDEX(data,,2)>=0,INDEX(data,,3)>=0,MOD(INDEX(data,,2),1)=0,MOD(INDEX(data,,3),1)=0),INDEX(valid,ROWS(valid),' + column + ')),"未回報")';
}

// Each number/category keeps its latest valid cumulative report, then sums 1–15.
function medicalNumberedTotalFormula(sheetName, columns, category, legacyNumber = 0, legacyValue = null) {
  const source = "'" + sheetName.replace(/'/g, "''") + "'!";
  const range = key => source + columns[key] + ':' + columns[key];
  const categoryText = category.replace(/"/g, '""');
  const fallback = Number.isSafeInteger(legacyValue) && legacyValue >= 0 ? legacyValue : null;
  const seed = fallback === null ? '0' : `IF(n=${legacyNumber},${fallback},0)`;
  const empty = fallback === null ? '"未回報"' : String(fallback);
  return `=LET(valid,ARRAYFORMULA(IFERROR((${range('category')}="${categoryText}")*REGEXMATCH(${range('number')},"^([1-9]|1[0-5])號$")*ISNUMBER(${range('quantity')})*(${range('quantity')}>=0)*(MOD(${range('quantity')},1)=0),FALSE)),IF(SUM(valid)=0,${empty},LET(ids,FILTER(${range('number')},valid),counts,FILTER(${range('quantity')},valid),SUM(MAP(SEQUENCE(15),LAMBDA(n,XLOOKUP(n&"號",ids,counts,${seed},0,-1)))))))`;
}

function findMedicalResponseSheet(ss, form) {
  const saved = PropertiesService.getScriptProperties().getProperty('MEDICAL_RESPONSE_SHEET_ID');
  if (saved) {
    const sheet = ss.getSheets().find(item => String(item.getSheetId()) === saved);
    if (!sheet) throw new Error('原衛勤回應分頁不存在，未猜測其他資料來源。');
    return sheet;
  }
  const matches = ss.getSheets().filter(sheet => {
    if (!sheet.getLastColumn()) return false;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return (headers.includes('醫療協處累計人數') && headers.includes('後送累計人數')) ||
      ['衛勤類別', '回報編號', '累計人數'].every(title => headers.includes(title));
  });
  if (matches.length !== 1) throw new Error('無法唯一辨識衛勤回應分頁。');
  PropertiesService.getScriptProperties().setProperty('MEDICAL_RESPONSE_SHEET_ID', String(matches[0].getSheetId()));
  return matches[0];
}

function setupMedicalSupport() {
  return upgradeMedicalNumbered();
}

function upgradeMedicalNumbered(legacyNumber) {
  if (legacyNumber !== undefined && (!Number.isInteger(legacyNumber) || legacyNumber < 0 || legacyNumber > 15)) {
    throw new Error('舊回報編號須為 1–15，或 0 表示只保留舊紀錄。');
  }
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
    const titles = ['衛勤類別', '回報編號', '累計人數'];
    const types = [FormApp.ItemType.MULTIPLE_CHOICE, FormApp.ItemType.LIST, FormApp.ItemType.TEXT];
    const items = form.getItems();
    const oldTitles = ['醫療協處累計人數', '後送累計人數'];
    for (const item of items) {
      const index = titles.indexOf(item.getTitle());
      const expected = index >= 0 ? types[index] : oldTitles.includes(item.getTitle()) ? FormApp.ItemType.TEXT : null;
      if (!expected || item.getType() !== expected || items.filter(other => other.getTitle() === item.getTitle()).length !== 1) {
        throw new Error('衛勤表單題目不符，保留現有表單與回應。');
      }
    }
    const stats = ss.getSheetByName(MEDICAL_STATS_SHEET) || ss.insertSheet(MEDICAL_STATS_SHEET);
    if (props.getProperty('MEDICAL_REPORT_MODE') !== 'numbered' && !props.getProperty('MEDICAL_LEGACY_CAPTURED')) {
      const oldValues = stats.getRange('A2:B2').getValues()[0];
      props.setProperty('MEDICAL_LEGACY_CAPTURED', 'true');
      props.setProperty('MEDICAL_LEGACY_NUMBER', String(legacyNumber || 0));
      ['MEDICAL_LEGACY_ASSISTANCE', 'MEDICAL_LEGACY_EVACUATION'].forEach((key, i) => {
        if (legacyNumber && Number.isSafeInteger(oldValues[i]) && oldValues[i] >= 0) props.setProperty(key, String(oldValues[i]));
      });
    }
    const accepting = form.isAcceptingResponses();
    form.setAcceptingResponses(false);
    try {
      const categoryItem = items.find(item => item.getTitle() === titles[0]);
      (categoryItem ? categoryItem.asMultipleChoiceItem() : form.addMultipleChoiceItem().setTitle(titles[0]))
        .setChoiceValues(['醫療協處', '醫療後送']).setRequired(true);
      const numberItem = items.find(item => item.getTitle() === titles[1]);
      (numberItem ? numberItem.asListItem() : form.addListItem().setTitle(titles[1]))
        .setChoiceValues(Array.from({ length: 15 }, (_, i) => (i + 1) + '號')).setRequired(true);
      const quantityItem = items.find(item => item.getTitle() === titles[2]);
      (quantityItem ? quantityItem.asTextItem() : form.addTextItem().setTitle(titles[2]))
        .setRequired(true).setValidation(FormApp.createTextValidation().requireTextMatchesPattern('^[0-9]{1,9}$').setHelpText('請填 0 或正整數（最多九位數）。').build());
      items.filter(item => oldTitles.includes(item.getTitle())).forEach(item => form.deleteItem(item));
      form.setDescription('選擇醫療協處或醫療後送，再選自己的1～15號編號，填該編號、該類別目前累計人數。相同編號與類別以最新回報取代舊值，再加總所有編號。例如1號100人、2號200人，該類別總計300人。衛勤人數不納入進離場總計。');
      form.setConfirmationMessage('已收到此編號的累計回報，看板會更新各編號加總。');
    } finally { form.setAcceptingResponses(accepting); }
    if (!destination) form.setDestination(FormApp.DestinationType.SPREADSHEET, TARGET_SPREADSHEET_ID);
    if (form.supportsAdvancedResponderPermissions()) form.setPublished(true);
    form.setAcceptingResponses(true);
    SpreadsheetApp.flush();
    const source = findMedicalResponseSheet(ss, form);
    let columns = null;
    for (let attempt = 0; attempt < 5; attempt++) {
      const header = source.getRange(1, 1, 1, source.getLastColumn()).getValues()[0];
      if (titles.every(title => header.filter(value => value === title).length === 1)) {
        columns = { category: responseColumnLetter(header.indexOf(titles[0])), number: responseColumnLetter(header.indexOf(titles[1])), quantity: responseColumnLetter(header.indexOf(titles[2])) };
        break;
      }
      Utilities.sleep(1000);
    }
    if (!columns) throw new Error('新表單回應欄位尚未更新，請重試 setupMedicalSupport；歷史回應保留。');
    const seedNumber = Number(props.getProperty('MEDICAL_LEGACY_NUMBER') || 0);
    const seed = key => {
      const value = props.getProperty(key);
      return seedNumber && value !== null ? Number(value) : null;
    };
    stats.getRange('A1:D1').setValues([['醫療協處', '醫療後送', '最新回報時間', '填報表單']]);
    const quotedSource = "'" + source.getName().replace(/'/g, "''") + "'!";
    stats.getRange('A2:C2').setFormulas([[
      medicalNumberedTotalFormula(source.getName(), columns, '醫療協處', seedNumber, seed('MEDICAL_LEGACY_ASSISTANCE')),
      medicalNumberedTotalFormula(source.getName(), columns, '醫療後送', seedNumber, seed('MEDICAL_LEGACY_EVACUATION')),
      '=IFERROR(MAX(FILTER(' + quotedSource + 'A:A,ISNUMBER(' + quotedSource + 'A:A))),"未回報")'
    ]]);
    stats.getRange('C2').setNumberFormat('yyyy/MM/dd HH:mm:ss');
    stats.getRange('D2').setValue(form.getPublishedUrl());
    props.setProperty('MEDICAL_REPORT_MODE', 'numbered');
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
    choices: form ? form.getItems().filter(item => [FormApp.ItemType.MULTIPLE_CHOICE, FormApp.ItemType.LIST].includes(item.getType())).map(item => ({ title: item.getTitle(), values: (item.getType() === FormApp.ItemType.LIST ? item.asListItem() : item.asMultipleChoiceItem()).getChoices().map(choice => choice.getValue()) })) : [],
    accepting: form ? form.isAcceptingResponses() : false,
    stats: stats ? stats.getRange('A1:D2').getDisplayValues() : null,
    formulas: stats ? stats.getRange('A2:C2').getFormulas()[0] : [],
    responseRows: ss.getSheets().filter(sheet => /回應|Responses/.test(sheet.getName())).map(sheet => ({ name: sheet.getName(), rows: sheet.getLastRow() })),
    peopleBoard: ss.getSheetByName('總即時戰情看板').getRange('A1:AD40').getDisplayValues()
  };
}
