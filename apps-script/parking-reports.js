// Two report forms; previous dedicated responses remain history/fallback sources.
const PARKING_DEDICATED_LOTS = ['嶺東科大'];
let parkingReportSourcesCache_ = null;

function parkingReportGroups_() {
  return [{ key: 'other', title: '其他停車場一次回報', lots: PARKING_LOTS.filter(lot => !PARKING_DEDICATED_LOTS.includes(lot.name)) }]
    .concat(PARKING_DEDICATED_LOTS.map((name, index) => ({
      key: 'dedicated' + index, title: name + '車位回報', lots: PARKING_LOTS.filter(lot => lot.name === name)
    })));
}

function parkingReportFields_(group) {
  return group.lots.flatMap(lot => ['car', 'motorcycle'].filter(field => lot[field] > 0).map(field => ({
    lot: lot.name, field, capacity: lot[field],
    title: lot.name + '｜目前剩餘' + (field === 'car' ? '汽車' : '機車') + '位'
  })));
}

// Integer range validation, including zero, without accepting decimals or over-capacity values.
function parkingCapacityPattern_(capacity) {
  const digits = String(capacity);
  const parts = ['0'];
  for (let length = 1; length < digits.length; length++) parts.push('[1-9]' + '[0-9]'.repeat(length - 1));
  for (let i = 0; i < digits.length; i++) {
    const low = i === 0 ? 1 : 0;
    const high = Number(digits[i]) - 1;
    if (high >= low) parts.push(digits.slice(0, i) + (high === low ? String(low) : '[' + low + '-' + high + ']') + '[0-9]'.repeat(digits.length - i - 1));
  }
  parts.push(digits);
  return '^(?:' + parts.join('|') + ')$';
}

function configureParkingReportForm_(form, group) {
  const fields = parkingReportFields_(group);
  const items = form.getItems();
  const existingByTitle = new Map();
  for (const item of items) {
    const title = item.getTitle();
    if (!fields.some(field => field.title === title) || item.getType() !== FormApp.ItemType.TEXT || existingByTitle.has(title)) {
      throw new Error('回報題目不符，未刪除或重建題目：' + group.title);
    }
    existingByTitle.set(title, item);
  }
  form.setTitle('「成功嶺營區開放」' + group.title)
    .setDescription(group.key === 'other'
      ? '由一位回報人員填寫以下 22 處停車場的剩餘車位，包含經貿6及經貿8，全部填完後按一次提交。嶺東科大請使用專用回報網址。已滿請填 0；總容量不是剩餘車位。'
      : '本網址只回報' + group.lots[0].name + '，直接填目前剩餘車位。已滿請填 0。');
  fields.forEach(field => {
    const existing = existingByTitle.get(field.title);
    const item = existing ? existing.asTextItem() : form.addTextItem();
    item.setTitle(field.title).setRequired(true)
      .setHelpText('請填 0～' + field.capacity + ' 的整數；沒有剩餘車位請填 0。')
      .setValidation(FormApp.createTextValidation().requireTextMatchesPattern(parkingCapacityPattern_(field.capacity))
        .setHelpText('請填 0～' + field.capacity + ' 的整數。').build());
  });
  form.setCollectEmail(false).setLimitOneResponsePerUser(false)
    .setConfirmationMessage('已收到本次剩餘車位回報。').setShowLinkToRespondAgain(true);
}

function parkingSplitActive_() {
  return PropertiesService.getScriptProperties().getProperty('PARKING_REPORT_MODE') === 'split';
}

function parkingReportSources_(ss) {
  const props = PropertiesService.getScriptProperties();
  return parkingReportGroups_().map(group => {
    const id = props.getProperty('PARKING_REPORT_SHEET_' + group.key);
    const sheet = ss.getSheets().find(sheet => String(sheet.getSheetId()) === id);
    if (!sheet) throw new Error('找不到回報分頁：' + group.title);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const fields = parkingReportFields_(group).map(field => {
      const matches = headers.map((title, index) => title === field.title ? index : -1).filter(index => index >= 0);
      if (matches.length !== 1) throw new Error('回報欄位缺少或重複：' + field.title);
      return Object.assign({}, field, { column: responseColumnLetter(matches[0]) });
    });
    return { group, sheet, fields };
  });
}

function parkingSplitFormula_(name, field, capacity, legacyFormula) {
  if (!parkingSplitActive_()) return legacyFormula;
  if (!parkingReportSourcesCache_) parkingReportSourcesCache_ = parkingReportSources_(SpreadsheetApp.openById(TARGET_SPREADSHEET_ID));
  const source = parkingReportSourcesCache_.find(source => source.fields.some(item => item.lot === name && item.field === field));
  if (!source) throw new Error('找不到場地回報來源：' + name);
  const column = source.fields.find(item => item.lot === name && item.field === field).column;
  const prefix = "'" + source.sheet.getName().replace(/'/g, "''") + "'!";
  const archivedKey = { '經貿6停車場': 'dedicated1', '經貿8停車場': 'dedicated2' }[name];
  if (archivedKey) {
    const id = PropertiesService.getScriptProperties().getProperty('PARKING_REPORT_SHEET_' + archivedKey);
    if (id) {
      const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      const archived = ss.getSheets().find(sheet => String(sheet.getSheetId()) === id);
      if (!archived) throw new Error('找不到原獨立回報分頁：' + name);
      const title = name + '｜目前剩餘' + (field === 'car' ? '汽車' : '機車') + '位';
      const headers = archived.getRange(1, 1, 1, archived.getLastColumn()).getValues()[0];
      if (headers.filter(value => value === title).length !== 1) throw new Error('原獨立回報欄位不符：' + name);
      legacyFormula = parkingBatchRemainingFormula_("'" + archived.getName().replace(/'/g, "''") + "'!",
        responseColumnLetter(headers.indexOf(title)), capacity, legacyFormula);
    }
  }
  // Old shared submissions have no cells for the newly merged fields.
  return parkingBatchRemainingFormula_(prefix, column, capacity, legacyFormula, Boolean(archivedKey));
}

function parkingBatchRemainingFormula_(prefix, column, capacity, legacyFormula, requireValue = false) {
  const time = prefix + 'A:A';
  const value = prefix + column + ':' + column;
  // Whole-column references include the first row that Google Forms inserts after setup.
  const eligible = requireValue ? `ISNUMBER(${time})*(${value}<>"")=1` : `ISNUMBER(${time})`;
  const latest = `XLOOKUP(TRUE,ARRAYFORMULA(${eligible}),ARRAYFORMULA(IF(${value}="","未回報",IFERROR(VALUE(${value}),"回報異常"))),"__NONE__",0,-1)`;
  return `=LET(latest,${latest},IF(latest="__NONE__",${legacyFormula.slice(1)},IF(ISNUMBER(latest),IF(AND(latest>=0,latest<=${capacity},latest=INT(latest)),latest,"回報異常"),latest)))`;
}

function closeLegacyParkingForm_(form, sharedUrl) {
  // Preserve an existing working redirect instead of rewriting closed legacy
  // Forms, which can reject otherwise valid updates with "Invalid data".
  if (!form.isAcceptingResponses()) return;
  const message = '停車場回報請改用：' + sharedUrl;
  if (form.isAcceptingResponses()) form.setAcceptingResponses(false);
  if (form.getCustomClosedFormMessage() !== message) {
    try { form.setCustomClosedFormMessage(message); }
    catch (error) {
      // Some published Forms reject custom closed messages. Closure itself is
      // mandatory; the optional redirect must not leave another old form open.
      if (form.isAcceptingResponses()) throw error;
      Logger.log('表單已停用，自訂停用訊息未更新：' + error.message);
    }
  }
}

function setupParkingReports() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const paused = [];
  try {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const legacy = findResponseSource(ss, 'parking');
    const legacyForm = FormApp.openByUrl(legacy.sheet.getFormUrl());
    if (legacyForm.getDestinationId() !== TARGET_SPREADSHEET_ID ||
        legacyForm.getPublishedUrl() !== 'https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform') {
      throw new Error('原停車場表單不符，未切換回報來源。');
    }
    for (const group of parkingReportGroups_()) {
      const key = 'PARKING_REPORT_FORM_' + group.key;
      let id = props.getProperty(key);
      if (!id) {
        const created = FormApp.create('「成功嶺營區開放」' + group.title);
        id = created.getId();
        props.setProperty(key, id);
      }
      const form = FormApp.openById(id);
      let destination = null;
      try { destination = form.getDestinationId(); }
      catch (error) { if (!String(error.message).includes('no response destination')) throw error; }
      if (destination && destination !== TARGET_SPREADSHEET_ID) throw new Error('回報表單目的地不符：' + group.title);
      paused.push({ form, accepting: form.isAcceptingResponses() });
      form.setAcceptingResponses(false);
      configureParkingReportForm_(form, group);
      if (!destination) form.setDestination(FormApp.DestinationType.SPREADSHEET, TARGET_SPREADSHEET_ID);
      // New Forms support publishing explicitly; older Forms are already published.
      if (form.supportsAdvancedResponderPermissions()) form.setPublished(true);
      let matches = [];
      for (let attempt = 0; attempt < 10; attempt++) {
        SpreadsheetApp.flush();
        matches = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID).getSheets().filter(sheet => {
          const url = sheet.getFormUrl();
          return url && FormApp.openByUrl(url).getId() === form.getId();
        });
        if (matches.length === 1) break;
        Utilities.sleep(1000);
      }
      if (matches.length !== 1) throw new Error('等待回報分頁同步，請重試：' + group.title);
      props.setProperty('PARKING_REPORT_SHEET_' + group.key, String(matches[0].getSheetId()));
    }
    parkingReportSources_(ss); // Validate every field before activating the new routing.
    const registry = ss.getSheetByName('停車場回報入口') || ss.insertSheet('停車場回報入口');
    const oldRegistryRows = registry.getLastRow();
    const registryRows = [
      ['回報入口', '回報網址', '回應分頁ID'],
      ...parkingReportGroups_().map(group => [group.title,
        FormApp.openById(props.getProperty('PARKING_REPORT_FORM_' + group.key)).getPublishedUrl(),
        Number(props.getProperty('PARKING_REPORT_SHEET_' + group.key))])
    ];
    registry.getRange(1, 1, registryRows.length, 3).setValues(registryRows);
    if (oldRegistryRows > registryRows.length) registry.getRange(registryRows.length + 1, 1, oldRegistryRows - registryRows.length, 3).clearContent();
    props.setProperty('PARKING_REPORT_MODE', 'split');
    parkingReportSourcesCache_ = null;
    // Only the parking status formulas change; attendance and medical data are untouched.
    rebuildParkingStatusSheet(ss, legacy.sheet.getName().replace(/'/g, "''"), legacy.columns);
    closeLegacyParkingForm_(FormApp.openById(legacyForm.getId()),
      FormApp.openById(props.getProperty('PARKING_REPORT_FORM_other')).getPublishedUrl());
    for (const key of ['dedicated1', 'dedicated2']) {
      const id = props.getProperty('PARKING_REPORT_FORM_' + key);
      if (!id) continue;
      const form = FormApp.openById(id);
      if (form.getDestinationId() !== TARGET_SPREADSHEET_ID) throw new Error('原獨立表單目的地不符');
      closeLegacyParkingForm_(form, FormApp.openById(props.getProperty('PARKING_REPORT_FORM_other')).getPublishedUrl());
    }
    paused.forEach(item => item.form.setAcceptingResponses(item.accepting));
    SpreadsheetApp.flush();
    return inspectParkingReports();
  } finally {
    paused.forEach(item => item.form.setAcceptingResponses(item.accepting));
    lock.releaseLock();
  }
}

function inspectParkingReports() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const status = ss.getSheetByName('停車場現況');
  return {
    active: parkingSplitActive_(),
    forms: parkingReportGroups_().map(group => {
      const id = props.getProperty('PARKING_REPORT_FORM_' + group.key);
      const form = id ? FormApp.openById(id) : null;
      return { key: group.key, title: group.title, lots: group.lots.map(lot => lot.name),
        url: form ? form.getPublishedUrl() : null, accepting: form ? form.isAcceptingResponses() : false,
        items: form ? form.getItems().map(item => ({ id: item.getId(), title: item.getTitle(), required: item.asTextItem().isRequired() })) : [] };
    }),
    status: status ? status.getRange(1, 1, PARKING_LOTS.length + 1, 6).getDisplayValues() : null,
    responseRows: ss.getSheets().filter(sheet => sheet.getFormUrl()).map(sheet => ({ id: sheet.getSheetId(), rows: sheet.getLastRow() })),
    people: ss.getSheetByName('總即時戰情看板').getRange('A1:AD22').getDisplayValues(),
    medical: ss.getSheetByName('衛勤支援統計') ? ss.getSheetByName('衛勤支援統計').getRange('A1:D2').getDisplayValues() : null
  };
}
