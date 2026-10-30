/**
 * 上線前清除測試資料：試算表「運管中心 → 清除測試資料」。
 * 與主程式（含 TARGET_SPREADSHEET_ID、findResponseSource）一起安裝，勿覆蓋主程式。
 * 開啟試算表不會清除資料；必須按選單並確認。
 */
function onOpen() {
  // 開啟時只建立介面；資料來源與權限檢查留在使用者點選之後。
  SpreadsheetApp.getUi().createMenu('運管中心')
    .addItem('新增測試資料', 'runGenerateTestData')
    .addItem('清除測試資料', 'runClearAllData').addToUi();
}

// The registry is shared with the separately deployed reporting project.
function readParkingReportRegistry_(ss) {
  const registry = ss.getSheetByName('停車場回報入口');
  if (!registry) return [];
  const rowCount = registry.getLastRow();
  if (![3, 5].includes(rowCount)) throw new Error('停車場回報入口清單不完整，未修改資料。');
  const values = registry.getRange(1, 1, rowCount, 3).getValues();
  if (JSON.stringify(values[0]) !== JSON.stringify(['回報入口', '回報網址', '回應分頁ID'])) {
    throw new Error('停車場回報入口欄位不符，未修改資料。');
  }
  const titles = ['其他停車場一次回報', '嶺東科大車位回報'];
  if (rowCount === 5) titles.push('經貿6停車場車位回報', '經貿8停車場車位回報');
  const records = titles.map(title => {
    const matches = values.slice(1).filter(row => row[0] === title);
    if (matches.length !== 1 || !/^https:\/\/docs\.google\.com\/forms\/d\/e\/[A-Za-z0-9_-]+\/viewform$/.test(matches[0][1]) ||
        !Number.isSafeInteger(Number(matches[0][2])) || Number(matches[0][2]) < 0) {
      throw new Error('停車場回報入口設定不符：' + title);
    }
    return { title, url: matches[0][1], sheetId: Number(matches[0][2]), kind: 'parkingBatch' };
  });
  if (new Set(records.map(record => record.sheetId)).size !== titles.length || new Set(records.map(record => record.url)).size !== titles.length) {
    throw new Error('停車場回報入口重複，未修改資料。');
  }
  return records;
}

function prepareTestDataReset_(ss) {
  if (!ss || ss.getId() !== TARGET_SPREADSHEET_ID) throw new Error('不是指定的正式試算表，未清除資料。');
  const published = {
    people: 'https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform',
    parking: 'https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform'
  };
  const definitions = ['people', 'parking'].map(kind => ({ kind, title: kind === 'people' ? '人員疏運' : '舊停車場回報', url: published[kind],
    sheet: findResponseSource(ss, kind).sheet })).concat(readParkingReportRegistry_(ss).map(record => {
    const sheet = ss.getSheets().find(sheet => sheet.getSheetId && sheet.getSheetId() === record.sheetId);
    if (!sheet) throw new Error('找不到回報分頁：' + record.title);
    return Object.assign({}, record, { sheet });
  }));
  // Historical dedicated data remains a formula fallback after the merge and
  // must also be cleared, even though these links are absent from the registry.
  if (readParkingReportRegistry_(ss).length === 2) {
    const archived = [
      ['經貿6停車場車位回報', 'https://docs.google.com/forms/d/e/1FAIpQLSdjSlwTCnBK6UI9vWCMA2F_boLWDv4JFJTXorpjVaEY_xswWw/viewform'],
      ['經貿8停車場車位回報', 'https://docs.google.com/forms/d/e/1FAIpQLSchHlnE9hOUnYu8b3CUQbDKU1EO80hdOIRT3jkH-8D8iFCBgQ/viewform']
    ];
    for (const [title, url] of archived) {
      const matches = ss.getSheets().filter(sheet => sheet.getFormUrl && sheet.getFormUrl() && FormApp.openByUrl(sheet.getFormUrl()).getPublishedUrl() === url);
      if (matches.length !== 1) throw new Error('找不到原獨立回報來源：' + title);
      definitions.push({ kind: 'parkingArchive', title, url, sheet: matches[0] });
    }
  }
  const targets = definitions.map(definition => {
    const { kind, sheet, title } = definition;
    const formUrl = sheet.getFormUrl();
    if (!formUrl) throw new Error('回應分頁未連結表單：' + sheet.getName());
    const form = FormApp.openByUrl(formUrl);
    if (form.getDestinationId() !== ss.getId() || form.getPublishedUrl() !== definition.url ||
        (kind === 'parkingBatch' && (form.getTitle() !== '「成功嶺營區開放」' + title ||
          FormApp.openByUrl(definition.url).getId() !== form.getId()))) {
      throw new Error('表單與正式網址或試算表不符：' + sheet.getName());
    }
    const rows = Math.max(0, sheet.getLastRow() - 1);
    const range = rows ? sheet.getRange(2, 1, rows, sheet.getLastColumn()) : null;
    if (range && (!range.canEdit() || range.getFormulas().some(row => row.some(Boolean)))) {
      throw new Error('回應資料包含公式或無編輯權限，未清除：' + sheet.getName());
    }
    const responseIds = form.getResponses().map(response => response.getId());
    return { kind, title, sheet, form, rows, responseIds, accepting: form.isAcceptingResponses() };
  });
  if (new Set(targets.map(target => target.sheet.getSheetId())).size !== targets.length ||
      new Set(targets.map(target => target.form.getId())).size !== targets.length) throw new Error('資料來源重複，未清除資料。');
  const signature = JSON.stringify(targets.map(t => [t.sheet.getSheetId(), t.form.getId(), t.rows, t.responseIds]));
  return { targets, signature };
}

function runClearAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const preview = prepareTestDataReset_(ss);
    const counts = preview.targets.map(t => t.title +
      '：' + t.responseIds.length + ' 份回覆、試算表 ' + t.rows + ' 筆').join('\n');
    const answer = ui.alert('清除全部測試回報？', counts +
      '\n\n會先備份試算表，再清空以上 ' + preview.targets.length + ' 份表單回覆與回應資料。\n' +
      '表單題目、看板公式與原網址保留。表單原生回覆刪除後無法還原。', ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
    const result = clearConfirmedTestData_(ss, preview.signature);
    ui.alert('清除完成', preview.targets.length + ' 份表單與回應資料已清空。\n人員統計歸零，停車位等待重新回報。\n' +
      '試算表備份：' + result.backupUrl, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('清除未完成', error.message, ui.ButtonSet.OK);
  }
}

// 私有入口：没有確認當時的來源快照就不能清除；對話框關閉後才取得鎖定。
function clearConfirmedTestData_(ss, expectedSignature) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('已有清除作業進行中，請稍後再試。');
  const paused = [];
  let backupUrl = '';
  let failure = null;
  try {
    const plan = prepareTestDataReset_(ss);
    if (!expectedSignature || plan.signature !== expectedSignature) {
      throw new Error('回報資料已變更，請重新按選單確認。未清除資料。');
    }
    plan.targets.forEach(t => { paused.push(t); t.form.setAcceptingResponses(false); });
    SpreadsheetApp.flush();
    if (prepareTestDataReset_(ss).signature !== expectedSignature) {
      throw new Error('剛有新回報進入，請重新確認。未清除資料。');
    }
    backupUrl = ss.copy('上線前備份_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss')).getUrl();
    plan.targets.forEach(t => {
      t.form.deleteAllResponses();
      if (t.rows) t.sheet.getRange(2, 1, t.rows, t.sheet.getLastColumn()).clearContent();
    });
    SpreadsheetApp.flush();
    plan.targets.forEach(t => {
      if (t.form.getResponses().length || t.sheet.getLastRow() > 1) throw new Error('仍有回報資料，請檢查：' + t.sheet.getName());
    });
  } catch (error) {
    failure = error.message;
  } finally {
    paused.forEach(t => {
      try { t.form.setAcceptingResponses(t.accepting); }
      catch (error) { failure = (failure ? failure + '\n' : '') + '請手動確認表單收件狀態：' + t.sheet.getName(); }
    });
    lock.releaseLock();
  }
  if (failure) throw new Error(failure + (backupUrl ? '\n可能已部分清除，試算表備份：' + backupUrl : ''));
  return { backupUrl };
}
