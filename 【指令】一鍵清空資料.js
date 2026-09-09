/**
 * 上線前清除測試資料：試算表「運管中心 → 清除測試資料」。
 * 與主程式（含 TARGET_SPREADSHEET_ID、findResponseSource）一起安裝，勿覆蓋主程式。
 * 開啟試算表不會清除資料；必須按選單並確認。
 */
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss && ss.getId() === TARGET_SPREADSHEET_ID) {
    SpreadsheetApp.getUi().createMenu('運管中心')
      .addItem('清除測試資料', 'runClearAllData').addToUi();
  }
}

function prepareTestDataReset_(ss) {
  if (!ss || ss.getId() !== TARGET_SPREADSHEET_ID) throw new Error('不是指定的正式試算表，未清除資料。');
  const published = {
    people: 'https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform',
    parking: 'https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform'
  };
  const targets = ['people', 'parking'].map(kind => {
    const sheet = findResponseSource(ss, kind).sheet;
    const formUrl = sheet.getFormUrl();
    if (!formUrl) throw new Error('回應分頁未連結表單：' + sheet.getName());
    const form = FormApp.openByUrl(formUrl);
    if (form.getDestinationId() !== ss.getId() || form.getPublishedUrl() !== published[kind]) {
      throw new Error('表單與正式網址或試算表不符：' + sheet.getName());
    }
    const rows = Math.max(0, sheet.getLastRow() - 1);
    const range = rows ? sheet.getRange(2, 1, rows, sheet.getLastColumn()) : null;
    if (range && (!range.canEdit() || range.getFormulas().some(row => row.some(Boolean)))) {
      throw new Error('回應資料包含公式或無編輯權限，未清除：' + sheet.getName());
    }
    const responseIds = form.getResponses().map(response => response.getId());
    return { kind, sheet, form, rows, responseIds, accepting: form.isAcceptingResponses() };
  });
  if (targets[0].sheet.getSheetId() === targets[1].sheet.getSheetId() ||
      targets[0].form.getId() === targets[1].form.getId()) throw new Error('兩份資料來源重複，未清除資料。');
  const signature = JSON.stringify(targets.map(t => [t.sheet.getSheetId(), t.form.getId(), t.rows, t.responseIds]));
  return { targets, signature };
}

function runClearAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const preview = prepareTestDataReset_(ss);
    const counts = preview.targets.map(t => (t.kind === 'people' ? '表單 1 人員疏運' : '表單 2 停車場') +
      '：' + t.responseIds.length + ' 份回覆、試算表 ' + t.rows + ' 筆').join('\n');
    const answer = ui.alert('清除全部測試回報？', counts +
      '\n\n會先備份試算表，再清空兩份表單回覆與回應資料。\n' +
      '表單題目、看板公式與原網址保留。表單原生回覆刪除後無法還原。', ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) return;
    const result = clearConfirmedTestData_(ss, preview.signature);
    ui.alert('清除完成', '兩份表單與回應資料已清空。\n人員統計歸零，停車位等待重新回報。\n' +
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
