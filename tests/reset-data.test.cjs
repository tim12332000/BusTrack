const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const main = fs.readFileSync('【請複製這個】GoogleAppsScript完整終極版.js', 'utf8');
const reset = fs.readFileSync('【指令】一鍵清空資料.js', 'utf8');
const generate = fs.readFileSync('【指令】產生測試資料.js', 'utf8');
const targetId = '1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w';
const published = [
  'https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform',
  'https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform'
];

function setup(options = {}) {
  const calls = [], alerts = [];
  const headers = [
    ['Timestamp', '1. 方向', '2. 站點 / 門號', '3. 車號', '4. 人數', '5. 備註'],
    ['時間戳記', '1. 停車場區域', '2. 目前剩餘汽車停車位', '3. 目前剩餘機車停車位', options.missingNote ? '其他' : '4. 備註']
  ];
  const forms = published.map((url, i) => ({
    ids: ['response-' + i], accepting: i === 0,
    getId: () => 'form-' + i,
    getPublishedUrl: () => options.wrongForm && i === 1 ? 'wrong-form' : url,
    getDestinationId: () => targetId,
    getResponses() { return this.ids.map(id => ({ getId: () => id })); },
    isAcceptingResponses() { return this.accepting; },
    setAcceptingResponses(value) { calls.push('accept-' + i + '-' + value); this.accepting = value; },
    deleteAllResponses() {
      calls.push('delete-' + i);
      if (options.deleteFailure && i === 1) throw new Error('delete failed');
      this.ids = [];
    }
  }));
  const sheets = headers.map((header, i) => ({
    rows: 3, added: [], getSheetId: () => i + 10, getName: () => 'response-' + i,
    appendRow(row) {
      if (options.appendFailure && i === 1) throw new Error('append failed');
      this.added.push(row); this.rows++; calls.push('append-' + i);
    },
    getFormUrl: () => 'edit-' + i, getLastColumn: () => header.length,
    getLastRow() { return this.rows; },
    getRange(row, column, count, columns) {
      return {
        getValues: () => [header], canEdit: () => !options.protected,
        getFormulas: () => [[options.formula ? '=SUM(A1)' : '']],
        clearContent: () => {
          assert.deepEqual([row, column, count, columns], [2, 1, 2, header.length]);
          calls.push('clear-' + i); this.rows = 1;
        }
      };
    }
  }));
  const untouched = { getLastColumn: () => 0 };
  // 「停車場現況」分頁：B 欄名稱、E/F 欄容量，內容取自主程式場地名單。
  const lots = () => vm.runInContext('PARKING_LOTS', ctx);
  const status = {
    getLastRow: () => lots().length + 1,
    getRange: (row, column, count, columns) => {
      assert.deepEqual([row, column, count, columns], [2, 2, lots().length, 5]);
      return { getValues: () => lots().map(lot => [lot.name, '未回報', '未回報', lot.car || '', lot.motorcycle || '']) };
    }
  };
  const ss = {
    getSheetByName: name => name === '停車場現況' && !options.noStatus ? status : null,
    getId: () => options.wrongSheet ? 'backup-copy' : targetId,
    getSheets: () => [untouched, ...sheets],
    copy() { calls.push('backup'); if (options.backupFailure) throw new Error('backup failed'); return { getUrl: () => 'backup-url' }; }
  };
  const ui = {
    ButtonSet: { YES_NO: 'yes-no', OK: 'ok' }, Button: { YES: 'yes' },
    alert(title, text, buttons) {
      alerts.push({ title, text });
      if (buttons === 'yes-no' && options.newResponse) forms[0].ids.push('arrived-during-dialog');
      return options.cancel ? 'no' : 'yes';
    },
    createMenu: () => ({ addItem(label, handler) { calls.push('menu-' + handler); return this; }, addToUi() {} })
  };
  const ctx = vm.createContext({
    Logger: { log() {} }, SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => ui, flush() {} },
    FormApp: { openByUrl: url => forms[Number(url.split('-')[1])] },
    LockService: { getScriptLock: () => ({ tryLock: () => !options.busy, releaseLock: () => calls.push('unlock') }) },
    Utilities: { formatDate: () => 'test-time' }
  });
  vm.runInContext(main + '\n' + reset + '\n' + generate, ctx);
  return { ctx, ss, forms, sheets, calls, alerts };
}

function setupSplit(options = {}) {
  const p = setup();
  const titles = ['其他停車場一次回報', '嶺東科大車位回報', '經貿6停車場車位回報', '經貿8停車場車位回報'];
  const dedicated = options.twoEntries ? ['嶺東科大'] : ['嶺東科大', '經貿6停車場', '經貿8停車場'];
  const lots = Array.from(p.ctx.readParkingLots_(p.ss));
  const registryRows = titles.map((title, index) => [title, 'https://docs.google.com/forms/d/e/split-' + index + '/viewform', index + 100]);
  if (options.twoEntries) {
    registryRows[2][1] = 'https://docs.google.com/forms/d/e/1FAIpQLSdjSlwTCnBK6UI9vWCMA2F_boLWDv4JFJTXorpjVaEY_xswWw/viewform';
    registryRows[3][1] = 'https://docs.google.com/forms/d/e/1FAIpQLSchHlnE9hOUnYu8b3CUQbDKU1EO80hdOIRT3jkH-8D8iFCBgQ/viewform';
  }
  const batchHeaders = titles.map((title, index) => ['Timestamp', ...lots.filter(lot => index === 0 ? !dedicated.includes(lot.name) : lot.name + '車位回報' === title)
    .flatMap(lot => ['car', 'motorcycle'].filter(field => lot[field] > 0)
      .map(field => lot.name + '｜目前剩餘' + (field === 'car' ? '汽車' : '機車') + '位'))]);
  if (options.badHeader) batchHeaders[3][1] = 'unexpected';
  if (options.duplicateHeader) batchHeaders[0][2] = batchHeaders[0][1];
  if (options.badRegistry) registryRows[3][0] = 'unknown';
  if (options.duplicateSource) registryRows[3][2] = registryRows[0][2];
  const batchForms = titles.map((title, index) => ({
    ids: ['native-split-' + index], accepting: index !== 1,
    getId: () => 'split-form-' + index,
    getTitle: () => options.wrongTitle && index === 3 ? 'unrelated' : '「成功嶺營區開放」' + title,
    getPublishedUrl: () => registryRows[index][1],
    getDestinationId: () => options.wrongDestination && index === 3 ? 'other-spreadsheet' : targetId,
    getResponses() { return this.ids.map(id => ({ getId: () => id })); },
    isAcceptingResponses() { return this.accepting; },
    setAcceptingResponses(value) { p.calls.push('split-accept-' + index + '-' + value); this.accepting = value; },
    deleteAllResponses() { p.calls.push('split-delete-' + index); this.ids = []; }
  }));
  const batchSheets = batchHeaders.map((headers, index) => ({
    rows: 3, added: [], getSheetId: () => index + 100, getName: () => 'split-response-' + index,
    getFormUrl: () => 'split-edit-' + index, getLastColumn: () => headers.length,
    getLastRow() { return this.rows; },
    appendRow(row) { this.added.push(row); this.rows++; p.calls.push('split-append-' + index); },
    getRange(row, column, count, width) { return {
      getValues: () => [headers], canEdit: () => !(options.protectedSplit && index === 3),
      getFormulas: () => [[options.formulaSplit && index === 3 ? '=1' : '']],
      clearContent: () => { assert.deepEqual([row, column, count, width], [2, 1, 2, headers.length]); this.rows = 1; p.calls.push('split-clear-' + index); }
    }; }
  }));
  const registry = { getLastRow: () => options.twoEntries ? 3 : 5, getRange: () => ({ getValues: () => [['回報入口', '回報網址', '回應分頁ID'], ...registryRows.slice(0, options.twoEntries ? 2 : 4)] }) };
  const getSheetByName = p.ss.getSheetByName;
  const getSheets = p.ss.getSheets;
  const openByUrl = p.ctx.FormApp.openByUrl;
  p.ss.getSheetByName = name => name === '停車場回報入口' ? registry : getSheetByName(name);
  p.ss.getSheets = () => [...getSheets(), ...batchSheets];
  p.ctx.FormApp.openByUrl = url => {
    const index = registryRows.findIndex((row, index) => row[1] === url || url === 'split-edit-' + index);
    return index < 0 ? openByUrl(url) : batchForms[index];
  };
  if (options.wrongIdentity) {
    const open = p.ctx.FormApp.openByUrl;
    p.ctx.FormApp.openByUrl = url => url === registryRows[3][1] ? { getId: () => 'unrelated-form' } : open(url);
  }
  if (options.newSplitResponse) {
    const ui = p.ctx.SpreadsheetApp.getUi();
    const alert = ui.alert;
    ui.alert = (title, text, buttons) => { const result = alert(title, text, buttons); if (buttons === 'yes-no') batchForms[3].ids.push('new-arrival'); return result; };
    p.ctx.SpreadsheetApp.getUi = () => ui;
  }
  return Object.assign(p, { batchForms, batchSheets, batchHeaders, registryRows });
}

test('opening creates only the menu; cancelling performs no writes', () => {
  const p = setup({ cancel: true });
  p.ctx.onOpen();
  assert.deepEqual(p.calls, ['menu-runGenerateTestData', 'menu-runClearAllData']);
  p.calls.length = 0;
  p.ctx.runClearAllData();
  assert.deepEqual(p.calls, []);
});

test('menu initialization needs no spreadsheet access or authorization', () => {
  const p = setup();
  p.ctx.SpreadsheetApp.getActiveSpreadsheet = () => { throw new Error('spreadsheet access unavailable'); };
  p.ctx.onOpen();
  assert.deepEqual(p.calls, ['menu-runGenerateTestData', 'menu-runClearAllData']);
});

test('confirmed reset backs up first, clears both sources below headers, restores collection states', () => {
  const p = setup(); p.ctx.runClearAllData();
  assert.equal(p.alerts.at(-1).title, '清除完成');
  assert.ok(p.calls.indexOf('backup') < p.calls.indexOf('delete-0'));
  assert.deepEqual(p.calls.filter(c => /^(delete|clear)-/.test(c)), ['delete-0', 'clear-0', 'delete-1', 'clear-1']);
  assert.ok(p.forms.every(f => f.ids.length === 0));
  assert.deepEqual(p.forms.map(f => f.accepting), [true, false]);
  assert.ok(p.sheets.every(s => s.rows === 1));
  assert.equal(p.calls.at(-1), 'unlock');
});

test('wrong targets, formulas, protected cells and changed reports stop before deleting', () => {
  for (const option of ['wrongForm', 'wrongSheet', 'formula', 'protected', 'newResponse', 'busy']) {
    const p = setup({ [option]: true }); p.ctx.runClearAllData();
    assert.equal(p.alerts.at(-1).title, '清除未完成', option);
    assert.ok(!p.calls.some(c => /^(delete|clear)-/.test(c)), option);
  }
});

test('backup failure deletes nothing and restores forms', () => {
  const p = setup({ backupFailure: true }); p.ctx.runClearAllData();
  assert.ok(!p.calls.some(c => /^(delete|clear)-/.test(c)));
  assert.deepEqual(p.forms.map(f => f.accepting), [true, false]);
  assert.equal(p.alerts.at(-1).title, '清除未完成');
});

test('partial failure reports backup and never claims success', () => {
  const p = setup({ deleteFailure: true }); p.ctx.runClearAllData();
  assert.equal(p.alerts.at(-1).title, '清除未完成');
  assert.match(p.alerts.at(-1).text, /backup-url/);
  assert.deepEqual(p.forms.map(f => f.accepting), [true, false]);
  assert.equal(p.calls.at(-1), 'unlock');
});

test('private executor rejects missing confirmation snapshot', () => {
  const p = setup();
  assert.throws(() => p.ctx.clearConfirmedTestData_(p.ss), /重新按選單確認/);
  assert.deepEqual(p.calls, ['unlock']);
});

test('generation appends 14 people and 23 parking rows without deleting or overwriting existing data', () => {
  const p = setup(); p.ctx.runGenerateTestData();
  assert.equal(p.alerts.at(-1).title, '新增完成');
  assert.deepEqual(p.sheets.map(s => s.added.length), [14, 23]);
  assert.deepEqual(p.sheets.map(s => s.rows), [17, 26]);
  assert.ok(!p.calls.some(c => /^(clear|delete|accept)-/.test(c)));
  assert.ok(p.sheets.every(s => s.added.every(row => row.at(-1) === '測試資料')));
  assert.ok(p.forms.every(f => f.ids.length === 1), 'native form responses remain unchanged');
});

test('generation cancellation, lock contention and bad second-sheet schema add nothing', () => {
  for (const option of ['cancel', 'busy', 'wrongForm', 'missingNote', 'noStatus']) {
    const p = setup({ [option]: true }); p.ctx.runGenerateTestData();
    assert.ok(p.sheets.every(s => s.added.length === 0), option);
  }
});

test('generation reports partial additions honestly and releases the reset lock', () => {
  const p = setup({ appendFailure: true }); p.ctx.runGenerateTestData();
  assert.equal(p.alerts.at(-1).title, '新增未完成');
  assert.match(p.alerts.at(-1).text, /已新增 14 筆/);
  assert.equal(p.calls.at(-1), 'unlock');
});

test('parking test rows follow migrated headers and leave retired action/quantity columns empty', () => {
  const p = setup();
  const headers = ['Timestamp', '3. 車號', '1. 停車場區域', '2. 回報項目 / 動作', '3. 車輛數量 (輛)', '4. 備註', '2. 目前剩餘汽車停車位', '3. 目前剩餘機車停車位'];
  const lots = vm.runInContext('PARKING_LOTS', p.ctx);
  const rows = p.ctx.buildTestDataRows_('parking', { area: 'C', car: 'G', motorcycle: 'H' }, headers, new Date(), p.ctx.readParkingLots_(p.ss));
  assert.equal(rows.length, 23);
  assert.deepEqual(Array.from(rows[0].slice(1)), ['', 'CITYPARKING春安站', '', '', '測試資料', 53, 1]);
  // 名稱必須和主程式場地名單一致，「停車場現況」才比對得到；剩餘不可超過容量，無此車種填 0。
  assert.deepEqual(rows.map(row => row[2]), lots.map(lot => lot.name));
  rows.forEach((row, i) => {
    assert.ok(Number.isInteger(row[6]) && row[6] >= 0 && row[6] <= lots[i].car, lots[i].name);
    assert.ok(Number.isInteger(row[7]) && row[7] >= 0 && row[7] <= lots[i].motorcycle, lots[i].name);
  });
  assert.equal(rows[1][7], 0);
  assert.equal(rows[2][6], 0);
  assert.ok(rows.some(row => row[6] === 0) && rows.some(row => row[6] > 0));
  assert.ok(rows.every(row => Number.isFinite(row[0].getTime())));
});

test('editor-runnable parking generator writes to the sheet the status formulas read, without any dialog', () => {
  const p = setup();
  const lots = vm.runInContext('PARKING_LOTS', p.ctx);
  const headers = ['時間戳記', '1. 停車場區域', '2. 目前剩餘汽車停車位', '3. 目前剩餘機車停車位', '4. 備註'];
  const added = [], logs = [];
  const responses = { getLastColumn: () => headers.length, getRange: () => ({ getValues: () => [headers] }), appendRow: row => added.push(row) };
  const status = p.ss.getSheetByName('停車場現況');
  status.getRange = (row, column, count, columns) => row === 2 && column === 3 && count === undefined
    ? { getFormula: () => "=LET(latest,XLOOKUP(\"x\", 'Form Responses 2'!B2:B, 1))" }
    : { getValues: () => lots.map(lot => [lot.name, '', '', lot.car || '', lot.motorcycle || '']) };
  p.ctx.SpreadsheetApp.openById = id => (assert.equal(id, '1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w'),
    { getSheetByName: name => name === '停車場現況' ? status : name === 'Form Responses 2' ? responses : null });
  p.ctx.SpreadsheetApp.getUi = () => { throw new Error('Cannot call SpreadsheetApp.getUi() from this context.'); };
  p.ctx.Logger.log = text => logs.push(text);
  p.ctx.addParkingTestDataFromEditor();
  assert.equal(added.length, 23);
  assert.deepEqual(added.map(row => row[1]), Array.from(lots, lot => lot.name));
  assert.ok(added.every(row => row[4] === '測試資料'));
  assert.match(logs[0], /Form Responses 2.*23 筆/);
});

test('split reset clears all six sources only after backup and restores each collection state', () => {
  const p = setupSplit();
  p.ctx.runClearAllData();
  assert.equal(p.alerts.at(-1).title, '清除完成');
  assert.match(p.alerts.at(-1).text, /6 份表單/);
  assert.equal(p.calls.filter(call => /^(delete-|split-delete-)/.test(call)).length, 6);
  assert.ok(p.calls.indexOf('backup') < p.calls.indexOf('delete-0'));
  assert.ok(p.calls.indexOf('backup') < p.calls.indexOf('split-delete-0'));
  assert.ok([...p.sheets, ...p.batchSheets].every(sheet => sheet.rows === 1));
  assert.ok([...p.forms, ...p.batchForms].every(form => form.ids.length === 0));
  assert.deepEqual(p.batchForms.map(form => form.accepting), [true, false, true, true]);
  assert.deepEqual(p.forms.map(form => form.accepting), [true, false]);
});

test('split registry identity, destination, uniqueness and race failures prevent all deletion', () => {
  for (const option of ['badRegistry', 'duplicateSource', 'wrongTitle', 'wrongDestination', 'wrongIdentity', 'newSplitResponse', 'protectedSplit', 'formulaSplit']) {
    const p = setupSplit({ [option]: true });
    p.ctx.runClearAllData();
    assert.equal(p.alerts.at(-1).title, '清除未完成', option);
    assert.ok(!p.calls.some(call => /^(delete-|clear-|split-delete-|split-clear-)/.test(call)), option);
    assert.deepEqual(p.batchForms.map(form => form.accepting), [true, false, true, true], option);
  }
});

test('split generation appends 14 people and one row per new form, leaving legacy and native responses intact', () => {
  const p = setupSplit();
  p.ctx.runGenerateTestData();
  assert.equal(p.alerts.at(-1).title, '新增完成');
  assert.match(p.alerts.at(-1).text, /14 筆人員及 4 筆停車場/);
  assert.deepEqual(p.sheets.map(sheet => sheet.added.length), [14, 0]);
  assert.deepEqual(p.batchSheets.map(sheet => sheet.added.length), [1, 1, 1, 1]);
  assert.ok([...p.forms, ...p.batchForms].every(form => form.ids.length === 1));
  assert.ok(!p.calls.some(call => /delete|clear|accept/.test(call)));
  const lots = Array.from(p.ctx.readParkingLots_(p.ss));
  p.batchSheets.forEach((sheet, index) => {
    const row = sheet.added[0];
    assert.ok(Number.isFinite(row[0].getTime()));
    row.slice(1).forEach((value, fieldIndex) => {
      const title = p.batchHeaders[index][fieldIndex + 1];
      const lot = lots.find(lot => title.startsWith(lot.name + '｜'));
      const capacity = title.endsWith('汽車位') ? lot.car : lot.motorcycle;
      assert.equal(value, Math.floor(capacity * 0.4));
    });
  });
});

test('invalid last batch schema or registry prevents every test data append', () => {
  for (const option of ['badHeader', 'duplicateHeader', 'badRegistry', 'wrongIdentity', 'wrongTitle']) {
    const p = setupSplit({ [option]: true });
    p.ctx.runGenerateTestData();
    assert.equal(p.alerts.at(-1).title, '新增未完成', option);
    assert.ok([...p.sheets, ...p.batchSheets].every(sheet => sheet.added.length === 0), option);
  }
});

test('editor generator rejects split mode before changing any rows', () => {
  const p = setupSplit();
  p.ctx.SpreadsheetApp.openById = () => p.ss;
  assert.throws(() => p.ctx.addParkingTestDataFromEditor(), /四入口模式請使用/);
  assert.ok([...p.sheets, ...p.batchSheets].every(sheet => sheet.added.length === 0));
});

test('two-entry reset also clears archived Jingmao fallback data', () => {
  const p = setupSplit({ twoEntries: true });
  p.ctx.runClearAllData();
  assert.equal(p.alerts.at(-1).title, '清除完成');
  assert.ok([...p.sheets, ...p.batchSheets].every(sheet => sheet.rows === 1));
  assert.ok(p.batchForms.every(form => form.ids.length === 0));
});

test('two-entry generation writes only active forms, including merged Jingmao fields', () => {
  const p = setupSplit({ twoEntries: true });
  p.ctx.runGenerateTestData();
  assert.equal(p.alerts.at(-1).title, '新增完成');
  assert.deepEqual(p.batchSheets.map(sheet => sheet.added.length), [1, 1, 0, 0]);
  for (const name of ['經貿6停車場', '經貿8停車場']) {
    const index = p.batchHeaders[0].indexOf(name + '｜目前剩餘汽車位');
    assert.ok(index > 0);
    assert.ok(p.batchSheets[0].added[0][index] >= 0);
  }
});

test('two-entry reset refuses to leave an unidentified archive behind', () => {
  const p = setupSplit({ twoEntries: true });
  p.batchForms[2].getPublishedUrl = () => 'https://docs.google.com/forms/d/e/unrelated/viewform';
  p.ctx.runClearAllData();
  assert.equal(p.alerts.at(-1).title, '清除未完成');
  assert.ok(!p.calls.some(call => /^(delete-|clear-|split-delete-|split-clear-)/.test(call)));
});
