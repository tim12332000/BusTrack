const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const main = fs.readFileSync('【請複製這個】GoogleAppsScript完整終極版.js', 'utf8');
const reset = fs.readFileSync('【指令】一鍵清空資料.js', 'utf8');
const targetId = '1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w';
const published = [
  'https://docs.google.com/forms/d/e/1FAIpQLSeCDaMu9LlQhgwJKdzr6uCw2VX44ni5eO1Dn6gRePX4ur3dKw/viewform',
  'https://docs.google.com/forms/d/e/1FAIpQLSdNP01CZkqh5EeCkfmzrwQpQcPCw0fmXZmkQ50FVbvJrxUIPA/viewform'
];

function setup(options = {}) {
  const calls = [], alerts = [];
  const headers = [
    ['Timestamp', '1. 方向', '2. 站點 / 門號', '3. 車號', '4. 人數'],
    ['時間戳記', '1. 停車場區域', '2. 目前剩餘汽車停車位', '3. 目前剩餘機車停車位']
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
    rows: 3, getSheetId: () => i + 10, getName: () => 'response-' + i,
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
  const ss = {
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
  vm.runInContext(main + '\n' + reset, ctx);
  return { ctx, ss, forms, sheets, calls, alerts };
}

test('opening creates only the menu; cancelling performs no writes', () => {
  const p = setup({ cancel: true });
  p.ctx.onOpen();
  assert.deepEqual(p.calls, ['menu-runClearAllData']);
  p.calls.length = 0;
  p.ctx.runClearAllData();
  assert.deepEqual(p.calls, []);
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
