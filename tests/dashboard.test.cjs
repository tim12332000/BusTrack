const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const code = fs.readFileSync('【請複製這個】GoogleAppsScript完整終極版.js', 'utf8');
const peopleHeaders = ['Timestamp', '1. 方向', '2. 站點 / 門號', '3. 車號', '4. 人數', '5. 備註'];
const parkingHeaders = ['時間戳記', '1. 停車場區域', '2. 目前剩餘汽車停車位', '3. 目前剩餘機車停車位', '4. 備註'];
function sheet(name, headers = [], linked = null) {
  const writes = [];
  let maxColumns = 26;
  return {
    writes, getName: () => name, getLastColumn: () => headers.length,
    getFormUrl: () => linked, getMaxRows: () => 1000, getMaxColumns: () => maxColumns,
    insertColumnsAfter(position, count) { assert.equal(position, maxColumns); maxColumns += count; },
    getRange(...args) {
      let range;
      range = new Proxy({}, { get(_, method) {
        if (method === 'getValues') return () => [headers];
        if (method === 'getValue') return () => 0;
        return (...values) => { writes.push({ args, method, values }); return range; };
      }});
      return range;
    },
    clearConditionalFormatRules() {}, setHiddenGridlines() {}, setColumnWidth() {}, setRowHeight() {}
  };
}
function context(sheets) {
  const ss = { getSheets: () => sheets, getSheetByName: n => sheets.find(s => s.getName() === n),
    insertSheet(n) { const s = sheet(n); sheets.push(s); return s; }, getUrl: () => 'test' };
  const ctx = vm.createContext({ Logger: { log() {} }, SpreadsheetApp: { openById: () => ss, flush() {}, BorderStyle: { SOLID: 'solid' } } });
  vm.runInContext(code, ctx);
  return { ctx, ss };
}
test('screenshot order selects people by headers, not first English response tab', () => {
  const parking = sheet('Form Responses 2', parkingHeaders, 'form2');
  const old = sheet('表單回應 2', parkingHeaders);
  const people = sheet('Form Responses 1', peopleHeaders, 'form1');
  const { ctx, ss } = context([parking, old, people]);
  assert.equal(ctx.findResponseSource(ss, 'people').sheet, people);
  assert.equal(ctx.findResponseSource(ss, 'parking').sheet, parking);
});
test('ambiguous or missing sources stop without modifying responses or creating dashboard', () => {
  const sources = [sheet('one', parkingHeaders), sheet('two', parkingHeaders), sheet('people', peopleHeaders)];
  const { ctx } = context(sources);
  assert.throws(() => ctx.repairDashboard(), /無法唯一辨識/);
  assert.equal(sources.length, 3);
  assert.ok(sources.every(s => s.writes.length === 0));
  assert.throws(() => context([]).ctx.repairDashboard(), /無法唯一辨識/);
});
test('repair preserves response sheets, follows moved columns and escapes sheet names', () => {
  const headers = Array(27).fill('');
  headers[0] = 'Timestamp'; headers[6] = '1. 方向'; headers[7] = '2. 站點 / 門號';
  headers[8] = '3. 車號'; headers[26] = '4. 人數';
  const people = sheet("People's replies", headers);
  const parking = sheet('Form Responses 2', parkingHeaders);
  const { ctx, ss } = context([parking, people]);
  ctx.repairDashboard();
  ctx.repairDashboard();
  assert.equal(people.writes.length, 0);
  assert.equal(parking.writes.length, 0);
  const detail = JSON.stringify(ss.getSheetByName('各車即時明細').writes);
  assert.ok(detail.includes("'People''s replies'!AA:AA"));
  assert.ok(detail.includes("'People''s replies'!G:G"));
  const dashboard = ss.getSheetByName('總即時戰情看板');
  const formulas = dashboard.writes.filter(w => w.method === 'setFormula').map(w => w.values[0]);
  assert.ok(formulas.some(f => f.includes("SUMIFS('People''s replies'!AA:AA")));
  assert.ok(formulas.some(f => f.includes("'Form Responses 2'!D2:D") && f.includes('XLOOKUP')));
  assert.ok(!formulas.some(f => f.includes("SUMIFS('Form Responses 2'")));
});

test('parking reads new fields even when legacy quantity and action remain in response sheet', () => {
  const legacy = ['Timestamp', '3. 車號', '1. 停車場區域', '2. 回報項目 / 動作', '3. 車輛數量 (輛)', '4. 備註'];
  const { ctx, ss } = context([sheet('old', legacy), sheet('current', [...legacy, '2. 目前剩餘汽車停車位', '3. 目前剩餘機車停車位'])]);
  const source = ctx.findResponseSource(ss, 'parking');
  assert.equal(source.sheet.getName(), 'current');
  assert.equal(source.columns.car, 'G');
  assert.equal(source.columns.motorcycle, 'H');
  const f = ctx.parkingRemainingFormula('current', source.columns, 'car', '*嶺東科大', 90);
  assert.ok(f.includes('"*嶺東科大", \'current\'!C2:C'));
  assert.ok(f.includes('IF(\'current\'!G2:G="", "未回報"'));
  assert.ok(f.includes('"未回報", 0, -1)'));
  assert.ok(f.includes('latest<=90'));
  assert.ok(f.includes('latest=INT(latest)'));
});

test('form migration is idempotent and does not reuse old quantity as remaining spaces', () => {
  const { ctx } = context([]);
  let nextId = 1;
  const items = [];
  function item(title, type) {
    const entry = { id: nextId++, title, type,
      getTitle() { return this.title; }, getType() { return this.type; }, getId() { return this.id; },
      getIndex() { return items.indexOf(this); }, asTextItem() { return this; }, asMultipleChoiceItem() { return this; },
      setTitle(v) { this.title = v; return this; }, setRequired(v) { this.required = v; return this; },
      setChoiceValues(v) { this.choices = v; return this; }, setValidation(v) { this.validation = v; return this; }
    };
    items.push(entry); return entry;
  }
  const area = item('1. 停車場區域', 'choice');
  item('2. 回報項目 / 動作', 'choice');
  const quantity = item('3. 車輛數量 (輛)', 'text');
  const note = item('4. 備註', 'text');
  ctx.FormApp = { ItemType: { MULTIPLE_CHOICE: 'choice', TEXT: 'text' }, createTextValidation() {
    return { requireTextMatchesPattern(v) { this.pattern = v; return this; }, setHelpText() { return this; }, build() { return this.pattern; } };
  }};
  const form = { getItems: () => items.slice(), addTextItem: () => item('', 'text'), addMultipleChoiceItem: () => item('', 'choice'),
    deleteItem(i) { items.splice(items.indexOf(i), 1); }, moveItem(from, to) { items.splice(to, 0, items.splice(from, 1)[0]); }, setDescription() {}, setTitle() {} };
  ctx.configureParkingRemainingForm(form);
  const ids = items.map(i => i.id);
  ctx.configureParkingRemainingForm(form);
  assert.deepEqual(items.map(i => i.id), ids);
  assert.equal(items.length, 4);
  assert.equal(items[0], area); assert.equal(items[3], note);
  assert.ok(!items.includes(quantity));
  assert.deepEqual(Array.from(items[0].choices), ['五都日出', '新烏日', '*嶺東科大', '*台中科大', '水湳轉運站', '*經貿六', '*經貿八']);
  for (const i of items.slice(1, 3)) {
    assert.equal(i.required, true);
    const re = new RegExp(i.validation);
    assert.ok(re.test('0')); assert.ok(re.test('123'));
    for (const invalid of ['', '-1', '2.5', '三']) assert.ok(!re.test(invalid));
  }
});

test('seven-site capacity totals and formulas exclude unsupported vehicle types', () => {
  const { ctx, ss } = context([sheet('people', peopleHeaders), sheet('parking', parkingHeaders)]);
  assert.equal(vm.runInContext('PARKING_LOTS.reduce((sum, lot) => sum + lot.car, 0)', ctx), 3023);
  assert.equal(vm.runInContext('PARKING_LOTS.reduce((sum, lot) => sum + lot.motorcycle, 0)', ctx), 2773);
  assert.equal(ctx.parkingRemainingFormula('parking', {}, 'car', '*台中科大', 0), '="不提供"');
  assert.equal(ctx.parkingTotalFormula('car'), '=IF(COUNT(A34,E34,I34,Q34,U34,Y34)=6,SUM(A34,E34,I34,Q34,U34,Y34),"未完整回報")');
  assert.equal(ctx.parkingTotalFormula('motorcycle'), '=IF(COUNT(C34,G34,K34,O34,S34)=5,SUM(C34,G34,K34,O34,S34),"未完整回報")');
  ctx.repairDashboard();
  const dashboard = ss.getSheetByName('總即時戰情看板');
  assert.equal(dashboard.getMaxColumns(), 28);
  const seventhCar = dashboard.writes.find(w => w.method === 'setFormula' && w.args[0] === 34 && w.args[1] === 25);
  assert.ok(seventhCar.values[0].includes('"*經貿八"'));
  assert.ok(seventhCar.values[0].includes('latest<=482'));
  const seventhMoto = dashboard.writes.find(w => w.method === 'setFormula' && w.args[0] === 34 && w.args[1] === 27);
  assert.equal(seventhMoto.values[0], '="不提供"');
});
test('duplicate question columns are rejected instead of guessing', () => {
  const { ctx, ss } = context([sheet('people', [...peopleHeaders, '4. 人數'])]);
  assert.throws(() => ctx.findResponseSource(ss, 'people'), /無法唯一辨識/);
});
test('both screen copies request dashboard by name and inline JavaScript parses', () => {
  for (const name of ['index.html', '戰情大螢幕.html']) {
    const html = fs.readFileSync(name, 'utf8');
    assert.ok(html.includes('&sheet=${encodeURIComponent("總即時戰情看板")}'));
    assert.ok(!html.includes('&gid=0'));
    for (const [, script] of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/g)) new vm.Script(script);
  }
});
