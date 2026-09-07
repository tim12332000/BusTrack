const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const code = fs.readFileSync('【請複製這個】GoogleAppsScript完整終極版.js', 'utf8');
const peopleHeaders = ['Timestamp', '1. 方向', '2. 站點 / 門號', '3. 車號', '4. 人數', '5. 備註'];
const parkingHeaders = ['時間戳記', '1. 停車場區域', '2. 回報項目 / 動作', '3. 車輛數量 (輛)', '4. 備註'];
function sheet(name, headers = [], linked = null) {
  const writes = [];
  return {
    writes, getName: () => name, getLastColumn: () => headers.length,
    getFormUrl: () => linked, getMaxRows: () => 1000, getMaxColumns: () => 26,
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
  assert.ok(formulas.some(f => f.includes("SUMIFS('Form Responses 2'!D:D")));
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
