const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function context() {
  const properties = {};
  const ctx = vm.createContext({
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => properties[key] || null, setProperty: (key, value) => { properties[key] = value; } }) },
    FormApp: { ItemType: { TEXT: 'text' }, createTextValidation: () => ({
      requireTextMatchesPattern(pattern) { this.pattern = pattern; return this; }, setHelpText() { return this; }, build() { return this.pattern; }
    }) }
  });
  vm.runInContext(fs.readFileSync('【請複製這個】GoogleAppsScript完整終極版.js', 'utf8'), ctx);
  vm.runInContext(fs.readFileSync('apps-script/parking-reports.js', 'utf8'), ctx);
  return { ctx, properties };
}

test('two forms partition all 23 lots once; other includes Jingmao Six and Eight', () => {
  const { ctx } = context();
  const groups = ctx.parkingReportGroups_();
  assert.deepEqual(Array.from(groups, group => group.lots.length), [22, 1]);
  const names = Array.from(groups).flatMap(group => Array.from(group.lots, lot => lot.name));
  assert.equal(new Set(names).size, 23);
  assert.deepEqual(Array.from(groups.slice(1), group => group.lots[0].name), ['嶺東科大']);
  assert.ok(groups[0].lots.some(lot => lot.name === '經貿6停車場'));
  assert.ok(groups[0].lots.some(lot => lot.name === '經貿8停車場'));
  const fields = ctx.parkingReportFields_(groups[0]);
  assert.equal(new Set(Array.from(fields, field => field.title)).size, fields.length);
  assert.ok(fields.every(field => field.capacity > 0));
  assert.deepEqual(Array.from(ctx.parkingReportFields_(groups[1]), field => field.field), ['motorcycle']);
});

test('every supported capacity accepts exactly integers from zero to capacity', () => {
  const { ctx } = context();
  for (const group of ctx.parkingReportGroups_()) for (const field of ctx.parkingReportFields_(group)) {
    const pattern = new RegExp(ctx.parkingCapacityPattern_(field.capacity));
    for (let value = 0; value <= field.capacity + 20; value++) assert.equal(pattern.test(String(value)), value <= field.capacity, `${field.title}: ${value}`);
    for (const invalid of ['', ' ', '-1', '0.5', '1.0', 'abc', '1e2', '99999']) assert.equal(pattern.test(invalid), false);
  }
});

test('repeat configuration preserves question IDs and requires all supported values', () => {
  const { ctx } = context();
  const items = [];
  let nextId = 1;
  const form = new Proxy({
    getItems: () => items.slice(),
    addTextItem() {
      const item = new Proxy({ id: nextId++, title: '', getTitle() { return this.title; }, getType: () => 'text',
        asTextItem() { return this; }, setTitle(title) { this.title = title; return this; },
        setRequired(required) { this.required = required; return this; }, setValidation(pattern) { this.pattern = pattern; return this; }
      }, { get(target, key) { return key in target ? target[key] : () => item; } });
      items.push(item); return item;
    }
  }, { get(target, key) { return key in target ? target[key] : () => form; } });
  const group = ctx.parkingReportGroups_()[0];
  const previous = { ...group, lots: group.lots.filter(lot => !['經貿6停車場', '經貿8停車場'].includes(lot.name)) };
  ctx.configureParkingReportForm_(form, previous);
  const previousIds = items.map(item => item.id);
  ctx.configureParkingReportForm_(form, group);
  assert.deepEqual(items.slice(0, previousIds.length).map(item => item.id), previousIds);
  assert.equal(items.length, previousIds.length + 4, 'only the four Jingmao fields are appended');
  const ids = items.map(item => item.id);
  ctx.configureParkingReportForm_(form, group);
  assert.deepEqual(items.map(item => item.id), ids);
  assert.equal(items.length, ctx.parkingReportFields_(group).length);
  assert.ok(items.every(item => item.required && new RegExp(item.pattern).test('0')));
  items.push({ getTitle: () => 'unknown', getType: () => 'text' });
  assert.throws(() => ctx.configureParkingReportForm_(form, group), /題目不符/);
});

test('routing preserves legacy before activation and selects each exact response source after', () => {
  const { ctx, properties } = context();
  const legacy = ctx.parkingRemainingFormula('legacy', { area: 'B', car: 'C', motorcycle: 'D' }, 'car', 'CITYPARKING春安站', 81);
  assert.ok(!legacy.includes('__NONE__'));
  const sheets = Array.from(ctx.parkingReportGroups_(), (group, index) => {
    properties['PARKING_REPORT_SHEET_' + group.key] = String(index + 1);
    return { getSheetId: () => index + 1, getName: () => "batch'" + index,
      getLastColumn: () => ctx.parkingReportFields_(group).length + 1,
      getRange: () => ({ getValues: () => [['Timestamp', ...Array.from(ctx.parkingReportFields_(group), field => field.title)]] }) };
  });
  ctx.SpreadsheetApp = { openById: () => ({ getSheets: () => sheets }) };
  properties.PARKING_REPORT_MODE = 'split';
  const formula = ctx.parkingRemainingFormula('legacy', { area: 'B', car: 'C', motorcycle: 'D' }, 'car', 'CITYPARKING春安站', 81);
  assert.ok(formula.includes("'batch''0'!A:A"));
  assert.ok(formula.includes(legacy.slice(1)));
  assert.ok(formula.includes('latest<=81'));
  assert.ok(formula.includes('latest=INT(latest)'));
  assert.ok(formula.includes('"未回報"'));
  assert.ok(!formula.includes('A2:A'));
  const dedicated = ctx.parkingRemainingFormula('legacy', { area: 'B', car: 'C', motorcycle: 'D' }, 'motorcycle', '經貿8停車場', 703);
  assert.ok(dedicated.includes("'batch''0'!"));
  assert.equal(ctx.parkingRemainingFormula('legacy', {}, 'car', '嶺東科大', 0), '="不提供"');
  assert.throws(() => ctx.updateParkingLotChoices(), /兩個回報入口/);
  assert.throws(() => ctx.upgradeParkingRemaining(), /兩個回報入口/);
});

test('missing or duplicate batch columns stop without guessing another form', () => {
  const { ctx, properties } = context();
  const groups = ctx.parkingReportGroups_();
  const sheets = Array.from(groups, (group, index) => {
    properties['PARKING_REPORT_SHEET_' + group.key] = String(index + 1);
    const titles = Array.from(ctx.parkingReportFields_(group), field => field.title);
    return { getSheetId: () => index + 1, getLastColumn: () => titles.length,
      getRange: () => ({ getValues: () => index ? [titles] : [titles.slice(1)] }) };
  });
  assert.throws(() => ctx.parkingReportSources_({ getSheets: () => sheets }), /欄位缺少或重複/);
});

test('legacy closure closes before setting the message and repeated closure performs no writes', () => {
  const { ctx } = context();
  let accepting = true, message = '';
  const writes = [];
  const form = {
    isAcceptingResponses: () => accepting,
    setAcceptingResponses(value) { accepting = value; writes.push('closed'); },
    getCustomClosedFormMessage: () => message,
    setCustomClosedFormMessage(value) { assert.equal(accepting, false); message = value; writes.push('message'); }
  };
  ctx.closeLegacyParkingForm_(form, 'shared-url');
  ctx.closeLegacyParkingForm_(form, 'shared-url');
  assert.deepEqual(writes, ['closed', 'message']);
  assert.match(message, /停車場回報請改用/);
});

test('closed legacy form with the correct redirect needs no risky message rewrite', () => {
  const { ctx } = context();
  ctx.closeLegacyParkingForm_({
    isAcceptingResponses: () => false,
    getCustomClosedFormMessage: () => '四個入口。其他停車場：shared-url',
    setCustomClosedFormMessage() { throw new Error('Invalid data updating form'); },
    setAcceptingResponses() { throw new Error('Must remain closed'); }
  }, 'shared-url');
});

test('unsupported custom closed messages do not prevent retiring an old form', () => {
  const { ctx } = context();
  let accepting = true;
  const warnings = [];
  ctx.Logger = { log: message => warnings.push(message) };
  ctx.closeLegacyParkingForm_({
    isAcceptingResponses: () => accepting,
    setAcceptingResponses: value => { accepting = value; },
    getCustomClosedFormMessage: () => '',
    setCustomClosedFormMessage() { throw new Error('Invalid data updating form'); }
  }, 'shared-url');
  assert.equal(accepting, false);
  assert.equal(warnings.length, 1);
});

test('merged Jingmao fields fall back through their original dedicated sheet', () => {
  const { ctx, properties } = context();
  properties.PARKING_REPORT_MODE = 'split';
  const sheets = Array.from(ctx.parkingReportGroups_(), (group, index) => {
    properties['PARKING_REPORT_SHEET_' + group.key] = String(index + 1);
    const headers = ['Timestamp', ...Array.from(ctx.parkingReportFields_(group), field => field.title)];
    return { getSheetId: () => index + 1, getName: () => 'active-' + index,
      getLastColumn: () => headers.length, getRange: () => ({ getValues: () => [headers] }) };
  });
  for (const [index, name] of [[1, '經貿6停車場'], [2, '經貿8停車場']]) {
    properties['PARKING_REPORT_SHEET_dedicated' + index] = String(100 + index);
    sheets.push({ getSheetId: () => 100 + index, getName: () => 'archive-' + index,
      getLastColumn: () => 3, getRange: () => ({ getValues: () => [['Timestamp', name + '｜目前剩餘汽車位', name + '｜目前剩餘機車位']] }) });
    ctx.SpreadsheetApp = { openById: () => ({ getSheets: () => sheets }) };
    const formula = ctx.parkingSplitFormula_(name, 'car', 600, '=17');
    assert.ok(formula.includes("'active-0'!"));
    assert.ok(formula.includes("'archive-" + index + "'!B:B"));
    assert.ok(formula.includes('<>""'), 'old shared rows without merged fields must not hide history');
  }
});
