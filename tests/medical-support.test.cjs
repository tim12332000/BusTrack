const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function page() {
  const html = fs.readFileSync('index.html', 'utf8');
  const nodes = new Map([...html.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)].map(m => [m[1], { textContent: m[2], style: {}, hidden: true, classList: { toggle() {} }, remove() {} }]));
  const context = { window: {}, document: { getElementById: id => nodes.get(id), createElement: () => ({}), head: { appendChild() {} }, addEventListener() {} }, setInterval() {}, console };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  return { nodes, context, update: (a, b, url = '') => context.window.onMedicalData({ table: { rows: [{ c: [{ v: a }, { v: b }, null, { v: url }] }] } }) };
}

test('medical updates independently without changing attendance totals or rates', () => {
  const p = page();
  const rows = Array.from({ length: 40 }, () => ({ c: [] }));
  rows[38].c[0] = { v: '嶺東科大-寶文校區接駁統計' };
  rows[39].c = [{ v: 123 }, { v: 45 }];
  p.context.window.onGvizData({ table: { rows } });
  const totals = ['overallTotalIn', 'overallTotalOut', 'overallTotalRate', 'busTotalIn', 'walkTotalIn'];
  const before = totals.map(id => p.nodes.get(id).textContent);
  p.update(900, 800);
  assert.equal(p.nodes.get('medicalAssistance').textContent, '900');
  assert.equal(p.nodes.get('medicalEvacuation').textContent, '800');
  p.context.window.onGvizData({ table: { rows } });
  assert.deepEqual(totals.map(id => p.nodes.get(id).textContent), before);
  p.update(12, 3);
  assert.equal(p.nodes.get('medicalAssistance').textContent, '12');
  assert.deepEqual(totals.map(id => p.nodes.get(id).textContent), before);
});

test('medical preserves zero and rejects missing, negative, fractional and error values', () => {
  const p = page();
  for (const invalid of [null, undefined, '', ' ', -1, 1.5, '#REF!', '未回報']) {
    p.update(invalid, invalid);
    assert.equal(p.nodes.get('medicalAssistance').textContent, '未回報');
  }
  p.update(0, 0);
  assert.equal(p.nodes.get('medicalAssistance').textContent, '0');
  p.context.window.onMedicalData({ status: 'error' });
  assert.match(p.nodes.get('medicalStatus').textContent, /讀取失敗/);
  assert.equal(p.nodes.get('medicalAssistance').textContent, '0');
});

test('medical form link only accepts a Google Forms responder URL', () => {
  const p = page(), link = p.nodes.get('medicalFormLink');
  p.update(1, 0, 'javascript:alert(1)');
  assert.equal(link.href, undefined);
  const url = 'https://docs.google.com/forms/d/e/medical-test/viewform';
  p.update(1, 0, url);
  assert.equal(link.href, url);
  assert.equal(link.hidden, false);
});

test('medical formulas select the latest complete valid pair and escape sheet names', () => {
  const context = {};
  vm.runInNewContext(fs.readFileSync('apps-script/medical-support.js', 'utf8'), context);
  const formula = context.medicalLatestFormula("Medical's responses", 2);
  assert.match(formula, /'Medical''s responses'!A:C/);
  assert.match(formula, /ROW\('Medical''s responses'!A:A\)>1/);
  assert.doesNotMatch(formula, /![A-C]2:/);
  assert.match(formula, /ISNUMBER\(INDEX\(data,,2\)\)/);
  assert.match(formula, /MOD\(INDEX\(data,,3\),1\)=0/);
  assert.match(formula, /INDEX\(valid,ROWS\(valid\),2\)/);
  assert.doesNotMatch(formula, /SUM|總即時戰情看板|Form Responses 1/);
  assert.equal(fs.readFileSync('index.html', 'utf8'), fs.readFileSync('戰情大螢幕.html', 'utf8'));
});

test('numbered medical totals are independent by category and use latest values for numbers 1–15', () => {
  const context = {};
  vm.runInNewContext(fs.readFileSync('apps-script/medical-support.js', 'utf8'), context);
  const formula = context.medicalNumberedTotalFormula("Medical's responses", { category: 'D', number: 'E', quantity: 'F' }, '醫療協處');
  assert.match(formula, /'Medical''s responses'!D:D="醫療協處"/);
  assert.match(formula, /SEQUENCE\(15\)/);
  assert.match(formula, /XLOOKUP\(n&"號",ids,counts,0,0,-1\)/);
  assert.match(formula, /SUM\(MAP/);
  assert.doesNotMatch(formula, /![A-F]2:/);
  assert.doesNotMatch(formula, /總即時戰情看板|Form Responses 1/);
  const seeded = context.medicalNumberedTotalFormula('responses', { category: 'D', number: 'E', quantity: 'F' }, '醫療後送', 1, 11);
  assert.match(seeded, /XLOOKUP\(n&"號",ids,counts,IF\(n=1,11,0\),0,-1\)/);
});

test('invalid legacy assignments stop before any external modifications', () => {
  const context = {};
  vm.runInNewContext(fs.readFileSync('apps-script/medical-support.js', 'utf8'), context);
  for (const number of [-1, 16, 1.5, '1', null]) assert.throws(() => context.upgradeMedicalNumbered(number), /1–15/);
});
