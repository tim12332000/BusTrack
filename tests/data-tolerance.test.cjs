const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function page(file) {
  const html = fs.readFileSync(file, 'utf8');
  function element() {
    return { textContent: '', children: [], dataset: {}, style: { setProperty() {} },
      classList: { toggle() {} }, remove() {}, setAttribute() {},
      append(...nodes) { this.children.push(...nodes); },
      appendChild(node) { this.children.push(node); return node; },
      replaceChildren(...nodes) { this.children = nodes; } };
  }
  const nodes = new Map([...html.matchAll(/id="([^"]+)"/g)].map(m => [m[1], element()]));
  const warnings = [];
  const context = { window: {}, document: { getElementById: id => nodes.get(id), createElement: element,
    head: { appendChild() {} }, addEventListener() {} }, setInterval() {}, console: { warn: (...args) => warnings.push(args) } };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  return { nodes, warnings, context, text: id => nodes.get(id).textContent,
    update: rows => context.window.onGvizData({ table: { rows } }) };
}

function peopleRows(value) {
  const rows = Array.from({ length: 40 }, () => ({ c: [] }));
  rows[5].c[0] = { v: value };
  rows[5].c[6] = { v: 20 };
  rows[14].c[0] = { v: 30 };
  rows[38].c[0] = { v: '嶺東科大-寶文校區接駁統計' };
  rows[39].c = [{ v: 10 }, { v: 0 }];
  return rows;
}

for (const file of ['index.html', '戰情大螢幕.html']) {
  test(`${file}: totals remain numeric even when no station has a usable count`, () => {
    const p = page(file), rows = peopleRows('abc');
    for (const column of [0, 3, 6, 9, 12, 15, 18, 21]) rows[5].c[column] = { v: 'abc' };
    for (const column of [0, 4, 8, 12]) rows[14].c[column] = { v: 'abc' };
    rows[39].c = [{ v: '未回報' }, { v: '未回報' }];
    p.update(rows);
    for (const id of ['busTotalIn', 'busTotalOut', 'walkTotalIn', 'walkTotalOut', 'overallTotalIn', 'overallTotalOut']) {
      assert.equal(p.text(id), '0', id);
    }
    for (const id of ['busTotalRate', 'walkTotalRate', 'overallTotalRate']) assert.equal(p.text(id), '0.0%', id);
    assert.equal(p.text('bus1In'), '未回報');
    assert.equal(p.text('bus5In'), '未回報');
    assert.equal(p.text('overallNote'), '步行＋接駁車');
  });
  test(`${file}: invalid count is isolated, other totals survive, next report recovers`, () => {
    const p = page(file);
    for (const value of ['abc', '.', '12abc', '#REF!', Infinity, NaN, 1e308, {}, true]) {
      p.update(peopleRows(value));
      assert.equal(p.text('bus1In'), '未回報', String(value));
      assert.equal(p.text('bus2In'), '20');
      assert.equal(p.text('overallTotalIn'), '60');
      assert.equal(p.text('busTotalIn'), '30');
      assert.equal(p.text('busTotalRate'), '0.0%');
      assert.equal(p.text('overallTotalRate'), '0.0%');
      assert.equal(p.text('overallNote'), '步行＋接駁車');
      assert.doesNotMatch(p.text('liveText'), /更新失敗/);
      assert.equal(p.warnings.length, 0);
    }
    p.update(peopleRows(5));
    assert.equal(p.text('bus1In'), '5');
    assert.equal(p.text('overallTotalIn'), '65');
    assert.doesNotMatch(p.text('overallNote'), /異常/);
  });
  test(`${file}: walking and accessible errors stay local`, () => {
    const p = page(file), rows = peopleRows(5);
    rows[12].c[16] = { v: '♿ 復康巴士' };
    rows[14].c[0] = { v: 'abc' };
    rows[14].c[8] = { v: 40 };
    rows[14].c[16] = { v: Infinity };
    p.update(rows);
    assert.equal(p.text('walk1In'), '未回報');
    assert.equal(p.text('walk2In'), '40');
    assert.equal(p.text('accessibleIn'), '未回報');
    assert.equal(p.text('walkTotalIn'), '40');
    assert.equal(p.text('walkTotalRate'), '0.0%');
    assert.equal(p.text('overallTotalIn'), '75');
    assert.equal(p.text('overallTotalRate'), '0.0%');
    assert.equal(p.text('overallNote'), '步行＋接駁車');
    assert.equal(p.warnings.length, 0);
    rows[5] = null;
    assert.doesNotThrow(() => p.update(rows));
    assert.equal(p.text('bus1In'), '未回報');
    assert.equal(p.text('walk2In'), '40');
    assert.equal(p.warnings.length, 0);
  });
  test(`${file}: decimal and formatted numeric values still work`, () => {
    const p = page(file);
    for (const [value, expected] of [[1.5, '1.5'], ['1,200', '1,200'], [0, '0']]) {
      p.update(peopleRows(value));
      assert.equal(p.text('bus1In'), expected);
    }
  });
}

test('parking: corrupt rows and bad counts do not suppress valid lots or prevent recovery', () => {
  const p = page('parking.html');
  const row = (name, value) => ({ c: [null, { v: name }, { v: value }, { v: 0 }] });
  for (const value of ['abc', '.', '1.5', '#REF!', Infinity, NaN, {}, true, [12]]) {
    p.update([null, {}, { c: 'bad' }, row('CITYPARKING春安站', value), row('協弘停車場', 12)]);
    assert.equal(p.text('parkTotalCars'), '12');
    assert.doesNotMatch(p.text('parkingNote'), /未回報/);
    assert.doesNotMatch(p.text('liveText'), /更新失敗/);
    assert.equal(p.warnings.length, 0);
  }
  p.update([row('CITYPARKING春安站', 8), row('協弘停車場', 12)]);
  assert.equal(p.text('parkTotalCars'), '20');
});

for (const file of ['index.html', '戰情大螢幕.html', 'parking.html']) {
  test(`${file}: broken payload preserves prior display and valid refresh recovers`, () => {
    const p = page(file), parking = file === 'parking.html';
    const rows = parking ? [{ c: [null, { v: '協弘停車場' }, { v: 12 }] }] : peopleRows(5);
    const id = parking ? 'parkTotalCars' : 'overallTotalIn';
    p.update(rows);
    const before = p.text(id);
    for (const payload of [null, {}, { status: 'error' }, { table: { rows: [] } }, { table: { rows: {} } }, { table: { rows: 'bad' } }]) {
      assert.doesNotThrow(() => p.context.window.onGvizData(payload));
      assert.equal(p.text(id), before);
      assert.match(p.text('liveText'), /讀取失敗/);
      assert.equal(p.warnings.length, 0);
    }
    p.update(rows);
    assert.equal(p.text('liveText'), '');
  });
}

test('medical: malformed payload preserves prior counts and the next valid response recovers', () => {
  const p = page('index.html');
  const update = payload => p.context.window.onMedicalData(payload);
  const valid = { table: { rows: [{ c: [{ v: 12 }, { v: 3 }] }] } };
  update(valid);
  for (const payload of [null, {}, { table: { rows: {} } }, { table: { rows: [null] } }, { table: { rows: [{ c: 'bad' }] } }]) {
    assert.doesNotThrow(() => update(payload));
    assert.equal(p.text('medicalAssistance'), '12');
    assert.match(p.text('medicalStatus'), /讀取失敗/);
  }
  update({ table: { rows: [{ c: [{ v: true }, { v: 3 }] }] } });
  assert.equal(p.text('medicalAssistance'), '未回報');
  assert.equal(p.text('medicalEvacuation'), '3');
  update(valid);
  assert.equal(p.text('medicalAssistance'), '12');
  assert.doesNotMatch(p.text('medicalStatus'), /讀取失敗/);
});
