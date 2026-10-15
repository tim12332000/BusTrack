const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Execute the real parking callback and inspect individual rows as well as totals.
// These tests describe compatibility and fault isolation, without changing form policy.
function parkingPage() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'parking.html'), 'utf8');
  function element() {
    return { textContent: '', children: [], dataset: {}, className: '',
      style: { setProperty() {} }, classList: { toggle() {} },
      setAttribute() {}, remove() {},
      append(...children) { this.children.push(...children); },
      appendChild(child) { this.children.push(child); return child; },
      replaceChildren(...children) { this.children = children; } };
  }
  const nodes = new Map([...html.matchAll(/id="([^"]+)"/g)].map(([, id]) => [id, element()]));
  const warnings = [];
  const context = { window: {}, setInterval() {}, console: { warn: (...args) => warnings.push(args) },
    document: { getElementById: id => nodes.get(id), createElement: element,
      head: { appendChild() {} }, addEventListener() {} } };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  function descendants(node) {
    return [node, ...node.children.filter(child => typeof child === 'object').flatMap(descendants)];
  }
  return {
    warnings, nodes,
    update: rows => context.window.onGvizData({ table: { rows } }),
    total: type => nodes.get(type === 'cars' ? 'parkTotalCars' : 'parkTotalMotorcycles').textContent,
    count(name, type) {
      const row = descendants(nodes.get('parkingGroups')).find(node => node.dataset.lot === name);
      assert.ok(row, `lot still rendered: ${name}`);
      return row.children[type === 'cars' ? 1 : 2].children[0].textContent;
    },
    assertHealthy() {
      assert.equal(warnings.length, 0, 'no callback exception hidden by its catch block');
      assert.equal(nodes.get('liveText').textContent, '');
      for (const node of descendants(nodes.get('parkingGroups'))) {
        assert.doesNotMatch(String(node.textContent), /NaN|Infinity/);
        if (node.style.width !== undefined) {
          const width = Number.parseFloat(node.style.width);
          assert.ok(Number.isFinite(width) && width >= 0 && width <= 100, node.style.width);
        }
      }
    },
  };
}

const target = '水湳轉運站停車場'; // cars 614, motorcycles 1253
const anchor = 'CITYPARKING春安站'; // independent valid report: 8 cars, 5 motorcycles
const row = (name, cars, motorcycles) => ({ c: [null, { v: name }, { v: cars }, { v: motorcycles }] });
const invalidCases = [
  ['English', 'abc'], ['mixed number and English', '12abc'], ['decimal point', '.'],
  ['repeated decimal points', '..'], ['fraction string', '1.5'], ['fraction number', 1.5],
  ['negative integer', -1], ['negative fraction', '-0.5'], ['empty', ''], ['spaces', '  '],
  ['null', null], ['undefined', undefined], ['Chinese numeral', '十二'], ['fullwidth digits', '１２'],
  ['spreadsheet error', '#REF!'], ['formula text', '=1+1'], ['NaN number', NaN], ['NaN text', 'NaN'],
  ['positive infinity', Infinity], ['negative infinity', -Infinity], ['infinity text', 'Infinity'],
  ['unsafe integer', Number.MAX_SAFE_INTEGER + 1], ['extreme number', 1e308],
  ['boolean true', true], ['boolean false', false], ['object', {}], ['numeric array', [12]], ['empty array', []],
  ['repeated comma', '1,,2'], ['short comma group', '1,2'], ['leading comma', ',12'],
  ['trailing comma', '12,'], ['long comma group', '1,0000'], ['short final group', '1,200,3'],
  ['decimal with misplaced comma', '1,2.0'], ['comma only', ','],
];

for (const [type, capacity, anchorCount] of [['cars', 614, 8], ['motorcycles', 1253, 5]]) {
  const otherType = type === 'cars' ? 'motorcycles' : 'cars';
  const otherAnchor = type === 'cars' ? 5 : 8;
  const fixture = value => [row(anchor, 8, 5), row(target, type === 'cars' ? value : 7, type === 'motorcycles' ? value : 7)];
  for (const [label, value] of [...invalidCases, ['over capacity', capacity + 1]]) {
    test(`parking ${type}: ${label} stays local and next valid value recovers`, () => {
      const p = parkingPage();
      p.update(fixture(10));
      p.update(fixture(value));
      assert.equal(p.count(target, type), '未回報');
      assert.equal(p.total(type), String(anchorCount), 'bad value must not contribute or reuse a stale count');
      assert.equal(p.count(target, otherType), '7', 'other vehicle type keeps updating');
      assert.equal(p.total(otherType), String(7 + otherAnchor));
      p.assertHealthy();
      p.update(fixture(4));
      assert.equal(p.count(target, type), '4');
      assert.equal(p.total(type), String(4 + anchorCount));
      p.assertHealthy();
    });
  }
  const validCases = [0, '0', 12, '12', ' 12 ', '12.0', capacity, String(capacity)];
  if (type === 'motorcycles') validCases.push('1,200', ' 1,200 ', '1,200.0');
  for (const value of validCases) {
    test(`parking ${type}: compatible value ${JSON.stringify(value)} remains usable`, () => {
      const p = parkingPage();
      p.update(fixture(value));
      const count = Number(String(value).replace(/,/g, ''));
      assert.equal(p.count(target, type), count.toLocaleString('en-US'));
      assert.equal(p.total(type), (count + anchorCount).toLocaleString('en-US'));
      p.assertHealthy();
    });
  }
  // Ambiguous formats are observed here, not approved as new input rules.
  for (const [value, currentCount] of [['1e1', 10], ['0x10', 16]]) {
    test(`parking ${type}: characterize current conversion of ${JSON.stringify(value)} (policy deferred)`, () => {
      const p = parkingPage();
      p.update(fixture(value));
      assert.equal(p.count(target, type), String(currentCount));
      assert.equal(p.total(type), String(currentCount + anchorCount));
      p.assertHealthy();
    });
  }
}

test('parking numeric values do not create parking spaces for an unavailable vehicle type', () => {
  const p = parkingPage();
  p.update([row(anchor, 8, 5), row('嶺東科大', 999, 20)]);
  assert.equal(p.count('嶺東科大', 'cars'), '—');
  assert.equal(p.total('cars'), '8');
  assert.equal(p.total('motorcycles'), '25');
  p.assertHealthy();
});

test('parking missing cells and corrupt rows preserve valid numeric reports', () => {
  const p = parkingPage();
  p.update([null, {}, { c: 'bad' }, { c: [] }, { c: [null, { v: target }] }, row(anchor, 8, 5)]);
  assert.equal(p.count(target, 'cars'), '未回報');
  assert.equal(p.count(target, 'motorcycles'), '未回報');
  assert.equal(p.total('cars'), '8');
  assert.equal(p.total('motorcycles'), '5');
  p.assertHealthy();
});
