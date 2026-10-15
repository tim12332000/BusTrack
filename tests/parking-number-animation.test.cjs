const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function page() {
  function element() {
    return { textContent: '', children: [], dataset: {}, style: { setProperty() {} },
      classList: { toggle() {} }, setAttribute() {}, remove() {},
      append(...children) { this.children.push(...children); },
      appendChild(child) { this.children.push(child); },
      replaceChildren(...children) { this.children = children; } };
  }
  const html = fs.readFileSync(path.join(__dirname, '..', 'parking.html'), 'utf8');
  const nodes = new Map([...html.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)].map(([, id, value]) => [id, Object.assign(element(), { textContent: value })]));
  const frames = new Map();
  let time = 0, id = 0, reduced = false;
  const context = { window: {}, console, setInterval() {}, performance: { now: () => time },
    requestAnimationFrame: cb => { frames.set(++id, cb); return id; },
    cancelAnimationFrame: frame => frames.delete(frame), matchMedia: () => ({ matches: reduced }),
    document: { hidden: false, getElementById: key => nodes.get(key), createElement: element,
      head: { appendChild() {} }, addEventListener() {} } };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  return { context, frames, nodes,
    reduce() { reduced = true; },
    number: key => vm.runInContext(`parkingNumbers.get(${JSON.stringify(key)})`, context),
    tick(now) {
      time = now;
      const callbacks = [...frames.values()]; frames.clear();
      callbacks.forEach(cb => cb(now));
    },
    update(value) {
      context.window.onGvizData({ table: { rows: [{ c: [null, { v: 'CITYPARKING春安站' }, { v: value }, { v: 5 }] }] } });
      assert.equal(nodes.get('liveText').textContent, '', 'callback must finish without a swallowed exception');
    },
  };
}

test('parking totals, station sums and individual counts animate to the exact value', () => {
  const p = page(); p.update(40);
  const counters = [p.nodes.get('parkTotalCars'), p.number('group:4號門:cars'), p.number('lot:CITYPARKING春安站:cars')];
  p.tick(200);
  for (const counter of counters) assert.ok(Number(counter.textContent) > 0 && Number(counter.textContent) < 40);
  p.tick(800);
  for (const counter of counters) assert.equal(counter.textContent, '40');
  assert.equal(p.frames.size, 0);
});

test('parking refreshes reuse visible counters without restarting an identical animation', () => {
  const p = page(); p.update(40); p.tick(200);
  const counter = p.number('lot:CITYPARKING春安站:cars'), visible = counter.textContent;
  const frameIds = [...p.frames.keys()];
  p.update(40);
  assert.equal(p.number('lot:CITYPARKING春安站:cars'), counter);
  assert.equal(counter.textContent, visible);
  assert.deepEqual([...p.frames.keys()], frameIds);
  p.tick(800);
  assert.equal(counter.textContent, '40');
  assert.equal(p.frames.size, 0);
});

test('parking interrupted increases and decreases continue from the displayed value', () => {
  const p = page(); p.update(80); p.tick(200);
  const counter = p.number('lot:CITYPARKING春安站:cars'), visible = Number(counter.textContent);
  p.update(10);
  assert.equal(Number(counter.textContent), visible);
  p.tick(400);
  assert.ok(Number(counter.textContent) < visible && Number(counter.textContent) > 10);
  p.tick(1000);
  assert.equal(counter.textContent, '10');
  assert.equal(p.nodes.get('parkTotalCars').textContent, '10');
  assert.equal(p.frames.size, 0);
});

test('parking unknown counts cancel pending frames and later valid counts animate again', () => {
  const p = page(); p.update(40); p.tick(200); p.update('abc');
  const counter = p.number('lot:CITYPARKING春安站:cars');
  assert.equal(counter.textContent, '未回報');
  p.tick(1000);
  assert.equal(counter.textContent, '未回報');
  assert.equal(p.nodes.get('parkTotalCars').textContent, '0');
  p.update(40); p.tick(1200);
  assert.ok(Number(counter.textContent) > 0 && Number(counter.textContent) < 40);
  p.tick(1800); assert.equal(counter.textContent, '40');
});

test('parking respects reduced motion, hidden tabs and real zero', () => {
  for (const mode of ['reduced', 'hidden']) {
    const p = page(); p.update(40); p.tick(200);
    if (mode === 'reduced') p.reduce(); else p.context.document.hidden = true;
    p.update(12);
    assert.equal(p.nodes.get('parkTotalCars').textContent, '12');
    assert.equal(p.number('lot:CITYPARKING春安站:cars').textContent, '12');
    assert.equal(p.frames.size, 0);
    p.update(0);
    assert.equal(p.nodes.get('parkTotalCars').textContent, '0');
    assert.equal(p.number('lot:CITYPARKING春安站:cars').textContent, '0');
  }
});

test('parking repeated refreshes keep the counter registry bounded', () => {
  const p = page();
  for (let i = 0; i < 15; i++) { p.update(i); p.tick((i + 1) * 800); }
  assert.equal(vm.runInContext('parkingNumbers.size', p.context), 52);
  assert.equal(p.frames.size, 0);
});
