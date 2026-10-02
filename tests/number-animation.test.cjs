const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function animationPage() {
  const html = fs.readFileSync('index.html', 'utf8');
  const nodes = new Map([...html.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)]
    .map(m => [m[1], { textContent: m[2], style: {}, classList: { toggle() {}, remove() {} } }]));
  const frames = new Map();
  let time = 0, id = 0, reduced = false;
  const context = {
    window: {}, console, setInterval() {},
    document: { hidden: false, getElementById: key => nodes.get(key),
      createElement: () => ({}), head: { appendChild() {} }, addEventListener() {} },
    performance: { now: () => time },
    matchMedia: () => ({ matches: reduced }),
    requestAnimationFrame: callback => { frames.set(++id, callback); return id; },
    cancelAnimationFrame: key => frames.delete(key)
  };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  const element = nodes.get('overallTotalIn');
  element.textContent = '1';
  return { element, frames, context,
    set: text => context.setNumberText(element, text),
    reduce: () => { reduced = true; },
    tick: now => {
      time = now;
      const callbacks = [...frames.values()]; frames.clear();
      callbacks.forEach(callback => callback(now));
    } };
}

test('counts show intermediate values and land on the exact target', () => {
  const p = animationPage(); p.set('100');
  assert.equal(p.element.textContent, '1');
  p.tick(200);
  assert.ok(Number(p.element.textContent) > 1 && Number(p.element.textContent) < 100);
  p.tick(800);
  assert.equal(p.element.textContent, '100');
  assert.equal(p.frames.size, 0);
});

test('an interrupted count continues from its visible value and supports decreases', () => {
  const p = animationPage(); p.set('1,000'); p.tick(200);
  const visible = p.element.textContent;
  p.set('10');
  assert.equal(p.element.textContent, visible);
  assert.equal(p.frames.size, 1);
  p.tick(600);
  assert.ok(Number(p.element.textContent) < Number(visible.replace(/,/g, '')));
  p.tick(1000); assert.equal(p.element.textContent, '10');
});

test('unknown states cancel animation and reduced motion applies targets immediately', () => {
  const p = animationPage(); p.set('100'); p.tick(100);
  p.set('未回報'); p.tick(800);
  assert.equal(p.element.textContent, '未回報');
  assert.equal(p.frames.size, 0);
  p.set('100'); assert.equal(p.element.textContent, '100');
  p.reduce(); p.set('200');
  assert.equal(p.element.textContent, '200');
  assert.equal(p.frames.size, 0);
});

test('percentage decimals animate and identical refreshes do not restart the counter', () => {
  const p = animationPage(); p.element.textContent = '1.0%';
  p.set('55.5%'); p.tick(200); p.set('55.5%');
  assert.equal(p.frames.size, 1);
  p.tick(800); assert.equal(p.element.textContent, '55.5%');
});
