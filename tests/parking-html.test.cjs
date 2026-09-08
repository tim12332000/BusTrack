const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');

function page() {
  const nodes = new Map([...html.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)].map((m) => [m[1], {
    textContent: m[2], style: {}, classList: { add() {}, remove() {}, toggle() {} }, remove() {}
  }]));
  const context = { window: {}, document: {
    getElementById: (id) => nodes.get(id) || null,
    createElement: () => ({}), head: { appendChild() {} }, addEventListener() {}
  }, setInterval() {}, console };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  return { update: (rows) => context.window.onGvizData({ table: { rows } }), text: (id) => nodes.get(id).textContent };
}

function fixture({ legacy = false, missing = false, offset = 0 } = {}) {
  const rows = Array.from({ length: 40 + offset }, () => ({ c: [] }));
  const put = (r, c, v) => { rows[r].c[c] = { v }; };
  put(0, 0, 14); put(0, 8, 6);
  put(25 + offset, 0, legacy ? '汽車停車場' : '🅿️ 汽機車停車場剩餘車位');
  put(27 + offset, 0, 0); put(27 + offset, 8, missing ? '未完整回報' : 50);
  put(30 + offset, 0, legacy ? '各區停車場即時剩餘車位' : '各停車場即時剩餘車位（7處）');
  for (let i = 0; i < 7; i++) {
    put(33 + offset, i * 4, i === 3 ? '不提供' : 0);
    put(33 + offset, i * 4 + 2, i >= 5 ? '不提供' : missing && i === 4 ? '未回報' : 10);
  }
  return rows;
}

test('both HTML entry points remain identical and parking has no occupancy controls', () => {
  assert.equal(html, fs.readFileSync(path.join(__dirname, '../戰情大螢幕.html'), 'utf8'));
  assert.doesNotMatch(html, /id="park(?:TotalIn|TotalRate|MainBar|\d+(?:In|Remain|Bar|Status))"/);
});

test('new schema preserves zero remaining cars, independent motorcycle counts and people', () => {
  const p = page(); p.update(fixture({ offset: 3 }));
  assert.equal(p.text('parkTotalCars'), '0');
  assert.equal(p.text('parkTotalMotorcycles'), '50');
  assert.equal(p.text('park1Cars'), '0');
  assert.equal(p.text('park6Motorcycles'), '—');
  assert.equal(p.text('park4Cars'), '—');
  assert.equal(p.text('park7Cars'), '0');
  assert.equal(p.text('parkReported'), '7/7 處');
  assert.equal(p.text('busTotalIn'), '14');
  assert.equal(p.text('busTotalOut'), '6');
});

test('missing or invalid reports stay unknown and do not become zero or assumed capacity', () => {
  const p = page(); const rows = fixture({ missing: true });
  rows[33].c[4] = { v: '' }; rows[33].c[8] = { v: -1 };
  p.update(rows);
  assert.equal(p.text('park2Cars'), '未回報');
  assert.equal(p.text('park3Cars'), '未回報');
  assert.equal(p.text('park5Motorcycles'), '未回報');
  assert.equal(p.text('parkReported'), '4/7 處');
  assert.equal(p.text('parkTotalCars'), '未完整回報');
  assert.equal(p.text('parkTotalMotorcycles'), '未完整回報');
});

test('legacy cloud data clears previous parking values and asks for migration', () => {
  const p = page(); p.update(fixture()); p.update(fixture({ legacy: true }));
  assert.equal(p.text('park1Cars'), '未回報');
  assert.equal(p.text('parkTotalMotorcycles'), '未完整回報');
  assert.match(p.text('parkingNote'), /等待更新/);
  assert.equal(p.text('busTotalIn'), '14');
});

test('capacity stays secondary and never initializes current remaining availability', () => {
  const p = page();
  assert.equal(p.text('park1Cars'), '未回報');
  assert.equal(p.text('parkTotalCars'), '未完整回報');
  assert.match(html, /總容量 3,023 位/);
  assert.match(html, /總容量 2,773 位/);
  for (const name of ['五都日出', '新烏日', '*嶺東科大', '*台中科大', '水湳轉運站', '*經貿六', '*經貿八']) {
    assert.ok(html.includes('🅿️ ' + name + '</div>'));
  }
});

test('each vehicle total ignores unsupported slots but requires its applicable reports', () => {
  const p = page(); const rows = fixture();
  rows[33].c[18] = { v: '未回報' };
  p.update(rows);
  assert.equal(p.text('parkTotalCars'), '0');
  assert.equal(p.text('parkTotalMotorcycles'), '未完整回報');
  assert.equal(p.text('parkReported'), '6/7 處');
  rows[33].c[18] = { v: 0 };
  rows[27].c[8] = { v: 40 };
  p.update(rows);
  assert.equal(p.text('park5Motorcycles'), '0');
  assert.equal(p.text('parkReported'), '7/7 處');
});
