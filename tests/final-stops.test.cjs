const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const html = fs.readFileSync('index.html', 'utf8');
function page() {
  const nodes = new Map([...html.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)].map(m => [m[1], { textContent: m[2], style: {}, remove() {} }]));
  const ctx = { window: {}, document: { getElementById: id => nodes.get(id), createElement: () => ({}), head: { appendChild() {} }, addEventListener() {} }, setInterval() {}, console };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], ctx);
  return { text: id => nodes.get(id).textContent, update: rows => ctx.window.onGvizData({ table: { rows } }) };
}
function fixture(inbound, outbound, offset = 0) {
  const rows = Array.from({ length: 50 }, () => ({ c: [] }));
  const put = (r, c, v) => { rows[r].c[c] = { v }; };
  put(0, 0, 99999); put(0, 8, 99999); // Stale combined totals must not override current stops.
  for (let i = 0; i < 4; i++) { put(5, i * 6, 10); put(5, i * 6 + 3, 5); }
  put(14, 0, 100); put(14, 4, 30); put(14, 8, 200); put(14, 12, 40);
  put(14, 16, 99999); put(14, 20, 99999); // Historic gate 4.
  if (inbound !== undefined) {
    put(38 + offset, 0, '嶺東科大-寶文校區接駁統計');
    put(39 + offset, 0, inbound); put(39 + offset, 1, outbound);
  }
  return rows;
}
test('exact final seven names appear in the requested order', () => {
  const names = [...html.matchAll(/class="card-header [^"]+">[^ ]+ ([^<]+)<\/div>/g)].map(m => m[1]);
  assert.deepEqual(names, ['成功車站', '新烏日車站', '水湳轉運站', '經貿六停車場', '嶺東科大-寶文校區', '1號門', '3號門']);
  assert.doesNotMatch(html, /id="walk3|130輛/);
});
test('new stop is located semantically, included once, and gate 4 excluded', () => {
  const p = page(); p.update(fixture(123, 23, 4));
  assert.equal(p.text('bus5In'), '123');
  assert.equal(p.text('busTotalIn'), '163');
  assert.equal(p.text('busTotalOut'), '43');
  assert.equal(p.text('walkTotalIn'), '300');
  assert.equal(p.text('walkTotalOut'), '70');
});
test('absent or invalid new-stop data stays unknown; real zero counts', () => {
  const p = page();
  for (const value of [undefined, '', null, -1, '未回報', '#ERROR!', 1.5]) {
    p.update(fixture(value, value));
    assert.equal(p.text('bus5In'), '未回報');
    assert.equal(p.text('busTotalIn'), '未完整回報');
  }
  p.update(fixture(0, 0));
  assert.equal(p.text('bus5In'), '0');
  assert.equal(p.text('busTotalIn'), '40');
  p.update(fixture(undefined));
  assert.equal(p.text('bus5In'), '未回報');
});
