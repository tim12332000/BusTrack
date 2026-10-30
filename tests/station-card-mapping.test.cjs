const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Sheet columns are the existing backend contract; card IDs are deliberately
// discovered by visible names so a renamed/reordered DOM cannot mask a swap.
const stations = [
  { name: '成功車站', group: 'bus', column: 0, inbound: 101, outbound: 11 },
  { name: '新烏日車站', group: 'bus', column: 6, inbound: 202, outbound: 22 },
  { name: '經貿六停車場', group: 'bus', column: 12, inbound: 303, outbound: 33 },
  { name: '水湳轉運站', group: 'bus', column: 18, inbound: 404, outbound: 44 },
  { name: '嶺東科大-寶文校區', group: 'lingdong', column: 0, inbound: 505, outbound: 55 },
  { name: '1號門', group: 'walk', column: 0, inbound: 606, outbound: 66 },
  { name: '3號門', group: 'walk', column: 8, inbound: 707, outbound: 77 },
  { name: '復康巴士', group: 'walk', column: 16, inbound: 808, outbound: 88 }
];

function dashboard(file) {
  const html = fs.readFileSync(path.join(process.env.TEST_DASHBOARD_SOURCE_DIR || '.', file), 'utf8');
  const nodes = new Map([...html.matchAll(/id="([^"]+)"/g)].map(match => [match[1], {
    textContent: '', style: {}, classList: { toggle() {} }, remove() {}
  }]));
  const errors = [];
  const ctx = { window: {}, document: {
    getElementById: id => nodes.get(id), createElement: () => ({}),
    head: { appendChild() {} }, addEventListener() {}
  }, setInterval() {}, console: { warn: (...args) => errors.push(args) } };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], ctx);
  const card = station => {
    const cards = html.split(/<div class="station-card(?:\s[^"]*)?">/).slice(1);
    const matches = cards.filter(part => part.match(/class="card-header[^>]*>([\s\S]*?)<\/div>/)?.[1].includes(station.name));
    assert.equal(matches.length, 1, `${station.name}: one visible card`);
    const ids = [...matches[0].matchAll(/class="card-col-num" id="([^"]+)"/g)].map(match => match[1]);
    assert.equal(ids.length, 2, `${station.name}: inbound and outbound fields`);
    return ids.map(id => nodes.get(id).textContent);
  };
  return { card, text: id => nodes.get(id).textContent, update(rows) {
    ctx.window.onGvizData({ status: 'ok', table: { rows } });
    assert.deepEqual(errors, [], 'rendering must not silently fail');
  } };
}

function fixture(values = stations, { offset = 0, textNumbers = false, sparse = false } = {}) {
  const rows = Array.from({ length: 45 + offset }, () => ({ c: [] }));
  const put = (row, column, value) => { rows[row + offset].c[column] = { v: value }; };
  put(6, 0, '接駁車站即時數據');
  put(7, 0, '🚌 成功車站');
  put(7, 6, '🚌 新烏日車站');
  put(7, 12, '🚌 經貿六停車場');
  put(7, 18, '🚌 水湳轉運站');
  put(8, 0, '進場');
  put(12, 0, '🚶 步行通道疏運概況');
  put(13, 0, '進場總人數');
  put(14, 0, 999999); // Stale combined totals must never replace station sums.
  put(17, 0, '步行通道即時數據');
  put(18, 0, '🚶 1號門');
  put(18, 8, '🚶 3號門');
  put(18, 16, '♿ 復康巴士');
  put(19, 0, '進場');
  put(38, 0, '嶺東科大-寶文校區接駁統計');
  for (const station of values) {
    const row = station.group === 'bus' ? 9 : station.group === 'walk' ? 20 : 39;
    const outColumn = station.column + (station.group === 'bus' ? 3 : station.group === 'walk' ? 4 : 1);
    const format = n => textNumbers ? String(n) : n;
    put(row, station.column, format(station.inbound));
    put(row, outColumn, format(station.outbound));
  }
  // Model GViz omitting blank rows without removing the three-row station blocks.
  return sparse ? rows.filter((row, index) => row.c.length || [8, 9, 19, 20].includes(index - offset)) : rows;
}

function assertCards(page, expected) {
  for (const station of expected) {
    assert.deepEqual(page.card(station), [String(station.inbound), String(station.outbound)], station.name);
  }
}

function assertTotals(page, expected) {
  const sum = (group, direction) => expected.filter(s => group === 'all' || (group === 'bus' ? s.group !== 'walk' : s.group === 'walk')).reduce((n, s) => n + s[direction], 0);
  for (const [prefix, group] of [['bus', 'bus'], ['walk', 'walk'], ['overall', 'all']]) {
    assert.equal(page.text(`${prefix}TotalIn`), sum(group, 'inbound').toLocaleString('en-US'));
    assert.equal(page.text(`${prefix}TotalOut`), sum(group, 'outbound').toLocaleString('en-US'));
  }
}

for (const file of ['index.html', '戰情大螢幕.html']) {
  for (const options of [{}, { offset: 4 }, { sparse: true }, { textNumbers: true }]) {
    test(`${file}: all eight named cards and totals match sheet data ${JSON.stringify(options)}`, () => {
      const page = dashboard(file);
      page.update(fixture(stations, options));
      assertCards(page, stations);
      assertTotals(page, stations);
    });
  }
  for (const changed of stations) {
    test(`${file}: refreshing ${changed.name} changes only its own inbound/outbound card`, () => {
      const page = dashboard(file);
      page.update(fixture());
      const next = stations.map(station => station === changed ? { ...station, inbound: 919, outbound: 19 } : station);
      page.update(fixture(next));
      assertCards(page, next);
      assertTotals(page, next);
      const zero = next.map(station => station.name === changed.name ? { ...station, inbound: 0, outbound: 0 } : station);
      page.update(fixture(zero));
      assertCards(page, zero);
      assertTotals(page, zero);
    });
  }
}
