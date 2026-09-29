const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');

const html = fs.readFileSync('parking.html', 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];

test('parking page is independent and requests the normalized parking sheet', () => {
  assert.ok(html.includes('href="index.html"'));
  assert.ok(html.includes("'停車場現況'"));
  assert.ok(!html.includes('id="busTotalIn"'));
  assert.match(script, /sheetName = document\.getElementById\('parkingGroups'\) \? '停車場現況'/);
});

test('photo parking list contains every named site and only listed capacity totals', () => {
  const list = [...script.matchAll(/"group": "([^"]+)", "name": "([^"]+)", "cars": (\d+), "motorcycles": (\d+)/g)]
    .map(([, group, name, cars, motorcycles]) => ({ group, name, cars: Number(cars), motorcycles: Number(motorcycles) }));
  assert.equal(list.length, 23);
  assert.deepEqual(list.slice(0, 3).map(lot => lot.name), ['CITYPARKING春安站', '協弘停車場', '嶺東科大']);
  assert.deepEqual(list.slice(-6).map(lot => lot.name), ['水湳轉運站停車場', '經貿6停車場', '經貿8停車場', '中央公園北側停車場', '臺中國際會展停車場', '綠美圖停車場']);
  assert.deepEqual(list.reduce((total, lot) => [total[0] + lot.cars, total[1] + lot.motorcycles], [0, 0]), [8305, 5693]);
  assert.ok(!html.includes('總容量 5,000'));
  assert.ok(!html.includes('總容量 1,400'));
});

test('unknown reports remain unknown and empty photo capacities display as unlisted', () => {
  assert.match(script, /capacity === 0 \? '—' : value === null \? '未回報'/);
  assert.match(script, /capacity \? `容量 \$\{formatNumber\(capacity\)\}` : '未列容量'/);
  assert.match(script, /isComplete \? formatNumber\(amount\) : '未完整回報'/);
});
