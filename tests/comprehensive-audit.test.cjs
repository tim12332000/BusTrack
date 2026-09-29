const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

test('shuttle and parking have independent valid entry pages', () => {
  const shuttle = fs.readFileSync('index.html', 'utf8');
  const parking = fs.readFileSync('parking.html', 'utf8');
  assert.equal(shuttle, fs.readFileSync('戰情大螢幕.html', 'utf8'));
  assert.ok(shuttle.includes('href="parking.html"'));
  assert.ok(parking.includes('href="index.html"'));
  assert.ok(!shuttle.includes('id="parkingGroups"'));
  assert.ok(parking.includes('id="parkingGroups"'));
  for (const page of [shuttle, parking]) {
    for (const tag of ['html', 'head', 'body', 'script']) {
      assert.equal((page.match(new RegExp(`<${tag}\\b`, 'g')) || []).length, (page.match(new RegExp(`</${tag}>`, 'g')) || []).length);
    }
    new vm.Script(page.match(/<script>([\s\S]*?)<\/script>/)[1]);
  }
});
