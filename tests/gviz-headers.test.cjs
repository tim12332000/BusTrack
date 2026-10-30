const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

for (const file of ['index.html', '戰情大螢幕.html']) {
  test(`${file}: dashboard request disables inferred headers to preserve station markers`, () => {
    const html = fs.readFileSync(file, 'utf8');
    const requests = [];
    const ctx = {
      window: {}, document: {
        getElementById: () => ({ textContent: '', remove() {} }), createElement: () => ({}),
        head: { appendChild: script => requests.push(script.src) }, addEventListener() {}
      }, setInterval() {}, console
    };
    vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], ctx);
    const request = requests.map(url => new URL(url)).find(url => url.searchParams.get('sheet') === '總即時戰情看板');
    assert.ok(request, 'dashboard request must be sent');
    assert.equal(request.searchParams.get('headers'), '0');
  });
}
