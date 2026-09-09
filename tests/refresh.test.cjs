const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');

function page(href, fetchResult = async () => ({ ok: true })) {
  const nodes = new Map([...html.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)].map(m => [m[1], {
    textContent: m[2], style: {}, classList: { add() {}, remove() {}, toggle() {} }, remove() {}
  }]));
  const requests = [], navigations = [];
  const context = {
    window: { location: { href, replace: url => navigations.push(url), reload: () => navigations.push(href) } },
    document: {
      getElementById: id => nodes.get(id) || null,
      createElement: () => ({}), head: { appendChild() {} }, addEventListener() {}
    },
    fetch: async (url, options) => { requests.push({ url, options }); return fetchResult(); },
    URL, AbortController, setTimeout: () => 1, clearTimeout() {}, setInterval() {}, console
  };
  vm.runInNewContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], context);
  return { refresh: () => context.forceRefresh(), requests, navigations, nodes };
}

test('force refresh purges the fixed page before navigating and preserves theme and anchor', async () => {
  let release;
  const p = page('https://raw.githack.com/tim12332000/BusTrack/main/index.html?theme=nordic&_refresh=old#parkingSection',
    () => new Promise(resolve => { release = resolve; }));
  const pending = p.refresh();
  await p.refresh();
  assert.equal(p.requests.length, 1, 'repeated taps must not send duplicate purge requests');
  assert.equal(p.navigations.length, 0, 'wait for purge before navigating');
  assert.equal(p.requests[0].options.method, 'DELETE');
  assert.equal(new URL(p.requests[0].url, 'https://raw.githack.com').searchParams.get('url'),
    'https://raw.githack.com/tim12332000/BusTrack/main/index.html');
  release({ ok: true });
  await pending;
  const url = new URL(p.navigations[0]);
  assert.equal(url.searchParams.get('theme'), 'nordic');
  assert.match(url.searchParams.get('_refresh'), /^\d+$/);
  assert.equal(url.hash, '#parkingSection');
});

test('failed purge keeps the existing board visible and enables retry', async () => {
  for (const result of [async () => ({ ok: false }), async () => { throw new Error('offline'); }]) {
    const p = page('https://raw.githack.com/tim12332000/BusTrack/main/index.html', result);
    await p.refresh();
    assert.equal(p.navigations.length, 0);
    assert.equal(p.nodes.get('fabRefresh').disabled, false);
    assert.equal(p.nodes.get('liveText').textContent, '刷新失敗，請重試');
  }
});

test('local files reload without calling the CDN purge service', async () => {
  const p = page('file:///C:/Bus/index.html');
  await p.refresh();
  assert.equal(p.requests.length, 0);
  assert.deepEqual(p.navigations, ['file:///C:/Bus/index.html']);
});
