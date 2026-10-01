"""停車場頁：23 處都要在一個畫面內看完，且各站總數在各種回報狀況下都正確。"""
from pathlib import Path
import json
from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
(root/'.omx').mkdir(exist_ok=True)
# 畫面順序：新烏日站（左欄）、4號門、水湳轉運站。
CAPACITY = ['4,936', '1,447', '140', '874', '3,229', '3,372']
results = []
with sync_playwright() as p:
    for engine in [p.chromium, p.webkit]:
        browser = engine.launch()
        page = browser.new_page(viewport={'width': 874, 'height': 402}, device_scale_factor=1)
        page.route('https://**', lambda route: route.abort())
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto((root/'parking.html').as_uri())
        for width, height in [(1920,1080),(1366,768),(1280,720),(956,440),(874,402),(844,320),(760,300),(667,375),(390,844),(375,667)]:
            page.set_viewport_size({'width': width, 'height': height})
            for mode in ['capacity', 'zero', 'missing', 'invalid', 'partial']:
                page.evaluate('''mode => {
                  const value = (lot, type, i) => mode === 'capacity' ? lot[type] : mode === 'zero' ? 0
                    : mode === 'invalid' ? (type === 'cars' ? lot[type] + 1 : -1)
                    : mode === 'partial' ? (i === 4 ? '' : lot[type]) : '未回報';
                  const rows = PARKING_LOTS.map((lot, i) => ({c: [{v: lot.group}, {v: lot.name},
                    {v: value(lot, 'cars', i)}, {v: value(lot, 'motorcycles', i)}]})).reverse();
                  window.onGvizData({table: {rows}}); window.scrollTo(0, 0);
                }''', mode)
                state = page.evaluate('''() => ({
                  rows: document.querySelectorAll('.lot-row[data-lot]').length,
                  capacity: [...document.querySelectorAll('.group-cap')].map(e => e.textContent.replace(/[^\\d,]/g, '')),
                  remaining: [...document.querySelectorAll('.group-metric b')].map(e => e.textContent),
                  pills: [...document.querySelectorAll('.group-metric .status-pill')].map(e => e.textContent),
                  missing: [...document.querySelectorAll('.group-name small')].map(e => e.textContent),
                  bars: document.querySelectorAll('.group-metric .capacity-bar .bar-fill').length,
                  total: document.getElementById('parkTotalCars').textContent
                })''')
                assert state['rows'] == 23, state
                assert state['capacity'] == CAPACITY, state
                assert state['bars'] == 6, state
                if mode == 'capacity':
                    assert state['remaining'] == CAPACITY and state['pills'] == ['充足'] * 6, state
                if mode == 'zero':
                    assert state['remaining'] == ['0'] * 6 and state['pills'] == ['已客滿'] * 6, state
                if mode in ['missing', 'invalid']:
                    assert state['remaining'] == ['未回報'] * 6 and state['pills'] == ['待回報'] * 6, state
                if mode == 'partial':
                    # 新烏日的「嘟嘟房高鐵臺中站」沒回報：該站仍顯示其餘場的加總，並標出 1 處未回報。
                    assert state['remaining'][0:2] == ['3,399', '711'], state
                    assert state['missing'][0] == '14 處 · 1 處未回報', state
                    assert state['total'] == '6,768', state
                result = page.evaluate('''() => ({
                  height: innerHeight,
                  scrollHeight: document.scrollingElement.scrollHeight,
                  bodyScroll: document.body.scrollHeight,
                  bottom: Math.max(...[...document.querySelectorAll('.lot-row')].map(e => e.getBoundingClientRect().bottom)),
                  overflow: document.scrollingElement.scrollWidth > innerWidth,
                  clipped: [...document.querySelectorAll('.lot-name,.lot-num,.group-name,.group-metric,.metric-value,.metric-capacity,h1')]
                    .filter(e => e.clientWidth > 0 && (e.scrollWidth > e.clientWidth + 1 || e.scrollHeight > e.clientHeight + 1)).map(e => e.textContent)
                })''')
                result.update(engine=engine.name, width=width, mode=mode)
                results.append(result)
                if mode == 'partial':
                    page.screenshot(path=str(root/'.omx'/f'parking-table-{engine.name}-{width}x{height}.png'))
                assert not errors, errors
                assert result['bottom'] <= height + 1 and result['scrollHeight'] <= height + 1 and result['bodyScroll'] <= height + 1, result
                assert not result['overflow'] and not result['clipped'], result
        browser.close()
(root/'.omx/parking-table-results.json').write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding='utf-8')
print(f'PASS: {len(results)} Chromium/WebKit cases; all 23 rows fit on one screen with capacity/zero/missing/invalid/partial data')
