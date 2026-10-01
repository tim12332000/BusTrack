from pathlib import Path
import json
from playwright.sync_api import sync_playwright

root=Path(__file__).resolve().parents[1]
results=[]
with sync_playwright() as p:
    for engine in [p.chromium,p.webkit]:
        browser=engine.launch()
        page=browser.new_page(viewport={'width':874,'height':402},device_scale_factor=1)
        page.route('https://**',lambda route:route.abort())
        errors=[]
        page.on('pageerror',lambda error:errors.append(str(error)))
        page.goto((root/'parking.html').as_uri())
        assert page.locator('.lot-row').count()==23
        assert all('未完整回報' in t for t in page.locator('.remaining').all_inner_texts())
        for width,height in [(1920,1080),(1280,720),(956,440),(874,402),(844,320),(760,300),(667,375),(390,844),(375,667)]:
            page.set_viewport_size({'width':width,'height':height})
            for mode in ['capacity','zero','missing','invalid']:
                page.evaluate('''mode => {
                  const rows=PARKING_LOTS.map(lot=>({c:[{v:lot.group},{v:lot.name},
                    {v:mode==='capacity'?lot.cars:mode==='zero'?0:mode==='invalid'?99999:'未回報'},
                    {v:mode==='capacity'?lot.motorcycles:mode==='zero'?0:mode==='invalid'?-1:'未回報'}]})).reverse();
                  window.onGvizData({table:{rows}}); window.scrollTo(0,0);
                }''',mode)
                assert page.locator('.lot-row').count()==23
                assert page.locator('.area-total').all_inner_texts()==['140','874','4,936','1,447','3,229','3,372']
                remaining=page.locator('.remaining').all_inner_texts()
                if mode=='zero': assert remaining==['剩餘 0']*6,remaining
                if mode in ['missing','invalid']: assert remaining==['剩餘未完整回報']*6,remaining
                if mode=='capacity': assert remaining==['剩餘 140','剩餘 874','剩餘 4,936','剩餘 1,447','剩餘 3,229','剩餘 3,372'],remaining
                result=page.evaluate('''() => {
                  const nodes=[...document.querySelectorAll('header,.parking-overview,.area-summary,.lot-row,.parking-legend')];
                  return {bottom:Math.max(...nodes.map(e=>e.getBoundingClientRect().bottom)),height:innerHeight,
                    hidden:document.querySelectorAll('.lot-row').length!==23||[...document.querySelectorAll('.lot-row')].some(e=>!e.getBoundingClientRect().height||!e.getBoundingClientRect().width),
                    overflow:document.scrollingElement.scrollWidth>innerWidth,
                    clipped:[...document.querySelectorAll('.lot-table th,.lot-table td,.area-total,h1')].filter(e=>e.scrollWidth>e.clientWidth+1).map(e=>e.textContent)};
                }''')
                result.update(engine=engine.name,width=width,mode=mode)
                results.append(result)
                if mode=='capacity':page.screenshot(path=str(root/'.omx'/f'parking-table-{engine.name}-{width}.png'))
                assert not errors,errors
                assert not result['hidden'] and not result['overflow'] and not result['clipped'],result
                assert result['bottom']<=height+1,result
        browser.close()
(root/'.omx/parking-table-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print(f'PASS: {len(results)} Chromium/WebKit cases; all 23 rows visible, capacity/zero/missing/invalid data checked')
