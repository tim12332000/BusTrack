"""The entire shuttle board must fit before any scrolling, including pending data."""
from pathlib import Path
import json
from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
results = []
with sync_playwright() as p:
    for engine in [p.chromium, p.webkit]:
        browser = engine.launch()
        page = browser.new_page(viewport={'width': 402, 'height': 874}, device_scale_factor=3,
                                is_mobile=True, has_touch=True, reduced_motion='reduce')
        page.route('https://**', lambda route: route.abort())
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto((root/'index.html').as_uri())
        for width, height in [(874,402),(956,440),(844,320),(760,300),(667,320)]:
            page.set_viewport_size({'width':width,'height':height})
            for pending in [True, False]:
                page.evaluate('''pending => {
                  const rows=Array.from({length:50},()=>({c:[]}));
                  const put=(r,c,v)=>rows[r].c[c]={v};
                  for(let i=0;i<4;i++){put(5,i*6,12345);put(5,i*6+3,6789);}
                  put(14,0,450);put(14,4,280);put(14,8,620);put(14,12,390);
                  put(12,16,'♿ 復康巴士');put(14,16,45);put(14,20,15);
                  if(!pending){put(38,0,'嶺東科大-寶文校區接駁統計');put(39,0,12345);put(39,1,6789);}
                  window.onGvizData({table:{rows}});
                  window.onMedicalData({table:{rows:[{c:[{v:pending?'未回報':123},{v:pending?'未回報':12},null,{v:'https://docs.google.com/forms/d/e/medical-test/viewform'}]}]}});
                  window.scrollTo(0,0);
                }''', pending)
                page.wait_for_timeout(100)
                result=page.evaluate('''() => {
                  const visibleBottom=visualViewport.height+visualViewport.offsetTop;
                  const cards=[...document.querySelectorAll('.station-card')];
                  return {width:innerWidth,height:innerHeight,visibleBottom,
                    landscape:matchMedia('(orientation: landscape)').matches,
                    boxes:[...document.querySelectorAll('header,.summary-banner,.cards-grid')].map(e=>({name:e.className,height:e.getBoundingClientRect().height,columns:getComputedStyle(e).gridTemplateColumns})),
                    bottom:Math.max(...cards.map(e=>e.getBoundingClientRect().bottom)),
                    hidden:cards.filter(e=>!e.getBoundingClientRect().height).length,
                    overflow:document.scrollingElement.scrollWidth>innerWidth,
                    clipped:[...document.querySelectorAll('.card-header,.card-col-num,.metric-value,.metric-label,h1')]
                      .filter(e=>e.scrollWidth>e.clientWidth+1||e.scrollHeight>e.clientHeight+1).map(e=>e.textContent)};
                }''')
                result.update(engine=engine.name,pending=pending)
                results.append(result)
                assert page.locator('.station-card').count() == 9
                assert page.locator('#accessibleIn').inner_text() == '45'
                assert page.locator('#walkTotalIn').inner_text() == '1,115'
                assert result['bottom'] <= result['visibleBottom']-4, result
                assert not result['hidden'] and not result['overflow'] and not result['clipped'], result
                assert not errors, errors
                if width==874 and pending:
                    page.screenshot(path=str(root/'.omx'/f'landscape-fit-{engine.name}.png'))
        browser.close()
(root/'.omx/landscape-fit.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
print(f'PASS: {len(results)} Chromium/WebKit cases; all nine cards fit without scrolling after rotation')
