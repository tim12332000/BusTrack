from pathlib import Path
import json
import os
from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
results = []
(root/'.omx').mkdir(exist_ok=True)
with sync_playwright() as p:
    browser = p.chromium.launch(executable_path=os.environ.get('CHROME_PATH', r'C:\Program Files\Google\Chrome\Application\chrome.exe'), headless=True)
    for width, height in [(844,390), (667,375), (844,320), (915,412), (390,844), (375,667), (1920,1080)]:
        page = browser.new_page(viewport={'width':width,'height':height}, device_scale_factor=1)
        page.route('https://**', lambda route: route.abort())
        page.goto((root/'index.html').as_uri(), wait_until='load')
        page.evaluate('''() => {
          const rows=Array.from({length:40},()=>({c:[]}));
          const put=(r,c,v)=>rows[r].c[c]={v};
          put(0,0,12345);put(0,8,6789);
          put(25,0,'🅿️ 汽機車停車場剩餘車位');
          put(27,0,1805);put(27,8,1780);
          put(30,0,'各停車場即時剩餘車位（7處）');
          for(let i=0;i<7;i++){put(33,i*4,i===3?'不提供':200);put(33,i*4+2,i>=5?'不提供':100);}
          window.onGvizData({table:{rows}});
        }''')
        page.wait_for_timeout(550)
        assert page.locator('h1').inner_text() == '「國防知性之旅-成功嶺營區開放」即時戰情中心'
        assert page.locator('#fabFlip').count() == 0
        page.screenshot(path=str(root/'.omx'/f'fit-{width}x{height}.png'), full_page=True)
        result = page.evaluate('''() => {
          const outside=[...document.querySelectorAll('.station-card, header, .floating-toolbar, .summary-banner')].map(e=>{
            const r=e.getBoundingClientRect();return {name:e.className,x:r.x,y:r.y,right:r.right,bottom:r.bottom};
          }).filter(r=>r.x < -1 || r.y < -1 || r.right>innerWidth+1 || r.bottom>innerHeight+1);
          const clipped=[...document.querySelectorAll('.card-header,.card-col-label,.card-col-num,.metric-value,.metric-label,.card-status-text,h1,.progress-pct,.parking-capacity')]
            .filter(e=>e.clientWidth>0 && (e.scrollWidth>e.clientWidth+1 || e.scrollHeight>e.clientHeight+1))
            .map(e=>({text:e.textContent.trim(),width:e.clientWidth,scrollWidth:e.scrollWidth}));
          const cropped=[...document.querySelectorAll('.station-card .card-col-num,.station-card .parking-capacity,.station-card .card-status-text')].filter(e=>{
            const r=e.getBoundingClientRect(),p=e.closest('.station-card').getBoundingClientRect();
            return r.bottom>p.bottom+0.5 || r.right>p.right+0.5 || r.x<p.x-0.5;
          }).map(e=>e.textContent);
          return {width:innerWidth,height:innerHeight,scrollHeight:document.scrollingElement.scrollHeight,scrollWidth:document.scrollingElement.scrollWidth,outside,clipped,cropped};
        }''')
        results.append(result)
        page.close()
    browser.close()
(root/'.omx/fit-results.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
for result in results:
    assert result['scrollHeight'] <= result['height'] + 1, result
    assert result['scrollWidth'] <= result['width'] + 1, result
    assert not result['outside'] and not result['clipped'] and not result['cropped'], result
print(f"PASS: {len(results)} viewports fit all cards and labels without scrolling or cropping")
