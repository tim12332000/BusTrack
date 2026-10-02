from pathlib import Path
import json
import os
from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
results = []
(root/'.omx').mkdir(exist_ok=True)
with sync_playwright() as p:
    browser = p.chromium.launch(executable_path=os.environ.get('CHROME_PATH', r'C:\Program Files\Google\Chrome\Application\chrome.exe'), headless=True)
    for filename, title in [('index.html', '進（離）場人數統計'), ('parking.html', '停車場現況')]:
      for width, height in [(956,440), (874,402), (844,390), (667,375), (844,320), (915,412), (390,844), (375,667), (1920,1080)]:
        page = browser.new_page(viewport={'width':width,'height':height}, device_scale_factor=1)
        page.route('https://**', lambda route: route.abort())
        page.goto((root/filename).as_uri(), wait_until='load')
        page.evaluate('''() => {
          const rows=Array.from({length:40},()=>({c:[]}));
          const put=(r,c,v)=>rows[r].c[c]={v};
          put(0,0,12345);put(0,8,6789);
          put(36,0,'嶺東科大-寶文校區接駁統計');put(37,0,1500);put(37,1,700);
          if (document.getElementById('parkingGroups')) {
            PARKING_LOTS.forEach((lot,i) => { put(i,0,lot.group);put(i,1,lot.name);put(i,2,lot.cars);put(i,3,lot.motorcycles); });
          } else {
            put(25,0,'🅿️ 汽機車停車場剩餘車位');
            put(27,0,1805);put(27,8,1780);
            put(30,0,'各停車場即時剩餘車位（7處）');
            for(let i=0;i<7;i++){put(33,i*4,i===3?'不提供':200);put(33,i*4+2,i>=5?'不提供':100);}
          }
          window.onGvizData({table:{rows}});
        }''')
        page.wait_for_timeout(550)
        assert page.locator('h1').inner_text() == '「國防知性之旅-成功嶺營區開放」' + title
        if filename == 'index.html':
            assert page.locator('.station-card').count() == 8
            assert page.locator('#bus5In').inner_text() == '1,500'
            assert page.locator('#walk3In').count() == 0
        else:
            assert page.locator('#parkingGroups .lot-row[data-lot]').count() == 23
            assert page.locator('#parkingGroups .lot-group').count() == 3
            assert page.locator('#parkingGroups .group-metric .capacity-bar').count() == 6
            # 停車場頁要一個畫面看完全部，不可垂直捲動。
            assert page.evaluate('document.scrollingElement.scrollHeight <= innerHeight + 1 && document.body.scrollHeight <= innerHeight + 1'), (width, height)
            assert page.locator('#parkTotalCars').inner_text() == '8,305'
            assert page.locator('#parkTotalMotorcycles').inner_text() == '5,693'
        assert page.locator('#fabFlip').count() == 0
        assert page.locator('#fabRotate').count() == 0
        assert page.locator('#fabRefresh').inner_text() == '↻ 強制刷新'
        if filename == 'index.html' and page.evaluate('document.scrollingElement.scrollHeight > innerHeight'):
            page.locator('.station-card').last.scroll_into_view_if_needed()
            assert page.locator('.station-card').last.evaluate('(e) => { const r = e.getBoundingClientRect(); return r.top < innerHeight && r.bottom > 0; }')
            page.evaluate('window.scrollTo(0, 0)')
        page.screenshot(path=str(root/'.omx'/f'{filename}-fit-{width}x{height}.png'), full_page=True)
        result = page.evaluate('''() => {
          const outside=[...document.querySelectorAll('.station-card, .lot-group, header, .floating-toolbar, .summary-banner')].map(e=>{
            const r=e.getBoundingClientRect();return {name:e.className,x:r.x,y:r.y,right:r.right,bottom:r.bottom};
          }).filter(r=>r.x < -1 || r.y < -1 || r.right>innerWidth+1);
          const clipped=[...document.querySelectorAll('.card-header,.card-col-label,.card-col-num,.metric-value,.metric-label,.card-status-text,h1,.progress-pct,.parking-capacity,.lot-name,.lot-num,.group-name,.group-metric,.metric-capacity')]
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
    # Separate pages allow vertical scrolling on small screens.
    assert result['scrollWidth'] <= result['width'] + 1, result
    assert not result['outside'] and not result['clipped'] and not result['cropped'], result
print(f"PASS: {len(results)} viewports fit all cards and labels without horizontal overflow or cropping")
