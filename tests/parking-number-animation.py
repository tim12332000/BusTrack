"""Verify real animation frames and counter identity in Chromium and WebKit."""
from pathlib import Path
from playwright.sync_api import sync_playwright

root = Path(__file__).resolve().parents[1]
with sync_playwright() as p:
    for engine in [p.chromium, p.webkit]:
        browser = engine.launch()
        page = browser.new_page(reduced_motion='no-preference')
        page.route('https://**', lambda route: route.abort())
        errors = []
        page.on('pageerror', lambda error: errors.append(str(error)))
        page.goto((root / 'parking.html').as_uri())
        page.evaluate('''() => {
          window.report = value => window.onGvizData({table:{rows:[{c:[null,{v:'CITYPARKING春安站'},{v:value},{v:5}]}]}});
          window.report(10);
        }''')
        page.wait_for_function("document.getElementById('parkTotalCars').textContent === '10'")
        page.evaluate('''() => {
          window.counter = parkingNumbers.get('lot:CITYPARKING春安站:cars');
          window.report(80);
        }''')
        page.wait_for_function("Number(window.counter.textContent) > 10 && Number(window.counter.textContent) < 80")
        assert page.evaluate("document.contains(window.counter)")
        page.evaluate('window.report(80)')
        assert page.evaluate("window.counter === parkingNumbers.get('lot:CITYPARKING春安站:cars') && document.contains(window.counter)")
        page.wait_for_function("window.counter.textContent === '80' && document.getElementById('parkTotalCars').textContent === '80'")
        page.evaluate("window.report('abc')")
        assert page.evaluate("window.counter.textContent === '未回報'")
        page.wait_for_function("document.getElementById('parkTotalCars').textContent === '0'")
        page.evaluate('window.report(20)')
        page.wait_for_function("window.counter.textContent === '20'")
        page.emulate_media(reduced_motion='reduce')
        page.evaluate('window.report(5)')
        assert page.evaluate("window.counter.textContent === '5' && numberAnimations.size === 0")
        assert not errors, errors
        browser.close()
        print(f'PASS: {engine.name} intermediate frames, DOM identity, exact targets, invalid input, recovery, reduced motion')
