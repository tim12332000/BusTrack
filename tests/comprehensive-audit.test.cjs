const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const indexPath = path.join(__dirname, '../index.html');
const screenPath = path.join(__dirname, '../戰情大螢幕.html');
const indexHtml = fs.readFileSync(indexPath, 'utf8');
const screenHtml = fs.readFileSync(screenPath, 'utf8');

test('1. 雙入口檔案 (index.html 與 戰情大螢幕.html) 內容 100% 絕對一致', () => {
  assert.equal(indexHtml, screenHtml, '兩份 HTML 檔案內容必須完全相同');
});

test('2. HTML 標籤結構語法完整性驗證 (不得有遺漏閉合或多餘標籤)', () => {
  const doctypeCount = (indexHtml.match(/<!DOCTYPE html>/gi) || []).length;
  assert.equal(doctypeCount, 1, 'DOCTYPE 宣告必須恰好為 1 次');

  const htmlOpen = (indexHtml.match(/<html\b[^>]*>/gi) || []).length;
  const htmlClose = (indexHtml.match(/<\/html>/gi) || []).length;
  assert.equal(htmlOpen, 1, '<html> 開啟標籤必須恰好 1 個');
  assert.equal(htmlClose, 1, '</html> 閉合標籤必須恰好 1 個');

  const headOpen = (indexHtml.match(/<head\b[^>]*>/gi) || []).length;
  const headClose = (indexHtml.match(/<\/head>/gi) || []).length;
  assert.equal(headOpen, 1, '<head> 開啟標籤必須恰好 1 個');
  assert.equal(headClose, 1, '</head> 閉合標籤必須恰好 1 個');

  const styleOpen = (indexHtml.match(/<style\b[^>]*>/gi) || []).length;
  const styleClose = (indexHtml.match(/<\/style>/gi) || []).length;
  assert.equal(styleOpen, 1, '<style> 開啟標籤必須恰好 1 個');
  assert.equal(styleClose, 1, '</style> 閉合標籤必須恰好 1 個');

  const bodyOpen = (indexHtml.match(/<body\b[^>]*>/gi) || []).length;
  const bodyClose = (indexHtml.match(/<\/body>/gi) || []).length;
  assert.equal(bodyOpen, 1, '<body> 開啟標籤必須恰好 1 個');
  assert.equal(bodyClose, 1, '</body> 閉合標籤必須恰好 1 個');

  const scriptOpen = (indexHtml.match(/<script\b[^>]*>/gi) || []).length;
  const scriptClose = (indexHtml.match(/<\/script>/gi) || []).length;
  assert.equal(scriptOpen, 1, '<script> 開啟標籤必須恰好 1 個');
  assert.equal(scriptClose, 1, '</script> 閉合標籤必須恰好 1 個');
});

test('3. CSS 樣式大括號對稱性與關鍵媒體查詢檢驗', () => {
  const styleMatch = indexHtml.match(/<style\b[^>]*>([\s\S]*?)<\/style>/i);
  assert.ok(styleMatch, '必須能成功擷取 <style> 內容');
  const css = styleMatch[1];

  let depth = 0;
  for (let i = 0; i < css.length; i++) {
    if (css[i] === '{') depth++;
    else if (css[i] === '}') depth--;
    assert.ok(depth >= 0, `CSS 大括號閉合異常 (在索引 ${i} 處發現多餘閉合括號)`);
  }
  assert.equal(depth, 0, 'CSS 所有開括號與閉括號必須完全配對成雙，不可遺漏閉合');

  // 關鍵響應式媒體查詢
  assert.ok(css.includes('@media (max-width: 900px), (orientation: landscape) and (max-height: 520px)'), '緊湊排版必須涵蓋手機直向與低高度橫向');
  assert.ok(css.includes('@media (orientation: landscape) and (max-height: 350px)'), '必須包含超薄橫向媒體查詢');
  assert.ok(css.includes('@media (orientation: portrait) and (max-width: 900px)'), '必須包含手機直向媒體查詢');
  assert.ok(css.includes('overflow-y: auto'), '橫向模式必須具備縱向捲動保護');
});

test('4. 前端所有業務 DOM ID 存在性檢驗 (不可漏掉任何一個)', () => {
  const requiredIds = [
    // 頂部
    'liveBadge', 'liveText', 'currentTime', 'fabFullscreen', 'fabRefresh',
    // 接駁車大盤
    'busTotalIn', 'busTotalOut', 'busTotalRate', 'busMainBar',
    // 4 大接駁站卡片
    'bus1In', 'bus1Out', 'bus1Status', 'bus1Bar',
    'bus2In', 'bus2Out', 'bus2Status', 'bus2Bar',
    'bus3In', 'bus3Out', 'bus3Status', 'bus3Bar',
    'bus4In', 'bus4Out', 'bus4Status', 'bus4Bar',
    // 步行門大盤
    'walkTotalIn', 'walkTotalOut', 'walkTotalRate', 'walkMainBar',
    // 3 大步行門卡片
    'walk1In', 'walk1Out', 'walk1Status', 'walk1Bar',
    'walk2In', 'walk2Out', 'walk2Status', 'walk2Bar',
    'walk3In', 'walk3Out', 'walk3Status', 'walk3Bar',
    // 停車場大盤
    'parkingSection', 'parkingNote', 'parkTotalCars', 'parkTotalMotorcycles', 'parkReported',
    // 7 處停車場卡片
    'park1Cars', 'park1Motorcycles',
    'park2Cars', 'park2Motorcycles',
    'park3Cars', 'park3Motorcycles',
    'park4Cars', 'park4Motorcycles',
    'park5Cars', 'park5Motorcycles',
    'park6Cars', 'park6Motorcycles',
    'park7Cars', 'park7Motorcycles'
  ];

  for (const id of requiredIds) {
    const re = new RegExp(`id="${id}"`);
    assert.ok(re.test(indexHtml), `DOM 中必須包含 id="${id}"`);
  }
});

test('5. 模擬 JavaScript 全功能運作 (即時數據解析、數值格式化、邊界處理)', () => {
  const dom = new Map();
  for (const m of indexHtml.matchAll(/id="([^"]+)"[^>]*>([^<]*)/g)) {
    dom.set(m[1], {
      textContent: m[2],
      style: {},
      classList: {
        set: new Set(),
        add(c) { this.set.add(c); },
        remove(c) { this.set.delete(c); },
        toggle(c, force) {
          if (force !== undefined) { if (force) this.set.add(c); else this.set.delete(c); }
          else { if (this.set.has(c)) this.set.delete(c); else this.set.add(c); }
        }
      }
    });
  }

  const scriptMatch = indexHtml.match(/<script>([\s\S]*?)<\/script>/);
  assert.ok(scriptMatch, '必須存在 <script>');
  const scriptContent = scriptMatch[1];

  const fakeDocument = {
    getElementById: (id) => dom.get(id) || null,
    documentElement: { requestFullscreen() {}, webkitRequestFullscreen() {} },
    exitFullscreen() {},
    webkitExitFullscreen() {},
    createElement: () => ({ remove() {} }),
    head: { appendChild() {} },
    addEventListener() {}
  };

  const fakeWindow = {};
  const context = vm.createContext({
    window: fakeWindow,
    document: fakeDocument,
    setInterval() {},
    clearInterval() {},
    Date,
    Math,
    Number,
    String,
    parseFloat,
    isNaN,
    console
  });

  // 執行腳本，確認語法與函式初始化無拋錯
  vm.runInContext(scriptContent, context);

  // 驗證內部工具函式
  assert.equal(context.formatNumber(1234567), '1,234,567');
  assert.equal(context.formatNumber(0), '0');

  const r1 = context.calcRate(100, 50);
  assert.equal(r1.pctText, '50.0%');
  assert.equal(r1.remain, 50);

  const r2 = context.calcRate(0, 0);
  assert.equal(r2.pctText, '0.0%');

  // 驗證 onGvizData 正常資料接收
  const mockRows = Array.from({ length: 45 }, () => ({ c: [] }));
  const put = (r, c, v) => { mockRows[r].c[c] = { v }; };

  // 接駁車大盤 & 站點標題與數值
  put(0, 0, 4500); put(0, 8, 3000);
  put(2, 0, '🚌 接駁車站即時數據');
  // b1: 成功車站
  put(5, 0, 1000); put(5, 3, 800);
  // b2: 新烏日
  put(5, 6, 2000); put(5, 9, 1200);
  // b3: 經貿六
  put(5, 12, 1000); put(5, 15, 600);
  // b4: 水湳
  put(5, 18, 500); put(5, 21, 400);

  // 步行大盤 & 站點
  put(7, 0, '🚶 步行通道疏運概況');
  put(9, 0, 1500); put(9, 8, 1200);
  put(11, 0, '🚶 步行通道即時數據');
  put(14, 0, 600); put(14, 4, 500);
  put(14, 8, 500); put(14, 12, 400);
  put(14, 16, 400); put(14, 20, 300);

  // 停車場
  put(25, 0, '🅿️ 汽機車停車場剩餘車位');
  put(27, 0, 1200); put(27, 8, 1500);
  put(30, 0, '各停車場即時剩餘車位（7處）');
  for (let i = 0; i < 7; i++) {
    put(33, i * 4, i === 3 ? '不提供' : 100);
    put(33, i * 4 + 2, i >= 5 ? '不提供' : 150);
  }

  fakeWindow.onGvizData({ status: 'ok', table: { rows: mockRows } });

  // 斷言更新後的 DOM 數值
  assert.equal(dom.get('busTotalIn').textContent, '4,500');
  assert.equal(dom.get('busTotalOut').textContent, '3,000');
  assert.equal(dom.get('busTotalRate').textContent, '66.7%');
  assert.equal(dom.get('bus1In').textContent, '1,000');
  assert.equal(dom.get('bus1Out').textContent, '800');
  assert.equal(dom.get('walkTotalIn').textContent, '1,500');
  assert.equal(dom.get('park1Cars').textContent, '100');
  assert.equal(dom.get('park4Cars').textContent, '—');
  assert.equal(dom.get('park6Motorcycles').textContent, '—');
  assert.equal(dom.get('parkReported').textContent, '7/7 處');
});
