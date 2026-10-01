/**
 * 運管中心 → 新增測試資料。附加至目前兩張回應分頁，不覆蓋現有資料。
 * 資料供看板展示，不建立 Google Forms 原生回覆。
 * 需與主程式來源辨識函式及「一鍵清空資料」一起安裝。
 */
function runGenerateTestData() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert('新增測試資料？',
    '會新增 14 筆人員進離場，及「停車場現況」每個場地各 1 筆剩餘車位測試資料。\n' +
    '全部標記「測試資料」，保留現有回報；看板統計會包含這批資料。\n' +
    '資料直接寫入試算表，不會增加 Google 表單的原生回覆數。', ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    ui.alert('請稍後再試', '已有資料作業進行中。', ui.ButtonSet.OK);
    return;
  }
  let added = 0;
  let parkingCount = 0;
  let failure = null;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lots = readParkingLots_(ss);
    parkingCount = lots.length;
    const plan = prepareTestDataReset_(ss);
    const batches = plan.targets.map(target => {
      const source = findResponseSource(ss, target.kind);
      const headers = target.sheet.getRange(1, 1, 1, target.sheet.getLastColumn()).getValues()[0];
      return { sheet: target.sheet, rows: buildTestDataRows_(target.kind, source.columns, headers, new Date(), lots) };
    });
    // 由 Sheets 附加至尾端，避免過期列號覆蓋同時進來的新回報。
    batches.forEach(batch => batch.rows.forEach(row => { batch.sheet.appendRow(row); added++; }));
    SpreadsheetApp.flush();
  } catch (error) {
    failure = error.message;
  } finally {
    lock.releaseLock();
  }
  if (failure) {
    ui.alert('新增未完成', '已新增 ' + added + ' 筆。\n' + failure, ui.ButtonSet.OK);
  } else {
    ui.alert('新增完成', '已新增 14 筆人員及 ' + parkingCount + ' 筆停車場測試資料，看板會自動更新。', ui.ButtonSet.OK);
  }
}

// 場地名稱與容量直接讀「停車場現況」分頁：停車場網頁讀的就是這張表，名稱一定比對得上。
function readParkingLots_(ss) {
  const sheet = ss.getSheetByName('停車場現況');
  if (!sheet || sheet.getLastRow() < 2) throw new Error('找不到「停車場現況」分頁，未新增資料。');
  // B 欄場地名稱，E、F 欄汽車與機車容量（空白代表沒有該車種）。
  return sheet.getRange(2, 2, sheet.getLastRow() - 1, 5).getValues()
    .filter(row => String(row[0]).trim() !== '')
    .map(row => ({ name: String(row[0]).trim(), car: Number(row[3]) || 0, motorcycle: Number(row[4]) || 0 }));
}

function buildTestDataRows_(kind, columns, headers, now, lots) {
  const positions = {};
  Object.keys(columns).forEach(key => {
    positions[key] = columns[key].split('').reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0) - 1;
  });
  for (const field of [{ key: 'time', pattern: /^(時間戳記|時間標記|Timestamp)$/i }, { key: 'note', pattern: /備註/ }]) {
    const matches = headers.map((h, i) => field.pattern.test(String(h).trim()) ? i : -1).filter(i => i >= 0);
    if (matches.length !== 1) throw new Error('時間或備註欄位無法唯一辨識，未新增資料。');
    positions[field.key] = matches[0];
  }
  let records;
  if (kind === 'people') {
    records = ['🚌 成功車站', '🚌 新烏日台鐵站', '🚌 經貿六停車場', '🚌 水湳轉運站', '🚶 1號門', '🚶 3號門', '🚶 4號門']
      .reduce((all, station, i) => {
        const bus = i < 4 ? '1 號車' : '🚶 步行通道';
        const quantity = i < 4 ? 40 + i : 200 + i * 10;
        return all.concat([{ station, bus, direction: '進場', quantity },
          { station, bus, direction: '離場', quantity: Math.floor(quantity * 0.6) }]);
      }, []);
  } else {
    // 剩餘比例輪流取不同值，讓充足、吃緊、告急、客滿都看得到；沒有該車種填 0。
    const ratios = [0.65, 0.4, 0.2, 0.05, 0, 0.85, 0.5, 0.3];
    records = lots.map((lot, i) => ({
      area: lot.name,
      car: Math.round(lot.car * ratios[i % ratios.length]),
      motorcycle: Math.round(lot.motorcycle * ratios[(i + 3) % ratios.length])
    }));
  }
  return records.map((record, index) => {
    const row = Array(headers.length).fill('');
    const values = Object.assign({ time: new Date(now.getTime() + index), note: '測試資料' }, record);
    Object.keys(values).forEach(key => {
      if (!Number.isInteger(positions[key]) || positions[key] < 0 || positions[key] >= headers.length) {
        throw new Error('測試資料欄位不完整：' + key);
      }
      row[positions[key]] = values[key];
    });
    return row;
  });
}

/**
 * 在 Apps Script 編輯器直接執行用：上方選這個函式 → 按「執行」。
 * 只新增停車場測試資料（每個場地 1 筆），不跳視窗；結果看下方「執行紀錄」。
 */
function addParkingTestDataFromEditor() {
  const ss = SpreadsheetApp.openById('1SOb3pPSJoxGorKtGzcQuYh3FgNAN3UGD68TE5qR679w');
  const lots = readParkingLots_(ss);
  // 寫進「停車場現況」公式正在讀的那張回應分頁，網頁才看得到。
  const formula = ss.getSheetByName('停車場現況').getRange(2, 3).getFormula();
  const match = formula.match(/'([^']+)'!/);
  const sheet = match && ss.getSheetByName(match[1]);
  if (!sheet) throw new Error('看不出「停車場現況」讀哪一張回應分頁，未新增資料。');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const letter = pattern => {
    const index = headers.findIndex(h => pattern.test(String(h).trim()));
    if (index < 0) throw new Error('「' + match[1] + '」找不到欄位：' + pattern.source);
    let result = '';
    for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) result = String.fromCharCode(65 + (n - 1) % 26) + result;
    return result;
  };
  const columns = { area: letter(/停車場區域/), car: letter(/剩餘汽車/), motorcycle: letter(/剩餘機車/) };
  const rows = buildTestDataRows_('parking', columns, headers, new Date(), lots);
  rows.forEach(row => sheet.appendRow(row));
  SpreadsheetApp.flush();
  Logger.log('已在「' + match[1] + '」新增 ' + rows.length + ' 筆停車場測試資料（備註：測試資料）。');
}
