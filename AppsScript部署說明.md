# Apps Script 遠端執行

## 上線前清除測試資料

同一個「運管中心」選單也提供「新增測試資料」：確認後附加 14 筆人員進離場、7 筆停車場剩餘車位資料，涵蓋全部站點，備註皆標記「測試資料」。保留原有資料；重複按會繼續累加。資料直接寫入目前使用的回應分頁，不建立 Google Forms 原生回覆。正式上線前仍須清除測試資料。

使用電腦瀏覽器重新開啟原試算表，選「運管中心 → 清除測試資料」，核對兩份表單的筆數後按「是」。首次使用需依 Google 畫面授權。

- 開啟試算表或取消確認不會清除資料。
- 先備份整份試算表，再清空兩份正式表單的原生回覆與兩張目前使用的回應分頁；保留標題列、看板公式、表單題目與網址。
- 清除期间暫停兩份表單收件，結束時恢復各自原本的收件狀態。人員統計歸零，停車位等待重新回報。
- 備份保存試算表紀錄；Google Forms 的原生回覆刪除後無法還原。若中途失敗，提示會保留備份網址，不能視為全部成功。
- 自訂選單請用電腦版 Google 試算表操作。

選單使用綁定原試算表的獨立清理專案，識別資訊在 `apps-script/reset-menu-project.json`。原有遠端修復專案保留。
清理來源為 `【指令】一鍵清空資料.js`，新增測試資料來源為 `【指令】產生測試資料.js`；部署時一併帶入主程式的目標試算表常數、回應來源設定、`responseColumnLetter` 與 `findResponseSource`，不安裝初始化／重建表單功能。
該綁定專案只使用 `spreadsheets` 與 `forms` scopes，沒有公開的網頁清除入口。

驗證：`node --test tests/*.test.cjs`。安裝只更新程式，不執行 `runClearAllData`；正式清除需由操作人員確認。

## 已完成

- 專案：`BUS`
- Script ID：`1Pm6eAGUA9BLxx6ULDL8pCXLgXSwOpemoYKddNb_V7eYq-x-KtFUW-KTV`
- API 部署 ID：`AKfycbzLVk6GC2xk0UaTXP3ObQqirZhqOY1NCNMHHYuzkJgM4YUHi8HdNRHNVRgYdPj0mOx3eQ`
- 首次部署版本：1，2026-09-08。
- 目前已驗證部署版本：2，2026-09-08。
- 標準 Cloud 專案：`gen-lang-client-0465156039`，專案編號 `505010910759`；Apps Script API 已啟用。
- clasp 具名登入：`bus`，使用此 Cloud 專案建立的 Desktop OAuth client。憑證與 refresh token 僅保存在本機，未提交 Git。
- 已經由部署 API 讀回確認：`EXECUTION_API`，存取權 `MYSELF`。
- 部署 manifest 保存於 `apps-script/appsscript.json`。工作副本為 `.omx/parking-upgrade/`，僅包含主程式與 manifest，不要將專案根目錄所有測試／歷史腳本一併上傳。

## 執行方式與驗證

在 `.omx/parking-upgrade/` 工作副本執行：

```powershell
npx --yes @google/clasp --user bus run upgradeParkingRemaining
```

日常僅修公式時將函式名稱改為 `repairDashboard`。函式沒有回傳值時，成功輸出為 `null`；仍須核對表單／試算表，不能只憑 CLI exit code 判斷成功。

本次已成功遠端執行 `upgradeParkingRemaining`。實際讀回確認：

- 原表單網址保留，七處場地、剩餘汽車位、剩餘機車位題目存在。
- 舊的回報動作、車輛增減數量題目已移除。
- 總看板標記為七處版本，容量 3,023 汽車位／2,773 機車位。
- 無該車種顯示「不提供」，其餘新制尚未回報；沒有讀到公式錯誤。
- 未為驗證而向正式表單提交假資料；新制回報到看板的完整資料流程仍需實際回報驗證。

manifest 明確列出 `spreadsheets`、`forms`、`drive.readonly` 三個執行 scopes。更改腳本使用的服務時，須同步檢查 manifest 與重新授權需求。

## 重建連線時注意

Google 官方要求腳本和 OAuth client 使用同一個標準 Cloud 專案。原本的 `default` 登入使用 Google 共用的 client，曾回傳 HTTP 403 `PERMISSION_DENIED`；遠端執行應使用 `--user bus`。

重建登入時使用 `clasp --user bus login --creds <私人位置的client JSON> --use-project-scopes --include-clasp-scopes`。不要把 client JSON 或 `.clasprc.json` 提交至版本庫。

Apps Script 與 Cloud 專案的連結、一般 Desktop OAuth client 建立，官方公開 API 未提供完整替代介面，需使用 Google 設定頁。`.clasp.json` 的 `projectId` 不會代替腳本的 Cloud 專案綁定。

切換預設 Cloud 專案後無法切回原預設專案，應先核對現有關聯。不要公開 API 執行權限或建立匿名執行入口來繞過 OAuth。

本次嘗試啟動可自動操作的 Chrome 設定視窗遭環境政策拒絕，未建立瀏覽器連線；此操作不得改用另一種命令重試繞過。

官方參考：[遠端執行要求](https://developers.google.com/apps-script/api/how-tos/execute)、[Cloud 專案關聯](https://developers.google.com/apps-script/guides/cloud-platform-projects)、[clasp run 設定](https://github.com/google/clasp/blob/master/docs/run.md)。
