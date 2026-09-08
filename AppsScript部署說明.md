# Apps Script 遠端執行

## 已完成

- 專案：`BUS`
- Script ID：`1Pm6eAGUA9BLxx6ULDL8pCXLgXSwOpemoYKddNb_V7eYq-x-KtFUW-KTV`
- API 部署 ID：`AKfycbzLVk6GC2xk0UaTXP3ObQqirZhqOY1NCNMHHYuzkJgM4YUHi8HdNRHNVRgYdPj0mOx3eQ`
- 首次部署版本：1，2026-09-08。
- 已經由部署 API 讀回確認：`EXECUTION_API`，存取權 `MYSELF`。
- 部署 manifest 保存於 `apps-script/appsscript.json`。工作副本為 `.omx/parking-upgrade/`，僅包含主程式與 manifest，不要將專案根目錄所有測試／歷史腳本一併上傳。

## 尚未完成

現有 clasp 登入使用 Google 共用的 OAuth client。直接使用部署 ID 呼叫 `scripts.run`，回傳 HTTP 403 `PERMISSION_DENIED`。這不是部署不存在；不得將部署建立成功宣稱為遠端執行成功。

Google 官方要求腳本和 OAuth client 使用同一個標準 Cloud 專案。接續步驟：

1. 在 Apps Script 專案設定確認目前綁定的 GCP 專案。若已有適合的標準專案，優先沿用。
2. 在該 Cloud 專案啟用 Apps Script API，配置 Google Auth platform，建立 Desktop app OAuth client。
3. 下載 client JSON 至工作區外的私人位置，不提交到 Git。
4. 使用同一個 Cloud 專案的 client 登入；需涵蓋整份腳本的 Sheets、Forms、Drive 讀取權限，以及 clasp 管理權限。採用具名使用者（例如 `bus`）保留現有管理登入。
5. 以新登入實際呼叫 `upgradeParkingRemaining`，檢查執行結果及表單與看板，再記錄成功。

Apps Script 與 Cloud 專案的連結、一般 Desktop OAuth client 建立，官方公開 API 未提供完整替代介面，需使用 Google 設定頁。`.clasp.json` 的 `projectId` 不會代替腳本的 Cloud 專案綁定。

切換預設 Cloud 專案後無法切回原預設專案，應先核對現有關聯。不要公開 API 執行權限或建立匿名執行入口來繞過 OAuth。

本次嘗試啟動可自動操作的 Chrome 設定視窗遭環境政策拒絕，未建立瀏覽器連線；此操作不得改用另一種命令重試繞過。

官方參考：[遠端執行要求](https://developers.google.com/apps-script/api/how-tos/execute)、[Cloud 專案關聯](https://developers.google.com/apps-script/guides/cloud-platform-projects)、[clasp run 設定](https://github.com/google/clasp/blob/master/docs/run.md)。
