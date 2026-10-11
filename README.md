# 🤿 潛水旅費動態結算器

網址：https://qwe0975169776-dotcom.github.io/dive-calculator/

- `index.html`：網頁（單一檔案）。單價在 `const PRICES`；Google Sheet 網址在 `const SHEET_URL`
- `Code.gs`：試算表伺服器程式（Apps Script 啟動程式會自動從這裡下載，改這裡就會自動更新）
- `apps-script/loader.gs`：只需貼一次到 Apps Script 的啟動程式
- `tests/mock-server.js`：模擬 Google 服務，在本機測試 Code.gs

資料存在試算表的「名單」分頁；本機瀏覽器也會留一份備份，連不上時先存本機、之後自動補上傳。兩台裝置同時修改會自動合併。
