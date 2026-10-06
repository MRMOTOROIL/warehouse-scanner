# 倉庫出貨掃碼防呆系統

這是從既有 Apps Script 前端搬到 Vercel 的版本。

## 架構
手機/掃碼槍 → Vercel `index.html` → `/api/scan` → Apps Script Web App → Google Sheet

## Vercel 環境變數
- `APPS_SCRIPT_URL`: Apps Script Web App `/exec` URL
- `APPS_SCRIPT_API_TOKEN`: 可選；若 Apps Script Script Properties 設定 `API_TOKEN`，這裡填相同值

## Apps Script
將 `Code.gs` 更新後，部署為 Web App：
- Execute as: Me
- Who has access: 依你的實際帳號/網路環境設定；若 Vercel 需要公開呼叫，需允許匿名存取


## 目前版本的操作方式

- 固定使用 Bluetooth HID 掃碼槍，不顯示手機相機／掃描模式選擇。
- 所有掃描輸入限制為數字。
- 尚未掃託運單時，只接受 12 碼純數字託運單號。
- 畫面只保留「託運單號」與「商品」清單；商品完成時顯示 ✓，不顯示「應掃／已掃」統計數字。
- 例如 `2*福14` 會建立福14需求數量 2，必須掃到福14兩次才會打 ✓。
- 含「箱」的項目不需要逐件掃描。

## 重要：自動學習寫入 Google 試算表

`learnProduct` 的寫入動作是在 Apps Script 端執行。更新 GitHub 後，還需要把最新 `Code.gs` 貼回 Google Apps Script 並重新部署 Web App 新版本。

Apps Script 端現在提供 `doPost`，並支援託運單查詢、商品對照、商品學習與完成紀錄等 API action。

部署時請確認：執行身分為「我」，存取權允許 Vercel 呼叫；更新部署後，Vercel 的 `APPS_SCRIPT_URL` 仍指向該 `/exec` URL。
