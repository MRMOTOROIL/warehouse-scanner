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
