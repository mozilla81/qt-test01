# QT TEST01 - ERP 報價系統 (Google Apps Script)

這是一個以 **Google Apps Script (GAS) + Google 試算表 (Google Sheets) + 現代響應式 Web App 表單** 構建的輕量級企業報價 ERP 系統。

---

## 🌟 核心功能

1. **報價單主檔與明細管理**
   - 自動生成編號：支援依年月日自動流水號（如 `QT-20261118-001`）。
   - 動態品項輸入：支援多品項即時增刪，輸入數量、單價、折扣自動計算行小計。
   - 自動金額試算：未稅小計、5% 營業稅額、含稅總額即時試算。
2. **客戶與產品資料庫整合**
   - 內建常用客戶資料庫（名稱、統編、窗口、電話、Email、地址），支援下拉快速帶入。
   - 內建產品/服務價目庫，一鍵快速加入報價明細。
3. **雲端試算表後端儲存**
   - 主檔表 (`Quotations`) 與明細表 (`Quotation_Items`) 關聯儲存。
   - 具備並發防呆鎖定機制 (`LockService`)，保障資料完整性。
4. **報價歷程與檢視列印**
   - 歷史報價單快速查詢、狀態追蹤（草稿、成單、作廢）。
   - 內建報價單預覽與列印版型。

---

## 📁 專案結構

```text
qt-test01/
├── src/
│   ├── Code.gs             # GAS 後端邏輯 (API、試算表讀寫、自動編號、資料庫初始化)
│   ├── index.html          # Web App 現代化響應式報價表單前端
│   └── appsscript.json     # Apps Script 專案設定檔 (時區: Asia/Taipei)
├── .gitignore              # Git 忽略設定
└── README.md               # 專案說明與部署指南
```

---

## 🚀 快速上手與部署步驟 (Google Apps Script)

### 步驟 1：建立 Google 試算表
1. 前往 [Google 雲端硬碟](https://drive.google.com/)，建立一份新的 Google 試算表，命名為 **「ERP報價系統_資料庫」**。
2. 在頂部選單點選 **「擴充功能」 > 「Apps Script」**。

### 步驟 2：匯入程式碼
1. 將專案中的 `src/Code.gs` 內容複製並貼到 Apps Script 編輯器的 `程式碼.gs`。
2. 在編輯器中點選左側 **「+」 > 「HTML」**，建立一個名為 `index` 的檔案，並將專案中的 `src/index.html` 內容貼入。
3. 在專案設定中若需要可同步覆蓋 `appsscript.json`。

### 步驟 3：初始化資料庫
1. 在 Apps Script 編輯器上方函式下拉選單中選擇 **`initDatabase`**。
2. 點擊 **「執行」**（第一次執行需授權試算表讀寫權限）。
3. 執行完成後，返回 Google 試算表，你會看到自動建立了 4 個工作表：
   - `Quotations` (報價單主檔)
   - `Quotation_Items` (報價明細)
   - `Customers` (客戶資料庫)
   - `Products` (產品價目庫)

### 步驟 4：發布為 Web 應用程式 (Web App)
1. 在 Apps Script 編輯器右上角點擊 **「部署」 > 「新增部署作業」**。
2. 齒輪圖示選擇 **「網路應用程式」**：
   - **說明**：`ERP 報價系統 v1.0`
   - **以何者身分執行**：`我`
   - **誰可以存取**：`所有人` (或依公司內部權限設定)
3. 點擊 **「部署」**，即可獲得專屬的 Web 應用程式網址，打開即可開始開立報價單！

---

## 🔗 GitHub 連結與推送指南

如需將本專案同步至 GitHub：

1. **登入 GitHub CLI**：
   ```bash
   gh auth login
   ```
2. **在 GitHub 上建立遠端儲存庫並推動**：
   ```bash
   gh repo create "qt-test01" --public --source=. --remote=origin --push
   ```
   或使用標準 Git 指令：
   ```bash
   git remote add origin https://github.com/<您的使用者名稱>/qt-test01.git
   git push -u origin main
   ```
