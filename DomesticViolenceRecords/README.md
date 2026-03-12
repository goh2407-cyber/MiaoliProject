# 家暴服務紀錄系統

> 本專案使用 Google Apps Script 開發，配合 clasp 進行本機開發與部署。

## 專案結構

```
DomesticViolenceRecords/
├── Code.gs              # 主程式入口
├── Init.gs              # 試算表初始化
├── Auth.gs              # 權限管理
├── Lock.gs              # 鎖定與重試機制
├── RecordService.gs     # 服務紀錄 CRUD
├── OptionsService.gs    # 選項設定管理
├── ExportService.gs     # 匯出功能
├── StatsService.gs      # 統計功能
├── Utils.gs             # 工具函數
├── index.html           # 主頁面
├── styles.html          # CSS 樣式
├── scripts.html         # 前端邏輯
├── appsscript.json      # Apps Script 設定
└── README.md            # 本檔案
```

## 開發環境設置

### 1. 登入 clasp

```bash
clasp login
```

這會開啟瀏覽器讓你登入 Google 帳號。

### 2. 建立新專案

```bash
clasp create --type webapp --title "家暴服務紀錄系統"
```

### 3. 推送程式碼

```bash
clasp push
```

### 4. 開啟 Apps Script 編輯器

```bash
clasp open
```

### 5. 部署為 Web App

1. 在 Apps Script 編輯器中，點擊「部署」→「新增部署」
2. 選擇「Web 應用程式」
3. 設定執行身分為「使用者」
4. 設定存取權限為「組織內使用者」
5. 點擊「部署」

## 首次使用

1. 部署後，開啟 Web App URL
2. 系統會自動建立 Google Sheets
3. 第一位使用者會自動成為管理員
4. 管理員可在「系統管理」頁籤新增其他使用者

## 功能說明

- **服務紀錄管理**：新增、編輯、刪除、查詢服務紀錄
- **篩選功能**：依日期、社工、個案姓名篩選
- **匯出功能**：匯出為 Excel 或 PDF
- **使用者管理**：管理員可新增/停用使用者
- **權限控制**：管理者/使用者兩種角色

## 注意事項

- 多人同時使用時，系統會自動處理並發衝突
- 刪除紀錄會顯示確認對話框
- 所有紀錄都會記錄建立者和修改者
