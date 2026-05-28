# site-log-tool：工地日誌工具

手機版工地日報系統，記錄每日施工項目、工班人數、照片，並自動計算各工項完成百分比。

---

## 系統架構

```
手機瀏覽器（書籤）
        ↓
GitHub Pages（index.html）
  https://hsienchun-chen.github.io/site-log-tool/
        ↓  GET（案件清單、工項進度）
        ↓  POST（送出記錄 + 照片）
Google Apps Script Web App
  https://script.google.com/macros/s/AKfycbwiRu6a2GLKlQLvvCoJygKjRYOyS7mZeWTBlVFsGWU8300c9YKWZHjKIQA7MhniVQTWxA/exec
        ↓
  ├── Google Sheets（InfoData）   → 文字記錄
  └── Google Drive（照片資料夾） → 照片檔案
```

---

## 檔案結構

```
site-log-tool/
├── index.html          ← 手機表單（GitHub Pages 根目錄）
├── docs/index.html     ← 同上（備份，GitHub Pages 備用路徑）
├── gas/Code.js         ← GAS 後端完整程式碼（部署參考用）
├── .nojekyll           ← 停用 Jekyll，讓 GitHub Pages 直接服務靜態檔
└── CLAUDE.md           ← 本檔案
```

---

## Google 服務 ID

| 服務 | ID |
|------|-----|
| InfoData Google Sheets | `1a4N2m72WqDKDbOeFVKmW2YfmNVA-vNKApQjzGwQGDbE` |
| Drive 上傳資料夾（含毛利管制表 + 照片） | `11zjIQX0hI79OnYdLpoRZaH-0m5QhlK6s` |
| GAS Web App 網址 | 見上方架構圖 |

---

## InfoData 試算表結構

### 工作表：案件清單
| 案號 | 案件名稱 | 負責人 | 狀態 | 金額 | 備註 | 最後更新 |
由 `syncCaseList()` 從 `@案件列表.xlsx` 的 `115年` 工作表同步。
狀態為「在建」的案件才會出現在表單下拉選單。

### 工作表：工項清單
| 案號 | 案件名稱 | 類別 | 項目 | 單位 | 總數量 |
由 `syncWorkItems()` 從各案件的 `毛利管制表.xlsx` 的 `預估` 工作表解析。

### 工作表：每日記錄
| 時間戳 | 日期 | 案號 | 案件名稱 | 類別 | 項目 | 今日數量 | 工種人數 | 備註 | 照片資料夾 |
每次送出表單自動新增一列（每個工項一列）。

---

## GAS 函式說明

| 函式 | 用途 | 何時執行 |
|------|------|---------|
| `initSheets()` | 在 InfoData 建立三個工作表 | 第一次設定時執行一次 |
| `syncCaseList()` | 從 `@案件列表.xlsx` 同步案件清單 | 每次更新案件時執行 |
| `syncWorkItems()` | 從所有毛利管制表同步工項清單（分批，每次 5 份）| 每次更新工項時執行 |
| `resetSyncProgress()` | 清除 syncWorkItems 的分批進度，重頭跑 | syncWorkItems 需要重跑時 |
| `doGet(e)` | 回傳案件清單 / 工項進度 JSON（供表單讀取） | 由表單自動呼叫 |
| `doPost(e)` | 接收表單送出、寫入記錄、儲存照片 | 由表單自動呼叫 |

### syncWorkItems 分批機制
- 每次處理 5 份毛利管制表（避免 GAS 6 分鐘逾時）
- 跑完提示「還剩 N 份」，再執行一次繼續
- 17 份約需執行 3-4 次

---

## 毛利管制表解析邏輯

來源：NAS `Z:\2.在案\[案件資料夾]\採購\[案名]_毛利管制表.xlsx`

解析規則（`預估` 工作表）：
- 第 6 列是欄位標題，第 7 列開始是資料
- **Col B（類別）**：空白時沿用上一個值（fill-forward）
- **Col F（售價數量）**：有數字 = 有效工項；None = 成本子項，跳過
- 遇到「小計」「總計」「毛利率」文字即停止

---

## 案件列表解析邏輯

來源：NAS `Z:\@案件列表.xlsx`，工作表 `115年`（民國年自動計算）

區段識別：
- `在案列表` → 狀態標記「在建」
- `結案中列表` → 狀態標記「結案中」
- `提案、報價中`、`已結案` → 不匯入表單

欄位對應：案號(C)、案件名稱(E)、負責人(D)、金額(F)、備註(G)

---

## NAS 批次複製腳本

路徑：`renovation-workflow/copy_毛利管制表.ps1`

功能：把 NAS `Z:\2.在案\` 所有毛利管制表複製到 `ClaudeCode\InfoData\`，方便統一上傳至 Google Drive。

```powershell
# 在 PowerShell 執行
cd C:\Users\HsienChun\iCloudDrive\ClaudeCode\renovation-workflow
.\copy_毛利管制表.ps1
```

---

## 表單功能說明

### 使用流程
1. 選擇案件（下拉選單）
2. 勾選今日施工項目，輸入今日完成數量
   - 可篩選「全部 / 未完成 / 已完成」
   - 每項顯示累積完成 %（進度條）
3. 勾選工種並填人數
4. 拍照上傳（最多 10 張，自動壓縮至 1280px）
5. 填備註（選填）
6. 按「送出記錄」→ 自動跳出施工日報

### 施工日報（送出後）
顯示：案件、日期、今日施工項目、進場工班、照片
- 截圖後可直接傳 LINE
- 按「列印 / 存 PDF」可儲存正式記錄

---

## 完成進度計算

```
某工項完成 % = 每日記錄中該工項的數量加總 ÷ 毛利管制表的總數量 × 100%
```

顏色標示：
- 🔵 藍色（< 40%）：施工初期
- 🟠 橘色（40-79%）：進行中
- 🟢 綠色（≥ 80%）：接近完成

---

## 定期維護作業

| 時機 | 動作 |
|------|------|
| 新案件開案 | 更新 NAS `@案件列表.xlsx` → 上傳至 Drive → 執行 `syncCaseList()` |
| 工項有變動（追加、變更） | 更新 NAS 毛利管制表 → 執行 `copy_毛利管制表.ps1` → 上傳至 Drive → 執行 `resetSyncProgress()` + `syncWorkItems()` |
| 案件結案 | 在 `@案件列表.xlsx` 移至「已結案」→ 執行 `syncCaseList()` |

---

## 待開發功能

- [ ] 查詢模式：選日期 + 案件，查看歷史記錄
- [ ] 照片另存至獨立 Drive 資料夾（與上傳資料夾分開）
- [ ] 多人識別（記錄填寫人）
- [ ] 統整至主要 Notion 系統
