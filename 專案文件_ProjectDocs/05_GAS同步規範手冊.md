# Google Apps Script 與 Clasp 同步規範手冊 (v1.0)

> [!IMPORTANT]
> **AI 助手調度指令**：未來任何 AI 在處理此專案的「部署」或「同步」工作前，必須優先讀取此文件。

## 🎯 核心宗旨
本專案 (emedu-Slides-Sync) 使用 `clasp` 進行本地開發與雲端同步。為避免 Google Apps Script 緩慢的快取機制、特殊的資料傳輸限制以及檔案副檔名衝突，必須嚴格執行以下規範。

---

## 🚫 禁令與規範 (The "Never" List)

### 1. 嚴禁直接回傳 Date 物件
*   **故障現象**：按鈕點擊後卡在「處理中...」，伺服器無報錯日誌，但前端 `withSuccessHandler` 沒被啟動。
*   **原因**：`google.script.run` 非同步回傳不支援 JavaScript `Date` 物件。
*   **規範**：所有時間資訊必須在伺服器端轉為 **ISO 字串** 或 **時間戳記**。
    ```javascript
    // ❌ 錯誤
    return { timestamp: new Date() };
    // ✅ 正確
    return { timestamp: new Date().toISOString() };
    ```

### 2. 禁止 `.js` 與 `.gs` 檔案重疊
*   **故障現象**：推送代碼後，雲端運行結果依然是舊版。
*   **原因**：GAS 雲端環境若同時存在 `Main.js` 與 `Main.gs`，會造成執行衝突。
*   **規範**：
    - 本地開發環境嚴禁保留 `.js` 檔。
    - 每次大規模同步前，建議在雲端編輯器手動檢查有無「同名不同副檔名」的檔案。

### 3. 本地目錄純淨化 (Source Root)
*   **規範**：所有實體程式碼必須位於 `src/` 目錄。
*   **原因**：若根目錄與 `src/` 同時存在同名檔案，開發環境會陷入混亂，`clasp push` 也會產生多重映射衝突。
*   **操作**：確認根目錄無冗餘代碼檔案，且 `.clasp.json` 中的 `rootDir` 已正確指向 `src`。

### 4. 編碼守則 (Encoding Guard)
*   **規範**：嚴禁使用可能破壞 UTF-8 (無 BOM) 格式的批次指令進行版本取代。
*   **後果**：中文註解會損壞成為亂碼（Mojibake），導致 `clasp` 推送後語法報錯。
*   **推薦**：使用 VS Code 的全域搜尋取代功能。

---

## 🛠️ 強制部署標準程序 (SOP)

每當本地完成程式碼修改，準備推送到雲端時，請按以下順序執行：

### 第一步：環境清理
在同步目錄執行以下指令，清理損壞的映射或是舊檔案：
```powershell
# 清除所有舊的 js 映射
rm *.js
```

### 第二步：精確推送
使用強制推送參數，確保寫入權限與檔案覆蓋：
```powershell
clasp push -f
```

### [NEW] 第三步：Script ID 強制核對 (Alignment Check)
在執行 `clasp push` 前，務必開啟瀏覽器中的 GAS 編輯器，核對網址列中的 ID (`projects/XXXXX/edit`) 是否與 `.clasp.json` 內的 `scriptId` 完全一致。
> [!WARNING]
> 若 ID 不對，代碼會被推送到「另一個同名專案」中，導致您目前開啟的編輯器內容永遠不會更新。

### 第四步：雲端版本定錨 (管理部署)
推送到雲端後，**必須**手動進入 Google Apps Script 網頁介面執行：
1.  點擊 **「部署」** -> **「管理部署」**。
2.  點擊當前 Web App 的 **「編輯 (鉛筆圖示)」**。
3.  **版本** 欄位選取 **「新版本」**。
4.  點擊 **「部署」**。
> [!TIP]
> 僅執行 `clasp push` 不一定會讓「已發布」的 Web App 更新，必須手動產出「新版本」才能繞過 Google 伺服器的快取。

---

## 🔍 故障排除檢查清單 (Troubleshooting)

| 症狀 | 可能原因 | 排除動作 |
|:---:|:---|:---|
| **按鈕轉圈無反應** | 1. 伺服器 ReferenceError<br>2. 傳回了 Date 物件 | 檢查 `window.onerror` 攔截到的錯誤訊息；檢查 `apiLogin` 回傳值。 |
| **標題版本號不對** | 讀取來源為試算表名稱(舊) | 在 `Main.gs` 中檢查 `template.activityName` 是否已硬編碼為最新版。 |
| **程式碼更動沒生效** | 雲端殘留 `.js` 同名檔案 | 到雲端編輯器手動刪除同名 `.js` 檔案。 |
| **Clasp Push 失敗** | `appsscript.json` 編碼損毀 | 重新以 UTF-8 無 BOM 格式建立該 JSON 檔案。 |

---

## 🤖 AI 開發者指南
若你是 AI 助手，在輔助開發此專案時：
- **開發規範**：在 `Main.gs` 與 `UI_Action.html` 中的版號 (Footer/Title) 必須同步更新，防止使用者混淆。
- **偵測義務**：若出現同步失敗回報，應優先執行 `ls -R` 檢查有無隱藏的 `.js` 檔案。

---
**版本紀錄**：2026-03-04 | v1.0 | 建立於 v10.1.6 核級修復後。
