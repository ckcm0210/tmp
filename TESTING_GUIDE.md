# 測試指引 - Inspect Mode 基礎功能

## **測試目標**

驗證模式管理器和主程式整合是否正常運作，確保：
1. 現有功能完全保持
2. 模式切換功能正常
3. Always On Top 功能正常
4. 視窗大小調整正確

---

## **測試前準備**

### **備份確認**
- ✅ `main_backup.py` - 原始 main.py 的備份
- ✅ `main_with_mode_manager.py` - 新版本的副本
- ✅ `main.py` - 當前使用的版本

### **檔案檢查**
確認以下檔案存在：
- ✅ `core/mode_manager.py`
- ✅ `main.py` (已更新)
- ✅ `formula_comparator.py`
- ✅ `ui/workspace_view.py`

---

## **測試步驟**

### **測試 1: 基本啟動測試**

#### **執行步驟**
1. 打開命令提示字元
2. 切換到專案目錄：
   ```
   cd C:\Users\user\Excel_tools_develop\Excel_tools_develop_v38
   ```
3. 執行程式：
   ```
   python main.py
   ```

#### **預期結果**
- ✅ 程式正常啟動，無錯誤訊息
- ✅ 視窗標題：「Excel Tools - Integrated」
- ✅ 視窗大小：900x1100
- ✅ 頂部出現新的控制按鈕列：
  - 「Switch to Inspect Mode」按鈕
  - 「Always On Top: OFF」按鈕
  - 「Mode: Normal」標籤
- ✅ 下方有兩個標籤頁：
  - 「Excel Formula Comparator」
  - 「Workspace」

#### **如果失敗**
- 檢查錯誤訊息
- 確認 `core/mode_manager.py` 是否存在
- 檢查 Python 路徑是否正確

---

### **測試 2: Normal Mode 功能測試**

#### **執行步驟**
1. 確認程式在 Normal Mode（預設）
2. 測試 Excel Formula Comparator 標籤頁
3. 測試 Workspace 標籤頁
4. 嘗試使用現有功能

#### **預期結果**
- ✅ Excel Formula Comparator 功能完全正常
- ✅ Workspace 功能完全正常
- ✅ 所有按鈕和功能與原版相同
- ✅ 無任何功能缺失或異常

#### **如果失敗**
- 對比原版功能
- 檢查是否有組件初始化問題

---

### **測試 3: 模式切換測試**

#### **執行步驟**
1. 點擊「Switch to Inspect Mode」按鈕
2. 觀察視窗變化
3. 點擊「Switch to Normal Mode」按鈕
4. 觀察是否回到原狀態

#### **預期結果**

**切換到 Inspect Mode 時**：
- ✅ 視窗大小變為：1200x600
- ✅ 視窗標題變為：「Excel Tools - Inspect Mode」
- ✅ 按鈕文字變為：「Switch to Normal Mode」
- ✅ 模式標籤變為：「Mode: Inspect」
- ✅ 出現佔位內容：
  ```
  INSPECT MODE
  
  Dual-pane worksheet comparison
  coming soon...
  
  This will include:
  • Left and Right worksheet panels
  • Simplified interface
  • Focus on single cell analysis
  ```

**切換回 Normal Mode 時**：
- ✅ 視窗大小恢復：900x1100
- ✅ 視窗標題恢復：「Excel Tools - Integrated」
- ✅ 按鈕文字恢復：「Switch to Inspect Mode」
- ✅ 模式標籤恢復：「Mode: Normal」
- ✅ 原有標籤頁恢復正常

#### **如果失敗**
- 檢查視窗大小是否正確調整
- 確認標籤頁內容是否正確切換
- 檢查按鈕文字是否更新

---

### **測試 4: Always On Top 功能測試**

#### **執行步驟**
1. 點擊「Always On Top: OFF」按鈕
2. 嘗試切換到其他應用程式
3. 再次點擊按鈕關閉功能
4. 在兩種模式下都測試此功能

#### **預期結果**
- ✅ 點擊後按鈕文字變為：「Always On Top: ON」
- ✅ 視窗保持在所有其他視窗之上
- ✅ 再次點擊後按鈕文字變為：「Always On Top: OFF」
- ✅ 視窗恢復正常層級行為
- ✅ 在 Normal 和 Inspect 模式下都正常運作

#### **如果失敗**
- 檢查 Windows 系統是否支援 topmost 屬性
- 確認按鈕狀態是否正確更新

---

### **測試 5: 視窗大小和位置測試**

#### **執行步驟**
1. 在 Normal Mode 下調整視窗大小和位置
2. 切換到 Inspect Mode
3. 再切換回 Normal Mode
4. 檢查視窗狀態

#### **預期結果**
- ✅ 模式切換時視窗大小正確調整
- ✅ 視窗位置合理（不會跑到螢幕外）
- ✅ 視窗可以正常調整大小
- ✅ 切換模式不會造成視窗異常

---

## **測試報告格式**

請按以下格式回報測試結果：

```
測試日期：2025-01-XX
測試人員：[你的名字]

測試 1: 基本啟動測試
狀態：✅ 通過 / ❌ 失敗
備註：[如有問題請詳述]

測試 2: Normal Mode 功能測試
狀態：✅ 通過 / ❌ 失敗
備註：[如有問題請詳述]

測試 3: 模式切換測試
狀態：✅ 通過 / ❌ 失敗
備註：[如有問題請詳述]

測試 4: Always On Top 功能測試
狀態：✅ 通過 / ❌ 失敗
備註：[如有問題請詳述]

測試 5: 視窗大小和位置測試
狀態：✅ 通過 / ❌ 失敗
備註：[如有問題請詳述]

整體評價：
□ 所有功能正常，可以繼續下一階段
□ 有小問題但不影響繼續開發
□ 有重大問題需要修復

其他發現：
[任何額外的觀察或建議]
```

---

## **故障排除**

### **常見問題**

#### **問題 1: 程式無法啟動**
- 檢查 Python 環境
- 確認所有依賴模組已安裝
- 檢查檔案路徑是否正確

#### **問題 2: 模式切換無效果**
- 檢查 ModeManager 是否正確初始化
- 確認回調函數是否正確註冊
- 檢查視窗配置是否正確應用

#### **問題 3: 視窗大小異常**
- 檢查螢幕解析度
- 確認視窗大小配置是否合理
- 檢查是否有多螢幕設定影響

#### **問題 4: Always On Top 無效**
- 確認 Windows 版本支援
- 檢查是否有其他程式干擾
- 確認屬性設定是否正確

---

## **下一步行動**

### **如果測試通過**
- 更新進度日誌
- 開始階段 3：UI 組件開發
- 開始實現真正的 Inspect Mode 功能

### **如果測試失敗**
- 記錄具體問題
- 修復發現的錯誤
- 重新測試直到通過

---

**測試指引版本**: 1.0  
**創建日期**: 2025-01-XX  
**適用階段**: 階段 2 完成後的驗證測試