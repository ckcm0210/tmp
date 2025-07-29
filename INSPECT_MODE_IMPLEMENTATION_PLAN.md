# Inspect Mode 功能實施計劃書

## **專案概述**

### **功能目標**
實現 Inspect Mode 功能，提供雙面板比較模式，用於深度分析和比較兩個相似的 Excel 檔案。

### **核心需求**
1. **模式切換**：Normal Mode ↔ Inspect Mode 雙向切換
2. **介面簡化**：隱藏非必要控制項，專注於單一儲存格分析
3. **雙面板設計**：左右兩個獨立的工作面板
4. **視窗優化**：縮小視窗大小，不阻礙日常工作
5. **Always On Top**：保持視窗在最頂層

---

## **行動守則 (Action Guidelines)**

### **A. 使用者指定的核心原則**
1. **進度報告**：在每個階段完成後更新進度
2. **小步進行，逐步驗證**：所有重構任務拆解成最小可驗證單元
3. **專注於程式碼遷移，而非修改**：主要任務是「搬移」程式碼到新結構
4. **檢查並修復依賴性**：每次遷移後立即檢查並修復模組導入錯誤

### **B. 額外核心原則**
1. **保持系統可運作**：每完成一個微小步驟後，系統都應處於可運作狀態
2. **測試驅動重構**：在關鍵重構點後進行驗證
3. **手動版本備份**：每完成一個成功步驟後進行資料夾備份
4. **提供詳細測試指引**：每次修改後提供清晰的驗證步驟
5. **日誌驅動工作流程**：執行操作前先將行動提議及理由更新至日誌

### **C. 業界專業重構準則**
1. **不要混合重構與功能開發**：一次只做一件事
2. **先理解，後重構**：不重構自己不理解的程式碼
3. **有明確的重構目標**：目標是實現 Inspect Mode 的 MVC 架構
4. **嚴格區分移動與修改**：任何「修改」都必須被明確聲明並獲得批准

---

## **檔案結構設計**

### **現有檔案結構 (✓ 表示現有檔案)**
```
Excel_tools_develop_v38/
├── ✓ !run_me.ipynb
├── ✓ ACTION_GUIDELINES.md
├── ✓ excel_utils.py
├── ✓ formula_comparator.py
├── ✓ main.py                           # 需修改：支援新架構
├── ✓ PROGRESS_LOG.md
├── ✓ PROJECT_STRUCTURE_ANALYSIS.md
├── ✓ REFACTORING_GUIDE.md
├── ✓ untitled.txt
├── ✓ worksheet_export.py
├── ✓ worksheet_refresh.py
├── ✓ worksheet_summary.py
├── ✓ worksheet_tree.py
├── ✓ worksheet_ui.py
├── 📄 INSPECT_MODE_IMPLEMENTATION_PLAN.md  # 新增：本文檔
│
├── core/
│   ├── ✓ data_processor.py
│   ├── ✓ excel_connector.py
│   ├── ✓ excel_scanner.py
│   ├── ✓ formula_classifier.py
│   ├── ✓ link_analyzer.py
│   ├── ✓ models.py
│   ├── 🆕 mode_manager.py              # 新增：模式管理器
│   ├── 🆕 dual_pane_controller.py      # 新增：雙面板控制器
│   └── 🆕 window_manager.py            # 新增：視窗管理器
│
├── ui/
│   ├── ✓ summary_window.py
│   ├── ✓ visualizer.py
│   ├── ✓ workspace_view.py
│   ├── modes/                          # 新增：模式相關UI
│   │   ├── 🆕 __init__.py
│   │   ├── 🆕 base_mode.py             # 新增：模式基類
│   │   ├── 🆕 normal_mode.py           # 新增：正常模式UI
│   │   └── 🆕 inspect_mode.py          # 新增：檢查模式UI
│   ├── components/                     # 新增：UI組件
│   │   ├── 🆕 __init__.py
│   │   ├── 🆕 pane_widget.py           # 新增：面板組件
│   │   ├── 🆕 mode_switcher.py         # 新增：模式切換器
│   │   └── 🆕 top_controls.py          # 新增：頂部控制項
│   └── worksheet/                      # 現有目錄
│       ├── ✓ __init__.py
│       ├── ✓ controller.py
│       ├── ✓ tab_manager.py
│       └── ✓ view.py
│
└── utils/
    ├── ✓ excel_helpers.py
    ├── ✓ excel_io.py
    ├── ✓ helpers.py
    ├── ✓ range_optimizer.py
    └── 🆕 layout_manager.py             # 新增：布局管理器
```

### **圖例說明**
- ✓ **現有檔案**：已存在的檔案
- 🆕 **新增檔案**：需要創建的新檔案
- 📄 **文檔檔案**：說明文檔

---

## **實施階段計劃**

### **階段 1: 架構分析與設計 (2-3 iterations)**

#### **目標**
- 深入分析現有程式碼結構
- 設計新的 MVC 架構
- 規劃詳細的實施步驟

#### **具體任務**
1. **分析現有依賴關係**
   - 檢查 `main.py` 與其他模組的關係
   - 分析 `worksheet_ui.py` 的 UI 組件結構
   - 理解 `workspace_view.py` 的功能

2. **設計模式管理架構**
   - 規劃 `mode_manager.py` 的介面
   - 設計模式切換的狀態管理
   - 定義模式間的資料隔離策略

3. **規劃視窗布局**
   - Normal Mode: 900x1100 (現有大小)
   - Inspect Mode: 1200x600 (橫向雙面板)
   - Always On Top 功能設計

#### **驗證標準**
- [ ] 架構設計文檔完成
- [ ] 依賴關係圖清晰
- [ ] 實施步驟詳細且可執行

---

### **階段 2: 核心架構重構 (5-7 iterations)**

#### **目標**
- 創建核心管理組件
- 建立模式切換基礎架構
- 實現雙面板控制邏輯

#### **具體任務**

##### **2.1 創建模式管理器 (1-2 iterations)**
- 創建 `core/mode_manager.py`
- 實現模式狀態管理
- 定義模式切換介面

##### **2.2 創建視窗管理器 (1-2 iterations)**
- 創建 `core/window_manager.py`
- 實現視窗大小調整
- 實現 Always On Top 功能

##### **2.3 創建雙面板控制器 (2-3 iterations)**
- 創建 `core/dual_pane_controller.py`
- 實現獨立面板管理
- 確保面板間資料隔離

#### **驗證標準**
- [ ] 所有核心組件可成功導入
- [ ] 基本的模式切換邏輯運作
- [ ] 視窗管理功能正常

---

### **階段 3: UI 組件開發 (4-5 iterations)**

#### **目標**
- 開發模式相關的 UI 組件
- 實現具體的介面布局
- 整合現有功能到新架構

#### **具體任務**

##### **3.1 創建基礎 UI 組件 (1-2 iterations)**
- 創建 `ui/modes/base_mode.py`
- 創建 `ui/components/mode_switcher.py`
- 創建 `ui/components/top_controls.py`

##### **3.2 開發 Normal Mode UI (1-2 iterations)**
- 創建 `ui/modes/normal_mode.py`
- 遷移現有 UI 邏輯
- 確保向後兼容

##### **3.3 開發 Inspect Mode UI (1-2 iterations)**
- 創建 `ui/modes/inspect_mode.py`
- 創建 `ui/components/pane_widget.py`
- 實現雙面板布局

#### **驗證標準**
- [ ] Normal Mode 功能完全正常
- [ ] Inspect Mode 基本布局正確
- [ ] 模式切換無錯誤

---

### **階段 4: 功能整合與測試 (3-4 iterations)**

#### **目標**
- 整合所有新功能
- 全面測試系統穩定性
- 優化用戶體驗

#### **具體任務**

##### **4.1 功能整合 (1-2 iterations)**
- 整合 Always On Top 功能
- 完善模式切換邏輯
- 測試雙面板獨立性

##### **4.2 性能優化 (1 iteration)**
- 優化視窗切換性能
- 減少記憶體使用
- 改善響應速度

##### **4.3 全面測試 (1 iteration)**
- 測試所有功能組合
- 驗證錯誤處理
- 確保系統穩定性

#### **驗證標準**
- [ ] 所有功能正常運作
- [ ] 性能符合預期
- [ ] 無明顯錯誤或崩潰

---

## **風險評估與緩解策略**

### **高風險項目**

#### **1. UI 重構複雜度**
- **風險**：大量 UI 程式碼需要重新組織
- **緩解**：分階段實施，保持向後兼容
- **備案**：保留原有 UI 作為備用

#### **2. 雙面板資源管理**
- **風險**：兩個面板可能產生資源衝突
- **緩解**：獨立的控制器和資料管理
- **備案**：實現資源池管理

#### **3. 模式切換穩定性**
- **風險**：切換過程中可能出現狀態不一致
- **緩解**：充分測試所有切換場景
- **備案**：實現狀態恢復機制

### **緩解策略**

1. **版本備份**：每個階段完成後創建備份
2. **漸進式部署**：新功能作為可選模式添加
3. **回滾機制**：保持現有功能完全可用
4. **充分測試**：每個組件都要經過獨立測試

---

## **成功標準**

### **功能標準**
- [ ] Normal Mode 與現有功能完全一致
- [ ] Inspect Mode 正確實現雙面板布局
- [ ] 模式切換流暢無錯誤
- [ ] Always On Top 功能正常
- [ ] 視窗大小調整符合需求

### **品質標準**
- [ ] 程式碼結構清晰，易於維護
- [ ] 無明顯性能問題
- [ ] 錯誤處理完善
- [ ] 用戶體驗良好

### **技術標準**
- [ ] 遵循 MVC 架構原則
- [ ] 模組間耦合度低
- [ ] 程式碼可測試性高
- [ ] 文檔完整清晰

---

## **下一步行動**

### **立即執行**
1. 確認計劃書內容無誤
2. 開始階段 1：架構分析與設計
3. 創建第一個備份版本

### **等待確認**
- 視窗大小設計是否合適？
- 檔案結構規劃是否合理？
- 實施優先順序是否正確？

---

## **更新日誌**

| 日期 | 階段 | 狀態 | 備註 |
|------|------|------|------|
| 2025-01-XX | 計劃制定 | 完成 | 初始計劃書創建 |
| 2025-01-XX | 階段1-開始 | 完成 | 架構分析與設計完成 |
| | | | 創建 ARCHITECTURE_ANALYSIS.md 報告 |
| 2025-01-XX | 階段2A-開始 | 完成 | 模式管理器創建完成 |
| | | | core/mode_manager.py 已創建並測試 |
| 2025-01-XX | 階段2B-開始 | 完成 | main.py 整合模式管理器完成 |
| | | | 添加模式切換和 Always On Top 功能 |
| 2025-01-XX | 測試階段 | 完成 | 基本功能測試通過 |
| | | | 模式切換和視窗調整功能正常 |
| 2025-01-XX | 階段3-重做 | 完成 | 正確實現 Inspect Mode 功能 |
| | | | 重用現有 WorksheetController，隱藏不需要元素 |
| 2025-01-XX | 用戶反饋修正 | 完成 | 根據用戶要求完善 Inspect Mode |
| | | | 移除進度條、filters、不需要按鈕，調整布局 |
| 2025-01-XX | 功能測試 | 完成 | Inspect Mode 功能測試通過 |
| | | | 所有要求已正確實現並驗證 |
| 2025-01-XX | Close All Tabs | 完成 | 添加 Close All Tabs 按鈕功能 |
| | | | 與 Normal Mode 功能一致 |
| 2025-01-XX | 專案完成 | 完成 | Inspect Mode 完整實現完成 |
| | | | 所有功能正常運作 |

---

## **階段 1 進度記錄**

### **1.1 現有架構分析 (已完成)**

### **2.1 模式管理器開發 (已完成)**

#### **已完成的工作**
- ✅ 創建 `core/mode_manager.py`
- ✅ 實現 AppMode 枚舉 (NORMAL, INSPECT)
- ✅ 實現 ModeManager 類別
- ✅ 支援模式切換和視窗配置
- ✅ 實現 Always On Top 功能
- ✅ 添加模式切換回調機制
- ✅ 創建測試腳本驗證功能

#### **ModeManager 核心功能**
1. **模式管理**:
   - `switch_to_normal_mode()` / `switch_to_inspect_mode()`
   - `toggle_mode()` - 在兩種模式間切換
   - `is_normal_mode()` / `is_inspect_mode()` - 狀態檢查

2. **視窗配置**:
   - Normal Mode: 900x1100
   - Inspect Mode: 1200x600
   - 自動調整視窗標題和大小

3. **Always On Top**:
   - `set_always_on_top(enabled)` 
   - `toggle_always_on_top()`

4. **回調機制**:
   - `register_mode_switch_callback(callback)`
   - 支援多個組件監聽模式切換

### **2.2 主程式整合 (已完成)**

#### **已完成的工作**
- ✅ 創建 main.py 備份 (main_backup.py)
- ✅ 重構 main.py 為 ExcelToolsApp 類別
- ✅ 整合 ModeManager 到主程式
- ✅ 添加模式控制按鈕界面
- ✅ 實現模式切換邏輯
- ✅ 添加 Always On Top 功能
- ✅ 創建 Inspect Mode 佔位界面

#### **新增功能**
1. **模式控制界面**:
   - "Switch to Inspect Mode" / "Switch to Normal Mode" 按鈕
   - "Always On Top: ON/OFF" 切換按鈕
   - "Mode: Normal/Inspect" 狀態指示器

2. **ExcelToolsApp 類別**:
   - `setup_window()` - 視窗初始化
   - `setup_mode_controls()` - 模式控制界面
   - `setup_normal_mode()` - 正常模式 UI
   - `setup_inspect_mode()` - 檢查模式 UI (佔位)
   - `toggle_mode()` - 模式切換
   - `toggle_always_on_top()` - 頂層切換

3. **向後兼容性**:
   - 保持所有現有功能完整
   - Normal Mode 與原版完全相同
   - 新功能作為可選項添加

### **2.3 測試階段 (已完成)**

#### **測試結果**
- ✅ 基本啟動測試：通過
- ✅ Normal Mode 功能：完全正常
- ✅ 模式切換功能：正常運作
- ✅ Always On Top 功能：正常
- ✅ 視窗大小調整：正確 (900x1100 ↔ 1200x600)

#### **用戶反饋**
- ✅ 功能運作正常
- 🔄 需求調整：Inspect Mode 需要獨立的掃描功能
- 🔄 專注於單一儲存格掃描，不需要全 worksheet 掃描

### **3.1 Inspect Mode 掃描功能開發 (進行中)**

#### **新需求分析**
1. **獨立掃描功能**：
   - Inspect Mode 不依賴 Normal Mode 的掃描結果
   - 直接在 Inspect Mode 中提供掃描功能

2. **單一儲存格專注**：
   - 不掃描整個 worksheet
   - 只掃描用戶指定的單一儲存格
   - 類似 "Selected Range" 但更精確

3. **雙面板設計**：
   - 左右兩個獨立的掃描面板
   - 每個面板可以掃描不同檔案的不同儲存格
   - 便於比較分析

#### **已完成的開發（正確實現）**
- ✅ 創建 `ui/modes/inspect_mode.py`
- ✅ 實現 InspectMode 類別（重用現有組件）
- ✅ 實現 SimplifiedWorksheetController（繼承 WorksheetController）
- ✅ 整合到 main.py 的 Inspect Mode
- ✅ 創建雙面板布局（重用現有 UI）

#### **修正的實現方式**
- ✅ **遵循行動守則**：重用現有程式碼而非重新創作
- ✅ **繼承 WorksheetController**：保持所有現有功能
- ✅ **隱藏不需要的元素**：filters、summarize、export、import、reconnect 按鈕
- ✅ **保留核心功能**：go to reference、新 tab、details 等

#### **用戶反饋修正（已完成）**
- ✅ **移除進度條**：因為只掃描一個儲存格，不需要進度顯示
- ✅ **完全移除 Filters 區域**：address、formula、result、display_value 輸入框
- ✅ **移除所有不需要按鈕**：summarize、export、import、reconnect
- ✅ **調整 Formula List 大小**：只顯示一行，騰出更多空間
- ✅ **新增掃描按鈕**：類似 Normal Mode 的 Selected Range 功能
- ✅ **保持 Normal Mode 不變**：所有修改只影響 Inspect Mode

#### **新功能特點**
1. **雙面板布局**：
   - 使用 ttk.PanedWindow 實現可調整大小的左右面板
   - 每個面板獨立運作，互不干擾

2. **單一儲存格掃描**：
   - 用戶輸入儲存格地址（如 A1, B5）
   - 點擊 "Scan Cell" 掃描指定儲存格
   - 顯示詳細的儲存格分析結果

3. **Excel 連接管理**：
   - "Connect to Excel" 按鈕連接到活動的 Excel
   - 自動獲取當前工作簿和工作表
   - 顯示連接狀態和檔案資訊

4. **儲存格分析**：
   - 基本資訊：工作簿、工作表、儲存格地址
   - 數值資訊：顯示文字、計算值
   - 公式分析：公式內容、類型判斷
   - 支援外部連結、本地連結、純公式的識別

### **3.2 功能測試階段 (待進行)**

#### **已完成的分析**
- ✅ 檢查主要檔案結構
- ✅ 分析 main.py 的程式進入點
- ✅ 完成 ui/workspace_view.py 的架構分析
- ✅ 完成 ui/worksheet/controller.py 的控制邏輯分析
- ✅ 創建詳細架構分析報告 (ARCHITECTURE_ANALYSIS.md)

#### **發現的關鍵架構**
1. **主程式進入點** (`main.py`):
   - 使用 ttk.Notebook 管理兩個標籤頁
   - ExcelFormulaComparator 和 Workspace 兩個主要組件
   - 視窗大小: 900x1100，支援 topmost 屬性

2. **Workspace 類別**:
   - 獨立的工作區管理器，管理 Excel 工作簿
   - 使用自定義 AccumulateListbox 組件
   - 包含完整的工作簿操作功能

3. **WorksheetController 架構**:
   - 使用 MVC 模式：Controller + View + TabManager
   - 管理 Excel 連接、資料和 UI 狀態
   - 支援篩選、排序和多標籤詳細資訊

4. **模式切換最佳插入點**:
   - 在 main.py 的 ttk.Notebook 層級插入模式管理
   - 保持現有組件完整性，新增 Inspect Mode 作為第三個選項

#### **架構設計決策**
- ✅ 創建 ModeManager 管理模式切換
- ✅ 使用 DualPaneController 實現雙面板
- ✅ WindowManager 處理視窗大小和 Always On Top
- ✅ 保持現有程式碼完整性，採用組合而非修改策略

---

**文檔版本**: 1.0  
**最後更新**: 2025-01-XX  
**負責人**: Rovo Dev  
**審核狀態**: 待確認