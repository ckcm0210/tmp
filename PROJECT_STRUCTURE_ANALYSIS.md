# 專案結構分析與重構建議

## 檔案結構樹狀圖

以下是專案的結構分析，說明了每個主要檔案的職責和技術。

-   **`main.py`**
    -   **功能**: 應用程式的進入點。負責建立主視窗 (`tk.Tk`) 和筆記本介面 (`ttk.Notebook`)，並將 `ExcelFormulaComparator` 和 `Workspace` 兩個主要功能模組載入為不同的分頁。
    -   **技術**: `tkinter`, `ttk`.

-   **`formula_comparator.py`**
    -   **功能**: 提供了並排比較兩個 Excel 工作表公式的核心介面。它管理兩個獨立的 `WorksheetController` 實例（左和右），並處理它們之間的同步邏輯。
    -   **核心**: `ExcelFormulaComparator` 類別，負責佈局和高階邏輯（如 `scan_worksheet_full`, `sync_1_to_2`）。
    -   **技術**: `tkinter`, `ttk`, `win32com.client`.

-   **`worksheet_ui.py`**
    -   **功能**: 定義了單個工作表分析面板的完整 UI 介面 (`WorksheetUI` 類別)。包含所有視覺元件的建立、佈局和事件綁定，例如按鈕、篩選器輸入框、以及顯示公式的樹狀檢視 (`ttk.Treeview`)。
    -   **職責**: 純粹的視圖 (View) 層，負責顯示資料和接收使用者輸入。
    -   **技術**: `tkinter`, `ttk`.

-   **`ui/worksheet/controller.py`** (檔案未提供，但從程式碼推斷)
    -   **功能**: `WorksheetController` 類別，作為 MVC 模式中的控制器。它儲存應用程式的狀態（如 `all_formulas` 列表、篩選條件）並持有對 `view` (`WorksheetUI`) 和模型層 (核心邏輯) 的引用。
    -   **職責**: 處理業務邏輯，更新模型，並透過 `view` 來更新介面。

-   **`worksheet_tree.py`**
    -   **功能**: 包含與 `WorksheetUI` 中樹狀檢視 (`Treeview`) 互動的所有邏輯。
    -   **核心函式**:
        -   `apply_filter`: 根據使用者輸入的篩選條件過濾 `all_formulas` 列表並更新 `Treeview`。
        -   `sort_column`: 處理欄位排序。
        -   `on_select`, `on_double_click`: 處理使用者在 `Treeview` 上的選擇和雙擊事件。
        -   `go_to_reference_new_tab`: 在 Excel 中定位到參照的儲存格，並在 UI 中開啟新的詳情分頁。
    -   **技術**: `tkinter`, `win32com.client`, 正則表達式 (`re`).

-   **`core/` (核心商業邏輯)**
    -   **`excel_scanner.py`**:
        -   **功能**: 負責與 Excel 直接溝通以掃描和提取公式。
        -   **核心函式**: `refresh_data` (高階 UI 互動函式), `_get_formulas_from_excel` (純掃描邏輯)。
    -   **`excel_connector.py`**:
        -   **功能**: 管理與 Excel 應用程式的 COM 連線。
        -   **核心函式**: `reconnect_to_excel`, `activate_excel_window`。
    -   **`link_analyzer.py`**:
        -   **功能**: 負責解析公式字串，識別其中的本機、外部和儲存格參照。
        -   **核心函式**: `get_referenced_cell_values`。
    -   **`formula_classifier.py`**:
        -   **功能**: 根據公式內容將其分類為 `formula`, `local link` 或 `external link`。
    -   **`data_processor.py`**:
        -   **功能**: 為「Summarize External Links」功能提供資料處理，例如從 `Treeview` 中提取資料並找出獨特的外部連結。

-   **`utils/` (輔助工具)**
    -   **`excel_helpers.py`**:
        -   **功能**: 提供高階的 Excel 操作函式，這些函式通常由 UI 事件觸發。
        -   **核心函式**: `select_ranges_in_excel`, `replace_links_in_excel`。
    -   **`excel_io.py`**:
        -   **功能**: 提供與 Excel 檔案的底層 I/O 操作，特別是使用 `openpyxl` 或 `xlrd` 在非 COM 環境下讀取外部檔案的儲存格數值。
    -   **`range_optimizer.py`**:
        -   **功能**: 提供 Excel 範圍位址的解析和最佳化顯示功能，例如將 `A1, A2, A3` 顯示為 `A1:A3`。

-   **`ui/` (使用者介面)**
    -   **`summary_window.py`**:
        -   **功能**: 定義了點擊「Summarize External Links」後彈出的摘要視窗。
    -   **`visualizer.py`**:
        -   **功能**: 提供了將公式參照關係視覺化為圖表的功能。
        -   **技術**: `matplotlib`。

---

## 程式碼重構檢視與建議

經過近期的密集除錯和重構，程式碼的穩定性和結構已有顯著改善。特別是將核心邏輯（如掃描、連線）分離到 `core/` 目錄中，是正確的方向。然而，仍有以下幾個主要的改善空間：

### 1. 嚴格遵守職責分離 (Separation of Concerns)

-   **問題**: 我們遇到的絕大多數 `AttributeError` 都源於 `utils/` 和 `core/` 中的函式直接存取 `controller.view` 上的 UI 元件（如 `result_tree`, `progress_bar`）。這違反了分層架構的基本原則：**核心邏輯層不應該知道任何關於 UI 的細節**。
-   **建議**:
    -   修改 `utils/excel_helpers.py` 中的函式，例如 `replace_links_in_excel`。它應該被拆分為兩部分：
        1.  一個位於 `core/` 的**核心取代函式**，它只接收必要的資料（例如 `(舊連結, 新連結)` 的列表），執行取代操作，然後返回結果（成功數量、失敗數量）。這個函式**不應**包含任何 `messagebox` 或進度條更新。
        2.  保留在 `utils/excel_helpers.py` 或移至 UI 相關模組的**介面互動函式**，它負責呼叫核心函式，並根據返回的結果來更新進度條和顯示提示訊息。
    -   `core/excel_scanner.py` 中的 `refresh_data` 函式也混合了大量的 UI 更新邏輯。雖然它已經將掃描部分拆分到 `_get_formulas_from_excel`，但整個流程的 UI 控制依然很重。長遠來看，可以考慮使用回呼函式 (callback) 或事件驅動模型，讓核心邏輯在執行到特定階段時發出訊號，由 UI 層監聽並更新介面。

### 2. 拆分大型、複雜的檔案

-   **問題**: 有幾個檔案的職責過於繁重，導致程式碼行數過多，難以維護。
-   **建議**:
    -   **`worksheet_tree.py` (最需要重構的檔案)**:
        -   這個檔案目前超過 500 行，混合了 UI 事件處理 (`on_select`, `on_double_click`)、資料處理 (`apply_filter`) 和複雜的 Excel 互動邏輯 (`go_to_reference_new_tab`)。
        -   建議將其拆分為：
            -   `ui/tree_handler.py`: 專門處理 `Treeview` 的 UI 事件綁定和回呼函式。
            -   `core/filter_logic.py`: 負責 `apply_filter` 的核心篩選邏輯。
            -   `core/navigation.py`: 負責 `go_to_reference` 等與 Excel 互動的導航功能。
    -   **`utils/excel_helpers.py`**:
        -   如上所述，`replace_links_in_excel` 函式非常龐大（超過 200 行），包含了大量的 UI 邏輯、檔案 I/O 驗證和核心的取代操作。應將其核心的取代邏輯提取到 `core` 層。

### 3. 統一參數命名和函式呼叫風格

-   **問題**: 在不同的模組中，傳入的 `controller` 物件被賦予了不同的參數名，例如 `self`, `pane`, `controller`。這在除錯時造成了很大的困擾。
-   **建議**:
    -   在所有獨立的函式（非類別方法）中，統一使用 `controller` 作為接收 `WorksheetController` 實例的參數名。這將大大提高程式碼的可讀性。
    -   例如，`worksheet_export.py` 中的 `export_formulas_to_excel(self)` 應改為 `export_formulas_to_excel(controller)`，這點我們在除錯過程中已經修正了一部分，建議全面審查並統一。

總結來說，目前的重構已經打下了良好的基礎。下一步的重點應該是**嚴格劃清 UI 和核心邏輯的界線**，並**持續拆分大型模組**，使其職責更加單一。

---

### 4. 具體的重構執行計畫 (Action Plan)

根據以上分析，建議按以下優先級順序執行重構任務：

-   **P0 (最高優先級): 核心與 UI 的徹底分離**
    1.  **任務**: 重構 `utils/excel_helpers.py` 中的 `replace_links_in_excel` 函式。
        -   **步驟**:
            1.  在 `core/` 目錄下建立 `excel_writer.py` 或類似檔案。
            2.  建立一個新的核心函式 `core_replace_links(target_workbook, link_map)`，該函式只負責 Excel 操作，並返回一個包含結果的字典 (e.g., `{'success': 5, 'failed': 1}`)。
            3.  修改 `utils/excel_helpers.py` 中的原函式，使其呼叫 `core_replace_links`，並處理所有 UI 互動 (進度條、訊息框)。
        -   **目標**: 消除 `core` 層對 `tkinter` 的任何依賴。

    2.  **任務**: 重構 `core/excel_scanner.py` 中的 `refresh_data` 函式。
        -   **步驟**:
            1.  將 `refresh_data` 重新命名為 `ui_refresh_data` 並保留在 `WorksheetController` 中或其可存取的地方。
            2.  確保 `_get_formulas_from_excel` 是純粹的資料掃描器，不觸發任何 UI 更新。
            3.  由 `ui_refresh_data` 函式負責呼叫掃描器，並在獲取資料後更新 UI 元件。
        -   **目標**: 確保資料掃描過程可以被獨立呼叫和測試。

-   **P1 (次高優先級): 拆分 `worksheet_tree.py`**
    1.  **任務**: 將 `worksheet_tree.py` 的職責拆分出去。
        -   **步驟**:
            1.  建立 `ui/tree_handler.py`，將所有 `on_select`, `on_double_click`, `sort_column` 等事件處理函式移入。
            2.  建立 `core/filter_logic.py`，將 `apply_filter` 的邏輯移入。
            3.  建立 `core/navigation.py`，將 `go_to_reference_new_tab` 的邏輯移入。
            4.  `worksheet_tree.py` 本身可以廢棄，或只保留 `Treeview` 的初始化設定。
        -   **目標**: 讓每個模組的程式碼行數少於 200 行，且職責單一。

-   **P2 (中等優先級): 程式碼風格與一致性**
    1.  **任務**: 全域審查並統一 `controller` 的參數命名。
        -   **步驟**: 使用 IDE 的全域搜尋功能，找出所有將 `WorksheetController` 實例作為參數的函式，並將其參數名統一為 `controller`。
        -   **目標**: 提高程式碼可讀性，降低新進人員的理解成本。

### 5. 預期成果與衡量指標 (Expected Outcomes & Metrics)

-   **可維護性 (Maintainability)**:
    -   **指標**: 修改一個功能（例如：變更篩選邏輯）時，需要接觸的檔案數量從 3-4 個（混合了 UI 和邏輯）減少到 1-2 個（只需修改 `core/filter_logic.py`）。
    -   **指標**: 新功能開發時間縮短，因為業務邏輯和 UI 邏輯可以並行開發。

-   **穩定性 (Stability)**:
    -   **指標**: `AttributeError` 或因 UI 與邏輯耦合導致的 `NoneType` 錯誤數量顯著減少。
    -   **指標**: 核心邏輯可以被獨立測試，有助於在早期發現並修復錯誤。

-   **可測試性 (Testability)**:
    -   **指標**: `core/` 目錄下的所有函式都可以編寫單元測試 (Unit Test)，而無需模擬 (mock) 整個 `tkinter` UI。
    -   **目標**: 實現對核心業務邏輯至少 50% 的測試覆蓋率。

-   **開發者體驗 (Developer Experience)**:
    -   **指標**: 新進開發者理解專案結構所需的時間從數天縮短到數小時。
    -   **指標**: 由於職責清晰，除錯時定位問題的速度顯著加快。
