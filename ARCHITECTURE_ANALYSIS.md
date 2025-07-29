# 架構分析報告 - Inspect Mode 實施

## **現有架構深度分析**

### **1. 主程式架構 (main.py)**

#### **程式進入點結構**
```python
def main():
    root = tk.Tk()
    root.title("Excel Tools - Integrated")
    root.geometry("900x1100")  # 現有視窗大小
    
    notebook = ttk.Notebook(root)
    
    # 兩個主要標籤頁
    comparator_frame = ttk.Frame(notebook)
    notebook.add(comparator_frame, text='Excel Formula Comparator')
    ExcelFormulaComparator(comparator_frame, root)
    
    workspace_frame = ttk.Frame(notebook)
    notebook.add(workspace_frame, text='Workspace')
    Workspace(workspace_frame)
```

#### **關鍵發現**
- ✅ 使用 ttk.Notebook 管理標籤頁
- ✅ 兩個主要組件：ExcelFormulaComparator 和 Workspace
- ✅ 視窗大小固定為 900x1100
- ✅ 每個組件都接收 parent_frame 和 root 參數

---

### **2. Workspace 類別架構**

#### **類別結構**
```python
class Workspace:
    def __init__(self, parent_frame):
        self.parent = parent_frame
        self.root = parent_frame.winfo_toplevel()
        
        # 資料屬性
        self.file_names = []
        self.file_paths = []
        self.sheet_names = []
        self.active_cells = []
        self.showing_path = False
        self.target_captions = []
        
        self.setup_ui()
        self.show_names()
```

#### **UI 組件結構**
1. **標題區域** (`title_frame`)
   - 主標籤："Current Workspace:"
   - 計數標籤：顯示工作簿數量

2. **主要區域** (`main_frame`)
   - **左側**：檔案列表 (`AccumulateListbox`)
   - **右側**：操作按鈕區 (`side_btn_frame`)

3. **按鈕功能**
   - Refresh, Show Full Path, Activate Workbook
   - Minimize All, Save/Load Workspace
   - Save/Close Workbook 操作

#### **關鍵發現**
- ✅ Workspace 是獨立的工作區管理器
- ✅ 主要功能是管理 Excel 工作簿
- ✅ 使用自定義的 AccumulateListbox
- ✅ 與 Excel COM 對象交互

---

### **3. WorksheetController 架構**

#### **控制器結構**
```python
class WorksheetController:
    def __init__(self, parent_frame, root_app, pane_name):
        self.root = root_app
        self.pane_name = pane_name
        
        # Excel 連接
        self.xl = None
        self.workbook = None
        self.worksheet = None
        
        # 資料管理
        self.all_formulas = []
        self.cell_addresses = {}
        
        # UI 狀態
        self.use_openpyxl = tk.BooleanVar(value=True)
        self.show_formula = tk.BooleanVar(value=True)
        self.show_local_link = tk.BooleanVar(value=True)
        self.show_external_link = tk.BooleanVar(value=True)
        
        # 創建視圖和標籤管理器
        self.view = WorksheetView(parent_frame, self)
        self.tab_manager = TabManager(self.view.detail_notebook)
```

#### **關鍵發現**
- ✅ 使用 MVC 模式：Controller + View + TabManager
- ✅ 管理 Excel 連接和工作表資料
- ✅ 包含篩選和顯示狀態管理
- ✅ 支援多標籤詳細資訊顯示

---

## **Inspect Mode 架構設計**

### **4. 模式切換架構設計**

#### **建議的架構改動**
```
main.py (修改)
├── ModeManager (新增)
│   ├── current_mode: "normal" | "inspect"
│   ├── switch_mode()
│   └── get_window_config()
│
├── Normal Mode (重構現有)
│   ├── ExcelFormulaComparator
│   └── Workspace
│
└── Inspect Mode (新增)
    ├── LeftPane (WorksheetController)
    ├── RightPane (WorksheetController)
    ├── ModeControls (Always On Top, Switch Mode)
    └── CompactLayout
```

#### **視窗大小策略**
- **Normal Mode**: 900x1100 (維持現有)
- **Inspect Mode**: 1200x600 (橫向雙面板)
- **Always On Top**: 可切換的頂層顯示

---

### **5. 實施策略**

#### **階段 2A: 創建模式管理器**
```python
# core/mode_manager.py
class ModeManager:
    def __init__(self, root_window):
        self.root = root_window
        self.current_mode = "normal"
        self.mode_configs = {
            "normal": {"size": "900x1100", "layout": "single"},
            "inspect": {"size": "1200x600", "layout": "dual_pane"}
        }
    
    def switch_mode(self, new_mode):
        # 切換模式邏輯
        pass
    
    def apply_window_config(self, mode):
        # 應用視窗配置
        pass
```

#### **階段 2B: 創建雙面板控制器**
```python
# core/dual_pane_controller.py
class DualPaneController:
    def __init__(self, parent_frame):
        self.left_pane = WorksheetController(left_frame, root, "Left")
        self.right_pane = WorksheetController(right_frame, root, "Right")
    
    def create_layout(self):
        # 創建左右面板布局
        pass
```

#### **階段 2C: 創建視窗管理器**
```python
# core/window_manager.py
class WindowManager:
    def __init__(self, root_window):
        self.root = root_window
        self.always_on_top = False
    
    def set_always_on_top(self, enabled):
        # Always On Top 功能
        pass
    
    def resize_window(self, size):
        # 調整視窗大小
        pass
```

---

## **依賴關係分析**

### **現有依賴鏈**
```
main.py
├── formula_comparator.py (ExcelFormulaComparator)
└── ui.workspace_view.Workspace
    └── AccumulateListbox (自定義組件)

WorksheetController
├── ui.worksheet.view.WorksheetView
├── ui.worksheet.tab_manager.TabManager
└── 多個 worksheet_*.py 模組
```

### **新增依賴鏈 (建議)**
```
main.py (修改)
├── core.mode_manager.ModeManager (新增)
├── core.window_manager.WindowManager (新增)
└── ui.modes.normal_mode.NormalMode (重構)
    ├── formula_comparator.py
    └── ui.workspace_view.Workspace

main.py (Inspect Mode)
└── ui.modes.inspect_mode.InspectMode (新增)
    ├── core.dual_pane_controller.DualPaneController
    ├── ui.components.mode_switcher.ModeSwitcher
    └── ui.components.top_controls.TopControls
```

---

## **風險評估**

### **高風險區域**
1. **main.py 重構**
   - 風險：破壞現有程式進入點
   - 緩解：保持向後兼容，新增模式作為選項

2. **視窗大小切換**
   - 風險：UI 組件可能不適應新尺寸
   - 緩解：使用響應式布局設計

3. **雙面板資源管理**
   - 風險：兩個 WorksheetController 可能衝突
   - 緩解：確保完全獨立的資料和狀態

### **低風險區域**
1. **新檔案創建** - 不影響現有功能
2. **Always On Top** - 獨立功能
3. **模式切換 UI** - 可選功能

---

## **下一步行動建議**

### **立即執行 (階段 2A)**
1. 創建 `core/mode_manager.py`
2. 修改 `main.py` 支援模式管理
3. 測試基本的模式切換邏輯

### **後續執行**
1. 實現視窗管理器
2. 創建雙面板控制器
3. 開發 Inspect Mode UI

---

**分析完成時間**: 2025-01-XX  
**下一階段**: 開始實施階段 2A - 創建模式管理器