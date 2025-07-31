

import os
from pyvis.network import Network
import re

# --- 1. 設計一個更複雜、真實的依賴關係場景 ---
nodes_data = [
    {
        "id": 0, "shape": 'box',
        "short_label": "Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0",
        "color": "#e04141"
    },
    {
        "id": 1, "shape": 'box',
        "short_label": "Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0",
        "color": "#007bff"
    },
    {
        "id": 2, "shape": 'box',
        "short_label": "Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0",
        "color": "#007bff"
    },
    {
        "id": 3, "shape": 'box',
        "short_label": "Data!B3\nType: Value\nValue: 5225.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!B3\nType: Value\nValue: 5225.0",
        "color": "#28a745"
    },
    {
        "id": 4, "shape": 'box',
        "short_label": "Data!C1\nType: Value\nValue: SKU-001",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!C1\nType: Value\nValue: SKU-001",
        "color": "#28a745"
    },
    {
        "id": 5, "shape": 'box',
        "short_label": "RefData!B1\nType: Value\nValue: 500.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]RefData!B1\nType: Value\nValue: 500.0",
        "color": "#28a745"
    },
    {
        "id": 6, "shape": 'box',
        "short_label": "[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0",
        "full_label": "C:\\Linked Files\\[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0",
        "color": "#ff8c00"
    }
]

edges_data = [
    (0, 1), (0, 2), (0, 3),  # Final_Report!A1 depends on B1, B2, B3
    (1, 4), (1, 5),          # Data!B1 (VLOOKUP) depends on C1 and RefData!B1
    (2, 6)                   # Data!B2 depends on the external file
]

# --- 2. 使用 Pyvis 產生基礎圖表 ---
net = Network(height="90vh", width="100%", bgcolor="#ffffff", font_color="black", directed=True)
net.toggle_physics(False) # 關鍵：關閉物理引擎

# 將節點資料加入網路圖
for node_info in nodes_data:
    net.add_node(
        node_info["id"],
        label=node_info["short_label"], # 預設顯示簡化標籤
        shape=node_info["shape"],
        color=node_info["color"],
        # 將完整和簡化標籤儲存在節點的自訂屬性中，供 JavaScript 使用
        full_label=node_info["full_label"],
        short_label=node_info["short_label"]
    )

# 加入邊
for edge in edges_data:
    net.add_edge(edge[0], edge[1])

# 產生一個暫時的 HTML 檔案
temp_file = "temp_graph.html"
net.save_graph(temp_file)

# --- 3. 注入 HTML 和 JavaScript 以增加互動性 ---
with open(temp_file, 'r', encoding='utf-8') as f:
    html_content = f.read()

# 準備要注入的 HTML 控制項
checkbox_html = """
<div style='position: absolute; top: 10px; left: 10px; background: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; z-index: 1000;'>
  <label for='pathToggle' style='font-family: sans-serif; font-size: 14px;'>
    <input type='checkbox' id='pathToggle'>
    顯示完整檔案路徑
  </label>
</div>
"""

# 準備要注入的 JavaScript
javascript_injection = """
<script type='text/javascript'>
  // 等待網路圖完全初始化
  document.addEventListener('DOMContentLoaded', function() {
    var network = window.network; // 從全域範圍獲取 pyvis 建立的 network 物件
    var nodes = window.nodes;     // 獲取 nodes DataSet

    if (!network || !nodes) {
        console.error("Network or nodes not found!");
        return;
    }

    const pathToggle = document.getElementById('pathToggle');

    pathToggle.addEventListener('change', function() {
      const showFullPath = this.checked;
      let updatedNodes = [];

      nodes.forEach(node => {
        let newLabel = showFullPath ? node.full_label : node.short_label;
        updatedNodes.push({ id: node.id, label: newLabel });
      });

      // 一次性更新所有節點的標籤
      nodes.update(updatedNodes);
    });
  });
</script>
"""

# 將控制項和腳本注入到 HTML 中
# 將 checkbox 插入到 <body> 標籤之後
html_content = html_content.replace('<body>', '<body>\n' + checkbox_html)
# 將 JavaScript 插入到 </body> 標籤之前
html_content = html_content.replace('</body>', javascript_injection + '\n</body>')

# 寫入最終的 HTML 檔案
final_file = "final_interactive_graph.html"
with open(final_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

# 刪除暫存檔案
os.remove(temp_file)

print(f"Successfully generated final interactive graph at: {os.path.join(os.getcwd(), final_file)}")

