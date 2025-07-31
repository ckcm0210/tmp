

import os
from pyvis.network import Network
import json

# --- 1. 使用與上次相同的複雜、真實依賴關係場景 ---
nodes_data = [
    {
        "id": 0, "shape": 'box',
        "short_label": "Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0",
        "color": "#e04141", "level": 0 # 最頂層
    },
    {
        "id": 1, "shape": 'box',
        "short_label": "Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0",
        "color": "#007bff", "level": 1
    },
    {
        "id": 2, "shape": 'box',
        "short_label": "Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0",
        "color": "#007bff", "level": 1
    },
    {
        "id": 3, "shape": 'box',
        "short_label": "Data!B3\nType: Value\nValue: 5225.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!B3\nType: Value\nValue: 5225.0",
        "color": "#28a745", "level": 1
    },
    {
        "id": 4, "shape": 'box',
        "short_label": "Data!C1\nType: Value\nValue: SKU-001",
        "full_label": "C:\\My Reports\\[Master.xlsx]Data!C1\nType: Value\nValue: SKU-001",
        "color": "#28a745", "level": 2
    },
    {
        "id": 5, "shape": 'box',
        "short_label": "RefData!B1\nType: Value\nValue: 500.0",
        "full_label": "C:\\My Reports\\[Master.xlsx]RefData!B1\nType: Value\nValue: 500.0",
        "color": "#28a745", "level": 2
    },
    {
        "id": 6, "shape": 'box',
        "short_label": "[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0",
        "full_label": "C:\\Linked Files\\[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0",
        "color": "#ff8c00", "level": 2 # 最底層
    }
]

edges_data = [
    (0, 1), (0, 2), (0, 3),
    (1, 4), (1, 5),
    (2, 6)
]

# --- 2. 使用 Pyvis 產生基礎圖表 ---
net = Network(height="90vh", width="100%", bgcolor="#ffffff", font_color="black", directed=True)

# 關鍵：設定階層式佈局與箭頭樣式
options_str = '''
{
  "layout": {
    "hierarchical": {
      "enabled": true,
      "direction": "DU", 
      "sortMethod": "directed",
      "nodeSpacing": 150,
      "treeSpacing": 220
    }
  },
  "edges": {
    "smooth": {
      "type": "cubicBezier",
      "forceDirection": "vertical",
      "roundness": 0.4
    }
  },
  "physics": {
      "enabled": false
  }
}
'''
net.set_options(options_str)

# 將節點資料加入網路圖
for node_info in nodes_data:
    net.add_node(
        node_info["id"],
        label=node_info["short_label"],
        shape=node_info["shape"],
        color=node_info["color"],
        level=node_info["level"], # 指定層級
        full_label=node_info["full_label"],
        short_label=node_info["short_label"]
    )

# 加入邊
for edge in edges_data:
    net.add_edge(edge[0], edge[1])

# --- 3. 注入與上次相同的 HTML 和 JavaScript 以增加互動性 ---
temp_file = "temp_final_graph.html"
net.save_graph(temp_file)

with open(temp_file, 'r', encoding='utf-8') as f:
    html_content = f.read()

checkbox_html = """
<div style='position: absolute; top: 10px; left: 10px; background: #f8f9fa; padding: 10px; border: 1px solid #dee2e6; border-radius: 5px; z-index: 1000;'>
  <label for='pathToggle' style='font-family: sans-serif; font-size: 14px;'>
    <input type='checkbox' id='pathToggle'>
    顯示完整檔案路徑
  </label>
</div>
"""

javascript_injection = """
<script type='text/javascript'>
  document.addEventListener('DOMContentLoaded', function() {
    var network = window.network;
    var nodes = window.nodes;
    if (!network || !nodes) { return; }

    const pathToggle = document.getElementById('pathToggle');
    pathToggle.addEventListener('change', function() {
      const showFullPath = this.checked;
      let updatedNodes = [];
      nodes.forEach(node => {
        let newLabel = showFullPath ? node.full_label : node.short_label;
        updatedNodes.push({ id: node.id, label: newLabel });
      });
      nodes.update(updatedNodes);
    });
  });
</script>
"""

html_content = html_content.replace('<body>', '<body>\n' + checkbox_html)
html_content = html_content.replace('</body>', javascript_injection + '\n</body>')

final_file = "final_inverted_graph.html"
with open(final_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

os.remove(temp_file)

print(f"Successfully generated final inverted graph at: {os.path.join(os.getcwd(), final_file)}")

