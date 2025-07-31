

import os
from pyvis.network import Network
import json

# --- 1. 定義節點和邊的數據 ---
nodes_data = [
    # Level 0
    {"id": 0, "short_label": "Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0", "full_label": "C:\\My Reports\\[Master.xlsx]Final_Report!A1\nFormula: =SUM(Data!B1:Data!B3)\nValue: 7825.0", "color": "#e04141", "level": 0},
    # Level 1
    {"id": 1, "short_label": "Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0", "full_label": "C:\\My Reports\\[Master.xlsx]Data!B1\nFormula: =VLOOKUP(C1, RefData!A:B, 2, 0)\nValue: 500.0", "color": "#007bff", "level": 1},
    {"id": 2, "short_label": "Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0", "full_label": "C:\\My Reports\\[Master.xlsx]Data!B2\nFormula: =[External.xlsx]Sheet1!D5 * 1.05\nValue: 2100.0", "color": "#007bff", "level": 1},
    {"id": 3, "short_label": "Data!B3\nType: Value\nValue: 5225.0", "full_label": "C:\\My Reports\\[Master.xlsx]Data!B3\nType: Value\nValue: 5225.0", "color": "#28a745", "level": 1},
    # Level 2
    {"id": 4, "short_label": "Data!C1\nType: Value\nValue: SKU-001", "full_label": "C:\\My Reports\\[Master.xlsx]Data!C1\nType: Value\nValue: SKU-001", "color": "#28a745", "level": 2},
    {"id": 5, "short_label": "RefData!B1\nType: Value\nValue: 500.0", "full_label": "C:\\My Reports\\[Master.xlsx]RefData!B1\nType: Value\nValue: 500.0", "color": "#28a745", "level": 2},
    {"id": 6, "short_label": "[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0", "full_label": "C:\\Linked Files\\[External.xlsx]Sheet1!D5\nType: External Value\nValue: 2000.0", "color": "#ff8c00", "level": 2}
]
edges_data = [(0, 1), (0, 2), (0, 3), (1, 4), (1, 5), (2, 6)]

# --- 2. 手動計算階層式佈局的初始座標 ---
level_counts = {}
for node in nodes_data:
    level = node['level']
    if level not in level_counts:
        level_counts[level] = 0
    level_counts[level] += 1

node_positions = {}
level_y_step = 250
level_x_step = 350

current_level_counts = {level: 0 for level in level_counts}

for node in nodes_data:
    level = node['level']
    total_in_level = level_counts[level]
    current_index_in_level = current_level_counts[level]
    
    # 計算座標
    y = level * level_y_step
    x = (current_index_in_level - (total_in_level - 1) / 2.0) * level_x_step
    
    node['x'] = x
    node['y'] = y
    current_level_counts[level] += 1

# --- 3. 使用 Pyvis 產生圖表 ---
net = Network(height="90vh", width="100%", bgcolor="#ffffff", font_color="black", directed=True)

# 關鍵：不使用階層式佈局，直接設定全域選項
options_str = '''
{
  "interaction": {
    "dragNodes": true,
    "dragView": false,
    "zoomView": true
  },
  "physics": {
    "enabled": false
  },
  "nodes": {
    "font": {
      "align": "left"
    }
  },
  "edges": {
    "smooth": {
      "type": "cubicBezier",
      "forceDirection": "vertical",
      "roundness": 0.4
    }
  }
}
'''
net.set_options(options_str)

# 將節點資料（包含計算好的 x, y 座標）加入網路圖
for node_info in nodes_data:
    net.add_node(
        node_info["id"],
        label=node_info["short_label"],
        shape='box',
        color=node_info["color"],
        x=node_info['x'],
        y=node_info['y'],
        fixed=False, # 允許使用者拖曳
        full_label=node_info["full_label"],
        short_label=node_info["short_label"]
    )

for edge in edges_data:
    net.add_edge(edge[0], edge[1])

# --- 4. 注入 HTML 和 JavaScript ---
temp_file = "temp_stable_graph.html"
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

# 關鍵：使用最穩定的 JavaScript 版本
javascript_injection = """
<script type='text/javascript'>
  document.addEventListener('DOMContentLoaded', function() {
    var network = window.network;
    var nodes = window.nodes;
    if (!network || !nodes) { return; }

    const pathToggle = document.getElementById('pathToggle');
    pathToggle.addEventListener('change', function() {
      const showFullPath = this.checked;
      const currentPositions = network.getPositions();
      let updatedNodes = [];

      nodes.forEach(node => {
        let newLabel = showFullPath ? node.full_label : node.short_label;
        const position = currentPositions[node.id];
        
        updatedNodes.push({
          id: node.id,
          label: newLabel,
          x: position.x,
          y: position.y,
          fixed: true // 點擊後固定，防止微小移動
        });
      });
      
      nodes.update(updatedNodes);

      // 短暫延遲後，解除固定，恢復自由拖曳
      setTimeout(function() {
          let releaseNodes = [];
          nodes.forEach(node => {
              releaseNodes.push({id: node.id, fixed: false});
          });
          nodes.update(releaseNodes);
      }, 100);
    });
  });
</script>
"""

html_content = html_content.replace('<body>', '<body>\n' + checkbox_html)
html_content = html_content.replace('</body>', javascript_injection + '\n</body>')

final_file = "stable_interactive_graph.html"
with open(final_file, 'w', encoding='utf-8') as f:
    f.write(html_content)

os.remove(temp_file)

print(f"Successfully generated stable interactive graph at: {os.path.join(os.getcwd(), final_file)}")

