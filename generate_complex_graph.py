

from pyvis.network import Network
import os

# 創建一個網路圖物件，並設定為有向圖 (directed=True) 以顯示箭頭
net = Network(height="800px", width="100%", bgcolor="#f0f0f0", font_color="black", directed=True)

# 關鍵步驟：關閉物理引擎，實現節點的自由拖動和釘選
net.toggle_physics(False)

# --- 模擬一個更複雜、真實的依賴關係場景 ---

# 節點資料 (ID, Label, Title for Hover, Color, Shape)
nodes_data = [
    (0, 'Sheet1!D10', "<b>Sheet1!D10</b><br>Type: Formula<br>Formula: ='Financial Calcs'!C5 & \" Profit\"<br>Value: 5500 Profit", '#9727b0', 'ellipse'),
    (1, 'Financial Calcs!C5', "<b>Financial Calcs!C5</b><br>Type: Formula<br>Formula: =VLOOKUP(B5, 'Source Data'!A:B, 2, FALSE)<br>Value: 5500", '#007bff', 'ellipse'),
    (2, 'Financial Calcs!G5', "<b>Financial Calcs!G5</b><br>Type: Formula<br>Formula: =G3 * '[Pricing.xlsx]Rates'!A1<br>Value: 5462.5", '#007bff', 'ellipse'),
    (3, 'Sheet1!B5', "<b>Sheet1!B5</b><br>Type: Value<br>Value: PROD-XYZ", '#28a745', 'ellipse'),
    (4, 'Source Data!B2', "<b>Source Data!B2</b><br>Type: Value<br>Value: 5500", '#28a745', 'ellipse'),
    (5, 'Financial Calcs!G3', "<b>Financial Calcs!G3</b><br>Type: Value<br>Value: 1.15", '#28a745', 'ellipse'),
    (6, 'C:\\Excel Files\\[Pricing.xlsx]Rates!A1', "<b>C:\\Excel Files\\[Pricing.xlsx]Rates!A1</b><br>Type: External Value<br>Value: 4750.00", '#ff8c00', 'box')
]

# 邊資料 (Source ID -> Target ID)
edges_data = [
    (0, 1), # D10 -> C5
    # (0, 2), # D10 -> G5 (在這個新例子中，我們假設D10只用了C5)
    (1, 3), # C5 -> B5
    (1, 4), # C5 -> Source Data!B2 (VLOOKUP的結果來源)
    (2, 5), # G5 -> G3
    (2, 6)  # G5 -> [Pricing.xlsx]Rates!A1
]

# 添加節點和邊到圖中
for node in nodes_data:
    net.add_node(node[0], label=node[1], title=node[2], color=node[3], shape=node[4])

for edge in edges_data:
    net.add_edge(edge[0], edge[1])

# 產生 HTML 檔案
file_path = os.path.join(os.getcwd(), "complex_dependency_graph.html")
try:
    net.save_graph(file_path)
    print(f"Successfully generated complex graph at: {file_path}")
except Exception as e:
    print(f"Error generating graph: {e}")

