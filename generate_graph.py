
from pyvis.network import Network
import os

# 創建一個網路圖物件
net = Network(height="750px", width="100%", bgcolor="#222222", font_color="white", notebook=True, directed=True)

# 添加節點
# node_id, label, title (懸停文字), color
net.add_node(0, label="A3", title="<b>A3</b><br>Formula: =A1+A2<br>Value: 150", color="#007bff") # 藍色
net.add_node(1, label="A1", title="<b>A1</b><br>Formula: ='[External.xlsx]Sheet1'!$C$5<br>Value: 50", color="#ff8c00") # 橘色
net.add_node(2, label="A2", title="<b>A2</b><br>Type: Value<br>Value: 100", color="#28a745") # 綠色
net.add_node(3, label="[External.xlsx]Sheet1!C5", title="<b>[External.xlsx]Sheet1!C5</b><br>Type: External Value<br>Value: 50", color="#ff8c00", shape='box') # 橘色，方形

# 添加邊 (箭頭)
# source_node_id -> target_node_id
net.add_edge(0, 1)
net.add_edge(0, 2)
net.add_edge(1, 3)

# 產生 HTML 檔案
file_path = os.path.join(os.getcwd(), "dependency_graph_example.html")
try:
    net.save_graph(file_path)
    print(f"Successfully generated graph at: {file_path}")
except Exception as e:
    print(f"Error generating graph: {e}")

