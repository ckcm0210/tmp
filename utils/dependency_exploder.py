# -*- coding: utf-8 -*-
"""
Dependency Exploder - 公式依賴鏈遞歸分析器
"""

import re
import os
from utils.openpyxl_resolver import read_cell_with_resolved_references

class DependencyExploder:
    """公式依賴鏈爆炸分析器"""
    
    def __init__(self, max_depth=10):
        self.max_depth = max_depth
        self.visited_cells = set()
        self.circular_refs = []
    
    def explode_dependencies(self, workbook_path, sheet_name, cell_address, current_depth=0, root_workbook_path=None):
        """
        遞歸展開公式依賴鏈
        
        Args:
            workbook_path: Excel 檔案路徑
            sheet_name: 工作表名稱
            cell_address: 儲存格地址 (如 A1)
            current_depth: 當前遞歸深度
            
        Returns:
            dict: 依賴樹結構
        """
        # 創建唯一標識符
        cell_id = f"{workbook_path}|{sheet_name}|{cell_address}"
        
        # 檢查遞歸深度限制
        if current_depth >= self.max_depth:
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Max depth reached',
                'formula': None,
                'type': 'limit_reached',
                'children': [],
                'depth': current_depth,
                'error': 'Maximum recursion depth reached'
            }
        
        # 檢查循環引用
        if cell_id in self.visited_cells:
            self.circular_refs.append(cell_id)
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Circular reference',
                'formula': None,
                'type': 'circular_ref',
                'children': [],
                'depth': current_depth,
                'error': 'Circular reference detected'
            }
        
        # 標記為已訪問
        self.visited_cells.add(cell_id)
        
        try:
            # 讀取儲存格內容
            cell_info = read_cell_with_resolved_references(workbook_path, sheet_name, cell_address)
            
            if 'error' in cell_info:
                # 決定顯示格式
                current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
                if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                    filename = os.path.basename(workbook_path)
                    if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                        filename = filename.rsplit('.', 1)[0]
                    display_address = f"[{filename}]{sheet_name}!{cell_address}"
                else:
                    display_address = f"{sheet_name}!{cell_address}"
                
                return {
                    'address': display_address,
                    'workbook_path': workbook_path,
                    'sheet_name': sheet_name,
                    'cell_address': cell_address,
                    'value': 'Error',
                    'formula': None,
                    'type': 'error',
                    'children': [],
                    'depth': current_depth,
                    'error': cell_info['error']
                }
            
            # 基本節點信息
            original_formula = cell_info.get('formula')
            # 針對性修復雙反斜線問題
            fixed_formula = original_formula.replace('\\\\', '\\') if original_formula else None

            # 決定顯示格式：外部引用顯示為 [filename]sheet!cell，本地引用顯示為 sheet!cell
            # 使用 root_workbook_path 來判斷是否為外部引用
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                # 外部引用：顯示 [filename]sheet!cell 格式
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]  # 移除副檔名
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                # 本地引用：顯示 sheet!cell 格式
                display_address = f"{sheet_name}!{cell_address}"

            node = {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': cell_info.get('display_value', 'N/A'),
                'calculated_value': cell_info.get('calculated_value', 'N/A'),
                'formula': fixed_formula,
                'type': cell_info.get('cell_type', 'unknown'),
                'children': [],
                'depth': current_depth,
                'error': None
            }
            
            # 如果是公式，解析依賴關係
            if cell_info.get('cell_type') == 'formula' and cell_info.get('formula'):
                references = self.parse_formula_references(cell_info['formula'], workbook_path, sheet_name)
                
                # 遞歸展開每個引用
                for ref in references:
                    try:
                        child_node = self.explode_dependencies(
                            ref['workbook_path'],
                            ref['sheet_name'],
                            ref['cell_address'],
                            current_depth + 1,
                            root_workbook_path or workbook_path
                        )
                        node['children'].append(child_node)
                    except Exception as e:
                        # 添加錯誤節點
                        # 決定顯示格式
                        current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
                        if os.path.normpath(current_workbook_path) != os.path.normpath(ref['workbook_path']):
                            filename = os.path.basename(ref['workbook_path'])
                            if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                                filename = filename.rsplit('.', 1)[0]
                            error_display_address = f"[{filename}]{ref['sheet_name']}!{ref['cell_address']}"
                        else:
                            error_display_address = f"{ref['sheet_name']}!{ref['cell_address']}"
                        
                        error_node = {
                            'address': error_display_address,
                            'workbook_path': ref['workbook_path'],
                            'sheet_name': ref['sheet_name'],
                            'cell_address': ref['cell_address'],
                            'value': 'Error',
                            'formula': None,
                            'type': 'error',
                            'children': [],
                            'depth': current_depth + 1,
                            'error': str(e)
                        }
                        node['children'].append(error_node)
            
            # 移除已訪問標記（允許在不同分支中重複訪問）
            self.visited_cells.discard(cell_id)
            
            return node
            
        except Exception as e:
            # 移除已訪問標記
            self.visited_cells.discard(cell_id)
            
            # 決定顯示格式
            current_workbook_path = root_workbook_path if root_workbook_path else workbook_path
            if os.path.normpath(current_workbook_path) != os.path.normpath(workbook_path):
                filename = os.path.basename(workbook_path)
                if filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.xlsm'):
                    filename = filename.rsplit('.', 1)[0]
                display_address = f"[{filename}]{sheet_name}!{cell_address}"
            else:
                display_address = f"{sheet_name}!{cell_address}"
            
            return {
                'address': display_address,
                'workbook_path': workbook_path,
                'sheet_name': sheet_name,
                'cell_address': cell_address,
                'value': 'Error',
                'formula': None,
                'type': 'error',
                'children': [],
                'depth': current_depth,
                'error': str(e)
            }
    
    def parse_formula_references(self, formula, current_workbook_path, current_sheet_name):
        """
        解析公式中的所有引用
        
        Args:
            formula: 公式字符串
            current_workbook_path: 當前工作簿路徑
            current_sheet_name: 當前工作表名稱
            
        Returns:
            list: 引用列表
        """
        references = []
        
        if not formula or not formula.startswith('='):
            return references
        
        # 移除公式開頭的 = 號
        formula_content = formula[1:]
        
        try:
            # 預處理：標準化公式中的路徑，將雙反斜線轉為單反斜線
            normalized_formula = self._normalize_formula_paths(formula)
            
            # 1. 解析外部引用 (例如: 'C:\path\[file.xlsx]Sheet'!$A$1)
            external_pattern = r"'([^']*\[[^\]]+\][^']*)'!\$?([A-Z]+)\$?(\d+)"
            external_matches = re.findall(external_pattern, normalized_formula)
            
            # 創建一個副本來移除已處理的外部引用
            remaining_formula = normalized_formula
            
            for match in external_matches:
                full_ref, col, row = match
                if '[' in full_ref and ']' in full_ref:
                    path_part = full_ref.split('[')[0]
                    file_part = full_ref.split('[')[1].split(']')[0]
                    sheet_part = full_ref.split(']')[1] if ']' in full_ref else 'Sheet1'
                    
                    # 修復路徑中的雙反斜線問題 - 使用更直接的方法
                    # 直接組合路徑，然後用 normpath 處理所有斜線問題
                    raw_path = path_part + file_part
                    workbook_path = os.path.normpath(raw_path)
                    sheet_name = sheet_part
                    cell_address = f"{col}{row}"
                    
                    references.append({
                        'workbook_path': workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'type': 'external'
                    })
                    
                    # 從剩餘公式中移除這個外部引用，避免路徑被誤認為 cell address
                    external_ref_full = f"'{full_ref}'!${col}${row}"
                    remaining_formula = remaining_formula.replace(external_ref_full, "")
                    # 也處理沒有 $ 符號的情況
                    external_ref_no_dollar = f"'{full_ref}'!{col}{row}"
                    remaining_formula = remaining_formula.replace(external_ref_no_dollar, "")
            
            # 2. 解析本地絕對引用 (例如: Sheet1!A1, 工作表1!A1)
            # 使用移除外部引用後的公式
            normalized_content = remaining_formula[1:] if remaining_formula.startswith('=') else remaining_formula
            # 找到所有 ! 的位置
            exclamation_positions = [i for i, char in enumerate(normalized_content) if char == '!']
            
            for pos in exclamation_positions:
                # 向前找工作表名稱
                start = pos - 1
                
                # 檢查是否以單引號結尾
                if start >= 0 and normalized_content[start] == "'":
                    # 向前找到開始的單引號
                    quote_start = start - 1
                    while quote_start >= 0 and normalized_content[quote_start] != "'":
                        quote_start -= 1
                    
                    if quote_start >= 0:
                        sheet_name = normalized_content[quote_start + 1:start]
                    else:
                        continue
                else:
                    # 沒有單引號，向前找到邊界
                    while start >= 0 and normalized_content[start] not in "+'*/-()=,":
                        start -= 1
                    start += 1
                    sheet_name = normalized_content[start:pos]
                
                # 向後找 cell 地址
                remaining = normalized_content[pos + 1:]
                cell_match = re.match(r'\$?([A-Z]+)\$?(\d+)', remaining)
                
                if cell_match and sheet_name and '[' not in sheet_name and ']' not in sheet_name:
                    col, row = cell_match.groups()
                    cell_address = f"{col}{row}"
                    
                    references.append({
                        'workbook_path': current_workbook_path,
                        'sheet_name': sheet_name,
                        'cell_address': cell_address,
                        'type': 'local_absolute'
                    })
            
            # 3. 解析相對引用 (例如: A1, B5)
            # 使用移除外部引用後的公式
            relative_pattern = r"(?<![A-Za-z0-9_!'])([A-Z]+)(\d+)(?![A-Za-z0-9_])"
            relative_matches = re.findall(relative_pattern, normalized_content)
            
            # 獲取已存在的絕對引用，避免重複
            existing_refs = set()
            for ref in references:
                existing_refs.add(f"{ref['sheet_name']}!{ref['cell_address']}")
            
            for col, row in relative_matches:
                cell_address = f"{col}{row}"
                ref_key = f"{current_sheet_name}!{cell_address}"
                
                if ref_key not in existing_refs:
                    references.append({
                        'workbook_path': current_workbook_path,
                        'sheet_name': current_sheet_name,
                        'cell_address': cell_address,
                        'type': 'relative'
                    })
            
        except Exception as e:
            print(f"Warning: Error parsing formula references: {e}")
        
        return references
    
    def _normalize_formula_paths(self, formula):
        """
        標準化公式中的路徑，將雙反斜線轉為單反斜線
        
        Args:
            formula: 原始公式字符串
            
        Returns:
            str: 標準化後的公式字符串
        """
        if not formula:
            return formula
        
        # 使用正則表達式找到所有外部引用路徑並標準化
        def normalize_path_match(match):
            full_match = match.group(0)
            path_part = match.group(1)
            
            # 標準化路徑部分
            normalized_path = os.path.normpath(path_part)
            
            # 重建完整的引用
            return full_match.replace(path_part, normalized_path)
        
        # 匹配外部引用中的路徑部分
        external_ref_pattern = r"'([^']*\[[^\]]+\][^']*)'!"
        normalized_formula = re.sub(external_ref_pattern, normalize_path_match, formula)
        
        return normalized_formula
    
    def get_explosion_summary(self, root_node):
        """
        獲取爆炸分析摘要
        
        Args:
            root_node: 根節點
            
        Returns:
            dict: 摘要信息
        """
        def count_nodes(node):
            count = 1
            for child in node.get('children', []):
                count += count_nodes(child)
            return count
        
        def get_max_depth(node):
            if not node.get('children'):
                return node.get('depth', 0)
            return max(get_max_depth(child) for child in node['children'])
        
        def count_by_type(node, type_counts=None):
            if type_counts is None:
                type_counts = {}
            
            node_type = node.get('type', 'unknown')
            type_counts[node_type] = type_counts.get(node_type, 0) + 1
            
            for child in node.get('children', []):
                count_by_type(child, type_counts)
            
            return type_counts
        
        return {
            'total_nodes': count_nodes(root_node),
            'max_depth': get_max_depth(root_node),
            'type_distribution': count_by_type(root_node),
            'circular_references': len(self.circular_refs),
            'circular_ref_list': self.circular_refs
        }


def explode_cell_dependencies(workbook_path, sheet_name, cell_address, max_depth=10):
    """
    便捷函數：爆炸分析指定儲存格的依賴關係
    
    Args:
        workbook_path: Excel 檔案路徑
        sheet_name: 工作表名稱
        cell_address: 儲存格地址
        max_depth: 最大遞歸深度
        
    Returns:
        tuple: (依賴樹, 摘要信息)
    """
    exploder = DependencyExploder(max_depth=max_depth)
    dependency_tree = exploder.explode_dependencies(workbook_path, sheet_name, cell_address)
    summary = exploder.get_explosion_summary(dependency_tree)
    
    return dependency_tree, summary


# 測試函數
if __name__ == "__main__":
    # 測試用例
    test_workbook = r"C:\Users\user\Desktop\pytest\test.xlsx"
    test_sheet = "Sheet1"
    test_cell = "A1"
    
    try:
        tree, summary = explode_cell_dependencies(test_workbook, test_sheet, test_cell)
        print("Dependency Tree:")
        print(tree)
        print("\nSummary:")
        print(summary)
    except Exception as e:
        print(f"Test failed: {e}")