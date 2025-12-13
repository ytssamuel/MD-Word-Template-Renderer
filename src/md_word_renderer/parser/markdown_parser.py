"""
Markdown 解析器

解析特殊格式的 Markdown 檔案，格式為：編號. 欄位名稱 | 值
支援階層結構（透過縮排）
"""

import re
from pathlib import Path
from typing import Dict, List, Any, Optional

from .indent_detector import IndentDetector
from .escape_handler import EscapeHandler


class ParseError(Exception):
    """解析錯誤"""
    pass


class MarkdownParser:
    """
    解析特殊格式的 Markdown 檔案
    
    格式：編號. 欄位名稱 | 值
    支援階層結構（透過縮排）
    
    Example:
        >>> parser = MarkdownParser()
        >>> data = parser.parse("data.md")
        >>> print(data["系統名稱"])
        "範例系統"
    """
    
    # 解析單行的正則表達式
    # 匹配格式：編號. 欄位名稱 | 值
    LINE_PATTERN = re.compile(r'^(\d+)\.\s+([^|]+?)\s*\|\s*(.*)$')
    
    # 子項目格式（無管線符號）
    CHILD_PATTERN = re.compile(r'^(\d+)\.\s+(.+)$')
    
    def __init__(self, indent_size: int = 4):
        """
        初始化解析器
        
        Args:
            indent_size: 縮排空格數（預設 4）
        """
        self.indent_detector = IndentDetector(indent_size=indent_size)
        self.escape_handler = EscapeHandler()
        self.indent_size = indent_size
    
    def parse(self, filepath: str, encoding: str = 'utf-8') -> Dict[str, Any]:
        """
        解析 Markdown 檔案
        
        Args:
            filepath: .md 檔案路徑
            encoding: 檔案編碼（預設 utf-8）
            
        Returns:
            dict: 結構化的資料字典，支援名稱和編號雙重索引
            
        Raises:
            FileNotFoundError: 檔案不存在
            ParseError: 解析失敗
        """
        path = Path(filepath)
        if not path.exists():
            raise FileNotFoundError(f"檔案不存在: {filepath}")
        
        with open(path, 'r', encoding=encoding) as f:
            content = f.read()
        
        return self.parse_content(content)
    
    def parse_content(self, content: str) -> Dict[str, Any]:
        """
        解析 Markdown 內容字串
        
        Args:
            content: Markdown 內容
            
        Returns:
            dict: 結構化的資料字典
        """
        lines = content.split('\n')
        
        # 偵測縮排類型
        indent_type, indent_unit = self.indent_detector.detect(lines)
        
        # 解析所有行
        parsed_items = self._parse_lines(lines, indent_type, indent_unit)
        
        # 建立階層結構
        result = self._build_hierarchy(parsed_items)
        
        return result
    
    def _parse_lines(self, lines: List[str], 
                     indent_type: str, 
                     indent_unit: int) -> List[Dict[str, Any]]:
        """
        解析所有行
        
        Args:
            lines: 文字行列表
            indent_type: 縮排類型 ('space' 或 'tab')
            indent_unit: 每層縮排的單位數
            
        Returns:
            list: 解析後的項目列表
        """
        items = []
        
        for line_num, line in enumerate(lines, start=1):
            # 跳過空行和標題行
            if not line.strip() or line.strip().startswith('#'):
                continue
            
            # 計算縮排層級
            level = self.indent_detector.calculate_level(
                line, indent_type, indent_unit
            )
            
            # 去除前導空白
            stripped = line.strip()
            
            # 嘗試解析為主格式（編號. 名稱 | 值）
            match = self.LINE_PATTERN.match(stripped)
            if match:
                number, key, value = match.groups()
                items.append({
                    'line_num': line_num,
                    'level': level,
                    'number': number.strip(),
                    'key': key.strip(),
                    'value': self.escape_handler.unescape(value.strip()),
                    'children': []
                })
                continue
            
            # 嘗試解析為子項目格式（編號. 內容）
            child_match = self.CHILD_PATTERN.match(stripped)
            if child_match:
                number, value = child_match.groups()
                items.append({
                    'line_num': line_num,
                    'level': level,
                    'number': number.strip(),
                    'key': None,  # 子項目沒有 key
                    'value': self.escape_handler.unescape(value.strip()),
                    'children': []
                })
                continue
        
        return items
    
    def _build_hierarchy(self, items: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        建立階層關係
        
        使用堆疊演算法根據縮排層級建立父子關係
        
        Args:
            items: 解析後的項目列表
            
        Returns:
            dict: 結構化的資料字典，支援名稱和編號雙重存取
        """
        if not items:
            return {}
        
        result = {}
        # 堆疊存放 (level, item_ref, key_name) - item_ref 指向 children 列表
        stack: List[tuple] = []
        # 追蹤頂層欄位以便後續處理
        top_level_fields: Dict[str, Dict] = {}
        
        for item in items:
            level = item['level']
            
            # 建立子項目資料結構
            child_data = {
                'number': item['number'],
                'value': item['value'],
                'children': []
            }
            
            # 彈出所有層級 >= 當前層級的項目
            while stack and stack[-1][0] >= level:
                stack.pop()
            
            if level == 0 and item['key']:
                # 頂層項目（有 key 的主欄位）
                key_name = item['key']
                
                # 儲存引用以便後續處理
                top_level_fields[key_name] = {
                    'value': item['value'],
                    'children': child_data['children'],
                    'number': item['number']
                }
                
                # 建立編號索引
                result[f"#{item['number']}"] = {
                    'key': key_name,
                    'value': item['value'],
                    'children': child_data['children']
                }
                
                # 推入堆疊
                stack.append((level, child_data, key_name))
                
            elif stack:
                # 有父項目，添加為子項目
                parent_data = stack[-1][1]
                parent_data['children'].append(child_data)
                
                # 推入堆疊，以便自己也能有子項目
                stack.append((level, child_data, None))
            else:
                # 頂層但沒有 key（異常情況）
                stack.append((level, child_data, None))
        
        # 後處理：根據是否有子項目決定欄位值
        for key_name, field_info in top_level_fields.items():
            if field_info['value']:
                # 有值的欄位直接使用值
                result[key_name] = field_info['value']
            elif field_info['children']:
                # 沒有值但有子項目的欄位使用 children 列表
                result[key_name] = field_info['children']
            else:
                # 沒有值也沒有子項目，返回空字串
                result[key_name] = ""
        
        return result


def main():
    """測試用主函數"""
    import sys
    
    if len(sys.argv) < 2:
        print("用法: python markdown_parser.py <markdown_file>")
        sys.exit(1)
    
    parser = MarkdownParser()
    try:
        data = parser.parse(sys.argv[1])
        import json
        print(json.dumps(data, ensure_ascii=False, indent=2))
    except Exception as e:
        print(f"錯誤: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
