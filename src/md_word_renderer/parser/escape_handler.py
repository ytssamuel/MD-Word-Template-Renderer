"""
特殊字元處理器

處理 Markdown 資料中的特殊字元轉義
"""

import re
from typing import Dict


class EscapeHandler:
    """
    處理特殊字元轉義
    
    轉義規則：
    - \\| → |（管線符號）
    - \\n → 換行
    - \\\\ → \\（反斜線）
    - \\" → "（雙引號）
    - "" → "（雙引號內的引號）
    
    Example:
        >>> handler = EscapeHandler()
        >>> text = handler.unescape("這是\\|管線符號")
        >>> print(text)
        "這是|管線符號"
    """
    
    # 轉義序列對應表
    ESCAPE_SEQUENCES: Dict[str, str] = {
        '\\|': '|',
        '\\n': '\n',
        '\\\\': '\\',
        '\\"': '"',
        '\\t': '\t',
    }
    
    def __init__(self):
        """初始化處理器"""
        # 建立轉義序列的正則表達式
        # 優先處理較長的序列（如 \\\\ 在 \\ 之前）
        patterns = sorted(self.ESCAPE_SEQUENCES.keys(), key=len, reverse=True)
        escaped_patterns = [re.escape(p) for p in patterns]
        self.escape_pattern = re.compile('|'.join(escaped_patterns))
    
    def unescape(self, text: str) -> str:
        """
        還原轉義字元
        
        Args:
            text: 包含轉義序列的文字
            
        Returns:
            str: 還原後的文字
        """
        if not text:
            return text
        
        # 處理標準轉義序列
        def replace_escape(match):
            return self.ESCAPE_SEQUENCES.get(match.group(0), match.group(0))
        
        result = self.escape_pattern.sub(replace_escape, text)
        
        # 處理連續雙引號 "" → "
        result = result.replace('""', '"')
        
        return result
    
    def escape(self, text: str) -> str:
        """
        轉義特殊字元
        
        Args:
            text: 原始文字
            
        Returns:
            str: 轉義後的文字
        """
        if not text:
            return text
        
        result = text
        
        # 反向轉義（按順序處理避免衝突）
        result = result.replace('\\', '\\\\')
        result = result.replace('|', '\\|')
        result = result.replace('\n', '\\n')
        result = result.replace('\t', '\\t')
        result = result.replace('"', '\\"')
        
        return result
    
    def contains_special_chars(self, text: str) -> bool:
        """
        檢查是否包含需要轉義的特殊字元
        
        Args:
            text: 要檢查的文字
            
        Returns:
            bool: 是否包含特殊字元
        """
        special_chars = ['|', '\n', '\\', '"', '\t']
        return any(char in text for char in special_chars)
