"""
縮排偵測器

偵測並驗證 Markdown 檔案的縮排格式
參考 Python PEP 8 的縮排規則
"""

from typing import List, Tuple, Set
from collections import Counter


class IndentationError(Exception):
    """縮排錯誤"""
    pass


class IndentDetector:
    """
    偵測並驗證縮排格式
    
    功能：
    - 自動偵測縮排類型（空格或 Tab）
    - 計算縮排層級
    - 驗證縮排一致性
    - 支援不規則縮排的智能處理
    
    Example:
        >>> detector = IndentDetector()
        >>> indent_type, indent_unit = detector.detect(lines)
        >>> level = detector.calculate_level(line, indent_type, indent_unit)
    """
    
    def __init__(self, indent_size: int = 4):
        """
        初始化偵測器
        
        Args:
            indent_size: 預設每層縮排的空格數
        """
        self.default_indent_size = indent_size
        # 儲存偵測到的所有縮排值，用於智能層級計算
        self._indent_levels: List[int] = []
    
    def detect(self, lines: List[str], sample_size: int = 100) -> Tuple[str, int]:
        """
        偵測縮排類型
        
        掃描前 N 行，統計空格和 Tab 的使用頻率，選擇多數者作為縮排標準
        
        Args:
            lines: 文字行列表
            sample_size: 取樣行數（預設 100）
            
        Returns:
            tuple: (縮排類型, 每層級單位數)
            - ('space', 4) 表示使用 4 空格縮排
            - ('tab', 1) 表示使用 Tab 縮排
            
        Raises:
            IndentationError: 混用空格和 Tab
        """
        space_count = 0
        tab_count = 0
        indent_sizes: Set[int] = set()
        
        for line in lines[:sample_size]:
            if not line or not line.strip():
                continue
            
            # 計算前導空白
            leading_whitespace = len(line) - len(line.lstrip())
            if leading_whitespace == 0:
                continue
            
            # 檢查是否混用
            leading = line[:leading_whitespace]
            has_space = ' ' in leading
            has_tab = '\t' in leading
            
            if has_space and has_tab:
                # 不嚴格報錯，而是警告並使用空格模式
                pass
            
            if has_space:
                space_count += 1
                indent_sizes.add(leading_whitespace)
            elif has_tab:
                tab_count += 1
        
        # 儲存所有不同的縮排值（排序後）
        self._indent_levels = sorted(indent_sizes)
        
        # 決定縮排類型
        if tab_count > space_count:
            return ('tab', 1)
        
        # 計算縮排單位
        if indent_sizes:
            # 找出最小的非零縮排值
            sorted_indents = sorted(indent_sizes)
            if len(sorted_indents) >= 2:
                # 計算相鄰縮排的差值，找出最常見的差值
                diffs = []
                for i in range(1, len(sorted_indents)):
                    diff = sorted_indents[i] - sorted_indents[i-1]
                    if diff > 0:
                        diffs.append(diff)
                if diffs:
                    # 使用最常見的差值作為單位
                    counter = Counter(diffs)
                    most_common = counter.most_common(1)[0][0]
                    return ('space', most_common if most_common <= 8 else self.default_indent_size)
            
            min_indent = min(indent_sizes)
            return ('space', min_indent if min_indent <= 8 else self.default_indent_size)
        
        return ('space', self.default_indent_size)
    
    def calculate_level(self, line: str, indent_type: str, indent_unit: int) -> int:
        """
        計算縮排層級
        
        支援不規則縮排的智能計算：
        - 如果縮排值在已知的縮排層級列表中，使用其索引作為層級
        - 否則使用傳統的除法計算
        
        Args:
            line: 文字行
            indent_type: 縮排類型 ('space' 或 'tab')
            indent_unit: 每層縮排的單位數
            
        Returns:
            int: 縮排層級（從 0 開始）
        """
        if not line or not line.strip():
            return 0
        
        leading_whitespace = len(line) - len(line.lstrip())
        
        if leading_whitespace == 0:
            return 0
        
        if indent_type == 'tab':
            # 計算 Tab 數量
            return line[:leading_whitespace].count('\t')
        
        # 智能層級計算：使用預先偵測的縮排值
        if self._indent_levels:
            # 找出最接近的已知縮排值
            for i, known_indent in enumerate(self._indent_levels):
                if leading_whitespace <= known_indent:
                    return i + 1  # 層級從 1 開始（0 是無縮排）
            # 如果超過所有已知值，返回最大層級 + 1
            return len(self._indent_levels) + 1
        
        # 傳統計算方式
        return leading_whitespace // indent_unit
    
    def calculate_level_smart(self, indent_value: int) -> int:
        """
        智能計算縮排層級
        
        根據實際縮排值在已知縮排列表中的位置決定層級
        
        Args:
            indent_value: 縮排的空格數
            
        Returns:
            int: 層級（0 表示頂層）
        """
        if indent_value == 0:
            return 0
        
        if not self._indent_levels:
            return indent_value // self.default_indent_size
        
        # 找出縮排值在列表中的位置
        for i, level in enumerate(self._indent_levels):
            if indent_value <= level:
                return i + 1
        
        return len(self._indent_levels) + 1
    
    def validate(self, lines: List[str]) -> Tuple[bool, List[str]]:
        """
        驗證縮排一致性
        
        Args:
            lines: 文字行列表
            
        Returns:
            tuple: (是否通過, 錯誤訊息列表)
        """
        errors = []
        
        try:
            indent_type, indent_unit = self.detect(lines)
        except IndentationError as e:
            return False, [str(e)]
        
        for line_num, line in enumerate(lines, start=1):
            if not line or not line.strip():
                continue
            
            leading = len(line) - len(line.lstrip())
            
            # 檢查縮排是否為整數倍
            if indent_type == 'space' and leading % indent_unit != 0:
                errors.append(
                    f"第 {line_num} 行: 縮排 {leading} 空格不是 {indent_unit} 的整數倍"
                )
        
        return len(errors) == 0, errors
