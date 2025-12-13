"""
批次處理器

處理多個 Markdown 檔案的批次渲染
"""

import logging
from pathlib import Path
from typing import List, Dict, Any, Callable, Optional

from .file_utils import FileUtils


class BatchProcessor:
    """
    批次處理多個 Markdown 檔案
    
    功能：
    - 多檔案批次處理
    - 進度追蹤
    - 錯誤處理與繼續
    
    Example:
        >>> processor = BatchProcessor(verbose=True)
        >>> results = processor.process_files(
        ...     input_files=["file1.md", "file2.md"],
        ...     template_path="template.docx",
        ...     output_dir="output/",
        ...     process_func=render_single_file
        ... )
    """
    
    def __init__(self, verbose: bool = False, 
                 continue_on_error: bool = True):
        """
        初始化批次處理器
        
        Args:
            verbose: 是否顯示詳細訊息
            continue_on_error: 遇到錯誤時是否繼續處理
        """
        self.verbose = verbose
        self.continue_on_error = continue_on_error
        self.logger = logging.getLogger(__name__)
    
    def process_files(self, 
                     input_files: List[str],
                     template_path: str,
                     output_dir: str,
                     process_func: Callable[[str, str, str], None],
                     output_pattern: str = '{name}.docx') -> Dict[str, Any]:
        """
        批次處理檔案
        
        Args:
            input_files: 輸入檔案列表
            template_path: 模板路徑
            output_dir: 輸出目錄
            process_func: 處理函數，簽章為 (input, template, output) -> None
            output_pattern: 輸出檔名 pattern
            
        Returns:
            dict: 處理結果統計
            {
                'success': [...],
                'failed': [...],
                'total': int,
                'success_count': int,
                'failed_count': int
            }
        """
        results = {
            'success': [],
            'failed': [],
            'total': len(input_files),
            'success_count': 0,
            'failed_count': 0
        }
        
        if self.verbose:
            self.logger.info(f"開始批次處理，共 {len(input_files)} 個檔案")
        
        for i, input_file in enumerate(input_files, 1):
            try:
                output_file = FileUtils.generate_output_path(
                    input_file, output_dir, pattern=output_pattern.replace('.docx', '')
                )
                
                if self.verbose:
                    self.logger.info(f"[{i}/{len(input_files)}] 處理: {input_file}")
                
                process_func(input_file, template_path, output_file)
                
                results['success'].append({
                    'input': input_file,
                    'output': output_file
                })
                results['success_count'] += 1
                
                if self.verbose:
                    self.logger.info(f"✓ {input_file} → {output_file}")
                
            except Exception as e:
                error_info = {
                    'input': input_file,
                    'error': str(e)
                }
                results['failed'].append(error_info)
                results['failed_count'] += 1
                
                self.logger.error(f"✗ {input_file} - {str(e)}")
                
                if not self.continue_on_error:
                    break
        
        if self.verbose:
            self.logger.info(
                f"批次處理完成！成功: {results['success_count']}, "
                f"失敗: {results['failed_count']}"
            )
        
        return results
    
    def expand_file_patterns(self, patterns: List[str]) -> List[str]:
        """
        展開檔案 patterns
        
        Args:
            patterns: 檔案 pattern 列表（支援 glob）
            
        Returns:
            list: 展開後的檔案列表
        """
        files = []
        for pattern in patterns:
            expanded = FileUtils.find_files(pattern)
            if expanded:
                files.extend(expanded)
            elif Path(pattern).exists():
                # 如果不是 pattern，而是直接的檔案路徑
                files.append(pattern)
        
        # 去重並保持順序
        seen = set()
        unique_files = []
        for f in files:
            if f not in seen:
                seen.add(f)
                unique_files.append(f)
        
        return unique_files
