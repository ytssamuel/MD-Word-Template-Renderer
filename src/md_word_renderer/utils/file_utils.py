"""
檔案工具函數

提供檔案處理相關的工具函數
"""

import glob
from pathlib import Path
from typing import List, Optional


class FileUtils:
    """
    檔案處理工具類
    
    提供：
    - 檔案路徑處理
    - 批次檔案搜尋
    - 檔名生成
    """
    
    @staticmethod
    def find_files(pattern: str, recursive: bool = False) -> List[str]:
        """
        根據 pattern 搜尋檔案
        
        Args:
            pattern: 檔案 pattern（支援 glob 語法）
            recursive: 是否遞迴搜尋子目錄
            
        Returns:
            list: 符合條件的檔案路徑列表
        """
        if recursive:
            return list(glob.glob(pattern, recursive=True))
        return list(glob.glob(pattern))
    
    @staticmethod
    def ensure_directory(path: str) -> Path:
        """
        確保目錄存在
        
        Args:
            path: 目錄路徑
            
        Returns:
            Path: 目錄路徑物件
        """
        directory = Path(path)
        directory.mkdir(parents=True, exist_ok=True)
        return directory
    
    @staticmethod
    def generate_output_path(input_file: str, 
                             output_dir: str, 
                             extension: str = '.docx',
                             pattern: str = '{name}') -> str:
        """
        生成輸出檔案路徑
        
        Args:
            input_file: 輸入檔案路徑
            output_dir: 輸出目錄
            extension: 輸出副檔名
            pattern: 檔名 pattern（{name} 為原檔名）
            
        Returns:
            str: 輸出檔案路徑
        """
        input_path = Path(input_file)
        output_path = Path(output_dir)
        
        # 確保輸出目錄存在
        output_path.mkdir(parents=True, exist_ok=True)
        
        # 生成檔名
        base_name = input_path.stem
        output_name = pattern.format(name=base_name)
        
        # 確保有正確的副檔名
        if not output_name.endswith(extension):
            output_name += extension
        
        return str(output_path / output_name)
    
    @staticmethod
    def get_file_extension(filepath: str) -> str:
        """
        取得檔案副檔名
        
        Args:
            filepath: 檔案路徑
            
        Returns:
            str: 副檔名（包含點號）
        """
        return Path(filepath).suffix
    
    @staticmethod
    def is_markdown_file(filepath: str) -> bool:
        """
        檢查是否為 Markdown 檔案
        
        Args:
            filepath: 檔案路徑
            
        Returns:
            bool: 是否為 .md 檔案
        """
        return Path(filepath).suffix.lower() == '.md'
    
    @staticmethod
    def is_word_file(filepath: str) -> bool:
        """
        檢查是否為 Word 檔案
        
        Args:
            filepath: 檔案路徑
            
        Returns:
            bool: 是否為 .docx 檔案
        """
        return Path(filepath).suffix.lower() == '.docx'
