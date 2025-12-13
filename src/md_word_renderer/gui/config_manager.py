"""
設定管理器

處理應用程式設定的載入、儲存和管理
"""

import json
from pathlib import Path
from typing import Any, Dict, Optional


class ConfigManager:
    """
    設定管理器
    
    處理應用程式設定的持久化儲存
    """
    
    # 設定檔路徑
    CONFIG_DIR = Path.home() / ".md_word_renderer"
    CONFIG_FILE = CONFIG_DIR / "config.json"
    
    # 預設設定
    DEFAULT_CONFIG = {
        # 外觀設定
        "theme": "dark",           # "dark", "light", "system"
        "scale": 100,              # 縮放比例
        
        # 預設路徑
        "default_template_dir": "",
        "default_output_dir": "",
        "last_markdown_dir": "",
        "last_template_file": "",
        
        # 處理選項
        "validate_before_convert": True,
        "open_after_convert": True,
        "continue_on_error": True,
        
        # 批次處理
        "batch_file_pattern": "*.md",
        
        # 視窗設定
        "window_width": 900,
        "window_height": 700,
        "window_x": None,
        "window_y": None,
    }
    
    def __init__(self):
        """初始化設定管理器"""
        self._config: Dict[str, Any] = {}
        self.load()
    
    def load(self) -> Dict[str, Any]:
        """
        載入設定
        
        Returns:
            dict: 設定字典
        """
        try:
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    saved_config = json.load(f)
                    # 合併預設設定和已儲存設定
                    self._config = {**self.DEFAULT_CONFIG, **saved_config}
            else:
                self._config = self.DEFAULT_CONFIG.copy()
        except Exception:
            self._config = self.DEFAULT_CONFIG.copy()
        
        return self._config
    
    def save(self) -> bool:
        """
        儲存設定
        
        Returns:
            bool: 是否成功
        """
        try:
            self.CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self._config, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        取得設定值
        
        Args:
            key: 設定鍵
            default: 預設值
            
        Returns:
            設定值
        """
        return self._config.get(key, default)
    
    def set(self, key: str, value: Any) -> None:
        """
        設定值
        
        Args:
            key: 設定鍵
            value: 設定值
        """
        self._config[key] = value
    
    def update(self, updates: Dict[str, Any]) -> None:
        """
        批量更新設定
        
        Args:
            updates: 更新的設定字典
        """
        self._config.update(updates)
    
    def reset(self) -> None:
        """重設為預設值"""
        self._config = self.DEFAULT_CONFIG.copy()
    
    @property
    def config(self) -> Dict[str, Any]:
        """取得完整設定字典"""
        return self._config.copy()
