"""
設定檔載入器

載入並管理 YAML 設定檔
"""

import yaml
from pathlib import Path
from typing import Dict, Any, Optional


class ConfigLoader:
    """
    載入並管理設定檔
    
    支援：
    - YAML 格式設定檔
    - 預設值
    - CLI 參數覆寫
    
    Example:
        >>> config = ConfigLoader("config.yaml")
        >>> template = config.get("template")
        >>> merged = config.merge_with_args(cli_args)
    """
    
    def __init__(self, config_path: Optional[str] = None):
        """
        初始化設定載入器
        
        Args:
            config_path: 設定檔路徑，若為 None 則使用預設設定
        """
        self.config_path = config_path
        self.config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """
        載入 YAML 設定檔
        
        Returns:
            dict: 設定字典
        """
        if self.config_path is None:
            return self._default_config()
        
        path = Path(self.config_path)
        if not path.exists():
            return self._default_config()
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                loaded = yaml.safe_load(f) or {}
                # 合併預設設定
                default = self._default_config()
                return self._deep_merge(default, loaded)
        except yaml.YAMLError as e:
            raise ValueError(f"設定檔格式錯誤: {e}")
    
    def _default_config(self) -> Dict[str, Any]:
        """
        預設設定
        
        Returns:
            dict: 預設設定字典
        """
        return {
            'template': 'template.docx',
            'output_dir': 'output/',
            'parsing': {
                'indent_size': 4,
                'allow_mixed_indent': False,
                'encoding': 'utf-8'
            },
            'rendering': {
                'show_errors': True,
                'error_format': "[ERROR: 變數 '{var}' 不存在]",
                'preserve_styles': True
            },
            'validation': {
                'enabled': False,
                'schema_path': None,
                'strict_mode': False
            },
            'batch': {
                'show_progress': True,
                'continue_on_error': True,
                'output_pattern': '{name}.docx'
            },
            'logging': {
                'level': 'INFO',
                'file_output': False,
                'log_file': 'logs/renderer.log'
            }
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """
        取得設定值
        
        支援點記法存取，如 "parsing.indent_size"
        
        Args:
            key: 設定鍵名
            default: 預設值
            
        Returns:
            設定值或預設值
        """
        keys = key.split('.')
        value = self.config
        
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        
        return value
    
    def set(self, key: str, value: Any) -> None:
        """
        設定值
        
        Args:
            key: 設定鍵名
            value: 設定值
        """
        keys = key.split('.')
        config = self.config
        
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]
        
        config[keys[-1]] = value
    
    def merge_with_args(self, cli_args: Dict[str, Any]) -> Dict[str, Any]:
        """
        合併 CLI 參數（CLI 優先）
        
        Args:
            cli_args: 命令列參數字典
            
        Returns:
            dict: 合併後的設定
        """
        merged = self.config.copy()
        
        for key, value in cli_args.items():
            if value is not None:
                merged[key] = value
        
        return merged
    
    def _deep_merge(self, base: Dict[str, Any], 
                    override: Dict[str, Any]) -> Dict[str, Any]:
        """
        深度合併字典
        
        Args:
            base: 基礎字典
            override: 覆寫字典
            
        Returns:
            dict: 合併後的字典
        """
        result = base.copy()
        
        for key, value in override.items():
            if key in result and isinstance(result[key], dict) and isinstance(value, dict):
                result[key] = self._deep_merge(result[key], value)
            else:
                result[key] = value
        
        return result
    
    def save(self, output_path: Optional[str] = None) -> None:
        """
        儲存設定至檔案
        
        Args:
            output_path: 輸出路徑，若為 None 則使用原始路徑
        """
        path = output_path or self.config_path
        if path is None:
            raise ValueError("請指定輸出路徑")
        
        output = Path(path)
        output.parent.mkdir(parents=True, exist_ok=True)
        
        with open(output, 'w', encoding='utf-8') as f:
            yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
