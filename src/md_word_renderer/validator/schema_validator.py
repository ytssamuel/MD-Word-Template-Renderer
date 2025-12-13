"""
Schema 驗證器

使用 JSON Schema 驗證 Markdown 解析後的資料結構
"""

import json
from pathlib import Path
from typing import Dict, Any, Tuple, List, Optional

from jsonschema import validate, ValidationError, Draft7Validator


class SchemaValidator:
    """
    驗證 Markdown 資料結構
    
    使用 JSON Schema 確保資料符合預期格式
    
    Example:
        >>> validator = SchemaValidator("schema.json")
        >>> is_valid, errors = validator.validate(data)
        >>> if not is_valid:
        ...     print(errors)
    """
    
    def __init__(self, schema_path: Optional[str] = None):
        """
        初始化驗證器
        
        Args:
            schema_path: JSON Schema 檔案路徑，若為 None 則使用預設 Schema
        """
        self.schema = self._load_schema(schema_path)
        self.validator = Draft7Validator(self.schema)
    
    def validate(self, data: Dict[str, Any]) -> Tuple[bool, List[str]]:
        """
        驗證資料結構
        
        Args:
            data: 解析後的資料字典
            
        Returns:
            tuple: (是否通過, 錯誤訊息列表)
        """
        errors = []
        
        for error in self.validator.iter_errors(data):
            error_path = '.'.join(str(p) for p in error.path) if error.path else 'root'
            errors.append(f"[{error_path}] {error.message}")
        
        return len(errors) == 0, errors
    
    def validate_strict(self, data: Dict[str, Any]) -> None:
        """
        嚴格驗證，失敗時拋出異常
        
        Args:
            data: 解析後的資料字典
            
        Raises:
            ValidationError: 驗證失敗
        """
        validate(instance=data, schema=self.schema)
    
    def _load_schema(self, schema_path: Optional[str]) -> Dict[str, Any]:
        """
        載入 JSON Schema
        
        Args:
            schema_path: Schema 檔案路徑
            
        Returns:
            dict: Schema 定義
        """
        if schema_path is None:
            return self._default_schema()
        
        path = Path(schema_path)
        if not path.exists():
            raise FileNotFoundError(f"Schema 檔案不存在: {schema_path}")
        
        with open(path, 'r', encoding='utf-8') as f:
            return json.load(f)
    
    def _default_schema(self) -> Dict[str, Any]:
        """
        預設 Schema 定義
        
        Returns:
            dict: 基本的 Schema 結構
        """
        return {
            "$schema": "http://json-schema.org/draft-07/schema#",
            "type": "object",
            "additionalProperties": True,
            "description": "Markdown 資料結構 Schema"
        }
