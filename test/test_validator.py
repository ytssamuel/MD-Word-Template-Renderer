"""
Schema 驗證器測試
"""

import pytest
import sys
from pathlib import Path

# 添加 src 到 path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from md_word_renderer.validator import SchemaValidator


class TestSchemaValidator:
    """Schema 驗證器測試"""
    
    def setup_method(self):
        self.validator = SchemaValidator()
    
    def test_validate_empty_data(self):
        """測試空資料驗證"""
        is_valid, errors = self.validator.validate({})
        assert is_valid
        assert len(errors) == 0
    
    def test_validate_simple_data(self):
        """測試簡單資料驗證"""
        data = {
            "系統名稱": "測試系統",
            "版本": "1.0"
        }
        is_valid, errors = self.validator.validate(data)
        assert is_valid
    
    def test_validate_with_children(self):
        """測試包含子項目的資料"""
        data = {
            "異動內容": [
                {"number": "1", "value": "項目1", "children": []},
                {"number": "2", "value": "項目2", "children": [
                    {"number": "1", "value": "子項目"}
                ]}
            ]
        }
        is_valid, errors = self.validator.validate(data)
        assert is_valid
    
    def test_load_custom_schema(self):
        """測試載入自訂 Schema"""
        schema_path = Path(__file__).parent.parent / 'src' / 'md_word_renderer' / 'validator' / 'schemas' / 'markdown_schema.json'
        
        if schema_path.exists():
            validator = SchemaValidator(str(schema_path))
            is_valid, errors = validator.validate({"test": "value"})
            assert is_valid
    
    def test_load_nonexistent_schema(self):
        """測試載入不存在的 Schema"""
        with pytest.raises(FileNotFoundError):
            SchemaValidator("nonexistent_schema.json")


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
