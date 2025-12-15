"""
Markdown 解析器測試
"""

import pytest
import sys
from pathlib import Path

# 添加 src 到 path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from md_word_renderer.parser import MarkdownParser, IndentDetector, EscapeHandler


class TestEscapeHandler:
    """特殊字元處理器測試"""
    
    def setup_method(self):
        self.handler = EscapeHandler()
    
    def test_unescape_pipe(self):
        """測試管線符號轉義"""
        assert self.handler.unescape("這是\\|管線") == "這是|管線"
    
    def test_unescape_newline(self):
        """測試換行轉義"""
        assert self.handler.unescape("第一行\\n第二行") == "第一行\n第二行"
    
    def test_unescape_backslash(self):
        """測試反斜線轉義"""
        assert self.handler.unescape("路徑\\\\文件") == "路徑\\文件"
    
    def test_unescape_quote(self):
        """測試引號轉義"""
        assert self.handler.unescape('他說:\\"你好\\"') == '他說:"你好"'
    
    def test_unescape_double_quote(self):
        """測試連續雙引號"""
        assert self.handler.unescape('""引號""') == '"引號"'
    
    def test_unescape_empty(self):
        """測試空字串"""
        assert self.handler.unescape("") == ""
        assert self.handler.unescape(None) is None


class TestIndentDetector:
    """縮排偵測器測試"""
    
    def setup_method(self):
        self.detector = IndentDetector(indent_size=4)
    
    def test_detect_space_indent(self):
        """測試空格縮排偵測"""
        lines = [
            "1. 第一項 | 值",
            "    1. 子項目",
            "        1. 孫項目"
        ]
        indent_type, indent_unit = self.detector.detect(lines)
        assert indent_type == 'space'
        assert indent_unit == 4
    
    def test_detect_tab_indent(self):
        """測試 Tab 縮排偵測"""
        lines = [
            "1. 第一項 | 值",
            "\t1. 子項目",
            "\t\t1. 孫項目"
        ]
        indent_type, indent_unit = self.detector.detect(lines)
        assert indent_type == 'tab'
        assert indent_unit == 1
    
    def test_calculate_level(self):
        """測試層級計算"""
        assert self.detector.calculate_level("無縮排", 'space', 4) == 0
        assert self.detector.calculate_level("    四空格", 'space', 4) == 1
        assert self.detector.calculate_level("        八空格", 'space', 4) == 2
    
    def test_validate_consistent(self):
        """測試一致性驗證"""
        lines = [
            "1. 項目 | 值",
            "    1. 子項目",
            "    2. 子項目2"
        ]
        is_valid, errors = self.detector.validate(lines)
        assert is_valid
        assert len(errors) == 0


class TestMarkdownParser:
    """Markdown 解析器測試"""
    
    def setup_method(self):
        self.parser = MarkdownParser()
    
    def test_parse_simple_line(self):
        """測試簡單行解析"""
        content = "1. 系統名稱 | 範例系統"
        result = self.parser.parse_content(content)
        
        assert "系統名稱" in result
        assert result["系統名稱"] == "範例系統"
    
    def test_parse_empty_value(self):
        """測試空值解析"""
        content = "1. 中介軟體 | "
        result = self.parser.parse_content(content)
        
        assert "中介軟體" in result
        assert result["中介軟體"] == ""
    
    def test_parse_hierarchy(self):
        """測試階層解析"""
        content = """16. 異動內容-測試案例 | 
    1. 調整供應商登入之密碼功能
       1. TC001：新增供應商使用者
       2. TC002：超過最長使用設定"""
        
        result = self.parser.parse_content(content)
        
        assert "異動內容-測試案例" in result
        # 應該是列表類型
        items = result["異動內容-測試案例"]
        assert isinstance(items, list)
    
    def test_parse_with_escape(self):
        """測試包含轉義字元的解析"""
        content = "1. 特殊欄位 | 這是\\|管線符號"
        result = self.parser.parse_content(content)
        
        assert result["特殊欄位"] == "這是|管線符號"
    
    def test_parse_numbered_access(self):
        """測試編號存取"""
        content = """1. 系統名稱 | 範例系統
2. 版本號 | v1.0"""
        result = self.parser.parse_content(content)
        
        assert "#1" in result
        assert result["#1"]["key"] == "系統名稱"
        assert result["#1"]["value"] == "範例系統"


class TestMarkdownParserIntegration:
    """Markdown 解析器整合測試"""
    
    def setup_method(self):
        self.parser = MarkdownParser()
        # 測試資料路徑
        self.test_file = Path(__file__).parent.parent / 'referance' / 'sample_data.md'
    
    def test_parse_real_file(self):
        """測試解析範例檔案"""
        if not self.test_file.exists():
            pytest.skip(f"測試檔案不存在: {self.test_file}")
        
        result = self.parser.parse(str(self.test_file))
        
        # 基本欄位檢查
        assert "系統名稱" in result
        assert "範例系統" in result["系統名稱"]
        
        # 列表欄位檢查
        assert "異動內容-測試案例" in result
        
        # 編號索引檢查
        assert "#1" in result


class TestImageParsing:
    """圖片解析測試"""
    
    def setup_method(self):
        self.parser = MarkdownParser()
    
    def test_parse_image_syntax(self):
        """測試解析圖片語法"""
        content = """1. 測試案例 | 
   1. 截圖1
      1. ![image1](path/to/image1.png)
      2. ![image2](path/to/image2.png)"""
        
        result = self.parser.parse_content(content)
        
        assert "測試案例" in result
        test_cases = result["測試案例"]
        assert len(test_cases) == 1
        
        # 檢查子項目
        first_child = test_cases[0]
        assert first_child["value"] == "截圖1"
        assert len(first_child["children"]) == 2
        
        # 檢查圖片項目
        img1 = first_child["children"][0]
        assert img1["type"] == "image"
        assert img1["value"] == "image1"
        assert "path/to/image1.png" in img1["image_path"]
    
    def test_parse_image_with_alt_text(self):
        """測試解析帶有 alt 文字的圖片"""
        content = """1. 測試 | 
   1. ![這是描述文字](images/test.png)"""
        
        result = self.parser.parse_content(content)
        
        test_items = result["測試"]
        img = test_items[0]
        assert img["type"] == "image"
        assert img["value"] == "這是描述文字"
        assert img["image_alt"] == "這是描述文字"
    
    def test_parse_mixed_content(self):
        """測試解析混合內容（文字和圖片）"""
        content = """1. 步驟 | 
   1. 第一步
   2. ![截圖](img/step1.png)
   3. 第三步"""
        
        result = self.parser.parse_content(content)
        
        steps = result["步驟"]
        assert len(steps) == 3
        assert steps[0]["type"] == "text"
        assert steps[0]["value"] == "第一步"
        assert steps[1]["type"] == "image"
        assert steps[2]["type"] == "text"
        assert steps[2]["value"] == "第三步"
    
    def test_parse_sample_data_with_images(self):
        """測試解析含圖片的範例資料"""
        test_file = Path(__file__).parent.parent / 'referance' / 'sample_data.md'
        if not test_file.exists():
            pytest.skip(f"測試檔案不存在: {test_file}")
        
        result = self.parser.parse(str(test_file))
        
        # 檢查異動內容中的圖片
        test_cases = result.get("異動內容-測試案例", [])
        assert len(test_cases) > 0
        
        # 找到包含圖片的子項目
        first_case = test_cases[0]
        assert first_case["value"] == "調整供應商登入之密碼功能"
        
        # 檢查是否有圖片類型的子項目
        has_images = False
        for child in first_case.get("children", []):
            for sub_child in child.get("children", []):
                if sub_child.get("type") == "image":
                    has_images = True
                    assert "image_path" in sub_child
                    assert sub_child["image_path"].endswith(".png")
                    break
        
        assert has_images, "應該找到至少一個圖片"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
