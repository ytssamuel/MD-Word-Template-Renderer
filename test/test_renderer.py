"""
Word 渲染器測試
"""

import pytest
import sys
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock

# 添加 src 到 path
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from md_word_renderer.renderer import WordRenderer, RenderErrorHandler


class TestRenderErrorHandler:
    """渲染錯誤處理器測試"""
    
    def setup_method(self):
        self.handler = RenderErrorHandler()
    
    def test_format_error(self):
        """測試錯誤訊息格式化"""
        msg = self.handler.format_error("不存在的變數")
        assert "不存在的變數" in msg
        assert "ERROR" in msg
    
    def test_format_error_disabled(self):
        """測試停用錯誤顯示"""
        handler = RenderErrorHandler(show_errors=False)
        msg = handler.format_error("變數")
        assert msg == ""
    
    def test_custom_format(self):
        """測試自訂錯誤格式"""
        handler = RenderErrorHandler(error_format="找不到: {var}")
        msg = handler.format_error("測試變數")
        assert msg == "找不到: 測試變數"
    
    def test_error_tracking(self):
        """測試錯誤追蹤"""
        self.handler.format_error("var1")
        self.handler.format_error("var2")
        
        errors = self.handler.get_errors()
        assert len(errors) == 2
        assert self.handler.has_errors()
    
    def test_clear_errors(self):
        """測試清除錯誤"""
        self.handler.format_error("var1")
        self.handler.clear_errors()
        
        assert not self.handler.has_errors()
        assert len(self.handler.get_errors()) == 0


class TestWordRenderer:
    """Word 渲染器測試"""
    
    def setup_method(self):
        self.renderer = WordRenderer()
    
    def test_init(self):
        """測試初始化"""
        assert self.renderer.template is None
        assert self.renderer.show_errors is True
    
    def test_load_template_not_found(self):
        """測試載入不存在的模板"""
        with pytest.raises(FileNotFoundError):
            self.renderer.load_template("not_exist.docx")
    
    def test_render_without_template(self):
        """測試未載入模板就渲染"""
        from md_word_renderer.renderer.word_renderer import RenderError
        
        with pytest.raises(RenderError, match="請先使用 load_template"):
            self.renderer.render({"test": "value"})
    
    @patch('md_word_renderer.renderer.word_renderer.DocxTemplate')
    def test_load_template_success(self, mock_docx):
        """測試成功載入模板（mock）"""
        # 建立臨時測試檔案
        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            temp_path = f.name
        
        try:
            self.renderer.load_template(temp_path)
            mock_docx.assert_called_once_with(temp_path)
        finally:
            Path(temp_path).unlink(missing_ok=True)
    
    @patch('md_word_renderer.renderer.word_renderer.DocxTemplate')
    def test_render_success(self, mock_docx):
        """測試渲染成功（mock）"""
        mock_template = MagicMock()
        mock_docx.return_value = mock_template
        
        import tempfile
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            temp_path = f.name
        
        try:
            self.renderer.load_template(temp_path)
            self.renderer.render({"系統名稱": "測試系統"})
            
            mock_template.render.assert_called_once()
        finally:
            Path(temp_path).unlink(missing_ok=True)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
