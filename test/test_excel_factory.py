"""
測試 renderer/factory.py 的格式推斷與 renderer 構建邏輯
"""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from md_word_renderer.renderer.factory import (
    build_renderer,
    detect_format,
    output_extension_for,
    SUPPORTED_FORMATS,
)
from md_word_renderer.renderer.word_renderer import WordRenderer


try:
    from md_word_renderer.renderer.excel_renderer import ExcelRenderer as _Excel
    HAS_EXCEL = True
except Exception:
    HAS_EXCEL = False


class TestDetectFormat:
    def test_docx_returns_docx(self, tmp_path):
        p = tmp_path / "t.docx"
        p.write_bytes(b"")
        assert detect_format(str(p)) == "docx"

    def test_xlsx_returns_xlsx(self, tmp_path):
        p = tmp_path / "t.xlsx"
        p.write_bytes(b"")
        assert detect_format(str(p)) == "xlsx"

    def test_unknown_raises(self, tmp_path):
        p = tmp_path / "t.txt"
        p.write_bytes(b"")
        with pytest.raises(ValueError):
            detect_format(str(p))

    def test_path_object_accepted(self, tmp_path):
        p = tmp_path / "t.docx"
        p.write_bytes(b"")
        assert detect_format(p) == "docx"

    def test_case_insensitive(self, tmp_path):
        p = tmp_path / "T.XLSX"
        p.write_bytes(b"")
        assert detect_format(str(p)) == "xlsx"


class TestBuildRenderer:
    def test_auto_with_docx_returns_word_renderer(self, tmp_path):
        p = tmp_path / "t.docx"
        p.write_bytes(b"")
        r = build_renderer(template_path=str(p))
        assert isinstance(r, WordRenderer)

    def test_explicit_docx_returns_word_renderer(self, tmp_path):
        p = tmp_path / "t.docx"
        p.write_bytes(b"")
        r = build_renderer(template_path=str(p), format_hint="docx")
        assert isinstance(r, WordRenderer)

    @pytest.mark.skipif(not HAS_EXCEL, reason="openpyxl 不可用")
    def test_auto_with_xlsx_returns_excel_renderer(self, tmp_path):
        p = tmp_path / "t.xlsx"
        p.write_bytes(b"")
        r = build_renderer(template_path=str(p))
        assert r.__class__.__name__ == "ExcelRenderer"

    @pytest.mark.skipif(not HAS_EXCEL, reason="openpyxl 不可用")
    def test_explicit_xlsx_returns_excel_renderer(self, tmp_path):
        p = tmp_path / "t.xlsx"
        p.write_bytes(b"")
        r = build_renderer(template_path=str(p), format_hint="xlsx")
        assert r.__class__.__name__ == "ExcelRenderer"

    def test_auto_without_template_raises(self):
        with pytest.raises(ValueError):
            build_renderer(format_hint="auto")

    def test_unsupported_format_raises(self):
        with pytest.raises(ValueError):
            build_renderer(format_hint="pdf")


class TestOutputExtensionFor:
    def test_docx(self, tmp_path):
        p = tmp_path / "t.docx"
        p.write_bytes(b"")
        assert output_extension_for(str(p)) == ".docx"

    def test_xlsx(self, tmp_path):
        p = tmp_path / "t.xlsx"
        p.write_bytes(b"")
        assert output_extension_for(str(p)) == ".xlsx"


def test_supported_formats_constant():
    assert set(SUPPORTED_FORMATS) == {"docx", "xlsx"}
