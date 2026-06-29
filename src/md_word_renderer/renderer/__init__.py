"""
Word / Excel 渲染模組

提供 Word 模板渲染（基於 docxtpl + Jinja2）與 Excel 樣板渲染（基於 openpyxl）。
支援圖片插入功能。
"""

from .word_renderer import WordRenderer
from .error_handler import RenderErrorHandler
from .image_handler import ImageHandler

__all__ = [
    "WordRenderer",
    "RenderErrorHandler",
    "ImageHandler",
    "ExcelRenderer",
    "ExcelTemplateEngine",
    "build_renderer",
    "detect_format",
]


def __getattr__(name):
    """Lazily expose ExcelRenderer so environments without ``openpyxl`` don't break."""
    if name == "ExcelRenderer":
        from .excel_renderer import ExcelRenderer  # noqa: WPS433
        return ExcelRenderer
    if name == "LayoutConfig":
        from .excel_layout import LayoutConfig  # noqa: WPS433
        return LayoutConfig
    if name == "ExcelTemplateEngine":
        from .excel_template_engine import ExcelTemplateEngine  # noqa: WPS433
        return ExcelTemplateEngine
    if name == "build_renderer":
        from .factory import build_renderer  # noqa: WPS433
        return build_renderer
    if name == "detect_format":
        from .factory import detect_format  # noqa: WPS433
        return detect_format
    raise AttributeError(name)
