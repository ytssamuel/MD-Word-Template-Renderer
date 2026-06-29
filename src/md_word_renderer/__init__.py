"""
MD-Word Template Renderer

將結構化的 Markdown 資料渲染至 Word / Excel 模板，實現資料驅動的文件生成系統。
"""

__version__ = "2.2.1"
__author__ = "ytssamuel"

from .parser.markdown_parser import MarkdownParser
from .renderer.word_renderer import WordRenderer
from .validator import SchemaValidator

__all__ = ["MarkdownParser", "WordRenderer", "SchemaValidator"]


def __getattr__(name):
    if name == "ExcelRenderer":
        from .renderer.excel_renderer import ExcelRenderer
        return ExcelRenderer
    if name == "build_renderer":
        from .renderer.factory import build_renderer
        return build_renderer
    raise AttributeError(name)
