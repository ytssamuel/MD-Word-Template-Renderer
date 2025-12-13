"""
MD-Word Template Renderer

將結構化的 Markdown 資料渲染至 Word 模板，實現資料驅動的文件生成系統。
"""

__version__ = "1.0.0"
__author__ = "SpeedBOT Team"

from .parser.markdown_parser import MarkdownParser
from .renderer.word_renderer import WordRenderer
from .validator import SchemaValidator

__all__ = ["MarkdownParser", "WordRenderer", "SchemaValidator"]
