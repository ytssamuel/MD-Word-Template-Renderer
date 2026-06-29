"""
CLI 命令列工具

提供命令列介面執行 MD → Word / Excel 轉換
"""

from .main import main, cli, create_parser, process_one, resolve_format

__all__ = ['main', 'cli', 'create_parser', 'process_one', 'resolve_format']


def __getattr__(name):
    """Lazily expose names added for Excel support without changing the public list."""
    if name == 'ExcelRenderer':
        from ..renderer.excel_renderer import ExcelRenderer  # noqa: WPS433
        return ExcelRenderer
    raise AttributeError(name)
