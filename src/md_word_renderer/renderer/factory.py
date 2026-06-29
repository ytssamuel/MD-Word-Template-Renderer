"""
Renderer factory

依據樣板副檔名或明確的 ``format_hint`` 建立 ``WordRenderer`` 或 ``ExcelRenderer``。
"""

from pathlib import Path
from typing import Optional, Union

from .word_renderer import WordRenderer
from .excel_renderer import ExcelRenderer
from .excel_layout import LayoutConfig


SUPPORTED_FORMATS = ("docx", "xlsx")


def detect_format(template_path: Union[str, Path]) -> str:
    """從樣板副檔名推斷格式：``.docx`` → ``"docx"``、``.xlsx`` → ``"xlsx"``"""
    path = Path(template_path)
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return "docx"
    if suffix == ".xlsx":
        return "xlsx"
    raise ValueError(
        f"無法從副檔名 {suffix!r} 推斷格式；支援的副檔名：.docx / .xlsx"
    )


def build_renderer(
    template_path: Optional[Union[str, Path]] = None,
    format_hint: str = "auto",
    layout: Optional[LayoutConfig] = None,
) -> Union[WordRenderer, ExcelRenderer]:
    """
    建立合適的 renderer

    Args:
        template_path: 樣板檔路徑；當 ``format_hint == "auto"`` 時必填
        format_hint: ``"auto"`` / ``"docx"`` / ``"xlsx"``
        layout: Excel 樣板用，Word 樣板忽略

    Returns:
        :class:`WordRenderer` 或 :class:`ExcelRenderer`
    """
    fmt = (format_hint or "auto").lower()

    if fmt == "auto":
        if template_path is None:
            raise ValueError("format_hint='auto' 時必須提供 template_path")
        fmt = detect_format(template_path)

    if fmt not in SUPPORTED_FORMATS:
        raise ValueError(f"不支援的格式: {fmt!r}；可用值：{SUPPORTED_FORMATS}")

    if fmt == "docx":
        return WordRenderer()

    # xlsx
    return ExcelRenderer(layout=layout)


def output_extension_for(template_path: Union[str, Path]) -> str:
    """樣板副檔名 → 輸出副檔名（``.docx`` → ``.docx``、``.xlsx`` → ``.xlsx``）"""
    fmt = detect_format(template_path)
    return f".{fmt}"
