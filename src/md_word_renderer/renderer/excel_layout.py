"""
Excel 樣板佈局設定

定義 ``LayoutConfig`` 與樣板 metadata 載入邏輯。
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional


EXCEL_FORBIDDEN_SHEET_CHARS = set('/\\?*[]:')


def sanitize_sheet_name(name, fallback: str = "Sheet"):
    """
    將任意字串清洗成合法的 Excel 工作表名稱

    Excel 規則：最長 31 字元、不可包含 ``/ \\ ? * [ ] :``，不可為空。

    Args:
        name: 原始名稱
        fallback: 清洗後結果為空（或全底線）時使用的替代名稱

    Returns:
        str: 合法的工作表名稱
    """
    if name is None:
        return fallback

    cleaned = "".join("_" if ch in EXCEL_FORBIDDEN_SHEET_CHARS else ch for ch in name)
    cleaned = cleaned.strip().strip("'")

    if not cleaned or all(ch == "_" for ch in cleaned):
        return fallback

    return cleaned[:31]


@dataclass
class HeaderSheetConfig:
    """基本資訊工作表設定"""

    enabled: bool = True
    name: str = "基本資訊"

    def sheet_name(self):
        return sanitize_sheet_name(self.name, fallback="Header")


@dataclass
class ImageConfig:
    """圖片嵌入設定"""

    max_width_px: int = 480
    max_height_px: int = 360


@dataclass
class TemplateEngineConfig:
    """Jinja2 模板引擎設定（v2.2.1 新增）"""

    enabled: bool = True
    auto_flatten_lists: bool = True  # 對沒有 {% for %} 的 list sheet 自動攤平
    missing_variable: str = "silent"  # silent / keep / error_format
    error_format: str = "[ERROR: 變數 '{var}' 不存在]"

    def __post_init__(self):
        valid = {"silent", "keep", "error_format"}
        if self.missing_variable not in valid:
            raise ValueError(
                f"excel.template_engine.missing_variable 必須為 {valid} 之一，"
                f"收到 {self.missing_variable!r}"
            )


@dataclass
class LayoutConfig:
    """
    Excel 樣板佈局設定

    從 ``config.yaml`` 的 ``excel:`` 區段載入；未指定時使用預設值。
    """

    default_template: str = "templates/excel/sample_template.xlsx"
    list_sheet_naming: str = "{key}"
    header_sheet: HeaderSheetConfig = field(default_factory=HeaderSheetConfig)
    extra_columns: List[str] = field(default_factory=lambda: ["source_field", "path", "depth"])
    image: ImageConfig = field(default_factory=ImageConfig)
    auto_fit_columns: bool = True
    template_engine: TemplateEngineConfig = field(default_factory=TemplateEngineConfig)

    # ------------------------------------------------------------------ helpers

    def resolve_list_sheet_name(self, key, used=None):
        raw = self.list_sheet_naming.format(key=key)
        base = sanitize_sheet_name(raw, fallback="List")
        if not used:
            return base

        candidate = base
        counter = 2
        while candidate in used:
            suffix = f"_{counter}"
            candidate = f"{base[: 31 - len(suffix)]}{suffix}"
            counter += 1
        return candidate

    @classmethod
    def from_mapping(cls, mapping):
        if not mapping:
            return cls()

        header_raw = mapping.get("header_sheet") or {}
        header = HeaderSheetConfig(
            enabled=bool(header_raw.get("enabled", True)),
            name=str(header_raw.get("name", "基本資訊")),
        )

        image_raw = mapping.get("image") or {}
        image = ImageConfig(
            max_width_px=int(image_raw.get("max_width_px", 480)),
            max_height_px=int(image_raw.get("max_height_px", 360)),
        )

        extra = mapping.get("extra_columns")
        if not extra:
            extra = ["source_field", "path", "depth"]

        te_raw = mapping.get("template_engine") or {}
        te = TemplateEngineConfig(
            enabled=bool(te_raw.get("enabled", True)),
            auto_flatten_lists=bool(te_raw.get("auto_flatten_lists", True)),
            missing_variable=str(te_raw.get("missing_variable", "silent")),
            error_format=str(te_raw.get("error_format", "[ERROR: 變數 '{var}' 不存在]")),
        )

        return cls(
            default_template=str(mapping.get("default_template", cls.default_template)),
            list_sheet_naming=str(mapping.get("list_sheet_naming", "{key}")),
            header_sheet=header,
            extra_columns=list(extra),
            image=image,
            auto_fit_columns=bool(mapping.get("auto_fit_columns", True)),
            template_engine=te,
        )


# Backwards-compatible alias for older tests / callers
def load_layout_from_mapping(mapping):
    return LayoutConfig.from_mapping(mapping)
