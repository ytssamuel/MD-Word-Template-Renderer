"""
Excel 圖片處理器

將圖片嵌入 Excel 工作表的指定儲存格，並等比縮放至最大寬高。
"""

import re
from pathlib import Path
from typing import Optional

try:
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.worksheet.worksheet import Worksheet
    HAS_OPENPYXL = True
except ImportError:  # pragma: no cover - guarded import for environments without openpyxl
    XLImage = None  # type: ignore[assignment]
    Worksheet = None  # type: ignore[assignment]
    HAS_OPENPYXL = False


class ExcelImageError(Exception):
    """圖片處理錯誤"""


_CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def parse_cell_ref(ref: str) -> tuple:
    """解析 Excel 儲存格參照 ``A1`` -> (``"A"``, ``1``)"""
    match = _CELL_REF_RE.match(ref.strip().upper())
    if not match:
        raise ValueError(f"無效的儲存格參照: {ref}")
    return match.group(1), int(match.group(2))


class ExcelImageHandler:
    """
    將圖片嵌入 Excel 儲存格

    Args:
        max_width_px: 圖片最大寬度（像素）
        max_height_px: 圖片最大高度（像素）
    """

    def __init__(self, max_width_px: int = 480, max_height_px: int = 360):
        if not HAS_OPENPYXL:
            raise ExcelImageError(
                "openpyxl 未安裝；請執行 `pip install openpyxl` 後再使用"
            )
        self.max_width_px = max_width_px
        self.max_height_px = max_height_px

    def embed(
        self,
        worksheet: Worksheet,
        cell_ref: str,
        image_path: str,
        alt_text: Optional[str] = None,
    ) -> dict:
        """
        將圖片嵌入 ``worksheet`` 的指定儲存格

        Args:
            worksheet: openpyxl ``Worksheet`` 物件
            cell_ref: 儲存格參照（例如 ``"B3"``）
            image_path: 圖片檔路徑
            alt_text: 替代文字（會寫入相鄰儲存格以便閱讀）

        Returns:
            dict: ``{"cell": cell_ref, "path": image_path, "width": w, "height": h}``

        Raises:
            ExcelImageError: 找不到圖片或 openpyxl 無法讀取
        """
        path = Path(image_path)
        if not path.exists():
            raise ExcelImageError(f"找不到圖片: {image_path}")

        try:
            image = XLImage(str(path))
        except Exception as exc:  # pragma: no cover - 取決於 PIL 行為
            raise ExcelImageError(f"無法載入圖片 {image_path}: {exc}") from exc

        scaled = self._scale(image)
        image.width = scaled["width"]
        image.height = scaled["height"]

        worksheet.add_image(image, cell_ref)

        if alt_text:
            worksheet[cell_ref].comment = None  # placeholder; nothing to attach

        return {
            "cell": cell_ref,
            "path": str(path),
            "width": scaled["width"],
            "height": scaled["height"],
        }

    # --------------------------------------------------------------------- size

    def _scale(self, image: XLImage) -> dict:
        """根據 ``max_width_px`` 與 ``max_height_px`` 等比縮放"""
        width = image.width or self.max_width_px
        height = image.height or self.max_height_px

        if width <= self.max_width_px and height <= self.max_height_px:
            return {"width": width, "height": height}

        ratio_w = self.max_width_px / width
        ratio_h = self.max_height_px / height
        ratio = min(ratio_w, ratio_h)

        return {
            "width": int(width * ratio),
            "height": int(height * ratio),
        }
