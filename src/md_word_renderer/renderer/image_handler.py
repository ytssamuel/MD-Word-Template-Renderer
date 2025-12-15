"""
圖片處理模組

處理 Markdown 中的圖片，支援 docxtpl 的 InlineImage
"""

from pathlib import Path
from typing import Dict, Any, Optional, List, TYPE_CHECKING
import os

if TYPE_CHECKING:
    from docxtpl import DocxTemplate

try:
    from docxtpl import InlineImage
    from docx.shared import Mm, Cm, Inches, Pt
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False


class ImageHandler:
    """
    處理圖片插入到 Word 文件
    
    功能：
    - 驗證圖片檔案存在
    - 建立 docxtpl InlineImage 物件
    - 支援自訂圖片尺寸
    - 處理資料結構中的圖片欄位
    
    Example:
        >>> handler = ImageHandler(template)
        >>> handler.set_max_width(Cm(15))
        >>> processed_data = handler.process_data(data)
    """
    
    # 支援的圖片格式
    SUPPORTED_FORMATS = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
    
    def __init__(self, template: Optional['DocxTemplate'] = None, 
                 max_width: Optional[Any] = None,
                 max_height: Optional[Any] = None):
        """
        初始化圖片處理器
        
        Args:
            template: docxtpl 模板物件（用於建立 InlineImage）
            max_width: 圖片最大寬度（docx.shared 單位，如 Cm(15)）
            max_height: 圖片最大高度（docx.shared 單位）
        """
        self.template = template
        self.max_width = max_width
        self.max_height = max_height
        self._missing_images: List[str] = []
    
    def set_template(self, template: 'DocxTemplate') -> None:
        """設定模板物件"""
        self.template = template
    
    def set_max_width(self, width: Any) -> None:
        """設定圖片最大寬度"""
        self.max_width = width
    
    def set_max_height(self, height: Any) -> None:
        """設定圖片最大高度"""
        self.max_height = height
    
    def get_missing_images(self) -> List[str]:
        """取得找不到的圖片路徑列表"""
        return self._missing_images.copy()
    
    def clear_missing_images(self) -> None:
        """清除找不到的圖片記錄"""
        self._missing_images.clear()
    
    def validate_image(self, image_path: str) -> bool:
        """
        驗證圖片檔案是否存在且格式支援
        
        Args:
            image_path: 圖片檔案路徑
            
        Returns:
            bool: 圖片是否有效
        """
        path = Path(image_path)
        
        # 檢查檔案是否存在
        if not path.exists():
            return False
        
        # 檢查副檔名
        if path.suffix.lower() not in self.SUPPORTED_FORMATS:
            return False
        
        return True
    
    def create_inline_image(self, image_path: str, 
                           width: Optional[Any] = None,
                           height: Optional[Any] = None) -> Optional['InlineImage']:
        """
        建立 docxtpl InlineImage 物件
        
        Args:
            image_path: 圖片檔案路徑
            width: 圖片寬度（覆蓋預設值）
            height: 圖片高度（覆蓋預設值）
            
        Returns:
            InlineImage: 圖片物件，若失敗則返回 None
        """
        if not HAS_DOCX:
            raise ImportError("需要安裝 docxtpl 和 python-docx")
        
        if self.template is None:
            raise ValueError("請先設定 template")
        
        if not self.validate_image(image_path):
            self._missing_images.append(image_path)
            return None
        
        # 決定尺寸（優先使用參數，否則使用預設值）
        img_width = width or self.max_width
        img_height = height or self.max_height
        
        try:
            return InlineImage(
                self.template,
                image_path,
                width=img_width,
                height=img_height
            )
        except Exception as e:
            print(f"警告: 無法載入圖片 {image_path}: {e}")
            self._missing_images.append(image_path)
            return None
    
    def process_data(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        處理資料字典，將圖片路徑轉換為 InlineImage 物件
        
        遞迴處理整個資料結構，將 type='image' 的項目中的
        image_path 轉換為可用於模板渲染的 InlineImage 物件
        
        Args:
            data: 解析後的資料字典
            
        Returns:
            dict: 處理後的資料字典（包含 InlineImage 物件）
        """
        if not HAS_DOCX:
            return data
        
        self.clear_missing_images()
        return self._process_value(data)
    
    def _process_value(self, value: Any) -> Any:
        """
        遞迴處理值
        
        Args:
            value: 任意值（可能是 dict、list 或其他）
            
        Returns:
            處理後的值
        """
        if isinstance(value, dict):
            return self._process_dict(value)
        elif isinstance(value, list):
            return self._process_list(value)
        else:
            return value
    
    def _process_dict(self, d: Dict[str, Any]) -> Dict[str, Any]:
        """
        處理字典
        
        如果字典是圖片類型，則建立 InlineImage
        
        Args:
            d: 字典
            
        Returns:
            dict: 處理後的字典
        """
        result = {}
        
        for key, value in d.items():
            if key == 'children' and isinstance(value, list):
                # 處理 children 列表
                result[key] = self._process_list(value)
            elif isinstance(value, (dict, list)):
                result[key] = self._process_value(value)
            else:
                result[key] = value
        
        # 如果是圖片類型，建立 InlineImage
        if result.get('type') == 'image' and 'image_path' in result:
            image_path = result['image_path']
            inline_image = self.create_inline_image(image_path)
            if inline_image:
                result['image'] = inline_image
            else:
                # 圖片不存在時的替代文字
                result['image'] = f"[圖片無法載入: {Path(image_path).name}]"
        
        return result
    
    def _process_list(self, lst: List[Any]) -> List[Any]:
        """
        處理列表
        
        Args:
            lst: 列表
            
        Returns:
            list: 處理後的列表
        """
        return [self._process_value(item) for item in lst]


# 便捷函數
def process_images_in_data(data: Dict[str, Any], 
                          template: 'DocxTemplate',
                          max_width: Optional[Any] = None,
                          max_height: Optional[Any] = None) -> Dict[str, Any]:
    """
    處理資料中的圖片（便捷函數）
    
    Args:
        data: 解析後的資料字典
        template: docxtpl 模板物件
        max_width: 圖片最大寬度
        max_height: 圖片最大高度
        
    Returns:
        dict: 處理後的資料字典
    """
    handler = ImageHandler(template, max_width, max_height)
    return handler.process_data(data)
