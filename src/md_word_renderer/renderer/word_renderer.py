"""
Word 渲染器

使用 docxtpl 將資料渲染至 Word 模板
支援圖片插入功能
"""

from pathlib import Path
from typing import Dict, Any, Optional

from docxtpl import DocxTemplate

from .error_handler import RenderErrorHandler
from .image_handler import ImageHandler

try:
    from docx.shared import Cm, Mm
    HAS_DOCX_SHARED = True
except ImportError:
    HAS_DOCX_SHARED = False


class RenderError(Exception):
    """渲染錯誤"""
    pass


class WordRenderer:
    """
    使用 docxtpl 渲染 Word 模板
    
    功能：
    - 載入 Word 模板（.docx）
    - 變數替換（Jinja2 語法）
    - 條件渲染 {% if %}
    - 迴圈渲染 {% for %}
    - 圖片插入（InlineImage）
    - 錯誤處理
    
    Example:
        >>> renderer = WordRenderer()
        >>> renderer.load_template("template.docx")
        >>> renderer.render({"系統名稱": "測試系統"})
        >>> renderer.save("output.docx")
    """
    
    # 預設圖片最大寬度（15 公分）
    DEFAULT_IMAGE_WIDTH = Cm(15) if HAS_DOCX_SHARED else None
    
    def __init__(self, show_errors: bool = True, 
                 image_width: Optional[Any] = None,
                 image_height: Optional[Any] = None):
        """
        初始化渲染器
        
        Args:
            show_errors: 是否在輸出中顯示錯誤訊息
            image_width: 圖片寬度（docx.shared 單位，如 Cm(15)）
            image_height: 圖片高度（docx.shared 單位）
        """
        self.template: Optional[DocxTemplate] = None
        self.data: Optional[Dict[str, Any]] = None
        self.error_handler = RenderErrorHandler(show_errors=show_errors)
        self.show_errors = show_errors
        self.image_handler: Optional[ImageHandler] = None
        self.image_width = image_width or self.DEFAULT_IMAGE_WIDTH
        self.image_height = image_height
    
    def load_template(self, template_path: str) -> None:
        """
        載入 Word 模板
        
        Args:
            template_path: 模板檔案路徑（.docx）
            
        Raises:
            FileNotFoundError: 模板檔案不存在
            RenderError: 無法載入模板
        """
        path = Path(template_path)
        
        if not path.exists():
            raise FileNotFoundError(f"模板檔案不存在: {template_path}")
        
        if not path.suffix.lower() == '.docx':
            raise RenderError(f"不支援的檔案格式: {path.suffix}，請使用 .docx")
        
        try:
            self.template = DocxTemplate(template_path)
        except Exception as e:
            raise RenderError(f"無法載入模板: {e}")
    
    def render(self, data: Dict[str, Any]) -> None:
        """
        渲染模板
        
        Args:
            data: 解析後的資料字典
            
        Raises:
            RenderError: 渲染失敗
        """
        if self.template is None:
            raise RenderError("請先使用 load_template() 載入模板")
        
        self.data = data
        
        # 初始化圖片處理器
        self.image_handler = ImageHandler(
            template=self.template,
            max_width=self.image_width,
            max_height=self.image_height
        )
        
        # 處理資料中的圖片
        processed_data = self.image_handler.process_data(data)
        
        # 準備渲染上下文
        context = self._prepare_context(processed_data)
        
        try:
            self.template.render(context)
        except Exception as e:
            raise RenderError(f"渲染失敗: {e}")
    
    def get_missing_images(self) -> list:
        """取得渲染過程中找不到的圖片列表"""
        if self.image_handler:
            return self.image_handler.get_missing_images()
        return []
    
    def save(self, output_path: str) -> None:
        """
        儲存渲染後的文件
        
        Args:
            output_path: 輸出檔案路徑
            
        Raises:
            RenderError: 儲存失敗
        """
        if self.template is None:
            raise RenderError("請先載入並渲染模板")
        
        # 確保輸出目錄存在
        output = Path(output_path)
        output.parent.mkdir(parents=True, exist_ok=True)
        
        try:
            self.template.save(output_path)
        except Exception as e:
            raise RenderError(f"儲存失敗: {e}")
    
    def _prepare_context(self, data: Dict[str, Any]) -> Dict[str, Any]:
        """
        準備渲染上下文
        
        加入錯誤處理函數和輔助工具
        
        Args:
            data: 原始資料字典
            
        Returns:
            dict: 增強後的上下文字典
        """
        context = data.copy()
        
        # 將原始 data 字典也加入上下文
        # 這樣模板可以用 data["欄位名稱"] 存取含特殊字元（括號、斜線等）的欄位
        # 例如：{{data["需求依據(INC/PBI)"]}} 或 {{data["通知/公告方式"]}}
        context['data'] = data
        
        # 如果啟用錯誤顯示，使用 UndefinedSilent 或自訂處理
        # docxtpl/Jinja2 會自動處理未定義變數
        
        return context
    
    def render_to_file(self, data: Dict[str, Any], 
                       template_path: str, 
                       output_path: str) -> None:
        """
        一次完成載入、渲染、儲存
        
        便捷方法，適合單次使用
        
        Args:
            data: 資料字典
            template_path: 模板檔案路徑
            output_path: 輸出檔案路徑
        """
        self.load_template(template_path)
        self.render(data)
        self.save(output_path)
