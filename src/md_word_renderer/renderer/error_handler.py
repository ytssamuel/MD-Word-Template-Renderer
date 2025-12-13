"""
渲染錯誤處理器

處理模板渲染時的錯誤，提供友善的錯誤訊息
"""

from typing import Optional


class RenderErrorHandler:
    """
    處理模板渲染時的錯誤
    
    功能：
    - 格式化錯誤訊息
    - 記錄錯誤日誌
    - 提供預設值處理
    
    Example:
        >>> handler = RenderErrorHandler()
        >>> msg = handler.format_error("不存在的變數")
        >>> print(msg)
        "[ERROR: 變數 '不存在的變數' 不存在]"
    """
    
    DEFAULT_ERROR_FORMAT = "[ERROR: 變數 '{var}' 不存在]"
    
    def __init__(self, show_errors: bool = True, 
                 error_format: Optional[str] = None):
        """
        初始化錯誤處理器
        
        Args:
            show_errors: 是否顯示錯誤訊息
            error_format: 錯誤訊息格式（使用 {var} 作為變數名稱佔位符）
        """
        self.show_errors = show_errors
        self.error_format = error_format or self.DEFAULT_ERROR_FORMAT
        self.errors: list = []
    
    def format_error(self, variable_name: str) -> str:
        """
        格式化錯誤訊息
        
        Args:
            variable_name: 找不到的變數名稱
            
        Returns:
            str: 格式化後的錯誤訊息，若不顯示錯誤則返回空字串
        """
        if not self.show_errors:
            return ""
        
        error_msg = self.error_format.format(var=variable_name)
        self.errors.append({
            'variable': variable_name,
            'message': error_msg
        })
        
        return error_msg
    
    def handle(self, exception: Exception, context: str = "") -> None:
        """
        處理渲染異常
        
        Args:
            exception: 發生的異常
            context: 錯誤發生的上下文描述
        """
        error_info = {
            'type': type(exception).__name__,
            'message': str(exception),
            'context': context
        }
        self.errors.append(error_info)
    
    def get_errors(self) -> list:
        """
        取得所有錯誤
        
        Returns:
            list: 錯誤列表
        """
        return self.errors.copy()
    
    def clear_errors(self) -> None:
        """清除錯誤記錄"""
        self.errors.clear()
    
    def has_errors(self) -> bool:
        """
        檢查是否有錯誤
        
        Returns:
            bool: 是否有錯誤
        """
        return len(self.errors) > 0
