"""
GUI 錯誤處理器

提供友善的錯誤訊息和錯誤對話框
"""

import traceback
from typing import Optional, Dict, Any, Callable
from enum import Enum
import customtkinter as ctk
from tkinter import messagebox


class ErrorLevel(Enum):
    """錯誤等級"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


class ErrorCode(Enum):
    """錯誤代碼與友善訊息"""
    # 檔案相關
    FILE_NOT_FOUND = ("E001", "找不到檔案", "請確認檔案路徑是否正確，檔案是否存在。")
    FILE_READ_ERROR = ("E002", "無法讀取檔案", "檔案可能被其他程式鎖定，或沒有讀取權限。")
    FILE_WRITE_ERROR = ("E003", "無法寫入檔案", "請確認輸出路徑有寫入權限，或檔案未被其他程式開啟。")
    FILE_FORMAT_ERROR = ("E004", "檔案格式錯誤", "請確認檔案格式正確。")
    
    # Markdown 解析相關
    PARSE_SYNTAX_ERROR = ("E101", "Markdown 語法錯誤", "請檢查 Markdown 格式是否符合規範：\n- 格式：編號. 欄位名稱 | 值\n- 使用一致的縮排")
    PARSE_ENCODING_ERROR = ("E102", "檔案編碼錯誤", "請確認檔案使用 UTF-8 編碼。")
    PARSE_EMPTY_FILE = ("E103", "檔案內容為空", "Markdown 檔案沒有可解析的內容。")
    
    # 模板相關
    TEMPLATE_NOT_FOUND = ("E201", "找不到模板檔案", "請選擇有效的 Word 模板檔案 (.docx)。")
    TEMPLATE_INVALID = ("E202", "無效的模板格式", "模板檔案可能已損壞或不是有效的 Word 文件。")
    TEMPLATE_VARIABLE_ERROR = ("E203", "模板變數錯誤", "模板中的變數語法有誤，請檢查 {{ }} 標記是否完整。")
    TEMPLATE_RENDER_ERROR = ("E204", "模板渲染失敗", "資料與模板不匹配，請確認所有必要變數都有對應的值。")
    
    # 驗證相關
    VALIDATION_FAILED = ("E301", "資料驗證失敗", "資料不符合預期格式，請檢查必填欄位和資料類型。")
    VALIDATION_MISSING_FIELD = ("E302", "缺少必填欄位", "請確認所有必要的欄位都已填寫。")
    
    # 系統相關
    MEMORY_ERROR = ("E901", "記憶體不足", "請關閉其他程式後重試，或處理較小的檔案。")
    UNKNOWN_ERROR = ("E999", "發生未知錯誤", "請查看詳細錯誤訊息，或聯繫技術支援。")
    
    def __init__(self, code: str, title: str, suggestion: str):
        self.code = code
        self.title = title
        self.suggestion = suggestion


class GUIErrorHandler:
    """
    GUI 錯誤處理器
    
    提供統一的錯誤處理和友善的錯誤訊息
    """
    
    # 錯誤訊息映射
    ERROR_MAPPINGS: Dict[str, ErrorCode] = {
        "FileNotFoundError": ErrorCode.FILE_NOT_FOUND,
        "PermissionError": ErrorCode.FILE_WRITE_ERROR,
        "UnicodeDecodeError": ErrorCode.PARSE_ENCODING_ERROR,
        "ValueError": ErrorCode.PARSE_SYNTAX_ERROR,
        "MemoryError": ErrorCode.MEMORY_ERROR,
        "jinja2.exceptions.TemplateSyntaxError": ErrorCode.TEMPLATE_VARIABLE_ERROR,
        "jinja2.exceptions.UndefinedError": ErrorCode.TEMPLATE_RENDER_ERROR,
        "docx.opc.exceptions.PackageNotFoundError": ErrorCode.TEMPLATE_NOT_FOUND,
    }
    
    def __init__(self, parent: Optional[ctk.CTk] = None):
        """
        初始化錯誤處理器
        
        Args:
            parent: 父視窗
        """
        self.parent = parent
        self.error_log: list = []
    
    def handle_exception(
        self,
        exception: Exception,
        context: str = "",
        show_dialog: bool = True,
        log_callback: Optional[Callable[[str], None]] = None
    ) -> ErrorCode:
        """
        處理例外
        
        Args:
            exception: 例外物件
            context: 錯誤發生的上下文
            show_dialog: 是否顯示對話框
            log_callback: 日誌回呼函數
            
        Returns:
            ErrorCode: 對應的錯誤代碼
        """
        # 取得錯誤類型
        error_type = type(exception).__name__
        error_code = self.ERROR_MAPPINGS.get(error_type, ErrorCode.UNKNOWN_ERROR)
        
        # 特殊處理某些錯誤
        error_msg = str(exception)
        if "No such file" in error_msg or "找不到" in error_msg:
            error_code = ErrorCode.FILE_NOT_FOUND
        elif "Permission denied" in error_msg:
            error_code = ErrorCode.FILE_WRITE_ERROR
        elif "codec" in error_msg or "decode" in error_msg:
            error_code = ErrorCode.PARSE_ENCODING_ERROR
        elif "template" in error_msg.lower():
            error_code = ErrorCode.TEMPLATE_RENDER_ERROR
        
        # 記錄錯誤
        error_entry = {
            "code": error_code.code,
            "type": error_type,
            "message": error_msg,
            "context": context,
            "traceback": traceback.format_exc()
        }
        self.error_log.append(error_entry)
        
        # 日誌回呼
        if log_callback:
            log_callback(f"❌ [{error_code.code}] {error_code.title}: {error_msg}")
        
        # 顯示對話框
        if show_dialog:
            self.show_error_dialog(error_code, error_msg, context)
        
        return error_code
    
    def show_error_dialog(
        self,
        error_code: ErrorCode,
        detail: str = "",
        context: str = ""
    ) -> None:
        """
        顯示錯誤對話框
        
        Args:
            error_code: 錯誤代碼
            detail: 詳細錯誤訊息
            context: 上下文
        """
        # 組合訊息
        message = f"錯誤代碼: {error_code.code}\n\n"
        message += f"❌ {error_code.title}\n\n"
        
        if context:
            message += f"發生位置: {context}\n\n"
        
        message += f"💡 建議:\n{error_code.suggestion}\n\n"
        
        if detail:
            # 截斷過長的詳細訊息
            if len(detail) > 200:
                detail = detail[:200] + "..."
            message += f"詳細資訊:\n{detail}"
        
        messagebox.showerror("錯誤", message)
    
    def show_warning(self, title: str, message: str) -> None:
        """顯示警告對話框"""
        messagebox.showwarning(title, message)
    
    def show_info(self, title: str, message: str) -> None:
        """顯示資訊對話框"""
        messagebox.showinfo(title, message)
    
    def confirm(self, title: str, message: str) -> bool:
        """顯示確認對話框"""
        return messagebox.askyesno(title, message)
    
    def get_error_log(self) -> list:
        """取得錯誤日誌"""
        return self.error_log.copy()
    
    def clear_error_log(self) -> None:
        """清除錯誤日誌"""
        self.error_log.clear()


class ErrorDialog(ctk.CTkToplevel):
    """
    自訂錯誤對話框
    
    提供更詳細的錯誤資訊顯示
    """
    
    def __init__(
        self,
        parent: ctk.CTk,
        error_code: ErrorCode,
        detail: str = "",
        traceback_info: str = ""
    ):
        super().__init__(parent)
        
        self.title("錯誤詳情")
        self.geometry("500x400")
        self.resizable(False, False)
        
        self._create_widgets(error_code, detail, traceback_info)
        
        self.focus()
        self.grab_set()
    
    def _create_widgets(
        self,
        error_code: ErrorCode,
        detail: str,
        traceback_info: str
    ) -> None:
        """建立 UI"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        # 標題區
        header = ctk.CTkFrame(self, fg_color="red")
        header.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkLabel(
            header,
            text=f"❌ {error_code.title}",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="white"
        ).pack(pady=10)
        
        ctk.CTkLabel(
            header,
            text=f"錯誤代碼: {error_code.code}",
            text_color="white"
        ).pack(pady=5)
        
        # 建議區
        suggestion_frame = ctk.CTkFrame(self)
        suggestion_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        
        ctk.CTkLabel(
            suggestion_frame,
            text="💡 解決建議:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=10, pady=5)
        
        ctk.CTkLabel(
            suggestion_frame,
            text=error_code.suggestion,
            wraplength=450,
            justify="left"
        ).pack(anchor="w", padx=10, pady=5)
        
        # 詳細資訊
        detail_frame = ctk.CTkFrame(self)
        detail_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        detail_frame.grid_columnconfigure(0, weight=1)
        detail_frame.grid_rowconfigure(1, weight=1)
        
        ctk.CTkLabel(
            detail_frame,
            text="📋 詳細資訊:",
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        
        detail_text = ctk.CTkTextbox(detail_frame, wrap="word")
        detail_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        if detail:
            detail_text.insert("1.0", f"錯誤訊息:\n{detail}\n\n")
        if traceback_info:
            detail_text.insert("end", f"追蹤資訊:\n{traceback_info}")
        
        detail_text.configure(state="disabled")
        
        # 按鈕
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkButton(
            btn_frame,
            text="複製錯誤訊息",
            command=lambda: self._copy_error(detail, traceback_info)
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_frame,
            text="關閉",
            command=self.destroy
        ).pack(side="right", padx=10)
    
    def _copy_error(self, detail: str, traceback_info: str) -> None:
        """複製錯誤訊息到剪貼簿"""
        error_text = f"{detail}\n\n{traceback_info}"
        self.clipboard_clear()
        self.clipboard_append(error_text)
        messagebox.showinfo("已複製", "錯誤訊息已複製到剪貼簿")
