"""
設定視窗

應用程式設定介面
"""

import customtkinter as ctk
from pathlib import Path
from tkinter import filedialog, messagebox

from .config_manager import ConfigManager


class SettingsWindow(ctk.CTkToplevel):
    """
    設定視窗
    
    提供應用程式設定介面
    """
    
    def __init__(self, parent: ctk.CTk, config_manager: ConfigManager):
        super().__init__(parent)
        
        self.config_manager = config_manager
        self.parent = parent
        
        # 設定視窗
        self.title("設定")
        self.geometry("500x450")
        self.resizable(False, False)
        
        # 建立 UI
        self._create_widgets()
        
        # 載入設定
        self._load_settings()
        
        # 聚焦此視窗
        self.focus()
        self.grab_set()
    
    def _create_widgets(self) -> None:
        """建立所有 UI 元件"""
        self.grid_columnconfigure(0, weight=1)
        
        # 標籤頁
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        # 外觀設定
        self._create_appearance_tab()
        
        # 路徑設定
        self._create_paths_tab()
        
        # 處理選項
        self._create_options_tab()
        
        # 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkButton(
            btn_frame,
            text="儲存",
            width=100,
            command=self._save_settings
        ).pack(side="right", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="取消",
            width=100,
            command=self.destroy
        ).pack(side="right", padx=5)
        
        ctk.CTkButton(
            btn_frame,
            text="重設為預設值",
            width=120,
            command=self._reset_settings
        ).pack(side="left", padx=5)
    
    def _create_appearance_tab(self) -> None:
        """建立外觀設定標籤"""
        tab = self.tabview.add("🎨 外觀")
        tab.grid_columnconfigure(1, weight=1)
        
        # 主題選擇
        ctk.CTkLabel(tab, text="主題:").grid(
            row=0, column=0, sticky="w", padx=10, pady=10
        )
        self.theme_var = ctk.StringVar()
        self.theme_menu = ctk.CTkOptionMenu(
            tab,
            variable=self.theme_var,
            values=["dark", "light", "system"]
        )
        self.theme_menu.grid(row=0, column=1, sticky="w", padx=10, pady=10)
        
        # 縮放比例
        ctk.CTkLabel(tab, text="縮放比例:").grid(
            row=1, column=0, sticky="w", padx=10, pady=10
        )
        self.scale_var = ctk.IntVar()
        self.scale_slider = ctk.CTkSlider(
            tab,
            from_=75,
            to=150,
            variable=self.scale_var,
            number_of_steps=15
        )
        self.scale_slider.grid(row=1, column=1, sticky="ew", padx=10, pady=10)
        
        self.scale_label = ctk.CTkLabel(tab, text="100%")
        self.scale_label.grid(row=1, column=2, padx=10)
        
        self.scale_var.trace_add("write", self._update_scale_label)
    
    def _create_paths_tab(self) -> None:
        """建立路徑設定標籤"""
        tab = self.tabview.add("📁 路徑")
        tab.grid_columnconfigure(1, weight=1)
        
        # 預設模板目錄
        ctk.CTkLabel(tab, text="預設模板目錄:").grid(
            row=0, column=0, sticky="w", padx=10, pady=10
        )
        self.template_dir_var = ctk.StringVar()
        ctk.CTkEntry(tab, textvariable=self.template_dir_var).grid(
            row=0, column=1, sticky="ew", padx=5, pady=10
        )
        ctk.CTkButton(
            tab, text="...", width=40,
            command=lambda: self._browse_dir(self.template_dir_var)
        ).grid(row=0, column=2, padx=10, pady=10)
        
        # 預設輸出目錄
        ctk.CTkLabel(tab, text="預設輸出目錄:").grid(
            row=1, column=0, sticky="w", padx=10, pady=10
        )
        self.output_dir_var = ctk.StringVar()
        ctk.CTkEntry(tab, textvariable=self.output_dir_var).grid(
            row=1, column=1, sticky="ew", padx=5, pady=10
        )
        ctk.CTkButton(
            tab, text="...", width=40,
            command=lambda: self._browse_dir(self.output_dir_var)
        ).grid(row=1, column=2, padx=10, pady=10)
    
    def _create_options_tab(self) -> None:
        """建立處理選項標籤"""
        tab = self.tabview.add("⚙️ 選項")
        
        # 轉換前驗證
        self.validate_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            tab,
            text="轉換前驗證資料格式",
            variable=self.validate_var
        ).pack(anchor="w", padx=20, pady=10)
        
        # 轉換後開啟
        self.open_after_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            tab,
            text="轉換完成後自動開啟檔案",
            variable=self.open_after_var
        ).pack(anchor="w", padx=20, pady=10)
        
        # 錯誤時繼續
        self.continue_error_var = ctk.BooleanVar()
        ctk.CTkCheckBox(
            tab,
            text="批次處理時遇到錯誤繼續執行",
            variable=self.continue_error_var
        ).pack(anchor="w", padx=20, pady=10)
        
        # 批次處理檔案模式
        ctk.CTkLabel(tab, text="批次處理檔案模式:").pack(
            anchor="w", padx=20, pady=(20, 5)
        )
        self.pattern_var = ctk.StringVar()
        ctk.CTkEntry(tab, textvariable=self.pattern_var, width=200).pack(
            anchor="w", padx=20, pady=5
        )
        ctk.CTkLabel(
            tab,
            text="例如: *.md, data_*.md",
            text_color="gray"
        ).pack(anchor="w", padx=20)
    
    def _browse_dir(self, var: ctk.StringVar) -> None:
        """瀏覽目錄"""
        path = filedialog.askdirectory()
        if path:
            var.set(path)
    
    def _update_scale_label(self, *args) -> None:
        """更新縮放標籤"""
        self.scale_label.configure(text=f"{self.scale_var.get()}%")
    
    def _load_settings(self) -> None:
        """載入設定"""
        self.theme_var.set(self.config_manager.get("theme", "dark"))
        self.scale_var.set(self.config_manager.get("scale", 100))
        self.template_dir_var.set(self.config_manager.get("default_template_dir", ""))
        self.output_dir_var.set(self.config_manager.get("default_output_dir", ""))
        self.validate_var.set(self.config_manager.get("validate_before_convert", True))
        self.open_after_var.set(self.config_manager.get("open_after_convert", True))
        self.continue_error_var.set(self.config_manager.get("continue_on_error", True))
        self.pattern_var.set(self.config_manager.get("batch_file_pattern", "*.md"))
    
    def _save_settings(self) -> None:
        """儲存設定"""
        self.config_manager.update({
            "theme": self.theme_var.get(),
            "scale": self.scale_var.get(),
            "default_template_dir": self.template_dir_var.get(),
            "default_output_dir": self.output_dir_var.get(),
            "validate_before_convert": self.validate_var.get(),
            "open_after_convert": self.open_after_var.get(),
            "continue_on_error": self.continue_error_var.get(),
            "batch_file_pattern": self.pattern_var.get()
        })
        
        if self.config_manager.save():
            # 套用主題
            ctk.set_appearance_mode(self.theme_var.get())
            
            # 套用縮放
            scale = self.scale_var.get() / 100
            ctk.set_widget_scaling(scale)
            
            messagebox.showinfo("成功", "設定已儲存！\n部分設定需要重新啟動程式才會生效。")
            self.destroy()
        else:
            messagebox.showerror("錯誤", "儲存設定失敗")
    
    def _reset_settings(self) -> None:
        """重設為預設值"""
        if messagebox.askyesno("確認", "確定要重設為預設值嗎？"):
            self.config_manager.reset()
            self._load_settings()
