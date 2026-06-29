"""
主視窗

MD-Word/Excel Template Renderer GUI 主視窗實現
"""

import customtkinter as ctk
from pathlib import Path
from threading import Thread
from typing import Optional, Callable
import tkinter as tk
from tkinter import filedialog, messagebox

from .config_manager import ConfigManager
from .error_handler import GUIErrorHandler, ErrorCode
from md_word_renderer.renderer.factory import build_renderer, detect_format, output_extension_for


class MainWindow(ctk.CTk):
    """
    主視窗類別
    
    提供單一文件轉換的主要操作介面
    """
    
    def __init__(self):
        super().__init__()

        # 設定管理器
        self.config_manager = ConfigManager()

        # 載入設定
        self._apply_settings()

        # 設定視窗
        self.title("MD-Word/Excel Template Renderer")
        self._configure_window()
        
        # 檔案路徑變數
        self.markdown_path = ctk.StringVar()
        self.template_path = ctk.StringVar()
        self.output_path = ctk.StringVar()
        
        # 建立 UI
        self._create_widgets()
        
        # 綁定關閉事件
        self.protocol("WM_DELETE_WINDOW", self._on_closing)
    
    def _apply_settings(self) -> None:
        """套用設定"""
        # 主題
        theme = self.config_manager.get("theme", "dark")
        ctk.set_appearance_mode(theme)
        ctk.set_default_color_theme("blue")
        
        # 縮放
        scale = self.config_manager.get("scale", 100) / 100
        ctk.set_widget_scaling(scale)
    
    def _configure_window(self) -> None:
        """設定視窗大小和位置"""
        width = self.config_manager.get("window_width", 900)
        height = self.config_manager.get("window_height", 700)
        
        x = self.config_manager.get("window_x")
        y = self.config_manager.get("window_y")
        
        if x is not None and y is not None:
            self.geometry(f"{width}x{height}+{x}+{y}")
        else:
            self.geometry(f"{width}x{height}")
            self.center_window()
        
        self.minsize(700, 500)
    
    def center_window(self) -> None:
        """置中視窗"""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")
    
    def _create_widgets(self) -> None:
        """建立所有 UI 元件"""
        # 設定 grid 權重
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        # 頂部工具列
        self._create_toolbar()
        
        # 檔案選擇區
        self._create_file_section()
        
        # 預覽區
        self._create_preview_section()
        
        # 狀態列
        self._create_status_bar()
    
    def _create_toolbar(self) -> None:
        """建立工具列"""
        toolbar = ctk.CTkFrame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        # 批次處理按鈕（多個 MD + 一個模板）
        btn_batch = ctk.CTkButton(
            toolbar,
            text="📁 批次處理",
            width=100,
            command=self._open_batch_window
        )
        btn_batch.pack(side="left", padx=5)
        
        # 多模板批次處理按鈕（一個 MD + 多個模板）
        btn_multi_template = ctk.CTkButton(
            toolbar,
            text="📋 多模板批次",
            width=110,
            command=self._open_multi_template_window
        )
        btn_multi_template.pack(side="left", padx=5)
        
        # 設定按鈕
        btn_settings = ctk.CTkButton(
            toolbar,
            text="⚙️ 設定",
            width=80,
            command=self._open_settings
        )
        btn_settings.pack(side="left", padx=5)
        
        # 模板預覽按鈕
        btn_template_preview = ctk.CTkButton(
            toolbar,
            text="🔍 模板預覽",
            width=100,
            command=self._open_template_preview
        )
        btn_template_preview.pack(side="left", padx=5)
        
        # 驗證按鈕
        btn_validate = ctk.CTkButton(
            toolbar,
            text="✅ 驗證",
            width=80,
            command=self._validate_data
        )
        btn_validate.pack(side="left", padx=5)
        
        # 說明按鈕
        btn_help = ctk.CTkButton(
            toolbar,
            text="❓ 說明",
            width=80,
            command=self._show_help
        )
        btn_help.pack(side="left", padx=5)
        
        # 主題切換
        self.theme_switch = ctk.CTkSwitch(
            toolbar,
            text="深色模式",
            command=self._toggle_theme
        )
        self.theme_switch.pack(side="right", padx=10)
        
        # 設定初始狀態
        if self.config_manager.get("theme") == "dark":
            self.theme_switch.select()
    
    def _create_file_section(self) -> None:
        """建立檔案選擇區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        frame.grid_columnconfigure(1, weight=1)
        
        # Markdown 檔案
        ctk.CTkLabel(frame, text="📄 Markdown 檔案:").grid(
            row=0, column=0, sticky="w", padx=10, pady=5
        )
        ctk.CTkEntry(frame, textvariable=self.markdown_path).grid(
            row=0, column=1, sticky="ew", padx=5, pady=5
        )
        ctk.CTkButton(
            frame, text="瀏覽...", width=80,
            command=lambda: self._browse_file("markdown")
        ).grid(row=0, column=2, padx=10, pady=5)
        
        # 模板檔案
        self.template_label = ctk.CTkLabel(frame, text="📋 Word/Excel 模板:")
        self.template_label.grid(
            row=1, column=0, sticky="w", padx=10, pady=5
        )
        ctk.CTkEntry(frame, textvariable=self.template_path).grid(
            row=1, column=1, sticky="ew", padx=5, pady=5
        )

        template_btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        template_btn_frame.grid(row=1, column=2, padx=5, pady=5)

        ctk.CTkButton(
            template_btn_frame, text="瀏覽", width=60,
            command=lambda: self._browse_file("template")
        ).pack(side="left", padx=2)

        ctk.CTkButton(
            template_btn_frame, text="🔍", width=30,
            command=self._open_template_preview
        ).pack(side="left", padx=2)

        # 輸出檔案
        ctk.CTkLabel(frame, text="💾 輸出檔案:").grid(
            row=2, column=0, sticky="w", padx=10, pady=5
        )
        ctk.CTkEntry(frame, textvariable=self.output_path).grid(
            row=2, column=1, sticky="ew", padx=5, pady=5
        )
        ctk.CTkButton(
            frame, text="瀏覽...", width=80,
            command=lambda: self._browse_file("output")
        ).grid(row=2, column=2, padx=10, pady=5)

        # 轉換按鈕
        self.btn_convert = ctk.CTkButton(
            frame,
            text="🚀 開始轉換",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=40,
            command=self._start_conversion
        )
        self.btn_convert.grid(row=3, column=0, columnspan=3, pady=15)

        # 格式標籤（顯示目前偵測到的輸出格式）
        self.format_label = ctk.CTkLabel(
            frame,
            text="🔎 輸出格式: (尚未選擇樣板)",
            anchor="w",
        )
        self.format_label.grid(row=4, column=0, columnspan=3, sticky="w", padx=10)

        # 當使用者變更樣板路徑時，重新整理格式標籤與輸出預設
        self.template_path.trace_add("write", lambda *_: self._on_template_changed())
    
    def _create_preview_section(self) -> None:
        """建立預覽區"""
        # 標籤頁
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        
        # 資料預覽
        tab_data = self.tabview.add("📊 資料預覽")
        tab_data.grid_columnconfigure(0, weight=1)
        tab_data.grid_rowconfigure(0, weight=1)
        
        self.data_preview = ctk.CTkTextbox(tab_data, wrap="word")
        self.data_preview.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # 日誌
        tab_log = self.tabview.add("📝 處理日誌")
        tab_log.grid_columnconfigure(0, weight=1)
        tab_log.grid_rowconfigure(0, weight=1)
        
        self.log_text = ctk.CTkTextbox(tab_log, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
    
    def _create_status_bar(self) -> None:
        """建立狀態列"""
        self.status_frame = ctk.CTkFrame(self, height=30)
        self.status_frame.grid(row=3, column=0, sticky="ew", padx=5, pady=5)
        self.status_frame.grid_columnconfigure(0, weight=1)
        
        # 狀態文字
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="就緒",
            anchor="w"
        )
        self.status_label.grid(row=0, column=0, sticky="w", padx=10)
        
        # 進度條
        self.progress_bar = ctk.CTkProgressBar(self.status_frame)
        self.progress_bar.grid(row=0, column=1, sticky="e", padx=10)
        self.progress_bar.set(0)
    
    def _browse_file(self, file_type: str) -> None:
        """
        瀏覽檔案
        
        Args:
            file_type: 檔案類型 ("markdown", "template", "output")
        """
        if file_type == "markdown":
            initial_dir = self.config_manager.get("last_markdown_dir", "")
            path = filedialog.askopenfilename(
                title="選擇 Markdown 檔案",
                filetypes=[("Markdown 檔案", "*.md"), ("所有檔案", "*.*")],
                initialdir=initial_dir or None
            )
            if path:
                self.markdown_path.set(path)
                self.config_manager.set("last_markdown_dir", str(Path(path).parent))
                self._load_markdown_preview(path)
                self._auto_set_output_path(path)
                
        elif file_type == "template":
            initial_dir = self.config_manager.get("default_template_dir", "")
            path = filedialog.askopenfilename(
                title="選擇 Word/Excel 模板",
                filetypes=[
                    ("Word 文件", "*.docx"),
                    ("Excel 樣板", "*.xlsx"),
                    ("所有檔案", "*.*"),
                ],
                initialdir=initial_dir or None
            )
            if path:
                self.template_path.set(path)
                self.config_manager.set("last_template_file", path)
                self._on_template_changed()

        elif file_type == "output":
            initial_dir = self.config_manager.get("default_output_dir", "")
            ext = self._detect_template_extension()
            filetypes = {
                ".docx": [("Word 文件", "*.docx")],
                ".xlsx": [("Excel 活頁簿", "*.xlsx")],
            }.get(ext, [("Word 文件", "*.docx"), ("Excel 活頁簿", "*.xlsx")])
            path = filedialog.asksaveasfilename(
                title="儲存輸出檔案",
                filetypes=filetypes,
                defaultextension=ext or ".docx",
                initialdir=initial_dir or None
            )
            if path:
                self.output_path.set(path)

    def _on_template_changed(self) -> None:
        """樣板路徑變更時同步：格式標籤 / 輸出檔名預設"""
        template = self.template_path.get()
        if not template:
            if hasattr(self, "format_label"):
                self.format_label.configure(text="🔎 輸出格式: (尚未選擇樣板)")
            return

        try:
            fmt = detect_format(template)
        except ValueError:
            if hasattr(self, "format_label"):
                self.format_label.configure(text="🔎 輸出格式: (不支援的副檔名)")
            return

        if hasattr(self, "format_label"):
            self.format_label.configure(text=f"🔎 輸出格式: {fmt}")

        # 若輸出檔案是空的或副檔名與新格式不符，自動修正
        ext = f".{fmt}"
        current = self.output_path.get()
        if not current:
            md = self.markdown_path.get()
            if md:
                mp = Path(md)
                output_dir = self.config_manager.get("default_output_dir")
                if output_dir:
                    new_output = Path(output_dir) / f"{mp.stem}_output{ext}"
                else:
                    new_output = mp.with_suffix(ext)
                self.output_path.set(str(new_output))
        else:
            cp = Path(current)
            if cp.suffix.lower() != ext:
                self.output_path.set(str(cp.with_suffix(ext)))

    def _detect_template_extension(self) -> Optional[str]:
        template = self.template_path.get()
        if not template:
            return None
        try:
            return f".{detect_format(template)}"
        except ValueError:
            return None

    def _auto_set_output_path(self, markdown_path: str) -> None:
        """根據 Markdown 路徑自動設定輸出路徑（沿用樣板的副檔名）"""
        if not self.output_path.get():
            md_path = Path(markdown_path)
            output_dir = self.config_manager.get("default_output_dir")
            ext = self._detect_template_extension() or ".docx"
            if output_dir:
                output = Path(output_dir) / f"{md_path.stem}_output{ext}"
            else:
                output = md_path.with_suffix(ext)
            self.output_path.set(str(output))
    
    def _load_markdown_preview(self, path: str) -> None:
        """載入 Markdown 預覽"""
        try:
            # 匯入 parser
            from md_word_renderer.parser.markdown_parser import MarkdownParser
            
            parser = MarkdownParser()
            data = parser.parse(path)
            
            # 顯示解析結果
            self.data_preview.delete("1.0", "end")
            self.data_preview.insert("1.0", "解析結果:\n" + "-" * 50 + "\n\n")
            
            for key, value in data.items():
                self.data_preview.insert("end", f"📌 {key}: {value}\n")
            
            self._log(f"✅ 已載入 Markdown 檔案: {path}")
            self._log(f"   解析到 {len(data)} 個欄位")
            
        except ImportError:
            self.data_preview.delete("1.0", "end")
            self.data_preview.insert("1.0", "⚠️ 無法載入解析模組")
        except Exception as e:
            self._log(f"❌ 解析錯誤: {str(e)}")
    
    def _start_conversion(self) -> None:
        """開始轉換"""
        # 驗證輸入
        if not self.markdown_path.get():
            messagebox.showwarning("警告", "請選擇 Markdown 檔案")
            return
        if not self.template_path.get():
            messagebox.showwarning("警告", "請選擇 Word 模板")
            return
        if not self.output_path.get():
            messagebox.showwarning("警告", "請指定輸出路徑")
            return
        
        # 禁用按鈕
        self.btn_convert.configure(state="disabled")
        self.status_label.configure(text="轉換中...")
        self.progress_bar.set(0)
        
        # 在背景執行緒執行轉換
        thread = Thread(target=self._do_conversion, daemon=True)
        thread.start()
    
    def _do_conversion(self) -> None:
        """執行轉換（在背景執行緒）"""
        try:
            from md_word_renderer.parser.markdown_parser import MarkdownParser
            from md_word_renderer.cli.main import process_one

            # 解析
            self._update_progress(0.2, "解析 Markdown...")
            parser = MarkdownParser()
            data = parser.parse(self.markdown_path.get())

            # 驗證（可選）
            if self.config_manager.get("validate_before_convert", True):
                self._update_progress(0.4, "驗證資料...")
                # TODO: 加入驗證邏輯

            # 渲染（透過 factory 依樣板副檔名挑 renderer）
            fmt = detect_format(self.template_path.get())
            self._update_progress(0.6, f"渲染 {fmt.upper()} 文件...")
            process_one(
                input_path=self.markdown_path.get(),
                template_path=self.template_path.get(),
                output_path=self.output_path.get(),
                format_hint="auto",
                validate=bool(self.config_manager.get("validate_before_convert", True)),
                verbose=False,
            )

            # 完成
            self._update_progress(1.0, "完成!")
            self._log(f"✅ 轉換完成: {self.output_path.get()} (格式: {fmt})")

            # 開啟檔案（可選）
            if self.config_manager.get("open_after_convert", True):
                import os
                os.startfile(self.output_path.get())

            self.after(0, lambda: messagebox.showinfo("成功", "轉換完成!"))

        except ImportError as e:
            self._log(f"❌ 模組匯入錯誤: {str(e)}")
            self.after(0, lambda: messagebox.showerror("錯誤", f"模組匯入錯誤:\n{str(e)}"))
        except Exception as e:
            self._log(f"❌ 轉換失敗: {str(e)}")
            self.after(0, lambda: messagebox.showerror("錯誤", f"轉換失敗:\n{str(e)}"))
        finally:
            self.after(0, lambda: self.btn_convert.configure(state="normal"))
    
    def _update_progress(self, value: float, status: str) -> None:
        """更新進度"""
        self.after(0, lambda: self.progress_bar.set(value))
        self.after(0, lambda: self.status_label.configure(text=status))
        self._log(f"📊 {status}")
    
    def _log(self, message: str) -> None:
        """寫入日誌"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.after(0, lambda: self.log_text.insert("end", f"[{timestamp}] {message}\n"))
        self.after(0, lambda: self.log_text.see("end"))
    
    def _toggle_theme(self) -> None:
        """切換主題"""
        if self.theme_switch.get():
            ctk.set_appearance_mode("dark")
            self.config_manager.set("theme", "dark")
        else:
            ctk.set_appearance_mode("light")
            self.config_manager.set("theme", "light")
    
    def _open_batch_window(self) -> None:
        """開啟批次處理視窗"""
        from .batch_window import BatchWindow
        BatchWindow(self)
    
    def _open_settings(self) -> None:
        """開啟設定視窗"""
        from .settings_window import SettingsWindow
        SettingsWindow(self, self.config_manager)
    
    def _open_multi_template_window(self) -> None:
        """開啟多模板批次處理視窗"""
        from .multi_template_window import MultiTemplateWindow
        MultiTemplateWindow(self)
    
    def _show_help(self) -> None:
        """顯示說明"""
        help_text = """
MD-Word/Excel Template Renderer v2.2

使用步驟：
1. 選擇 Markdown 資料檔案
2. 選擇 Word 模板 (.docx) 或 Excel 樣板 (.xlsx)
3. 輸出格式會自動依樣板副檔名偵測
4. 指定輸出路徑
5. 點擊「開始轉換」

支援的 Markdown 格式：
- 編號. 欄位名稱 | 值（行內）
- 階層透過縮排表達

Word 模板語法（.docx）：
- {{ 變數名 }} - 插入變數
- {% if 條件 %} - 條件判斷
- {% for item in list %} - 迴圈

Excel 樣板語意（.xlsx）：
- 「基本資訊」sheet 自動填入純量欄位
- 每個清單欄位展開成獨立 sheet，每葉節點一列
- 圖片自動嵌入並等比縮放
- 樣板可放隱藏 LAYOUT sheet 覆寫預設

詳細說明請參考 doc/template_syntax_reference.md
        """
        messagebox.showinfo("使用說明", help_text.strip())
    
    def _open_template_preview(self) -> None:
        """開啟模板預覽視窗"""
        from .template_preview import TemplatePreviewWindow
        template = self.template_path.get()
        TemplatePreviewWindow(self, template)
    
    def _validate_data(self) -> None:
        """驗證資料"""
        md_path = self.markdown_path.get()
        if not md_path:
            messagebox.showwarning("警告", "請先選擇 Markdown 檔案")
            return
        
        try:
            from md_word_renderer.parser.markdown_parser import MarkdownParser
            from md_word_renderer.validator.schema_validator import SchemaValidator
            
            # 解析
            parser = MarkdownParser()
            data = parser.parse(md_path)
            
            # 驗證
            validator = SchemaValidator()
            is_valid, errors = validator.validate(data)
            
            if is_valid:
                # 顯示成功訊息和統計
                field_count = len(data)
                msg = f"✅ 資料驗證通過！\n\n"
                msg += f"📊 統計:\n"
                msg += f"  • 欄位數量: {field_count}\n"
                msg += f"  • 檔案: {Path(md_path).name}\n\n"
                msg += "資料格式正確，可以進行轉換。"
                messagebox.showinfo("驗證成功", msg)
                self._log(f"✅ 驗證通過: {field_count} 個欄位")
            else:
                # 顯示錯誤
                error_msg = "❌ 資料驗證失敗\n\n問題列表:\n"
                for i, err in enumerate(errors[:10], 1):  # 最多顯示 10 個
                    error_msg += f"{i}. {err}\n"
                if len(errors) > 10:
                    error_msg += f"\n... 還有 {len(errors) - 10} 個問題"
                messagebox.showerror("驗證失敗", error_msg)
                self._log(f"❌ 驗證失敗: {len(errors)} 個問題")
                
        except ImportError as e:
            messagebox.showerror("錯誤", f"無法載入驗證模組:\n{str(e)}")
        except Exception as e:
            error_handler = GUIErrorHandler(self)
            error_handler.handle_exception(e, context="資料驗證", log_callback=self._log)
    
    def _on_closing(self) -> None:
        """關閉視窗時的處理"""
        # 儲存視窗位置
        self.config_manager.set("window_width", self.winfo_width())
        self.config_manager.set("window_height", self.winfo_height())
        self.config_manager.set("window_x", self.winfo_x())
        self.config_manager.set("window_y", self.winfo_y())
        
        # 儲存設定
        self.config_manager.save()
        
        # 關閉
        self.destroy()
    
    def run(self) -> None:
        """執行應用程式"""
        self.mainloop()


def main():
    """主程式入口"""
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main()
