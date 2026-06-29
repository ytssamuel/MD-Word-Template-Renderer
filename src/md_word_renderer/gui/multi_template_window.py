"""
多模板批次處理視窗

使用多個模板渲染同一份 Markdown 資料
"""

import customtkinter as ctk
from pathlib import Path
from threading import Thread
from typing import List, Tuple, Optional
import tkinter as tk
from tkinter import filedialog, messagebox

from .config_manager import ConfigManager
from md_word_renderer.renderer.factory import build_renderer, detect_format


class MultiTemplateWindow(ctk.CTkToplevel):
    """
    多模板批次處理視窗
    
    1個 MD 檔 + n個模板 → n個 Word 檔
    """
    
    def __init__(self, parent: ctk.CTk):
        super().__init__(parent)
        
        # 設定管理器
        self.config_manager = ConfigManager()
        
        # 設定視窗
        self.title("多模板批次處理")
        self.geometry("850x650")
        self.minsize(650, 450)
        
        # 模板清單
        self.template_list: List[Tuple[str, str]] = []  # (path, status)
        
        # 路徑變數
        self.markdown_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        self.prefix_var = ctk.StringVar()
        self.suffix_var = ctk.StringVar()
        
        # 建立 UI
        self._create_widgets()
        
        # 聚焦此視窗
        self.focus()
        self.grab_set()
    
    def _create_widgets(self) -> None:
        """建立所有 UI 元件"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        # 說明標籤
        info_label = ctk.CTkLabel(
            self,
            text="📋 多模板批次處理：使用一份 Markdown 資料搭配多個 Word / Excel 模板，產生多個對應格式的文件",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        info_label.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 5))
        
        # 設定區
        self._create_settings_section()
        
        # 模板清單區
        self._create_template_list_section()
        
        # 控制區
        self._create_control_section()
    
    def _create_settings_section(self) -> None:
        """建立設定區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
        frame.grid_columnconfigure(1, weight=1)
        
        # Markdown 檔案
        ctk.CTkLabel(frame, text="📄 Markdown 資料:").grid(
            row=0, column=0, sticky="w", padx=10, pady=8
        )
        ctk.CTkEntry(frame, textvariable=self.markdown_path).grid(
            row=0, column=1, sticky="ew", padx=5, pady=8
        )
        ctk.CTkButton(
            frame, text="瀏覽...", width=80,
            command=self._browse_markdown
        ).grid(row=0, column=2, padx=10, pady=8)
        
        # 輸出目錄
        ctk.CTkLabel(frame, text="📁 輸出目錄:").grid(
            row=1, column=0, sticky="w", padx=10, pady=8
        )
        ctk.CTkEntry(frame, textvariable=self.output_dir).grid(
            row=1, column=1, sticky="ew", padx=5, pady=8
        )
        ctk.CTkButton(
            frame, text="瀏覽...", width=80,
            command=self._browse_output_dir
        ).grid(row=1, column=2, padx=10, pady=8)
        
        # 檔名選項（第二列）
        option_frame = ctk.CTkFrame(frame, fg_color="transparent")
        option_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)
        
        ctk.CTkLabel(option_frame, text="檔名前綴:").pack(side="left", padx=10)
        ctk.CTkEntry(option_frame, textvariable=self.prefix_var, width=100).pack(side="left", padx=5)
        
        ctk.CTkLabel(option_frame, text="檔名後綴:").pack(side="left", padx=10)
        ctk.CTkEntry(option_frame, textvariable=self.suffix_var, width=100).pack(side="left", padx=5)
        
        ctk.CTkLabel(
            option_frame,
            text="(輸出: 前綴+模板名+後綴.<樣板副檔名>)",
            text_color="gray"
        ).pack(side="left", padx=10)
    
    def _create_template_list_section(self) -> None:
        """建立模板清單區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)
        
        # 標題
        title_frame = ctk.CTkFrame(frame, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        ctk.CTkLabel(
            title_frame,
            text="📋 Word/Excel 模板清單",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left", padx=5)
        
        # 工具列
        toolbar = ctk.CTkFrame(frame)
        toolbar.grid(row=0, column=0, sticky="e", padx=5, pady=5)
        
        ctk.CTkButton(
            toolbar, text="➕ 新增模板",
            width=100,
            command=self._add_templates
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            toolbar, text="📁 新增資料夾",
            width=100,
            command=self._add_folder
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            toolbar, text="❌ 移除選取",
            width=90,
            command=self._remove_selected
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            toolbar, text="🗑️ 清空",
            width=70,
            command=self._clear_list
        ).pack(side="left", padx=5)
        
        # 模板數量
        self.template_count_label = ctk.CTkLabel(toolbar, text="共 0 個模板")
        self.template_count_label.pack(side="right", padx=10)
        
        # 模板清單（使用 Treeview）
        list_frame = ctk.CTkFrame(frame)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        from tkinter import ttk
        
        # 設定樣式
        style = ttk.Style()
        style.configure(
            "Treeview",
            background="#2b2b2b",
            foreground="white",
            fieldbackground="#2b2b2b"
        )
        style.map("Treeview", background=[("selected", "#1f538d")])
        
        self.tree = ttk.Treeview(
            list_frame,
            columns=("name", "path", "status"),
            show="headings",
            selectmode="extended"
        )
        self.tree.heading("name", text="模板名稱")
        self.tree.heading("path", text="檔案路徑")
        self.tree.heading("status", text="狀態")
        self.tree.column("name", width=150)
        self.tree.column("path", width=400)
        self.tree.column("status", width=100)
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # 捲軸
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
    
    def _create_control_section(self) -> None:
        """建立控制區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        # 進度條
        self.progress_bar = ctk.CTkProgressBar(frame)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.set(0)
        
        # 狀態
        self.status_label = ctk.CTkLabel(frame, text="就緒 - 請新增模板並選擇 Markdown 資料")
        self.status_label.pack(pady=5)
        
        # 按鈕區
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        self.btn_start = ctk.CTkButton(
            btn_frame,
            text="🚀 開始批次轉換",
            width=150,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._start_batch
        )
        self.btn_start.pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_frame,
            text="關閉",
            width=100,
            command=self.destroy
        ).pack(side="left", padx=10)
    
    def _browse_markdown(self) -> None:
        """瀏覽 Markdown 檔案"""
        path = filedialog.askopenfilename(
            title="選擇 Markdown 檔案",
            filetypes=[("Markdown 檔案", "*.md"), ("所有檔案", "*.*")]
        )
        if path:
            self.markdown_path.set(path)
    
    def _browse_output_dir(self) -> None:
        """瀏覽輸出目錄"""
        path = filedialog.askdirectory(title="選擇輸出目錄")
        if path:
            self.output_dir.set(path)
    
    def _add_templates(self) -> None:
        """新增模板檔案"""
        paths = filedialog.askopenfilenames(
            title="選擇 Word/Excel 模板",
            filetypes=[
                ("Word 文件", "*.docx"),
                ("Excel 樣板", "*.xlsx"),
                ("所有檔案", "*.*"),
            ],
        )
        for path in paths:
            if path not in [item[0] for item in self.template_list]:
                name = Path(path).stem
                self.template_list.append((path, "待處理"))
                self.tree.insert("", "end", values=(name, path, "待處理"))
        self._update_count()

    def _add_folder(self) -> None:
        """新增資料夾中的所有模板（同時收集 .docx 與 .xlsx）"""
        folder = filedialog.askdirectory(title="選擇模板資料夾")
        if folder:
            folder_path = Path(folder)
            for ext in ("*.docx", "*.xlsx"):
                for path in folder_path.glob(ext):
                    path_str = str(path)
                    if path_str not in [item[0] for item in self.template_list]:
                        name = path.stem
                        self.template_list.append((path_str, "待處理"))
                        self.tree.insert("", "end", values=(name, path_str, "待處理"))
            self._update_count()
    
    def _remove_selected(self) -> None:
        """移除選取的項目"""
        selected = self.tree.selection()
        for item in selected:
            values = self.tree.item(item, "values")
            self.template_list = [t for t in self.template_list if t[0] != values[1]]
            self.tree.delete(item)
        self._update_count()
    
    def _clear_list(self) -> None:
        """清空清單"""
        if messagebox.askyesno("確認", "確定要清空模板清單嗎？"):
            self.template_list.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            self._update_count()
    
    def _update_count(self) -> None:
        """更新模板數量顯示"""
        self.template_count_label.configure(text=f"共 {len(self.template_list)} 個模板")
    
    def _start_batch(self) -> None:
        """開始批次處理"""
        # 驗證
        if not self.markdown_path.get():
            messagebox.showwarning("警告", "請先選擇 Markdown 資料檔案")
            return
        if not self.template_list:
            messagebox.showwarning("警告", "請先新增要使用的模板")
            return
        if not self.output_dir.get():
            messagebox.showwarning("警告", "請選擇輸出目錄")
            return
        
        # 禁用按鈕
        self.btn_start.configure(state="disabled")
        self.status_label.configure(text="處理中...")
        
        # 重設所有狀態
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            self.tree.item(item, values=(values[0], values[1], "待處理"))
        
        # 背景執行
        thread = Thread(target=self._do_batch, daemon=True)
        thread.start()
    
    def _do_batch(self) -> None:
        """執行批次處理"""
        try:
            from md_word_renderer.parser.markdown_parser import MarkdownParser

            # 解析 Markdown（只需一次）
            self.after(0, lambda: self.status_label.configure(text="解析 Markdown..."))
            parser = MarkdownParser()
            data = parser.parse(self.markdown_path.get())

            output_dir = Path(self.output_dir.get())
            output_dir.mkdir(parents=True, exist_ok=True)

            prefix = self.prefix_var.get()
            suffix = self.suffix_var.get()

            total = len(self.template_list)
            success = 0
            failed = 0

            for i, (template_path, _) in enumerate(self.template_list):
                try:
                    # 從樣板副檔名推格式
                    fmt = detect_format(template_path)
                    output_ext = f".{fmt}"

                    # 更新狀態
                    self._update_item_status(template_path, "處理中...")
                    self.after(0, lambda p=(i+1)/total: self.progress_bar.set(p))
                    self.after(0, lambda s=f"處理 {i+1}/{total}: {Path(template_path).stem}":
                               self.status_label.configure(text=s))

                    # 組合輸出檔名
                    template_name = Path(template_path).stem
                    output_name = f"{prefix}{template_name}{suffix}{output_ext}"
                    output_path = output_dir / output_name

                    # 透過 factory 挑 renderer
                    renderer = build_renderer(template_path=template_path, format_hint=fmt)
                    renderer.render_to_file(data, template_path, str(output_path))

                    # 成功
                    self._update_item_status(template_path, "✅ 完成")
                    success += 1

                except Exception as e:
                    error_msg = str(e)[:25] + "..." if len(str(e)) > 25 else str(e)
                    self._update_item_status(template_path, f"❌ {error_msg}")
                    failed += 1

                    if not self.config_manager.get("continue_on_error", True):
                        break
            
            # 完成
            self.after(0, lambda: self.progress_bar.set(1))
            self.after(0, lambda: self.status_label.configure(
                text=f"完成! 成功: {success}, 失敗: {failed}"
            ))
            self.after(0, lambda: messagebox.showinfo(
                "多模板批次處理完成",
                f"使用資料: {Path(self.markdown_path.get()).name}\n"
                f"總共處理: {total} 個模板\n"
                f"成功: {success}\n"
                f"失敗: {failed}\n\n"
                f"輸出目錄: {output_dir}"
            ))
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("錯誤", f"批次處理失敗:\n{str(e)}"))
            self.after(0, lambda: self.status_label.configure(text=f"錯誤: {str(e)}"))
        finally:
            self.after(0, lambda: self.btn_start.configure(state="normal"))
    
    def _update_item_status(self, path: str, status: str) -> None:
        """更新項目狀態"""
        def update():
            for item in self.tree.get_children():
                values = self.tree.item(item, "values")
                if values[1] == path:
                    self.tree.item(item, values=(values[0], values[1], status))
                    break
        self.after(0, update)
