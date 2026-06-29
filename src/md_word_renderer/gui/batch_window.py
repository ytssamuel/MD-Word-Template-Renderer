"""
批次處理視窗

處理多個 Markdown 檔案的批量轉換
"""

import customtkinter as ctk
from pathlib import Path
from threading import Thread
from typing import List, Tuple, Optional
import tkinter as tk
from tkinter import filedialog, messagebox

from .config_manager import ConfigManager
from md_word_renderer.renderer.factory import build_renderer, detect_format


class BatchWindow(ctk.CTkToplevel):
    """
    批次處理視窗
    
    提供多檔案批量轉換功能
    """
    
    def __init__(self, parent: ctk.CTk):
        super().__init__(parent)
        
        # 設定管理器
        self.config_manager = ConfigManager()
        
        # 設定視窗
        self.title("批次處理")
        self.geometry("800x600")
        self.minsize(600, 400)
        
        # 檔案清單
        self.file_list: List[Tuple[str, str]] = []  # (path, status)
        
        # 模板路徑
        self.template_path = ctk.StringVar()
        self.output_dir = ctk.StringVar()
        
        # 建立 UI
        self._create_widgets()
        
        # 聚焦此視窗
        self.focus()
        self.grab_set()
    
    def _create_widgets(self) -> None:
        """建立所有 UI 元件"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 設定區
        self._create_settings_section()
        
        # 檔案清單區
        self._create_file_list_section()
        
        # 控制區
        self._create_control_section()
    
    def _create_settings_section(self) -> None:
        """建立設定區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        frame.grid_columnconfigure(1, weight=1)
        
        # 模板選擇
        ctk.CTkLabel(frame, text="📋 Word/Excel 模板:").grid(
            row=0, column=0, sticky="w", padx=10, pady=5
        )
        ctk.CTkEntry(frame, textvariable=self.template_path).grid(
            row=0, column=1, sticky="ew", padx=5, pady=5
        )
        ctk.CTkButton(
            frame, text="瀏覽...", width=80,
            command=self._browse_template
        ).grid(row=0, column=2, padx=10, pady=5)

        self.batch_format_label = ctk.CTkLabel(
            frame, text="🔎 輸出格式: (尚未選擇樣板)", anchor="w"
        )
        self.batch_format_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=10)
        self.template_path.trace_add("write", lambda *_: self._refresh_template_label())
        # 輸出目錄再放第二列下方，避免擠壓
        ctk.CTkLabel(frame, text="📁 輸出目錄:").grid(
            row=2, column=0, sticky="w", padx=10, pady=5
        )
        ctk.CTkEntry(frame, textvariable=self.output_dir).grid(
            row=2, column=1, sticky="ew", padx=5, pady=5
        )
        ctk.CTkButton(
            frame, text="瀏覽...", width=80,
            command=self._browse_output_dir
        ).grid(row=2, column=2, padx=10, pady=5)

    def _create_file_list_section(self) -> None:
        """建立檔案清單區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(1, weight=1)
        
        # 工具列
        toolbar = ctk.CTkFrame(frame)
        toolbar.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        
        ctk.CTkButton(
            toolbar, text="➕ 新增檔案",
            command=self._add_files
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            toolbar, text="📁 新增資料夾",
            command=self._add_folder
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            toolbar, text="❌ 移除選取",
            command=self._remove_selected
        ).pack(side="left", padx=5)
        
        ctk.CTkButton(
            toolbar, text="🗑️ 清空清單",
            command=self._clear_list
        ).pack(side="left", padx=5)
        
        # 檔案數量
        self.file_count_label = ctk.CTkLabel(toolbar, text="共 0 個檔案")
        self.file_count_label.pack(side="right", padx=10)
        
        # 檔案清單（使用 Textbox 模擬，因為 CTk 沒有 Listbox）
        list_frame = ctk.CTkFrame(frame)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        # 使用 Treeview
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
            columns=("path", "status"),
            show="headings",
            selectmode="extended"
        )
        self.tree.heading("path", text="檔案路徑")
        self.tree.heading("status", text="狀態")
        self.tree.column("path", width=500)
        self.tree.column("status", width=100)
        self.tree.grid(row=0, column=0, sticky="nsew")
        
        # 捲軸
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)
    
    def _create_control_section(self) -> None:
        """建立控制區"""
        frame = ctk.CTkFrame(self)
        frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)
        
        # 進度條
        self.progress_bar = ctk.CTkProgressBar(frame)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.set(0)
        
        # 狀態
        self.status_label = ctk.CTkLabel(frame, text="就緒")
        self.status_label.pack(pady=5)
        
        # 按鈕區
        btn_frame = ctk.CTkFrame(frame)
        btn_frame.pack(pady=10)
        
        self.btn_start = ctk.CTkButton(
            btn_frame,
            text="🚀 開始批次轉換",
            width=150,
            command=self._start_batch
        )
        self.btn_start.pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_frame,
            text="關閉",
            width=100,
            command=self.destroy
        ).pack(side="left", padx=10)
    
    def _browse_template(self) -> None:
        """瀏覽模板檔案"""
        path = filedialog.askopenfilename(
            title="選擇 Word/Excel 模板",
            filetypes=[
                ("Word 文件", "*.docx"),
                ("Excel 樣板", "*.xlsx"),
                ("所有檔案", "*.*"),
            ],
        )
        if path:
            self.template_path.set(path)

    def _refresh_template_label(self) -> None:
        """樣板路徑變更時更新格式標籤"""
        template = self.template_path.get()
        if not template:
            if hasattr(self, "batch_format_label"):
                self.batch_format_label.configure(text="🔎 輸出格式: (尚未選擇樣板)")
            return
        try:
            fmt = detect_format(template)
            if hasattr(self, "batch_format_label"):
                self.batch_format_label.configure(text=f"🔎 輸出格式: {fmt}")
        except ValueError:
            if hasattr(self, "batch_format_label"):
                self.batch_format_label.configure(text="🔎 輸出格式: (不支援的副檔名)")
    
    def _browse_output_dir(self) -> None:
        """瀏覽輸出目錄"""
        path = filedialog.askdirectory(title="選擇輸出目錄")
        if path:
            self.output_dir.set(path)
    
    def _add_files(self) -> None:
        """新增檔案"""
        paths = filedialog.askopenfilenames(
            title="選擇 Markdown 檔案",
            filetypes=[("Markdown 檔案", "*.md"), ("所有檔案", "*.*")]
        )
        for path in paths:
            if path not in [item[0] for item in self.file_list]:
                self.file_list.append((path, "待處理"))
                self.tree.insert("", "end", values=(path, "待處理"))
        self._update_count()
    
    def _add_folder(self) -> None:
        """新增資料夾中的所有 Markdown 檔案"""
        folder = filedialog.askdirectory(title="選擇資料夾")
        if folder:
            pattern = self.config_manager.get("batch_file_pattern", "*.md")
            folder_path = Path(folder)
            for path in folder_path.glob(pattern):
                path_str = str(path)
                if path_str not in [item[0] for item in self.file_list]:
                    self.file_list.append((path_str, "待處理"))
                    self.tree.insert("", "end", values=(path_str, "待處理"))
            self._update_count()
    
    def _remove_selected(self) -> None:
        """移除選取的項目"""
        selected = self.tree.selection()
        for item in selected:
            values = self.tree.item(item, "values")
            self.file_list = [f for f in self.file_list if f[0] != values[0]]
            self.tree.delete(item)
        self._update_count()
    
    def _clear_list(self) -> None:
        """清空清單"""
        if messagebox.askyesno("確認", "確定要清空檔案清單嗎？"):
            self.file_list.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)
            self._update_count()
    
    def _update_count(self) -> None:
        """更新檔案數量顯示"""
        self.file_count_label.configure(text=f"共 {len(self.file_list)} 個檔案")
    
    def _start_batch(self) -> None:
        """開始批次處理"""
        # 驗證
        if not self.file_list:
            messagebox.showwarning("警告", "請先新增要處理的檔案")
            return
        if not self.template_path.get():
            messagebox.showwarning("警告", "請選擇 Word/Excel 模板")
            return
        if not self.output_dir.get():
            messagebox.showwarning("警告", "請選擇輸出目錄")
            return

        # 禁用按鈕
        self.btn_start.configure(state="disabled")

        # 背景執行
        thread = Thread(target=self._do_batch, daemon=True)
        thread.start()

    def _do_batch(self) -> None:
        """執行批次處理"""
        try:
            from md_word_renderer.parser.markdown_parser import MarkdownParser
            from md_word_renderer.renderer.factory import detect_format

            parser = MarkdownParser()
            template = self.template_path.get()
            fmt = detect_format(template)
            output_ext = f".{fmt}"
            output_dir = Path(self.output_dir.get())

            total = len(self.file_list)
            success = 0
            failed = 0

            for i, (md_path, _) in enumerate(self.file_list):
                try:
                    # 更新狀態
                    self._update_item_status(md_path, "處理中...")
                    self.after(0, lambda p=(i+1)/total: self.progress_bar.set(p))
                    self.after(0, lambda s=f"處理 {i+1}/{total}":
                               self.status_label.configure(text=s))

                    # 解析
                    data = parser.parse(md_path)

                    # 透過 factory 依樣板挑 renderer
                    renderer = build_renderer(template_path=template, format_hint=fmt)
                    output_name = Path(md_path).stem + output_ext
                    output_path = output_dir / output_name
                    renderer.render_to_file(data, template, str(output_path))

                    # 成功
                    self._update_item_status(md_path, "✅ 完成")
                    success += 1

                except Exception as e:
                    self._update_item_status(md_path, f"❌ {str(e)[:20]}")
                    failed += 1

                    if not self.config_manager.get("continue_on_error", True):
                        break
            
            # 完成
            self.after(0, lambda: self.progress_bar.set(1))
            self.after(0, lambda: self.status_label.configure(
                text=f"完成! 成功: {success}, 失敗: {failed}"
            ))
            self.after(0, lambda: messagebox.showinfo(
                "批次處理完成",
                f"總共處理: {total} 個檔案\n成功: {success}\n失敗: {failed}"
            ))
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("錯誤", str(e)))
        finally:
            self.after(0, lambda: self.btn_start.configure(state="normal"))
    
    def _update_item_status(self, path: str, status: str) -> None:
        """更新項目狀態"""
        def update():
            for item in self.tree.get_children():
                if self.tree.item(item, "values")[0] == path:
                    self.tree.item(item, values=(path, status))
                    break
        self.after(0, update)
