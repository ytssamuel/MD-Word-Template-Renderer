"""
模板預覽器

解析 Word 模板並顯示變數列表
"""

import re
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional
from dataclasses import dataclass
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import customtkinter as ctk
from tkinter import messagebox


@dataclass
class TemplateVariable:
    """模板變數資訊"""
    name: str               # 變數名稱
    var_type: str          # 類型: simple, condition, loop, filter
    context: str           # 出現的上下文
    count: int = 1         # 出現次數
    filters: List[str] = None  # 使用的過濾器
    
    def __post_init__(self):
        if self.filters is None:
            self.filters = []


class TemplateAnalyzer:
    """
    模板分析器
    
    解析 Word 模板中的 Jinja2 變數
    """
    
    # 正規表達式模式
    PATTERNS = {
        # 簡單變數: {{ variable }} 或 {{ variable|filter }}
        'simple': r'\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*(?:\.[a-zA-Z_][a-zA-Z0-9_]*)*)\s*(?:\|[^}]+)?\s*\}\}',
        # 條件語句: {% if condition %}
        'condition': r'\{%\s*if\s+([^%]+?)\s*%\}',
        # 迴圈: {% for item in list %}
        'loop': r'\{%\s*for\s+(\w+)\s+in\s+([a-zA-Z_][a-zA-Z0-9_]*(?:\.[a-zA-Z_][a-zA-Z0-9_]*)*)\s*%\}',
        # 帶過濾器的變數
        'filter': r'\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*(?:\.[a-zA-Z_][a-zA-Z0-9_]*)*)\s*\|([^}]+)\}\}',
    }
    
    def __init__(self):
        self.variables: Dict[str, TemplateVariable] = {}
        self.loop_vars: Set[str] = set()  # 迴圈變數，需要排除
        self.errors: List[str] = []
    
    def analyze_file(self, template_path: str) -> Dict[str, TemplateVariable]:
        """
        分析模板檔案
        
        Args:
            template_path: 模板檔案路徑
            
        Returns:
            Dict[str, TemplateVariable]: 變數字典
        """
        self.variables.clear()
        self.loop_vars.clear()
        self.errors.clear()
        
        path = Path(template_path)
        if not path.exists():
            self.errors.append(f"模板檔案不存在: {template_path}")
            return {}
        
        try:
            # 提取模板中的文字內容
            content = self._extract_text_from_docx(template_path)
            
            # 分析內容
            self._analyze_content(content)
            
            return self.variables
            
        except Exception as e:
            self.errors.append(f"分析模板時發生錯誤: {str(e)}")
            return {}
    
    def _extract_text_from_docx(self, docx_path: str) -> str:
        """
        從 docx 檔案提取文字內容
        
        Args:
            docx_path: docx 檔案路徑
            
        Returns:
            str: 提取的文字內容
        """
        content_parts = []
        
        with ZipFile(docx_path, 'r') as zip_file:
            # 讀取主要文件內容
            for xml_file in ['word/document.xml', 'word/header1.xml', 
                            'word/header2.xml', 'word/footer1.xml',
                            'word/footer2.xml']:
                try:
                    with zip_file.open(xml_file) as f:
                        tree = ET.parse(f)
                        root = tree.getroot()
                        
                        # 提取所有文字節點
                        for elem in root.iter():
                            if elem.text:
                                content_parts.append(elem.text)
                            if elem.tail:
                                content_parts.append(elem.tail)
                except KeyError:
                    continue  # 某些檔案可能不存在
        
        return ' '.join(content_parts)
    
    def _analyze_content(self, content: str) -> None:
        """
        分析內容中的變數
        
        Args:
            content: 模板內容
        """
        # 先處理迴圈，收集迴圈變數
        for match in re.finditer(self.PATTERNS['loop'], content):
            loop_var = match.group(1)
            list_var = match.group(2)
            self.loop_vars.add(loop_var)
            
            # 記錄迴圈列表變數
            self._add_variable(list_var, 'loop', "{{% for " + loop_var + " in " + list_var + " %}}")
        
        # 處理條件語句
        for match in re.finditer(self.PATTERNS['condition'], content):
            condition = match.group(1).strip()
            # 從條件中提取變數名
            var_matches = re.findall(r'\b([a-zA-Z_][a-zA-Z0-9_]*(?:\.[a-zA-Z_][a-zA-Z0-9_]*)*)\b', condition)
            for var in var_matches:
                if var not in self.loop_vars and var not in ['and', 'or', 'not', 'in', 'is', 'true', 'false', 'none']:
                    self._add_variable(var, 'condition', "{{% if " + condition + " %}}")
        
        # 處理帶過濾器的變數
        for match in re.finditer(self.PATTERNS['filter'], content):
            var_name = match.group(1)
            filters = match.group(2).strip()
            
            if var_name not in self.loop_vars:
                var = self._add_variable(var_name, 'filter', f"{{{{ {var_name}|{filters} }}}}")
                if var:
                    var.filters = [f.strip().split('(')[0] for f in filters.split('|')]
        
        # 處理簡單變數
        for match in re.finditer(self.PATTERNS['simple'], content):
            var_name = match.group(1)
            if var_name not in self.loop_vars:
                self._add_variable(var_name, 'simple', f"{{{{ {var_name} }}}}")
    
    def _add_variable(self, name: str, var_type: str, context: str) -> Optional[TemplateVariable]:
        """
        新增變數
        
        Args:
            name: 變數名稱
            var_type: 變數類型
            context: 上下文
            
        Returns:
            TemplateVariable: 變數物件
        """
        if name in self.variables:
            self.variables[name].count += 1
            return self.variables[name]
        else:
            var = TemplateVariable(name=name, var_type=var_type, context=context)
            self.variables[name] = var
            return var
    
    def get_variable_summary(self) -> Dict[str, List[str]]:
        """
        取得變數摘要
        
        Returns:
            Dict[str, List[str]]: 按類型分類的變數列表
        """
        summary = {
            'simple': [],
            'condition': [],
            'loop': [],
            'filter': []
        }
        
        for var in self.variables.values():
            summary[var.var_type].append(var.name)
        
        return summary


class TemplatePreviewWindow(ctk.CTkToplevel):
    """
    模板預覽視窗
    
    顯示模板中的變數列表
    """
    
    def __init__(self, parent: ctk.CTk, template_path: str = ""):
        super().__init__(parent)
        
        self.title("模板變數預覽")
        self.geometry("600x500")
        self.minsize(400, 300)
        
        self.analyzer = TemplateAnalyzer()
        self.template_path = template_path
        
        self._create_widgets()
        
        if template_path:
            self._analyze_template()
        
        self.focus()
        self.grab_set()
    
    def _create_widgets(self) -> None:
        """建立 UI"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 檔案選擇區
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        file_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(file_frame, text="模板檔案:").grid(
            row=0, column=0, padx=10, pady=10
        )
        
        self.path_var = ctk.StringVar(value=self.template_path)
        ctk.CTkEntry(file_frame, textvariable=self.path_var).grid(
            row=0, column=1, sticky="ew", padx=5, pady=10
        )
        
        ctk.CTkButton(
            file_frame, text="瀏覽...", width=80,
            command=self._browse_template
        ).grid(row=0, column=2, padx=5, pady=10)
        
        ctk.CTkButton(
            file_frame, text="分析", width=80,
            command=self._analyze_template
        ).grid(row=0, column=3, padx=10, pady=10)
        
        # 變數列表區
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        
        # 全部變數
        tab_all = self.tabview.add("📋 全部變數")
        tab_all.grid_columnconfigure(0, weight=1)
        tab_all.grid_rowconfigure(0, weight=1)
        
        self.all_vars_text = ctk.CTkTextbox(tab_all, wrap="word")
        self.all_vars_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # 依類型分類
        tab_by_type = self.tabview.add("🏷️ 依類型")
        tab_by_type.grid_columnconfigure(0, weight=1)
        tab_by_type.grid_rowconfigure(0, weight=1)
        
        self.type_vars_text = ctk.CTkTextbox(tab_by_type, wrap="word")
        self.type_vars_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # 範例資料
        tab_sample = self.tabview.add("📝 範例資料")
        tab_sample.grid_columnconfigure(0, weight=1)
        tab_sample.grid_rowconfigure(0, weight=1)
        
        self.sample_text = ctk.CTkTextbox(tab_sample, wrap="word")
        self.sample_text.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        # 狀態列
        self.status_label = ctk.CTkLabel(self, text="")
        self.status_label.grid(row=2, column=0, sticky="w", padx=10, pady=5)
        
        # 按鈕區
        btn_frame = ctk.CTkFrame(self)
        btn_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        ctk.CTkButton(
            btn_frame, text="複製變數列表",
            command=self._copy_variables
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_frame, text="匯出範例 Markdown",
            command=self._export_sample
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            btn_frame, text="關閉",
            command=self.destroy
        ).pack(side="right", padx=10)
    
    def _browse_template(self) -> None:
        """瀏覽模板檔案"""
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="選擇 Word 模板",
            filetypes=[("Word 文件", "*.docx"), ("所有檔案", "*.*")]
        )
        if path:
            self.path_var.set(path)
            self.template_path = path
    
    def _analyze_template(self) -> None:
        """分析模板"""
        path = self.path_var.get()
        if not path:
            messagebox.showwarning("警告", "請先選擇模板檔案")
            return
        
        # 分析
        variables = self.analyzer.analyze_file(path)
        
        if self.analyzer.errors:
            messagebox.showerror("分析錯誤", "\n".join(self.analyzer.errors))
            return
        
        # 更新狀態
        self.status_label.configure(text=f"找到 {len(variables)} 個變數")
        
        # 顯示全部變數
        self._display_all_variables(variables)
        
        # 顯示分類
        self._display_by_type(variables)
        
        # 生成範例
        self._generate_sample(variables)
    
    def _display_all_variables(self, variables: Dict[str, TemplateVariable]) -> None:
        """顯示全部變數"""
        self.all_vars_text.delete("1.0", "end")
        
        if not variables:
            self.all_vars_text.insert("1.0", "未找到任何變數")
            return
        
        self.all_vars_text.insert("1.0", "模板變數列表\n")
        self.all_vars_text.insert("end", "=" * 50 + "\n\n")
        
        for name, var in sorted(variables.items()):
            type_icon = {
                'simple': '📌',
                'condition': '❓',
                'loop': '🔄',
                'filter': '🔧'
            }.get(var.var_type, '📌')
            
            self.all_vars_text.insert("end", f"{type_icon} {name}\n")
            self.all_vars_text.insert("end", f"   類型: {var.var_type}\n")
            self.all_vars_text.insert("end", f"   使用: {var.context}\n")
            if var.filters:
                self.all_vars_text.insert("end", f"   過濾器: {', '.join(var.filters)}\n")
            self.all_vars_text.insert("end", f"   出現次數: {var.count}\n\n")
    
    def _display_by_type(self, variables: Dict[str, TemplateVariable]) -> None:
        """依類型顯示"""
        self.type_vars_text.delete("1.0", "end")
        
        summary = self.analyzer.get_variable_summary()
        
        type_names = {
            'simple': '📌 簡單變數',
            'condition': '❓ 條件變數',
            'loop': '🔄 迴圈變數',
            'filter': '🔧 過濾器變數'
        }
        
        for var_type, names in summary.items():
            if names:
                self.type_vars_text.insert("end", f"\n{type_names[var_type]}\n")
                self.type_vars_text.insert("end", "-" * 30 + "\n")
                for name in sorted(names):
                    self.type_vars_text.insert("end", f"  • {name}\n")
    
    def _generate_sample(self, variables: Dict[str, TemplateVariable]) -> None:
        """生成範例 Markdown"""
        self.sample_text.delete("1.0", "end")
        
        self.sample_text.insert("1.0", "# 範例資料\n\n")
        
        for i, (name, var) in enumerate(sorted(variables.items()), 1):
            # 生成範例值
            if var.var_type == 'loop':
                sample_value = "(列表資料)"
            elif '.' in name:
                sample_value = f"子欄位值"
            else:
                sample_value = f"範例{name}值"
            
            self.sample_text.insert("end", f"{i}. {name} | {sample_value}\n")
    
    def _copy_variables(self) -> None:
        """複製變數列表"""
        content = self.all_vars_text.get("1.0", "end")
        self.clipboard_clear()
        self.clipboard_append(content)
        messagebox.showinfo("已複製", "變數列表已複製到剪貼簿")
    
    def _export_sample(self) -> None:
        """匯出範例 Markdown"""
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            title="儲存範例 Markdown",
            filetypes=[("Markdown 檔案", "*.md")],
            defaultextension=".md"
        )
        if path:
            content = self.sample_text.get("1.0", "end")
            with open(path, 'w', encoding='utf-8') as f:
                f.write(content)
            messagebox.showinfo("成功", f"範例已儲存至:\n{path}")
