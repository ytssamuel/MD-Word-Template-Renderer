# Markdown to Word 模板渲染系統 - 實作計劃

## 專案資訊

**專案名稱：** MD-Word Template Renderer  
**版本：** v1.0  
**計劃日期：** 2025-12-12  
**預計開發時程：** Phase 1 (2-3週)

---

## 一、專案目標

建立一個 Python 工具，將結構化的 Markdown 資料渲染至 Word 模板，實現資料驅動的文件生成系統。

**核心流程：**
```
Markdown 資料 + Word 模板 = 渲染完成的 Word 文件
```

**未來擴充：**
```
Word 文件 + Word 模板 = Markdown 資料（逆向解析）
```

---

## 二、技術架構

### 2.1 技術棧選擇

| 元件 | 技術選型 | 理由 |
|------|----------|------|
| 程式語言 | Python 3.8+ | 豐富的文件處理生態系 |
| Word 模板引擎 | **docxtpl** | 基於 Jinja2，支援複雜模板邏輯，節省開發時間 |
| Word 文件處理 | python-docx | docxtpl 底層依賴 |
| Markdown 解析 | 自實作 | 格式特殊，需客製化解析邏輯 |
| 命令列介面 | argparse | Python 標準庫 |

### 2.2 系統架構圖

```
┌─────────────────────────────────────────────────────────┐
│                    使用者介面層                          │
│  ┌─────────────────┐      ┌──────────────────┐         │
│  │  CLI Interface  │      │  Python API      │         │
│  └────────┬────────┘      └────────┬─────────┘         │
└───────────┼──────────────────────────┼──────────────────┘
            │                          │
┌───────────┼──────────────────────────┼──────────────────┐
│           │      核心處理層           │                  │
│  ┌────────▼─────────┐      ┌────────▼─────────┐        │
│  │ Markdown Parser  │      │  Template Engine │        │
│  │                  │      │    (docxtpl)     │        │
│  │ - 階層解析       │      │ - 變數替換       │        │
│  │ - JSON 轉換      │      │ - 條件渲染       │        │
│  │ - 特殊字元處理   │      │ - 迴圈展開       │        │
│  └────────┬─────────┘      └────────┬─────────┘        │
└───────────┼──────────────────────────┼──────────────────┘
            │                          │
┌───────────┼──────────────────────────┼──────────────────┐
│           │       資料儲存層          │                  │
│  ┌────────▼─────────┐      ┌────────▼─────────┐        │
│  │  .md 檔案        │      │  .docx 模板      │        │
│  └──────────────────┘      └────────┬─────────┘        │
│                             ┌────────▼─────────┐        │
│                             │  .docx 輸出      │        │
│                             └──────────────────┘        │
└─────────────────────────────────────────────────────────┘
```

---

## 三、模板語法規範（基於 docxtpl + Jinja2）

### 3.1 語法設計決策

根據討論結果確定的設計：

| 項目 | 決策 | 說明 |
|------|------|------|
| 變數語法 | `{{變數}}` | 與 Vue/Handlebars 一致 |
| 點記法存取 | **支援** | `{{異動內容-測試案例.1.value}}` |
| 存取方式 | **兩者皆支援** | 欄位名稱 `{{系統名稱}}` 或編號 `{{#1}}` |
| 空值處理 | 顯示空白 | 不顯示預設文字 |
| 迴圈變數 | 標準 Jinja2 | `number`, `value`, `children`, `index` |
| 縮排控制 | **模板控制** | 在 Word 模板中手動設定格式 |
| 特殊字元 | 反斜線轉義 | `\|` → `|`，雙引號機制 `""` → `"` |
| 錯誤處理 | 顯示錯誤訊息 | `[ERROR: 變數不存在]` |

### 3.2 完整語法表

#### 3.2.1 基本變數替換

| 語法 | 說明 | 範例 |
|------|------|------|
| `{{變數名稱}}` | 用欄位名稱存取 | `{{系統名稱}}` |
| `{{#編號}}` | 用編號存取值 | `{{#1}}` |
| `{{#編號.key}}` | 用編號存取欄位名 | `{{#1.key}}` |
| `{{#編號.value}}` | 用編號存取值（明確） | `{{#1.value}}` |

**Word 模板範例：**
```
系統名稱：{{系統名稱}}
變更單號：{{變更單號}}
```

#### 3.2.2 點記法存取

| 語法 | 說明 |
|------|------|
| `{{變數.子項.值}}` | 多層級存取 |
| `{{變數.1.value}}` | 存取第一個子項的值 |
| `{{變數.1.children.0.value}}` | 存取孫項目 |

**Word 模板範例：**
```
第一個測試案例：{{異動內容-測試案例.1.value}}
```

#### 3.2.3 條件渲染（Jinja2 語法）

```django
{% if 變數名稱 %}
有值時顯示的內容：{{變數名稱}}
{% endif %}
```

**Word 模板範例：**
```
{% if 中介軟體 %}
中介軟體：{{中介軟體}}
{% endif %}

{% if 其他系統 %}
相關系統：{{其他系統}}
{% else %}
無相關系統
{% endif %}
```

#### 3.2.4 迴圈渲染（Jinja2 語法）

**單層迴圈：**
```django
{% for item in 異動內容-測試案例 %}
{{loop.index}}. {{item.value}}
{% endfor %}
```

**巢狀迴圈：**
```django
{% for item in 異動內容-測試案例 %}
{{item.number}}. {{item.value}}
  {% for child in item.children %}
  {{child.number}}. {{child.value}}
  {% endfor %}
{% endfor %}
```

**可用變數：**
- `item.number` - 項目編號
- `item.value` - 項目內容
- `item.children` - 子項目列表
- `loop.index` - 當前索引（從 1 開始）
- `loop.index0` - 當前索引（從 0 開始）
- `loop.first` - 是否第一項
- `loop.last` - 是否最後一項

#### 3.2.5 特殊字元轉義

**Markdown 檔案中：**
```
1. 特殊欄位 | 這是一個\|包含管線符號的值
2. 引號欄位 | 他說："這是引號"
3. 換行欄位 | 第一行\n第二行
```

**轉義規則：**
- `\|` → `|`
- `\n` → 換行
- `\\` → `\`
- `\"` → `"`
- `""` 內的 `"` → `"`（Python 字串機制）

---

## 四、資料結構設計

### 4.1 Markdown 解析輸出格式

**輸入 Markdown：**
```markdown
1. 系統名稱 | Super E-Billing
16. 異動內容-測試案例 | 
    1. 調整供應商登入之密碼功能
       1. TC001：新增供應商使用者
       2. TC002：超過最長使用設定
    2. 調整中華電信解析錯誤文字格式
```

**輸出 JSON：**
```json
{
  "系統名稱": "Super E-Billing",
  "#1": {
    "key": "系統名稱",
    "value": "Super E-Billing"
  },
  "異動內容-測試案例": [
    {
      "number": "1",
      "value": "調整供應商登入之密碼功能",
      "children": [
        {"number": "1", "value": "TC001：新增供應商使用者"},
        {"number": "2", "value": "TC002：超過最長使用設定"}
      ]
    },
    {
      "number": "2",
      "value": "調整中華電信解析錯誤文字格式",
      "children": []
    }
  ],
  "#16": {
    "key": "異動內容-測試案例",
    "value": "",
    "children": [...]
  }
}
```

**資料結構說明：**
1. **扁平化存取**：欄位名稱直接作為 key
2. **編號存取**：`#數字` 格式作為 key
3. **階層資料**：使用陣列 + children 表示
4. **雙重索引**：同時支援名稱和編號存取

### 4.2 階層解析規則

**縮排偵測邏輯（參考 Python PEP 8）：**

1. **優先偵測縮排類型**：
   - 掃描前 100 行
   - 統計空格和 Tab 的使用頻率
   - 選擇多數者作為縮排標準

2. **縮排層級計算**：
   ```python
   # 若使用空格（假設 4 空格為一層）
   level = leading_spaces // 4
   
   # 若使用 Tab
   level = leading_tabs
   
   # 混合使用（錯誤）
   if has_both_spaces_and_tabs:
       raise IndentationError("不可混用空格和 Tab")
   ```

3. **階層關係建立**：
   - 維護一個層級堆疊（stack）
   - 層級增加 → 成為上一項的 child
   - 層級減少 → 返回對應父層級
   - 層級相同 → 同層兄弟節點

---

## 五、系統模組設計

### 5.1 專案結構

```
md_word_renderer/
├── __init__.py
├── parser/
│   ├── __init__.py
│   ├── markdown_parser.py      # Markdown 解析器
│   ├── indent_detector.py      # 縮排偵測
│   └── escape_handler.py       # 特殊字元處理
├── renderer/
│   ├── __init__.py
│   ├── word_renderer.py        # Word 渲染引擎
│   └── error_handler.py        # 錯誤處理
├── validator/                   # 資料驗證模組（新增）
│   ├── __init__.py
│   ├── schema_validator.py     # Schema 驗證
│   └── schemas/
│       └── markdown_schema.json # JSON Schema 定義
├── config/                      # 設定檔處理（新增）
│   ├── __init__.py
│   ├── config_loader.py        # 設定檔載入器
│   └── default_config.yaml     # 預設設定
├── utils/
│   ├── __init__.py
│   ├── file_utils.py           # 檔案處理工具（新增）
│   └── batch_processor.py      # 批次處理器（新增）
├── templates/
│   ├── example_template.docx   # 範例模板 1
│   ├── simple_template.docx    # 簡單模板
│   └── full_template.docx      # 完整功能模板
├── tests/
│   ├── test_parser.py
│   ├── test_renderer.py
│   ├── test_validator.py       # 驗證測試（新增）
│   ├── test_batch.py           # 批次處理測試（新增）
│   └── fixtures/
│       ├── OH_20251020.md
│       ├── test_data_*.md      # 多個測試檔案
│       └── expected_output.json
├── examples/                    # 範例目錄（新增）
│   ├── basic_example.md
│   ├── config.yaml
│   └── README.md
├── cli.py                       # 命令列介面
├── main.py                      # 主程式
├── config.yaml                  # 使用者設定檔範例
└── requirements.txt
```

### 5.2 核心模組規格

#### 5.2.1 MarkdownParser

**檔案：** `parser/markdown_parser.py`

```python
class MarkdownParser:
    """
    解析特殊格式的 Markdown 檔案
    
    格式：編號. 欄位名稱 | 值
    支援階層結構（透過縮排）
    """
    
    def __init__(self):
        self.indent_detector = IndentDetector()
        self.escape_handler = EscapeHandler()
    
    def parse(self, filepath: str) -> dict:
        """
        解析 Markdown 檔案
        
        Args:
            filepath: .md 檔案路徑
            
        Returns:
            dict: 結構化的資料字典
            
        Raises:
            FileNotFoundError: 檔案不存在
            IndentationError: 縮排格式錯誤
            ParseError: 解析失敗
        """
        pass
    
    def _parse_line(self, line: str) -> dict:
        """解析單行資料"""
        pass
    
    def _build_hierarchy(self, items: list) -> dict:
        """建立階層關係"""
        pass
```

**主要功能：**
1. 讀取 Markdown 檔案
2. 偵測縮排類型（空格/Tab）
3. 解析每行的編號、欄位名、值
4. 處理特殊字元轉義
5. 建立階層結構
6. 生成雙重索引（名稱 + 編號）

#### 5.2.2 WordRenderer

**檔案：** `renderer/word_renderer.py`

```python
from docxtpl import DocxTemplate

class WordRenderer:
    """
    使用 docxtpl 渲染 Word 模板
    """
    
    def __init__(self):
        self.template = None
        self.data = None
        self.error_handler = ErrorHandler()
    
    def load_template(self, template_path: str):
        """載入 Word 模板"""
        self.template = DocxTemplate(template_path)
    
    def render(self, data: dict) -> None:
        """
        渲染模板
        
        Args:
            data: 解析後的資料字典
        """
        # 注入錯誤處理函數
        context = self._prepare_context(data)
        
        try:
            self.template.render(context)
        except Exception as e:
            # 處理渲染錯誤
            self.error_handler.handle(e)
    
    def save(self, output_path: str):
        """儲存渲染後的文件"""
        self.template.save(output_path)
    
    def _prepare_context(self, data: dict) -> dict:
        """準備渲染上下文，加入錯誤處理"""
        context = data.copy()
        context['_error'] = self.error_handler.format_error
        return context
```

**主要功能：**
1. 載入 Word 模板（.docx）
2. 使用 docxtpl 渲染
3. 錯誤變數處理（顯示 `[ERROR: 變數不存在]`）
4. 儲存輸出文件

#### 5.2.3 IndentDetector

**檔案：** `parser/indent_detector.py`

```python
class IndentDetector:
    """
    偵測並驗證縮排格式
    參考 Python PEP 8 的縮排規則
    """
    
    def detect(self, lines: list) -> tuple:
        """
        偵測縮排類型
        
        Returns:
            tuple: (縮排類型, 每層級空格數)
            - ('space', 4) 或
            - ('tab', 1)
        """
        pass
    
    def calculate_level(self, line: str, indent_type: str, indent_size: int) -> int:
        """計算縮排層級"""
        pass
    
    def validate(self, lines: list) -> bool:
        """驗證縮排一致性"""
        pass
```

#### 5.2.4 ErrorHandler

**檔案：** `renderer/error_handler.py`

```python
class ErrorHandler:
    """
    處理模板渲染時的錯誤
    """
    
    def format_error(self, variable_name: str) -> str:
        """
        格式化錯誤訊息
        
        Args:
            variable_name: 找不到的變數名稱
            
        Returns:
            str: "[ERROR: 變數 'xxx' 不存在]"
        """
        return f"[ERROR: 變數 '{variable_name}' 不存在]"
    
    def handle(self, exception: Exception):
        """處理渲染異常"""
        pass
```

#### 5.2.5 SchemaValidator

**檔案：** `validator/schema_validator.py`

```python
import jsonschema
from jsonschema import validate, ValidationError

class SchemaValidator:
    """
    驗證 Markdown 資料結構
    """
    
    def __init__(self, schema_path: str = None):
        """
        初始化驗證器
        
        Args:
            schema_path: JSON Schema 檔案路徑
        """
        self.schema = self._load_schema(schema_path)
    
    def validate(self, data: dict) -> tuple[bool, list]:
        """
        驗證資料結構
        
        Args:
            data: 解析後的資料字典
            
        Returns:
            tuple: (是否通過, 錯誤訊息列表)
        """
        try:
            validate(instance=data, schema=self.schema)
            return True, []
        except ValidationError as e:
            return False, [str(e)]
    
    def _load_schema(self, schema_path: str) -> dict:
        """載入 JSON Schema"""
        pass
```

#### 5.2.6 ConfigLoader

**檔案：** `config/config_loader.py`

```python
import yaml
from typing import Dict, Any

class ConfigLoader:
    """
    載入並管理設定檔
    """
    
    def __init__(self, config_path: str = "config.yaml"):
        self.config_path = config_path
        self.config = self._load_config()
    
    def _load_config(self) -> Dict[str, Any]:
        """
        載入 YAML 設定檔
        
        Returns:
            dict: 設定字典
        """
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except FileNotFoundError:
            return self._default_config()
    
    def _default_config(self) -> Dict[str, Any]:
        """預設設定"""
        return {
            'template': 'template.docx',
            'output_dir': 'output/',
            'indent_size': 4,
            'show_errors': True,
            'validate_schema': False
        }
    
    def get(self, key: str, default=None):
        """取得設定值"""
        return self.config.get(key, default)
    
    def merge_with_args(self, cli_args: dict) -> dict:
        """
        合併 CLI 參數（CLI 優先）
        
        Args:
            cli_args: 命令列參數字典
            
        Returns:
            dict: 合併後的設定
        """
        merged = self.config.copy()
        for key, value in cli_args.items():
            if value is not None:
                merged[key] = value
        return merged
```

#### 5.2.7 BatchProcessor

**檔案：** `utils/batch_processor.py`

```python
from typing import List, Callable
from pathlib import Path
import logging

class BatchProcessor:
    """
    批次處理多個 Markdown 檔案
    """
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
        self.logger = logging.getLogger(__name__)
    
    def process_files(self, 
                     input_files: List[str],
                     template_path: str,
                     output_dir: str,
                     process_func: Callable) -> dict:
        """
        批次處理檔案
        
        Args:
            input_files: 輸入檔案列表
            template_path: 模板路徑
            output_dir: 輸出目錄
            process_func: 處理函數
            
        Returns:
            dict: 處理結果統計
        """
        results = {
            'success': [],
            'failed': [],
            'total': len(input_files)
        }
        
        for i, input_file in enumerate(input_files, 1):
            try:
                output_file = self._generate_output_path(
                    input_file, output_dir
                )
                
                if self.verbose:
                    self.logger.info(f"[{i}/{len(input_files)}] 處理: {input_file}")
                
                process_func(input_file, template_path, output_file)
                results['success'].append(input_file)
                
            except Exception as e:
                self.logger.error(f"處理失敗: {input_file} - {str(e)}")
                results['failed'].append((input_file, str(e)))
        
        return results
    
    def _generate_output_path(self, input_file: str, output_dir: str) -> str:
        """
        生成輸出檔案路徑
        
        Args:
            input_file: 輸入檔案路徑
            output_dir: 輸出目錄
            
        Returns:
            str: 輸出檔案路徑
        """
        input_path = Path(input_file)
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        # 將 .md 改為 .docx
        output_filename = input_path.stem + '.docx'
        return str(output_path / output_filename)
```

---

## 六、實作步驟（Phase 1）

### Step 1: 環境建置與專案初始化

**時程：** 1 天

**任務：**
- [ ] 建立專案目錄結構
- [ ] 建立 `requirements.txt`
  ```
  python-docx>=0.8.11
  docxtpl>=0.16.7
  Jinja2>=3.1.2
  PyYAML>=6.0
  jsonschema>=4.17.0
  ```
- [ ] 初始化 Git repository
- [ ] 建立 `.gitignore`
- [ ] 建立虛擬環境
  ```bash
  python -m venv venv
  source venv/bin/activate  # macOS
  pip install -r requirements.txt
  ```

### Step 2: Markdown 解析器開發

**時程：** 4-5 天

**任務：**

**2.1 縮排偵測（1 天）**
- [ ] 實作 `IndentDetector.detect()`
- [ ] 實作 `IndentDetector.calculate_level()`
- [ ] 實作 `IndentDetector.validate()`
- [ ] 單元測試

**2.2 特殊字元處理（1 天）**
- [ ] 實作 `EscapeHandler`
- [ ] 處理 `\|`, `\n`, `\\`, `\"`
- [ ] 處理雙引號內的引號
- [ ] 單元測試

**2.3 單行解析（1 天）**
- [ ] 實作正則表達式匹配 `編號. 名稱 | 值`
- [ ] 解析編號、欄位名稱、值
- [ ] 處理空值情況
- [ ] 單元測試

**2.4 階層建立（2 天）**
- [ ] 實作堆疊演算法建立父子關係
- [ ] 實作 `_build_hierarchy()`
- [ ] 生成雙重索引（名稱 + `#編號`）
- [ ] 整合測試：完整解析 `OH_20251020.md`

### Step 3: Word 渲染器開發

**時程：** 3-4 天

**任務：**

**3.1 基本渲染（2 天）**
- [ ] 整合 docxtpl
- [ ] 實作 `WordRenderer.load_template()`
- [ ] 實作 `WordRenderer.render()`
- [ ] 實作 `WordRenderer.save()`
- [ ] 測試簡單變數替換

**3.2 錯誤處理（1 天）**
- [ ] 實作 `ErrorHandler`
- [ ] 在 Jinja2 上下文中注入錯誤處理
- [ ] 測試不存在的變數處理

**3.3 進階功能測試（1 天）**
- [ ] 測試條件渲染 `{% if %}`
- [ ] 測試迴圈渲染 `{% for %}`
- [ ] 測試巢狀迴圈
- [ ] 測試點記法存取

### Step 4: 資料驗證模組

**時程：** 2 天

**任務：**
- [ ] 定義 JSON Schema（`validator/schemas/markdown_schema.json`）
- [ ] 實作 `SchemaValidator` 類別
- [ ] 整合到解析流程
- [ ] 單元測試

### Step 5: 命令列介面（基本版）

**時程：** 1 天

**任務：**
- [ ] 實作 `cli.py` 使用 argparse
- [ ] 支援基本參數：
  ```
  --markdown, -m    Markdown 檔案路徑（單檔案）
  --template, -t    Word 模板路徑
  --output, -o      輸出檔案路徑
  --verbose, -v     顯示詳細資訊
  --validate        啟用資料驗證
  ```
- [ ] 實作錯誤訊息友善化

### Step 6: 批次處理功能

**時程：** 2 天

**任務：**
- [ ] 實作 `BatchProcessor` 類別
- [ ] 實作 `FileUtils` 工具函數
- [ ] 擴充 CLI 支援多檔案：
  ```
  --markdown, -m    支援多個檔案或通配符
  --output, -o      輸出目錄（多檔案時）
  ```
- [ ] 實作進度顯示（多檔案時）
- [ ] 單元測試

### Step 7: 設定檔支援

**時程：** 1 天

**任務：**
- [ ] 實作 `ConfigLoader` 類別
- [ ] 建立 `default_config.yaml` 範例
- [ ] CLI 整合設定檔載入
- [ ] CLI 參數優先於設定檔
- [ ] 單元測試

### Step 8: Word 模板範例製作

**時程：** 1-2 天

**任務：**
- [ ] 製作 `simple_template.docx`（基礎範例）
- [ ] 製作 `example_template.docx`（中等複雜度）
- [ ] 製作 `full_template.docx`（完整功能展示）
- [ ] 包含所有語法範例：
  - 簡單變數
  - 編號存取
  - 條件渲染
  - 單層迴圈
  - 巢狀迴圈
- [ ] 添加註解說明
- [ ] 建立模板使用指南

### Step 9: 整合測試與文件

**時程：** 2-3 天

**任務：**
- [ ] 端對端測試：`OH_20251020.md` + 模板 → 輸出
- [ ] 錯誤情境測試
- [ ] 效能測試（大型檔案）
- [ ] 撰寫 README.md
- [ ] 撰寫使用手冊
- [ ] 製作範例教學

---

## 七、Word 模板範例設計

### 7.1 模板檔案：example_template.docx

**模板內容：**

```
============================================
      變更單資料模板（示範檔案）
============================================

【基本資訊】
系統名稱：{{系統名稱}}
預定換版日期：{{預定換版日期}}
預定變更時段：{{預定變更時段}}
變更單號：{{變更單號}}
備援路徑：{{備援路徑}}
需求依據：{{需求依據(INC/PBI)}}

【相關系統】
資料庫：{{資料庫}}

{% if 中介軟體 %}
中介軟體：{{中介軟體}}
{% endif %}

{% if 網路(含防火牆) %}
網路設定：{{網路(含防火牆)}}
{% endif %}

{% if 其他系統 %}
其他系統：{{其他系統}}
{% endif %}

【通知資訊】
計畫性維護：{{計畫性維護}}
通知對象：{{通知對象}}
通知方式：{{通知/公告方式}}
通知內容：{{通知/公告內容}}

【異動說明】
通報單異動原因：{{通報單-異動原因}}

【測試案例】
{% for item in 異動內容-測試案例 %}
{{loop.index}}. {{item.value}}
{% if item.children %}
{% for child in item.children %}
   {{child.number}}. {{child.value}}
{% endfor %}
{% endif %}

{% endfor %}

【測試環境資訊】
測試日期：{{測試日期}}
主機名稱：{{主機名稱(IP)}}
作業系統：{{作業系統}}
瀏覽器：{{瀏覽器}}
測試網址：{{測試網址}}
測試參數：{{測試參數}}

============================================
                   結束
============================================
```

### 7.2 預期輸出範例

**使用 OH_20251020.md 渲染後：**

```
============================================
      變更單資料模板（示範檔案）
============================================

【基本資訊】
系統名稱：Super E-Billing 帳款管理系統
預定換版日期：_114_年_10_月_20_日
預定變更時段：_18:_00_-_20_:_00_
變更單號：CRQ000000179030
備援路徑：20251020_CRQ179030
需求依據：PBI000000061147

【相關系統】
資料庫：CRQ000000179000

【通知資訊】
計畫性維護：由連管科發函通知國內各單位
通知對象：數金部
通知方式：E-MAIL
通知內容：預計10/20(一)安排更版 - 中華電信驗章失敗重複發送

【異動說明】
通報單異動原因：中華電信驗章失敗及其他功能修復

【測試案例】
1. 調整供應商登入之密碼功能
   1. TC001：	新增供應商使用者
   2. TC002：	超過最長使用設定
   3. TC003：	未達最短天數設定
   4. TC004：   舊密碼差異檢核

2. 調整中華電信解析錯誤文字格式
   1. TC001：	中華電信上傳檔案格式錯誤
   2. TC002：	中華電信上傳檔案格式正確

3. 修復中華電信子公司受理分行無法使用「應付歷史下載」下載中華電信分支機構的帳款檔案
   1. TC001：	中華電信子公司上傳帳款

4. 修復113年紅隊演練行前檢核表項目-未妥善保存客戶敏感性文件
   1. TC001：	客戶上傳帳款檔，不儲存
   2. TC002：	客戶上傳薪資檔，不儲存
   3. TC003：	中華電信上傳檔案，不儲存
   4. TC004：	中華電信上傳檔案，開啟儲存設定

5. 中華電信驗章失敗重複發送
   1. TC001：	中華電信驗章失敗

【測試環境資訊】
測試日期：2025/10/11
主機名稱：10.1.7.197(Web)、10.1.7.198(AP)
作業系統：Windows 11
瀏覽器：Google Chrome
測試網址：https://snecomb.bot.com.tw/BOT_CIB_B2C_WEB/common/login/Login_1.faces
測試參數：統編:85905689 供應商: F207760507 使用者:user02

============================================
                   結束
============================================
```

---

## 八、命令列介面設計

### 8.1 基本使用

```bash
# 基本用法
python cli.py -m data.md -t template.docx -o output.docx

# 詳細輸出
python cli.py -m data.md -t template.docx -o output.docx -v

# 使用長參數名稱
python cli.py --markdown data.md --template template.docx --output result.docx
```

### 8.2 參數說明

```
Usage: cli.py [-h] -m MARKDOWN [MARKDOWN ...] -t TEMPLATE -o OUTPUT [-v] [--validate] [--config CONFIG]

Markdown to Word 模板渲染工具

Required arguments:
  -m, --markdown MARKDOWN [MARKDOWN ...]
                        Markdown 資料檔案路徑 (.md)，支援多檔案或通配符
  -t, --template TEMPLATE
                        Word 模板檔案路徑 (.docx)
  -o, --output OUTPUT   輸出路徑（單檔案時為檔案路徑，多檔案時為目錄）

Optional arguments:
  -h, --help            顯示此說明訊息
  -v, --verbose         顯示詳細處理資訊
  --validate            執行資料結構驗證
  --config CONFIG       指定設定檔路徑（預設: config.yaml）
  --version             顯示版本資訊

Examples:
  # 單檔案處理
  python cli.py -m OH_20251020.md -t template.docx -o result.docx
  
  # 多檔案批次處理
  python cli.py -m *.md -t template.docx -o output/
  python cli.py -m file1.md file2.md file3.md -t template.docx -o output/
  
  # 使用設定檔
  python cli.py --config my_config.yaml
  
  # 啟用驗證
  python cli.py -m data.md -t template.docx -o output.docx --validate -v
```

### 8.3 設定檔範例

**檔案：** `config.yaml`

```yaml
# Markdown to Word 模板渲染系統 - 設定檔

# 預設模板路徑
template: templates/default.docx

# 輸出目錄
output_dir: output/

# Markdown 解析設定
parsing:
  # 縮排大小（空格數）
  indent_size: 4
  # 是否允許混用 Tab 和空格
  allow_mixed_indent: false
  # 編碼
  encoding: utf-8

# 渲染設定
rendering:
  # 找不到變數時是否顯示錯誤訊息
  show_errors: true
  # 錯誤訊息格式
  error_format: "[ERROR: 變數 '{var}' 不存在]"
  # 是否保留 Word 模板原有樣式
  preserve_styles: true

# 驗證設定
validation:
  # 是否啟用資料結構驗證
  enabled: false
  # Schema 檔案路徑
  schema_path: validator/schemas/markdown_schema.json
  # 驗證失敗時是否中止
  strict_mode: false

# 批次處理設定
batch:
  # 是否顯示進度條
  show_progress: true
  # 失敗時是否繼續處理其他檔案
  continue_on_error: true
  # 輸出檔案命名規則（{name} 為原檔名）
  output_pattern: "{name}.docx"

# 日誌設定
logging:
  # 日誌等級：DEBUG, INFO, WARNING, ERROR
  level: INFO
  # 是否輸出到檔案
  file_output: false
  # 日誌檔案路徑
  log_file: logs/renderer.log

# 效能設定
performance:
  # 大檔案閾值（行數）
  large_file_threshold: 1000
  # 是否啟用快取
  enable_cache: false
```

**使用方式：**

```bash
# 1. 使用預設設定檔（config.yaml）
python cli.py --config

# 2. 指定設定檔路徑
python cli.py --config my_custom_config.yaml

# 3. 設定檔 + CLI 參數（CLI 參數優先）
python cli.py --config config.yaml -m data.md -o custom_output.docx

# 4. 覆寫特定設定
python cli.py --config config.yaml --validate -v
```

### 8.4 輸出訊息設計

**正常執行：**
```
[INFO] 讀取 Markdown 檔案: OH_20251020.md
[INFO] 偵測到縮排類型: 空格（每層 4 格）
[INFO] 解析完成，共 22 個欄位
[INFO] 載入 Word 模板: template.docx
[INFO] 開始渲染...
[INFO] 渲染完成
[INFO] 儲存至: output.docx
[SUCCESS] 處理完成！
```

**詳細模式（-v）：**
```
[INFO] 讀取 Markdown 檔案: OH_20251020.md
[DEBUG] 共 50 行
[DEBUG] 偵測縮排...
[INFO] 偵測到縮排類型: 空格（每層 4 格）
[DEBUG] 解析第 1 行: 1. 系統名稱 | Super E-Billing 帳款管理系統
[DEBUG] 解析第 2 行: 2. 預定換版日期 | _114_年_10_月_20_日
...
[INFO] 解析完成，共 22 個欄位
[DEBUG] 資料結構:
  - 系統名稱: "Super E-Billing 帳款管理系統"
  - 異動內容-測試案例: 5 個子項目
[INFO] 載入 Word 模板: template.docx
[DEBUG] 模板包含 15 個段落
[INFO] 開始渲染...
[DEBUG] 替換變數: 系統名稱
[DEBUG] 處理迴圈: 異動內容-測試案例
[INFO] 渲染完成
[INFO] 儲存至: output.docx
[SUCCESS] 處理完成！用時 0.35 秒
```

**錯誤處理：**
```
[ERROR] 檔案不存在: data.md
[ERROR] 模板格式錯誤: 無法開啟 template.docx
[ERROR] 解析失敗（第 10 行）: 格式不符合 "編號. 名稱 | 值"
[ERROR] 縮排錯誤（第 15 行）: 不可混用空格和 Tab
[WARNING] 模板引用不存在的變數: {{不存在的欄位}}
```

**多檔案批次處理：**
```
[INFO] 開始批次處理，共 5 個檔案
[INFO] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 0% 0/5

[INFO] [1/5] 處理: file1.md
[INFO] ✓ file1.md → output/file1.docx
[INFO] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 20% 1/5

[INFO] [2/5] 處理: file2.md
[INFO] ✓ file2.md → output/file2.docx
[INFO] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 40% 2/5

[INFO] [3/5] 處理: file3.md
[ERROR] ✗ file3.md - 解析失敗（第 5 行）
[INFO] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 60% 3/5

[INFO] [4/5] 處理: file4.md
[INFO] ✓ file4.md → output/file4.docx
[INFO] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 80% 4/5

[INFO] [5/5] 處理: file5.md
[INFO] ✓ file5.md → output/file5.docx
[INFO] ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ 100% 5/5

[SUCCESS] 批次處理完成！
成功: 4 個檔案
失敗: 1 個檔案
總計用時: 1.25 秒
```

**資料驗證輸出：**
```
[INFO] 啟用資料驗證
[INFO] 載入 Schema: validator/schemas/markdown_schema.json
[VALIDATE] 驗證 Markdown 資料結構...
[VALIDATE] ✓ 所有必要欄位存在
[VALIDATE] ✓ 欄位類型正確
[VALIDATE] ✓ 階層結構有效
[SUCCESS] 資料驗證通過
```

---

## 九、測試策略

### 9.1 單元測試

**測試檔案結構：**
```
tests/
├── test_markdown_parser.py
│   ├── test_parse_simple_line()
│   ├── test_parse_empty_value()
│   ├── test_parse_hierarchy()
│   ├── test_escape_characters()
│   └── test_indent_detection()
├── test_word_renderer.py
│   ├── test_load_template()
│   ├── test_render_simple_variable()
│   ├── test_render_loop()
│   ├── test_error_handling()
│   └── test_save_output()
└── fixtures/
    ├── simple.md              # 簡單測試資料
    ├── hierarchy.md           # 階層測試資料
    ├── special_chars.md       # 特殊字元測試
    └── template_test.docx     # 測試模板
```

**測試涵蓋率目標：** > 85%

### 9.2 整合測試

**測試案例：**

| 測試 ID | 測試項目 | 測試資料 | 預期結果 |
|---------|----------|----------|----------|
| IT-001 | 完整流程 | OH_20251020.md | 成功生成 Word |
| IT-002 | 空值處理 | 包含空值的 .md | 空值顯示為空白 |
| IT-003 | 特殊字元 | 包含 `\|` 等的 .md | 正確轉義 |
| IT-004 | 大型檔案 | 1000+ 行的 .md | 成功處理 |
| IT-005 | 錯誤變數 | 模板有不存在的變數 | 顯示錯誤訊息 |
| IT-006 | 巢狀迴圈 | 3 層階層資料 | 正確渲染 |
| IT-007 | 混合縮排 | 空格+Tab 混用 | 拋出錯誤 |

### 9.3 效能測試

**測試指標：**
- 小型檔案（< 100 行）：< 0.5 秒
- 中型檔案（100-500 行）：< 2 秒
- 大型檔案（> 500 行）：< 5 秒

---

## 十、交付清單

### 10.1 程式碼

- [x] 完整的 Python 專案原始碼
- [x] requirements.txt（含 PyYAML、jsonschema）
- [x] README.md（使用說明）
- [x] 命令列工具（cli.py）
- [ ] 設定檔範例（config.yaml）

### 10.2 文件

- [x] 實作計劃（本文件）
- [ ] API 文件（docstring）
- [ ] 使用手冊
- [ ] 模板語法參考
- [ ] 批次處理指南
- [ ] 設定檔說明

### 10.3 範例

- [ ] simple_template.docx（簡單模板）
- [ ] example_template.docx（中等複雜度模板）
- [ ] full_template.docx（完整功能模板）
- [x] OH_20251020.md（測試資料）
- [ ] 多個測試用 .md 檔案
- [ ] 渲染後的輸出範例
- [ ] 範例 config.yaml

### 10.4 測試

- [ ] 單元測試（pytest）
- [ ] 整合測試
- [ ] 批次處理測試
- [ ] 資料驗證測試
- [ ] 測試報告

---

## 十一、未來擴充規劃（Phase 2 & 3）

### Phase 2：進階模板功能

**預計時程：** 2-3 週

**功能項目：**
1. **表格渲染**
   - 動態生成表格行
   - 自動調整欄寬
   - 表格樣式套用

2. **圖片插入**
   - 支援本地圖片路徑
   - 支援 base64 圖片
   - 尺寸控制

3. **樣式控制**
   - 粗體、斜體、底線
   - 顏色設定
   - 字型大小

4. **自訂函數**
   - 日期格式化
   - 字串處理
   - 數學運算

### Phase 3：逆向解析

**預計時程：** 3-4 週

**核心概念：**
```
Word 文件 + Word 模板 = Markdown 資料
```

**實作挑戰：**
1. 從渲染後的 Word 反推變數值
2. 識別迴圈生成的內容
3. 重建階層結構
4. 處理格式變化

**應用場景：**
- 從現有文件提取資料
- 批次轉換 Word 為 Markdown
- 資料歸檔與管理

---

## 十二、風險與對策

### 12.1 技術風險

| 風險項目 | 影響程度 | 發生機率 | 對策 |
|----------|----------|----------|------|
| docxtpl 功能限制 | 高 | 中 | 提前驗證功能，必要時換用其他套件 |
| Word 格式相容性 | 中 | 中 | 測試多種 Word 版本 |
| 特殊字元處理複雜 | 中 | 高 | 參考成熟的轉義機制（如 Python） |
| 階層解析錯誤 | 高 | 中 | 完善的單元測試與錯誤處理 |
| 效能問題（大檔案） | 低 | 低 | 採用流式處理、優化演算法 |

### 12.2 需求風險

| 風險項目 | 對策 |
|----------|------|
| 需求變更 | 模組化設計，易於擴充 |
| 語法不夠彈性 | 採用 Jinja2，支援豐富語法 |
| 特殊格式需求 | 預留擴充介面 |

---

## 十三、開發準則

### 13.1 程式碼風格

- 遵循 PEP 8
- 使用 Black 格式化
- 使用 Flake8 檢查
- 型別提示（Type Hints）

### 13.2 Git 工作流程

**分支策略：**
```
main          # 穩定版本
├── develop   # 開發分支
│   ├── feature/parser     # 解析器開發
│   ├── feature/renderer   # 渲染器開發
│   └── feature/cli        # CLI 開發
└── hotfix    # 緊急修復
```

**Commit 訊息格式：**
```
<type>(<scope>): <subject>

<body>

<footer>
```

類型：
- `feat`: 新功能
- `fix`: 修復 bug
- `docs`: 文件更新
- `style`: 格式調整
- `refactor`: 重構
- `test`: 測試
- `chore`: 雜項

範例：
```
feat(parser): 實作階層解析功能

- 新增 _build_hierarchy() 方法
- 使用堆疊演算法建立父子關係
- 支援 3 層以上的巢狀結構

Closes #12
```

### 13.3 文件規範

**Docstring 格式（Google Style）：**
```python
def parse_line(line: str, indent_level: int) -> dict:
    """
    解析單行 Markdown 資料
    
    Args:
        line: 要解析的文字行
        indent_level: 縮排層級
        
    Returns:
        dict: 包含 number, key, value 的字典
        
    Raises:
        ParseError: 格式不符合規範時
        
    Example:
        >>> parse_line("1. 系統名稱 | Super E-Billing", 0)
        {'number': '1', 'key': '系統名稱', 'value': 'Super E-Billing', 'level': 0}
    """
    pass
```

---

## 十四、問題與討論

### Q1: 模板製作工具

**問題：** 是否需要提供視覺化的模板編輯器，幫助使用者製作模板？

**建議：** Phase 1 先提供完整的範例模板和文件，Phase 3 再考慮開發編輯器

**您的意見：** 不需要 因為模板直接是word檔案，在測試過程中可能需要建立範例模板

**✅ 決策：** 不開發視覺化編輯器，在開發過程中建立多個範例模板供參考

---

### Q2: 資料驗證

**問題：** 是否需要對 Markdown 資料進行結構驗證（Schema Validation）？

**建議：** 可選功能，提供 JSON Schema 定義檔

**您的意見：** 是

**✅ 決策：** 實作資料驗證功能，提供 JSON Schema 定義，CLI 支援 `--validate` 參數

---

### Q3: 多檔案支援

**問題：** 是否需要支援一次處理多個 Markdown 檔案？

**範例：**
```bash
python cli.py -m *.md -t template.docx -o output_dir/
```

**您的意見：** 是

**✅ 決策：** 支援多檔案批次處理，輸出檔名自動對應輸入檔名

---

### Q4: 設定檔支援

**問題：** 是否需要支援設定檔（如 config.yaml）來管理常用參數？

**範例 config.yaml：**
```yaml
template: templates/default.docx
output_dir: output/
indent_size: 4
show_errors: true
validate_schema: true
```

**您的意見：** 是

**✅ 決策：** 支援 config.yaml 設定檔，CLI 參數優先於設定檔

---

## 十五、開發時程總覽

### Phase 1 時程表

| 週次 | 任務 | 工作天數 | 交付項目 |
|------|------|----------|----------|
| W1 | 環境建置 + Markdown 解析器 + 驗證 | 5 | 可解析並驗證 .md 為 JSON |
| W2 | Word 渲染器 + 基本 CLI | 5 | 可渲染基本模板（單檔案） |
| W3 | 進階 CLI（多檔案、設定檔） | 3 | 支援批次處理 |
| W4 | 整合測試 + 文件 + 範例 | 3-5 | 完整可用版本 |

**預計總時程：** 3.5 - 4 週（因新增功能調整）

---

## 十六、驗收標準

### 16.1 功能驗收

- [ ] 可成功解析 OH_20251020.md
- [ ] 可正確渲染 example_template.docx
- [ ] 支援所有定義的模板語法
- [ ] 錯誤處理符合規格
- [ ] CLI 正常運作

### 16.2 品質驗收

- [ ] 單元測試涵蓋率 > 85%
- [ ] 所有整合測試通過
- [ ] 無 Critical/High 等級 bug
- [ ] 程式碼通過 Flake8 檢查
- [ ] 文件完整

### 16.3 效能驗收

- [ ] 處理 OH_20251020.md < 1 秒
- [ ] 記憶體使用 < 100MB
- [ ] 支援 1000+ 行檔案
- [ ] 批次處理 10 個檔案 < 5 秒

---

## 十七、開發前檢查清單

在開始開發前，請確認以下事項：

### 環境準備

- [ ] Python 3.8+ 已安裝
- [ ] Git 已安裝
- [ ] 已建立專案目錄
- [ ] 已建立虛擬環境
- [ ] 已安裝必要套件（requirements.txt）

### 資源準備

- [ ] 已取得測試資料（OH_20251020.md）
- [ ] 已準備 Word 軟體（用於製作模板）
- [ ] 已閱讀本計劃文件
- [ ] 已了解 docxtpl 和 Jinja2 基本語法

### 開發工具

- [ ] IDE/編輯器已設定（推薦 VS Code）
- [ ] Python 擴充套件已安裝
- [ ] Linter 已設定（Flake8）
- [ ] Formatter 已設定（Black）
- [ ] Git 已設定（user.name、user.email）

### 文件準備

- [ ] 已建立 README.md 框架
- [ ] 已建立 .gitignore
- [ ] 已建立 requirements.txt
- [ ] 已建立專案目錄結構

### 確認事項

- [ ] 已理解所有模板語法規範
- [ ] 已理解資料結構設計
- [ ] 已理解多檔案處理流程
- [ ] 已理解設定檔機制
- [ ] 已確認開發時程可行

**全部確認後，即可開始 Step 1 開發！**

---

## 十八、聯絡資訊與資源

### 技術參考文件

1. **docxtpl 官方文件**
   - https://docxtpl.readthedocs.io/

2. **Jinja2 模板語法**
   - https://jinja.palletsprojects.com/

3. **python-docx 文件**
   - https://python-docx.readthedocs.io/

4. **PEP 8 Style Guide**
   - https://www.python.org/dev/peps/pep-0008/

---

**計劃版本：** v2.0 Final  
**建立日期：** 2025-12-12  
**最後更新：** 2025-12-12  
**下次審查：** Phase 1 完成後

**計劃狀態：** ✅ 已確認，準備開發

**主要更新（v2.0）：**
- ✅ 新增資料驗證功能（JSON Schema）
- ✅ 新增多檔案批次處理
- ✅ 新增設定檔支援（config.yaml）
- ✅ 調整開發時程為 3.5-4 週
- ✅ 擴充專案結構（validator、config、utils 模組）
- ✅ 新增 3 個模板範例（simple、example、full）

---

## 附錄：快速開始指南（計劃確認後）

```bash
# 1. 克隆專案
git clone <repo-url>
cd md_word_renderer

# 2. 建立虛擬環境
python -m venv venv
source venv/bin/activate  # macOS/Linux
# 或 venv\Scripts\activate  # Windows

# 3. 安裝依賴
pip install -r requirements.txt

# 4. 執行範例
python cli.py \
  -m example/OH_20251020.md \
  -t templates/example_template.docx \
  -o output/result.docx \
  -v

# 5. 檢視結果
open output/result.docx  # macOS
```

此計劃文件將作為開發的藍圖，請確認後即可開始實作。
