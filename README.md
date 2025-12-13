# MD-Word Template Renderer

將結構化的 Markdown 資料渲染至 Word 模板，實現資料驅動的文件生成系統。

## 專案狀態

✅ **v1.0 完成** - 核心功能已實作

## 功能特色

- ✅ 解析特殊格式的 Markdown 資料（編號.名稱|值）
- ✅ 支援階層結構（透過縮排，含不規則縮排智慧偵測）
- ✅ 基於 Jinja2 的強大模板語法
- ✅ 資料結構驗證（JSON Schema）
- ✅ 批次處理多個檔案
- ✅ 命令列工具 (CLI)
- ✅ 完整測試套件（43 個測試）

## 快速開始

### 安裝

```bash
# 克隆專案
git clone <repo-url>
cd SpeedBOT

# 建立虛擬環境
python -m venv venv
venv\Scripts\activate    # Windows
# 或 source venv/bin/activate  # macOS/Linux

# 安裝依賴
pip install -r requirements.txt
```

### 基本使用

#### 命令列工具

```bash
# 單一檔案轉換
python md2word.py render input.md template.docx output.docx

# 批次轉換
python md2word.py batch ./inputs/ template.docx ./outputs/

# 驗證 Markdown 格式
python md2word.py validate input.md

# 顯示版本資訊
python md2word.py info

# 顯示說明
python md2word.py --help
```

#### Python API

```python
from md_word_renderer import MarkdownParser, WordRenderer

# 解析 Markdown
parser = MarkdownParser()
data = parser.parse('data.md')

# 渲染 Word
renderer = WordRenderer()
renderer.load_template('template.docx')
renderer.render(data)
renderer.save('output.docx')

# 或一行完成
renderer.render_to_file(data, 'template.docx', 'output.docx')
```

## Markdown 資料格式

```markdown
# 變更單資料模板

1. 系統名稱 | Super E-Billing 帳款管理系統
2. 預定換版日期 | 114年10月20日
3. 變更單號 | CRQ000000179030
16. 異動內容-測試案例 | 
    1. 調整供應商登入之密碼功能
       1. TC001：新增供應商使用者
       2. TC002：超過最長使用設定
    2. 調整中華電信解析錯誤文字格式
```

**格式規則：**
- 格式：`編號. 欄位名稱 | 值`
- 階層：使用縮排（空格或 Tab，支援不規則縮排）
- 特殊字元：使用反斜線轉義（`\|`, `\n`）
- 空值：`欄位名稱 |` （值為空）

## Word 模板語法

### 簡單變數

```jinja2
系統名稱：{{系統名稱}}
變更單號：{{變更單號}}
```

### 含特殊字元的欄位（括號、斜線等）

```jinja2
需求依據：{{data["需求依據(INC/PBI)"]}}
通知方式：{{data["通知/公告方式"]}}
```

### 條件渲染

```jinja2
{% if 中介軟體 %}
中介軟體：{{中介軟體}}
{% endif %}
```

### 列表渲染

```jinja2
{% for item in data["異動內容-測試案例"] %}
{{loop.index}}. {{item.value}}
  {% for child in item.children %}
  {{child.number}}. {{child.value}}
  {% endfor %}
{% endfor %}
```

## 解析後資料結構

```python
{
    "系統名稱": "Super E-Billing 帳款管理系統",
    "變更單號": "CRQ000000179030",
    "異動內容-測試案例": [
        {
            "number": "1",
            "key": "項目名稱",
            "value": "調整供應商登入之密碼功能",
            "children": [
                {"number": "1", "key": "子項目1", "value": "TC001：..."},
                {"number": "2", "key": "子項目2", "value": "TC002：..."}
            ]
        }
    ],
    "#1": {"key": "系統名稱", "value": "..."},  # 編號索引
    "#2": {"key": "變更單號", "value": "..."}
}
```

## 專案結構

```
SpeedBOT/
├── src/md_word_renderer/     # 主要模組
│   ├── parser/               # Markdown 解析器
│   │   ├── markdown_parser.py
│   │   ├── indent_detector.py
│   │   └── escape_handler.py
│   ├── renderer/             # Word 渲染引擎
│   │   ├── word_renderer.py
│   │   └── error_handler.py
│   ├── validator/            # 資料驗證
│   │   └── schema_validator.py
│   ├── cli/                  # 命令列工具
│   │   └── main.py
│   └── __init__.py
├── templates/                # Word 模板範例
│   ├── simple_template.docx
│   ├── example_template.docx
│   └── full_template.docx
├── test/                     # 測試套件
│   ├── test_parser.py
│   ├── test_renderer.py
│   ├── test_validator.py
│   └── test_cli.py
├── output/                   # 輸出目錄
├── doc/                      # 文件
├── md2word.py               # CLI 入口點
├── requirements.txt
└── README.md
```

## 執行測試

```bash
# 執行所有測試
python -m pytest test/ -v

# 執行端到端測試
python test_e2e.py

# 顯示覆蓋率
python -m pytest test/ --cov=src/md_word_renderer
```

## CLI 指令參考

### render - 單一檔案轉換

```bash
python md2word.py render <input.md> <template.docx> <output.docx> [options]

選項：
  -v, --verbose       顯示詳細資訊
  --no-validate       跳過資料驗證
```

### batch - 批次轉換

```bash
python md2word.py batch <input_dir> <template.docx> <output_dir> [options]

選項：
  -p, --pattern       檔案搜尋模式 (預設: *.md)
  -v, --verbose       顯示詳細資訊
  --continue-on-error 遇到錯誤時繼續處理
```

### validate - 驗證資料

```bash
python md2word.py validate <input.md> [options]

選項：
  -s, --schema        自訂 JSON Schema 檔案
```

## 技術棧

- **Python** 3.10+
- **docxtpl** 0.16+ - Word 模板引擎
- **Jinja2** 3.1+ - 模板語法
- **python-docx** 0.8+ - Word 文件處理
- **PyYAML** 6.0+ - 設定檔解析
- **jsonschema** 4.17+ - 資料驗證

## 授權

MIT License

## 開發日誌

### v1.0.0 (2024)

- ✅ Markdown 解析器（含不規則縮排支援）
- ✅ Word 渲染引擎
- ✅ 資料驗證（JSON Schema）
- ✅ 批次處理
- ✅ 命令列工具
- ✅ 完整測試套件（43 個測試通過）
- ✅ 範例模板

---

**開發狀態：** v1.0 已完成  
**測試結果：** 43 passed ✅
