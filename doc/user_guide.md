# MD-Word/Excel Template Renderer 使用手冊

> **作者：ytssamuel**  
> 適用版本：v2.2.1（Excel 樣板為主 + Jinja2 模板引擎）

## 目錄

1. [簡介](#簡介)
2. [安裝指南](#安裝指南)
3. [快速入門](#快速入門)
4. [Markdown 資料格式](#markdown-資料格式)
5. [Word 模板語法](#word-模板語法)
6. [Excel 樣板語意（v2.2.1）](#excel-樣板語意v2221)
7. [命令列工具](#命令列工具)
8. [Python API](#python-api)
9. [進階功能](#進階功能)
10. [疑難排解](#疑難排解)

---

## 簡介

MD-Word/Excel Template Renderer 是一個 Python 工具，能將特定格式的 Markdown 資料同時渲染至 **Word（.docx）** 與 **Excel（.xlsx）** 模板。主要用途：

- 自動化文件 / 報表生成
- 批次處理多個資料檔案
- 資料驅動的 Word 報告與 Excel 工作底稿
- 一份 Markdown 同時餵給 Word 與 Excel 兩條 pipeline

### 核心概念

1. **Markdown 資料來源** - 特殊格式的 `.md` 檔案，包含結構化資料
2. **Word / Excel 模板** - 使用 Jinja2 語法的 `.docx` 或 `.xlsx` 模板
3. **渲染輸出** - 合併資料與模板後的文件，CLI/GUI 會依樣板副檔名自動選擇對應 renderer

### 支援格式

| 格式 | 版本 | 模板語意 | 引擎 |
|---|---|---|---|
| Word (.docx) | v1.0+ | Jinja2 + docxtpl | `python-docx` + `docxtpl` |
| Excel (.xlsx) | v2.2+ | v2.2.1 起：Jinja2（`{{var}}` / `{% if %}` / `{% for %}`） | `openpyxl` |

---

## 安裝指南

### 系統需求

- Python 3.10 或更高版本
- Windows / macOS / Linux

### 安裝步驟

```bash
# 1. 建立虛擬環境
python -m venv venv

# 2. 啟動虛擬環境
# Windows:
venv\Scripts\activate
# macOS/Linux:
source venv/bin/activate

# 3. 安裝核心依賴
pip install -r requirements.txt

# 4. 安裝 GUI 依賴（選用）
pip install -r requirements-gui.txt
```

### 驗證安裝

```bash
python md2word.py info
```

應該會看到版本資訊（v2.2.1）、支援格式（Word + Excel）與功能說明。

---

## 快速入門

### 步驟 1：準備 Markdown 資料檔案

建立 `data.md`：

```markdown
# 資料

1. 系統名稱 | 帳款管理系統
2. 變更單號 | CRQ000001
3. 日期 | 2025/01/15
16. 異動內容-測試案例 |
    1. 登入功能測試
       1. 驗證帳號密碼
       2. 忘記密碼流程
```

### 步驟 2：準備模板

**Word 模板** `template.docx` 內容包含：

```
系統名稱：{{系統名稱}}
變更單號：{{變更單號}}
```

**Excel 樣板** `template.xlsx`（v2.2.1 起樣板為主）：

```
A1: 欄位    B1: 值
A2: 系統名稱 B2: {{系統名稱}}
A3: 變更單號 B3: {{變更單號}}
```

### 步驟 3：執行轉換

```bash
# Word
python md2word.py render data.md template.docx output.docx

# Excel（自動從副檔名偵測）
python md2word.py render data.md template.xlsx output.xlsx

# 明確指定格式
python md2word.py render data.md template.xlsx output.xlsx --format xlsx
```

### 步驟 4：檢視結果

打開 `output.docx` 或 `output.xlsx`，應該會看到 `{{變數}}` 已被實際值取代。

---

## Markdown 資料格式

### 基本格式

每行資料遵循格式：`編號. 欄位名稱 | 值`

```markdown
1. 系統名稱 | 範例系統
2. 變更單號 | CRQ000000100000
3. 備註 | 這是備註內容
```

### 階層結構

使用縮排建立階層關係（4 空格或 Tab）：

```markdown
16. 測試案例 |
    1. 功能測試 | 測試登入功能
        1. 子測試 A | 驗證帳號
        2. 子測試 B | 驗證密碼
    2. 效能測試 | 測試系統效能
```

解析結果會是巢狀結構，方便在模板中迴圈處理。

### 空值欄位

值為空時，只需保留分隔符號：

```markdown
8. 中介軟體 |
9. 網路設定 |
```

### 特殊字元轉義

如果值中包含 `|`，使用反斜線轉義：

```markdown
10. 公式 | x \| y = z
11. 多行內容 | 第一行\n第二行
```

### 支援的轉義

| 轉義序列 | 說明 |
|---------|------|
| `\|` | 管道符號 |
| `\n` | 換行 |
| `\\` | 反斜線 |
| `\'` | 單引號 |
| `\"` | 雙引號 |

### 圖片語法

使用標準 Markdown 圖片語法嵌入圖片：

```markdown
![替代文字](圖片路徑)
```

**範例：**

```markdown
16. 異動內容-測試案例 |
    1. 調整供應商登入之密碼功能
       1. 新增供應商使用者
          1. ![2025-11-01_104655](17. 異動內容-測試案例/1. 調整/image.png)
```

**重要說明：**
- 圖片路徑可以是相對路徑（相對於 Markdown 檔案所在目錄）
- 替代文字會被保存為圖片說明
- 支援的圖片格式：PNG、JPG、JPEG、GIF、BMP
- Word 模板：`{{item.image}}` 渲染為 InlineImage
- Excel 樣板：圖片自動嵌入並等比縮放（`excel.image.max_width_px` / `max_height_px` 控制上限）

**在 Word 模板中使用圖片：**

```jinja2
{% for item in data["異動內容-測試案例"] %}
  {% for child in item.children %}
    {% if child.type == "image" %}
      {{child.image}}    {# 輸出圖片本身 #}
    {% else %}
      {{child.value}}    {# 輸出純文字 #}
    {% endif %}
  {% endfor %}
{% endfor %}
```

---

## Word 模板語法

Word 模板使用 docxtpl + Jinja2，支援變數替換、條件、迴圈。

### 簡單變數

```jinja2
系統名稱：{{系統名稱}}
變更單號：{{變更單號}}
```

### 含特殊字元的欄位名稱

欄位名稱含括號、斜線等特殊字元時，需使用字典存取語法：

```jinja2
需求依據：{{data["需求依據(INC/PBI)"]}}
通知方式：{{data["通知/公告方式"]}}
主機 IP：{{data["主機名稱(IP)"]}}
```

### 條件渲染

```jinja2
{% if 中介軟體 %}
中介軟體設定：{{中介軟體}}
{% else %}
中介軟體：無
{% endif %}
```

### 迴圈渲染

處理列表資料（如測試案例）：

```jinja2
{% for item in data["異動內容-測試案例"] %}
{{loop.index}}. {{item.value}}
  {% for child in item.children %}
  - {{child.number}}. {{child.value}}
  {% endfor %}
{% endfor %}
```

### 迴圈變數

在 `{% for %}` 迴圈中可用的變數：

| 變數 | 說明 |
|------|------|
| `loop.index` | 從 1 開始的索引 |
| `loop.index0` | 從 0 開始的索引 |
| `loop.first` | 是否為第一項 |
| `loop.last` | 是否為最後一項 |
| `loop.length` | 列表總長度 |

### 預設值

```jinja2
{{中介軟體 | default("無")}}
```

---

## Excel 樣板語意（v2.2.1）

v2.2.1 起，Excel 樣板**全面走 Jinja2**，與 Word renderer 對齊。

### 基本變數替換（Phase 1）

樣板內 `{{變數}}` 自動被替換；特殊字元鍵用 `data["..."]` 形式：

```
A1: 欄位        B1: 值
A2: 系統名稱    B2: {{系統名稱}}
A3: 變更單號    B3: {{變更單號}}
A4: 需求依據    B4: {{data["需求依據(INC/PBI)"]}}
A5: 通知方式    B5: {{data["通知/公告方式"]}}
```

```jinja2
A6: 第一個欄位的 key   B6: {{data["#1"].key}}
```

> **不再自動 append**：v2.2.0 的「自動把純量欄位往 header sheet append」行為在 v2.2.1 拿掉。樣板上寫多少 `{{...}}`，產出就多少列；沒寫就不會自動加。`auto_flatten_lists: true`（預設）讓 v2.2.0 既有 list sheet auto-flatten 仍可用。

### 條件渲染（Phase 2）

```jinja2
{% if 資料庫 %}有 {{ 資料庫 }}{% else %}無{% endif %}
```

Jinja2 原生語意：條件為真時渲染對應分支；否則清空（silent 模式）。

### 列展開（Phase 3 — 單層 `{% for %}`）

在樣板的「整張 sheet」使用 `{% for %}` 與 `{% endfor %}` 標記，標記必須在**自己的 row**，body 在中間。

```
A1: 編號   B1: 內容           C1: 型別    D1: 子項數
A2: {% for case in data['#16'].children %}
A3: {{case.number}}   B3: {{case.value}}   C3: {{case.type or 'text'}}   D3: {{case.children|length}}
A4: {% endfor %}
```

渲染後行為：

- 標記 row（A2、A4）會被刪除
- body row（A3）會被複製 N 份（N = 清單長度，例如 5 cases）並 append 到 sheet 底部
- 每份綁定 `case` 為當前 item，可用 `{{case.number}}`、`{{case.value}}`、`{{case.children|length}}` 等
- `loop.index` / `loop.first` / `loop.last` 也可用

**限制（v2.2.1）**：
- 不支援巢狀 `{% for %}`（內層為 v2.2.2+ 規劃）
- 內層 for marker 與外層 body 必須在不同 row

### 缺變數行為

| 模式 | 行為 |
|---|---|
| `silent`（預設） | `{{missing}}` → 空字串 |
| `keep` | 保留 `{{missing}}` 字面字串（除錯用） |
| `error_format` | 套用 `error_format` 模板，例：`[ERROR: missing]` |

設定方式（`config.yaml`）：

```yaml
excel:
  template_engine:
    enabled: true
    auto_flatten_lists: true   # 向後相容 v2.2.0
    missing_variable: silent
    error_format: "[ERROR: 變數 '{var}' 不存在]"
```

### LAYOUT 隱藏工作表（覆寫 layout 設定）

樣板內可放一張隱藏的 `LAYOUT` sheet，欄位 `key | value`，可覆寫 layout 設定：

| key | 範例值 | 對應 LayoutConfig |
|---|---|---|
| `auto_fit_columns` | `True` | `auto_fit_columns` |
| `image_max_width_px` | `480` | `image.max_width_px` |
| `template_engine_enabled` | `True` | `template_engine.enabled` |
| `auto_flatten_lists` | `False` | `template_engine.auto_flatten_lists` |
| `missing_variable` | `silent` | `template_engine.missing_variable` |

LAYOUT sheet 在 renderer 完成後會自動從活頁簿移除，**不會出現在最終 .xlsx 內**。

### 完整範例：`templates/excel/with_lists_template.xlsx`

```
=== sheet: 基本資訊 ===
A1: 欄位       B1: 值
A2: 系統名稱   B2: {{系統名稱}}
A3: 變更單號   B3: {{變更單號}}

=== sheet: 異動內容-測試案例 ===
A1: 編號       B1: 內容          C1: 型別     D1: 子項數
A2: {% for case in data['#16'].children %}
A3: {{case.number}}  B3: {{case.value}}  C3: {{case.type or 'text'}}  D3: {{case.children|length}}
A4: {% endfor %}

=== sheet: 異動內容-子項 ===
A1: 父項編號   B1: 父項內容      C1: 子項數    D1: 第一個子項編號  E1: 第一個子項值
A2: {% for case in data['#16'].children %}
A3: {{case.number}}  B3: {{case.value}}  C3: {{case.children|length}}
    D3: {% if case.children %}{{case.children[0].number}}{% endif %}
    E3: {% if case.children %}{{case.children[0].value}}{% endif %}
A4: {% endfor %}
```

### 圖片嵌入

Markdown 內 `![alt](path)` 在 list 葉節點上會以 `type: image` 標記；Excel renderer 自動等比縮放後嵌入 B 欄的儲存格（`excel.image.max_width_px` / `max_height_px` 控制上限）。圖片找不到時會在「型別」欄寫入 `image-missing: ...`，不會讓整批失敗。

---

## 命令列工具

### 基本指令

```bash
python md2word.py <command> [options]
```

可用子命令：`render` / `batch` / `batch-templates` / `validate` / `info`

### render - 單一檔案轉換

```bash
python md2word.py render <input.md> <template> <output>

選項：
  -v, --verbose              顯示詳細處理資訊
  --no-validate              跳過資料驗證步驟
  --format {auto,docx,xlsx}  輸出格式；預設 auto 從樣板副檔名推斷

範例：
  # Word（自動推斷）
  python md2word.py render data.md template.docx output.docx -v

  # Excel（自動推斷）
  python md2word.py render data.md templates/excel/sample_template.xlsx output.xlsx

  # 明確指定格式（覆寫副檔名推斷）
  python md2word.py render data.md tpl.xlsx out.xlsx --format xlsx
```

### batch - 批次轉換（多 MD + 單模板）

將多個 Markdown 資料檔案使用同一個模板轉換。

```bash
python md2word.py batch <input_dir> <template> <output_dir>

選項：
  -p, --pattern           檔案搜尋模式 (預設: *.md)
  -v, --verbose           顯示詳細資訊
  --no-validate           跳過資料驗證
  --continue-on-error     遇到錯誤時繼續處理其他檔案
  --format {auto,docx,xlsx}  輸出格式；預設 auto 從樣板副檔名推斷

範例：
  python md2word.py batch ./inputs/ template.docx ./outputs/ -v
  python md2word.py batch ./data/ template.docx ./out/ -p "OH_*.md"
  python md2word.py batch ./data/ template.xlsx ./out/
```

### batch-templates - 多模板批次轉換（單 MD + 多模板）

將一份 Markdown 資料使用多個模板轉換，產生多個文件。**同資料夾內可同時混搭 `.docx` / `.xlsx`**，各自決定輸出格式。

```bash
python md2word.py batch-templates <input.md> <template_dir> <output_dir>

選項：
  -p, --pattern           模板搜尋模式 (預設 *.docx；v2.2 起 *.docx 與 *.xlsx 皆可)
  -v, --verbose           顯示詳細資訊
  --no-validate           跳過資料驗證
  --continue-on-error     遇到錯誤時繼續處理其他模板
  --prefix                輸出檔案名稱前綴
  --suffix                輸出檔案名稱後綴（在副檔名之前）
  --format {auto,docx,xlsx}  覆寫所有模板的格式

範例：
  # 基本用法
  python md2word.py batch-templates data.md ./templates/ ./outputs/

  # 混合 .docx 與 .xlsx 樣板（依樣板副檔名自動決定輸出）
  python md2word.py batch-templates data.md ./templates/ ./outputs/ --no-validate

  # 加入前綴和後綴
  python md2word.py batch-templates data.md ./templates/ ./outputs/ --prefix "2025_" --suffix "_final"

  # 詳細輸出
  python md2word.py batch-templates data.md ./templates/ ./outputs/ -v
```

**使用情境：**
- 同一份資料需要產生不同格式的 Word 報告 + Excel 工作底稿
- 不同部門需要同一份資料的不同視角文件

### validate - 驗證資料

```bash
python md2word.py validate <input.md>

選項：
  -s, --schema        自訂 JSON Schema 檔案

範例：
  python md2word.py validate data.md
  python md2word.py validate data.md -s custom_schema.json
```

### info - 顯示資訊

```bash
python md2word.py info
```

---

## Python API

### 統一入口：`build_renderer`

```python
from md_word_renderer.renderer.factory import build_renderer

renderer = build_renderer(template_path='template.docx')   # 自動選 WordRenderer
renderer = build_renderer(template_path='sample.xlsx')     # 自動選 ExcelRenderer
```

### 基本使用（Word）

```python
from md_word_renderer import MarkdownParser, WordRenderer

# 解析 Markdown
parser = MarkdownParser()
data = parser.parse('data.md')

# 渲染至 Word
renderer = WordRenderer()
renderer.load_template('template.docx')
renderer.render(data)
renderer.save('output.docx')
```

### 基本使用（Excel）

```python
from md_word_renderer import MarkdownParser
from md_word_renderer.renderer.factory import build_renderer

parser = MarkdownParser()
data = parser.parse('data.md')

renderer = build_renderer(template_path='templates/excel/sample_template.xlsx')
renderer.render_to_file(data, 'templates/excel/sample_template.xlsx', 'output.xlsx')
```

### 一行完成

```python
renderer = build_renderer(template_path='template.docx')
renderer.render_to_file(data, 'template.docx', 'output.docx')
```

### 資料驗證

```python
from md_word_renderer import SchemaValidator

validator = SchemaValidator()
is_valid, errors = validator.validate(data)

if not is_valid:
    for error in errors:
        print(f"驗證錯誤: {error}")
```

### 存取解析結果

```python
# 存取欄位值
system_name = data['系統名稱']

# 存取列表
test_cases = data['異動內容-測試案例']
for case in test_cases:
    print(f"{case['number']}. {case['value']}")
    for child in case['children']:
        print(f"  - {child['value']}")

# 透過編號存取
first_field = data['#1']  # {'key': '系統名稱', 'value': '...'}
```

### 處理圖片

解析含有圖片的資料後，圖片項目會有特殊的 `type` 屬性：

```python
from md_word_renderer import MarkdownParser

parser = MarkdownParser()
data = parser.parse('data_with_images.md')

# 遍歷尋找圖片
def find_images(items):
    for item in items:
        if item.get('type') == 'image':
            print(f"找到圖片: {item['image_path']}")
            print(f"  替代文字: {item['image_alt']}")
        if item.get('children'):
            find_images(item['children'])

test_cases = data.get('異動內容-測試案例', [])
find_images(test_cases)
```

### 渲染含圖片的 Word 資料

```python
from md_word_renderer import WordRenderer
from docx.shared import Cm

# 建立渲染器，設定圖片尺寸
renderer = WordRenderer(image_width=Cm(15), image_height=None)

renderer.load_template('template.docx')
renderer.render(data)

# 檢查是否有遺失的圖片
missing_images = renderer.get_missing_images()
if missing_images:
    print("警告：以下圖片檔案不存在：")
    for img in missing_images:
        print(f"  - {img}")

renderer.save('output.docx')
```

---

## 進階功能

### 自訂 Schema 驗證

建立 `custom_schema.json`：

```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "type": "object",
  "required": ["系統名稱", "變更單號"],
  "properties": {
    "系統名稱": {"type": "string", "minLength": 1},
    "變更單號": {"type": "string", "pattern": "^CRQ\\d+$"}
  }
}
```

使用：

```bash
python md2word.py validate data.md -s custom_schema.json
```

### 批次處理腳本（Word + Excel 混搭）

```python
from pathlib import Path
from md_word_renderer import MarkdownParser
from md_word_renderer.renderer.factory import build_renderer

parser = MarkdownParser()

input_dir = Path('inputs')
output_dir = Path('outputs')
output_dir.mkdir(exist_ok=True)

for md_file in input_dir.glob('*.md'):
    data = parser.parse(str(md_file))
    # 用同一份 Markdown 同時產 Word 與 Excel
    word = build_renderer(template_path='template.docx')
    word.render_to_file(data, 'template.docx', str(output_dir / f"{md_file.stem}.docx"))

    excel = build_renderer(template_path='template.xlsx')
    excel.render_to_file(data, 'template.xlsx', str(output_dir / f"{md_file.stem}.xlsx"))

    print(f"已處理: {md_file.name}")
```

### Excel 自訂 LayoutConfig

```python
from md_word_renderer.renderer.excel_renderer import ExcelRenderer
from md_word_renderer.renderer.excel_layout import LayoutConfig, TemplateEngineConfig

layout = LayoutConfig(
    template_engine=TemplateEngineConfig(
        enabled=True,
        auto_flatten_lists=False,   # 純樣板驅動
        missing_variable="silent",
        error_format="[ERROR: 變數 '{var}' 不存在]",
    ),
    image_max_width_px=300,   # 限制圖片寬度
)
renderer = ExcelRenderer(layout=layout)
```

---

## 疑難排解

### 常見錯誤

#### 1. 模板變數未定義

**錯誤訊息**：`'欄位名稱' is undefined`

**原因**：Markdown 中沒有該欄位，或欄位名稱不符

**解決方案**：
- 檢查 Markdown 檔案是否包含該欄位
- 確認欄位名稱完全一致（包含空格）
- 對於特殊字元欄位，使用 `data["欄位名稱"]` 語法
- Excel 可改 `excel.template_engine.missing_variable: "error_format"` 讓錯誤訊息可見

#### 2. 無法解析 Markdown

**錯誤訊息**：解析結果為空或欄位數量不對

**原因**：格式不符

**解決方案**：
- 確認每行格式為：`編號. 欄位名稱 | 值`
- 確認編號後有 `.` 和空格
- 確認欄位名稱和值之間有 ` | ` 分隔

#### 3. 階層結構錯誤

**錯誤訊息**：子項目沒有正確歸屬到父項目

**原因**：縮排不一致

**解決方案**：
- 統一使用空格或 Tab
- 確保子項目的縮排大於父項目

#### 4. Excel 樣板 `{{...}}` 沒被替換

**原因**：v2.2.0 行為沒有處理樣板文字；升級至 v2.2.1 後自動套模板引擎。

**解決方案**：
- 確認使用的是 v2.2.1 版（`python md2word.py info` 看版本）
- 若要關閉模板引擎：`excel.template_engine.enabled: false`（少用）
- 缺變數行為改 `excel.template_engine.missing_variable: "error_format"` 查看問題

#### 5. Excel `{% for %}` 沒展開

**原因**：`{% for %}` 標記被放在與 body 同一 row。

**解決方案**：
- `{% for %}` 與 `{% endfor %}` 必須各自在獨立 row
- body rows 在中間
- 巢狀 for 為 v2.2.2+ 規劃，v2.2.1 不支援

#### 6. openpyxl 找不到

**錯誤訊息**：`openpyxl does not support .xlsx file format...`

**解決方案**：
```bash
pip install "openpyxl>=3.1,<4"
```

### 檢視解析結果

```python
import json
from md_word_renderer import MarkdownParser

parser = MarkdownParser()
data = parser.parse('data.md')

# 印出完整資料結構
print(json.dumps(data, ensure_ascii=False, indent=2))
```

---

## 聯絡與支援

如有問題，請提交 Issue 或聯絡作者 **ytssamuel**。