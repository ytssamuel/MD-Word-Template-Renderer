# MD-Word Template Renderer 使用手冊

## 目錄

1. [簡介](#簡介)
2. [安裝指南](#安裝指南)
3. [快速入門](#快速入門)
4. [Markdown 資料格式](#markdown-資料格式)
5. [Word 模板語法](#word-模板語法)
6. [命令列工具](#命令列工具)
7. [Python API](#python-api)
8. [進階功能](#進階功能)
9. [疑難排解](#疑難排解)

---

## 簡介

MD-Word Template Renderer 是一個 Python 工具，用於將特定格式的 Markdown 資料渲染至 Word 模板。主要用途：

- 自動化文件生成
- 批次處理多個資料檔案
- 資料驅動的報告生成

### 核心概念

1. **Markdown 資料來源** - 特殊格式的 .md 檔案，包含結構化資料
2. **Word 模板** - 使用 Jinja2 語法的 .docx 模板
3. **渲染輸出** - 合併資料與模板後的 Word 文件

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

# 3. 安裝依賴
pip install -r requirements.txt
```

### 驗證安裝

```bash
python md2word.py info
```

應該會看到版本資訊和功能說明。

---

## 快速入門

### 步驟 1：準備 Markdown 資料檔案

建立 `data.md`：

```markdown
# 資料

1. 系統名稱 | 帳款管理系統
2. 變更單號 | CRQ000001
3. 日期 | 2025/01/15
```

### 步驟 2：準備 Word 模板

建立 `template.docx`，內容包含：

```
系統名稱：{{系統名稱}}
變更單號：{{變更單號}}
日期：{{日期}}
```

### 步驟 3：執行轉換

```bash
python md2word.py render data.md template.docx output.docx
```

### 步驟 4：檢視結果

打開 `output.docx`，應該會看到變數已被實際值取代。

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

### 支援的轉義：

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
- 圖片會自動轉換為 Word 中的 InlineImage

**解析結果範例：**

```json
{
  "type": "image",
  "number": "1",
  "value": "2025-11-01_104655",
  "image_path": "C:\\absolute\\path\\to\\image.png",
  "image_alt": "2025-11-01_104655",
  "image": "<InlineImage 物件，渲染時自動產生>",
  "children": []
}
```

**在模板中使用圖片：**
- `{{item.value}}` - 輸出替代文字（純文字）
- `{{item.image}}` - 輸出圖片本身（InlineImage 物件）

---

## Word 模板語法

模板使用 Jinja2 語法，支援變數替換、條件、迴圈等功能。

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

## 命令列工具

### 基本指令

```bash
python md2word.py <command> [options]
```

### render - 單一檔案轉換

```bash
python md2word.py render <input.md> <template.docx> <output.docx>

選項：
  -v, --verbose       顯示詳細處理資訊
  --no-validate       跳過資料驗證步驟

範例：
  python md2word.py render data.md template.docx output.docx -v
```

### batch - 批次轉換（多 MD + 單模板）

將多個 Markdown 資料檔案使用同一個模板轉換。

```bash
python md2word.py batch <input_dir> <template.docx> <output_dir>

選項：
  -p, --pattern       檔案搜尋模式 (預設: *.md)
  -v, --verbose       顯示詳細資訊
  --continue-on-error 遇到錯誤時繼續處理其他檔案

範例：
  python md2word.py batch ./inputs/ template.docx ./outputs/ -v
  python md2word.py batch ./data/ template.docx ./out/ -p "OH_*.md"
```

### batch-templates - 多模板批次轉換（單 MD + 多模板）

將一份 Markdown 資料使用多個模板轉換，產生多個不同的 Word 文件。

```bash
python md2word.py batch-templates <input.md> <template_dir> <output_dir>

選項：
  -p, --pattern       模板搜尋模式 (預設: *.docx)
  -v, --verbose       顯示詳細資訊
  --continue-on-error 遇到錯誤時繼續處理其他模板
  --prefix            輸出檔案名稱前綴
  --suffix            輸出檔案名稱後綴（在副檔名之前）

範例：
  # 基本用法
  python md2word.py batch-templates data.md ./templates/ ./outputs/
  
  # 加入前綴和後綴
  python md2word.py batch-templates data.md ./templates/ ./outputs/ --prefix "2025_" --suffix "_final"
  
  # 詳細輸出
  python md2word.py batch-templates data.md ./templates/ ./outputs/ -v
```

**使用情境：**
- 同一份資料需要產生不同格式的報告（例如：簡報摘要、詳細報告、客戶版本）
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

### 基本使用

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

### 一行完成

```python
renderer = WordRenderer()
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
from md_word_renderer import MarkdownParser, WordRenderer
from docx.shared import Cm

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

# 尋找測試案例中的圖片
test_cases = data.get('異動內容-測試案例', [])
find_images(test_cases)
```

### 渲染含圖片的資料

```python
from md_word_renderer import WordRenderer
from docx.shared import Cm

# 建立渲染器，設定圖片尺寸
renderer = WordRenderer(image_width=Cm(15), image_height=None)

# 渲染資料（圖片會自動處理）
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

### 批次處理腳本

```python
from pathlib import Path
from md_word_renderer import MarkdownParser, WordRenderer

parser = MarkdownParser()
renderer = WordRenderer()

input_dir = Path('inputs')
output_dir = Path('outputs')
output_dir.mkdir(exist_ok=True)

for md_file in input_dir.glob('*.md'):
    data = parser.parse(str(md_file))
    output_file = output_dir / f"{md_file.stem}.docx"
    renderer.render_to_file(data, 'template.docx', str(output_file))
    print(f"已處理: {md_file.name} -> {output_file.name}")
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

如有問題，請提交 Issue 或聯絡開發團隊。
