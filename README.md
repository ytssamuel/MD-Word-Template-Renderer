# MD-Word Template Renderer

<p align="center">
  <img src="assets/app_icon.png" alt="MD-Word Renderer Icon" width="128" height="128">
</p>

將結構化的 Markdown 資料渲染至 Word 模板，實現資料驅動的文件生成系統。

## 專案狀態

✅ **v1.0 已完成** - 核心功能穩定運行  
✅ **v2.0 已完成** - GUI 圖形介面全功能上線  
✅ **v2.1 已完成** - 圖片插入功能  
📦 **打包發布** - 獨立執行檔可直接使用

[![Tests](https://img.shields.io/badge/tests-47%20passed-success)](test/)  
[![Python](https://img.shields.io/badge/python-3.10%2B-blue)](https://www.python.org/)  
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

## 功能特色

### 🎯 核心功能 (v1.0)
- ✅ **智慧 Markdown 解析** - 支援階層結構和不規則縮排自動偵測
- ✅ **強大模板引擎** - 基於 Jinja2，支援條件、迴圈等進階語法
- ✅ **資料驗證** - JSON Schema 自動驗證資料完整性
- ✅ **雙重索引** - 支援名稱和編號兩種存取方式
- ✅ **批次處理** - 多 MD + 單模板 / 單 MD + 多模板
- ✅ **CLI 工具** - 完整的命令列介面，適合自動化
- ✅ **測試覆蓋** - 47 個單元測試確保品質

### 🖼️ 圖片插入 (v2.1 NEW)
- ✅ **Markdown 圖片語法** - 支援 `![alt](path)` 標準語法
- ✅ **自動路徑解析** - 相對路徑自動轉換為絕對路徑
- ✅ **Word InlineImage** - 圖片自動嵌入 Word 文件
- ✅ **尺寸控制** - 可自訂圖片寬高
- ✅ **多格式支援** - PNG、JPG、GIF、BMP 等

### 🖥️ 圖形介面 (v2.0)
- ✅ **現代化 GUI** - 基於 CustomTkinter 的美觀介面
- ✅ **三種處理模式** - 單檔轉換 / 批次處理 / 多模板批次
- ✅ **即時預覽** - 資料結構和模板變數即時檢視
- ✅ **進度視覺化** - 批次處理進度條和狀態顯示
- ✅ **設定管理** - 記住常用路徑和設定
- ✅ **錯誤處理** - 友善的錯誤提示和問題診斷
- ✅ **模板預覽** - 快速檢視模板中可用的變數

### 📦 打包發布
- ✅ **獨立執行檔** - 無需 Python 環境，開箱即用
- ✅ **雙版本支援** - GUI 圖形介面 + CLI 命令列
- ✅ **精美圖標** - 現代科技風格的應用程式圖標
- ✅ **自動化腳本** - 一鍵打包 GUI/CLI/Both 版本

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

# 安裝 GUI 依賴 (選用)
pip install -r requirements-gui.txt
```

### 基本使用

#### GUI 圖形介面 (v2.0)

```bash
python run_gui.py
```

#### 命令列工具

```bash
# 單一檔案轉換
python md2word.py render input.md template.docx output.docx

# 批次轉換（多個 MD 檔 + 一個模板 → 多個 Word）
python md2word.py batch ./inputs/ template.docx ./outputs/

# 多模板批次轉換（一個 MD 檔 + 多個模板 → 多個 Word）
python md2word.py batch-templates data.md ./templates/ ./outputs/

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

## 📚 完整文檔

本專案提供詳盡的技術文檔，涵蓋使用指南、開發指南和系統設計：

### 👤 使用者文檔
- 📘 **[使用手冊](doc/user_guide.md)** - 完整的安裝和使用指南
- 📗 **[模板語法參考](doc/template_syntax_reference.md)** - Jinja2 模板語法詳解

### 🛠️ 開發者文檔
- 👨‍💻 **[開發準備](doc/DEV_START.md)** - 環境設定和開發清單
- 🔍 **[需求規劃](doc/template_renderer_spec.md)** - 系統需求和架構設計
- 📋 **[實作計劃](doc/implementation_plan.md)** - 開發階段和模組設計
- 🎨 **[GUI 開發計劃](doc/gui_development_plan.md)** - 圖形介面設計細節

📖 **查看 [文檔目錄](doc/README.md) 了解更多**

---

## Markdown 資料格式

```markdown
# 變更單資料模板

1. 系統名稱 | 範例系統
2. 預定換版日期 | 113年01月15日
3. 變更單號 | CRQ000000100000
16. 異動內容-測試案例 | 
    1. 調整供應商登入之密碼功能
       1. TC001：新增供應商使用者
       2. TC002：超過最長使用設定
    2. 調整檔案解析錯誤文字格式
```

**格式規則：**
- 格式：`編號. 欄位名稱 | 值`
- 階層：使用縮排（空格或 Tab，支援不規則縮排）
- 特殊字元：使用反斜線轉義（`\|`, `\n`）
- 空值：`欄位名稱 |` （值為空）
- 圖片：`![alt文字](圖片路徑)` ✨ NEW

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

### 圖片渲染 ✨ NEW

```jinja2
{# 圖片項目的 type 會是 "image" #}
{% for item in data["異動內容-測試案例"] %}
  {% for child in item.children %}
    {% if child.type == "image" %}
      {{child.image}}  {# 輸出圖片 #}
    {% else %}
      {{child.value}}  {# 輸出文字 #}
    {% endif %}
  {% endfor %}
{% endfor %}
```

## 解析後資料結構

```python
{
    "系統名稱": "範例系統",
    "變更單號": "CRQ000000100000",
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
├── src/md_word_renderer/         # 主要模組
│   ├── parser/                   # Markdown 解析器
│   │   ├── markdown_parser.py    # 主解析器
│   │   ├── indent_detector.py    # 縮排偵測
│   │   └── escape_handler.py     # 特殊字元處理
│   ├── renderer/                 # Word 渲染引擎
│   │   ├── word_renderer.py      # 主渲染器
│   │   ├── image_handler.py      # 圖片處理 ✨ NEW
│   │   └── error_handler.py      # 錯誤處理
│   ├── validator/                # 資料驗證
│   │   └── schema_validator.py   # JSON Schema 驗證
│   ├── cli/                      # 命令列工具
│   │   └── main.py               # CLI 入口
│   ├── gui/                      # 圖形介面 ⭐ NEW
│   │   ├── main_window.py        # 主視窗
│   │   ├── batch_window.py       # 批次處理視窗
│   │   ├── multi_template_window.py  # 多模板視窗
│   │   ├── settings_window.py    # 設定視窗
│   │   ├── config_manager.py     # 設定管理
│   │   ├── error_handler.py      # GUI 錯誤處理
│   │   └── template_preview.py   # 模板預覽
│   └── __init__.py
├── test/                         # 測試套件
│   ├── test_parser.py            # 解析器測試
│   ├── test_renderer.py          # 渲染器測試
│   ├── test_validator.py         # 驗證器測試
│   └── test_cli.py               # CLI 測試
├── doc/                          # 完整技術文檔 📚
│   ├── README.md                 # 文檔導航
│   ├── user_guide.md             # 使用手冊
│   ├── template_syntax_reference.md  # 模板語法
│   ├── DEV_START.md              # 開發指南
│   ├── template_renderer_spec.md # 需求規劃
│   ├── implementation_plan.md    # 實作計劃
│   └── gui_development_plan.md   # GUI 開發計劃
├── assets/                       # 資源檔案 ⭐ NEW
│   ├── app_icon.ico              # 應用程式圖標
│   ├── app_icon.png              # PNG 圖標
│   └── README.md                 # 圖標說明
├── dist/                         # 打包產物 📦
│   ├── MD-Word-Renderer.exe      # GUI 執行檔
│   └── md2word.exe                  # CLI 執行檔
├── run_gui.py                    # GUI 啟動腳本 ⭐ NEW
├── md2word.py                    # CLI 入口點
├── create_icon.py                # 圖標生成腳本 ⭐ NEW
├── build_exe.py                  # 打包腳本 ⭐ NEW
├── md_word_renderer.spec         # GUI 打包配置 ⭐ NEW
├── md_word_renderer_cli.spec     # CLI 打包配置 ⭐ NEW
├── requirements.txt              # 核心依賴
├── requirements-gui.txt          # GUI 依賴 ⭐ NEW
└── README.md                     # 本文件
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
- **CustomTkinter** - GUI 圖形介面
- **Click** - CLI 命令列框架

## 打包為執行檔

本專案支援打包為獨立執行檔，無需安裝 Python 環境即可使用。

### 打包方式

```bash
# 方式 1: 使用 Python 腳本（推薦）
python build_exe.py --version gui   # 只打包 GUI 版本
python build_exe.py --version cli   # 只打包 CLI 版本
python build_exe.py --version both  # 打包兩個版本

# 方式 2: 直接使用 PyInstaller
pyinstaller md_word_renderer.spec       # GUI 版本
pyinstaller md_word_renderer_cli.spec   # CLI 版本
```

### 打包產物

| 版本 | 執行檔名稱 | 說明 | 大小 |
|------|-----------|------|------|
| GUI | `MD-Word-Renderer.exe` | 圖形介面版本，雙擊即可使用 | ~17 MB |
| CLI | `md2word.exe` | 命令列版本，用於腳本自動化 | ~12-15 MB |

### 使用打包後的執行檔

**GUI 版本：**
```bash
# 雙擊執行或命令列啟動
.\dist\MD-Word-Renderer.exe
```

**CLI 版本：**
```bash
# 單一檔案轉換
.\dist\md2word.exe render input.md template.docx output.docx

# 批次轉換
.\dist\md2word.exe batch ./inputs/ template.docx ./outputs/

# 多模板批次
.\dist\md2word.exe batch-templates data.md ./templates/ ./outputs/

# 查看說明
.\dist\md2word.exe --help
```

### 打包配置檔案

- `md_word_renderer.spec` - GUI 版本打包配置
- `md_word_renderer_cli.spec` - CLI 版本打包配置
- `build_exe.py` - 自動化打包腳本
- `assets/app_icon.ico` - 應用程式圖標

## 授權

MIT License

## 開發日誌

### v2.0.0 (2025-12-14) 🎉

**圖形介面完整上線**

- ✅ CustomTkinter 現代化 GUI
- ✅ 三種處理模式（單檔/批次/多模板）
- ✅ 即時資料與模板預覽
- ✅ 進度視覺化和錯誤處理
- ✅ 設定持久化管理
- ✅ 應用程式圖標設計（現代科技風格）
- ✅ PyInstaller 打包配置
- ✅ 獨立執行檔發布（GUI + CLI）
- ✅ 完整技術文檔（6 份文件）

### v1.0.0 (2024)

**核心功能穩定版**

- ✅ Markdown 解析器（含不規則縮排支援）
- ✅ Word 渲染引擎（docxtpl + Jinja2）
- ✅ 資料驗證（JSON Schema）
- ✅ 批次處理（多對一/一對多）
- ✅ 命令列工具（Click）
- ✅ 完整測試套件（43 個測試通過）
- ✅ 範例模板和文檔

---

**開發狀態：** ✅ v2.0 已完成  
**測試結果：** 43 passed ✅  
**文檔狀態：** 📚 完整  
**打包狀態：** 📦 可發布
