# 實作計劃 - MD→Excel 渲染支援（v2.2）

> 本文件記錄 v2.2 Excel 渲染支援的完整實作規劃。
> 對應決策日期：2025-12；對應版本：`v2.2.0`。

## 一、背景與目標

### 目標
在既有 **MD→Word** 渲染管線之外，增加 **MD→Excel** 渲染支援，使用者能用同一份 Markdown 資料 (`referance/sample_data.md`) 與 `.xlsx` 樣板產出對應的活頁簿。

### 範圍
- 新增 `ExcelRenderer`、`ExcelImageHandler`、`LayoutConfig`、`build_renderer` 工廠。
- CLI 既有 `render / batch / batch-templates` 加 `--format {docx,xlsx,auto}` 旗標；GUI 自動依樣板副檔名偵測格式。
- 提供樣板產生腳本與 `templates/excel/*.xlsx`。
- 既有功能（`MarkdownParser`、`SchemaValidator`、Word 路徑、47 個既有測試）**完全不退步**。

### 非目標
- Excel 公式、樞紐分析、圖表、條件式格式。
- 動態清單欄位行內任意 Jinja2 DSL。
- Excel 內嵌 Word、PDF 等其他格式。

## 二、可行性與決策彙總

### 可重用（不需改動）
| 元件 | 路徑 | 評估 |
|---|---|---|
| `MarkdownParser` | `src/md_word_renderer/parser/markdown_parser.py` | 完全可重用 |
| `EscapeHandler` | `src/md_word_renderer/parser/escape_handler.py` | 完全可重用 |
| `SchemaValidator` | `src/md_word_renderer/validator/schema_validator.py` | 完全可重用 |
| `ConfigLoader` | `src/md_word_renderer/config/config_loader.py` | 可重用；新增 `excel:` 區段 |
| argparse 參數解析骨架 | `src/md_word_renderer/cli/main.py` | 結構可直接擴充 |

### 設計決策（已與專案負責人確認）
| 議題 | 決定 |
|---|---|
| 樣板策略 | **C**：樣板 + 程式生成（renderer 載入 `.xlsx`，迴圈/清單程式化展開） |
| CLI 結構 | 既有 `render/batch/batch-templates` 加 `--format {docx,xlsx,auto}`，預設 `auto` |
| 清單展開策略 | 每個頂層 list 欄位 → 各自一張 sheet；純量欄位 → 第一張「基本資訊」sheet |
| GUI 行為 | 不加 dropdown，**自動依樣板副檔名**偵測 |
| 樣板資產 | 加 `scripts/create_excel_templates.py` 並產出 `templates/excel/*.xlsx` |
| 版本號 | `2.2.0`（延續 README `v2.1 圖片插入`） |

## 三、檔案改動清單

### 新增檔案
```
src/md_word_renderer/renderer/excel_renderer.py
src/md_word_renderer/renderer/excel_image_handler.py
src/md_word_renderer/renderer/excel_layout.py
src/md_word_renderer/renderer/factory.py
test/test_excel_renderer.py
test/test_excel_factory.py
test/test_cli_excel.py
scripts/create_excel_templates.py
templates/excel/sample_template.xlsx          # 由腳本產生
templates/excel/with_lists_template.xlsx      # 由腳本產生
CHANGELOG.md                                  # 新增
doc/implementation_plan_excel.md              # 本文件
```

### 修改檔案
```
src/md_word_renderer/__init__.py              # __version__ = "2.2.0" + 加入 ExcelRenderer
src/md_word_renderer/renderer/__init__.py     # + ExcelRenderer 重匯出
src/md_word_renderer/utils/batch_processor.py # output_pattern 解耦副檔名
src/md_word_renderer/utils/file_utils.py      # generate_output_path 接受 extension
src/md_word_renderer/cli/main.py              # --format 旗標；抽出 process_one
src/md_word_renderer/gui/main_window.py       # 樣板 filter 擴充；輸出格式自動
src/md_word_renderer/gui/batch_window.py      # 同上
src/md_word_renderer/gui/multi_template_window.py
src/md_word_renderer/gui/template_preview.py  # 抽出共用「列出可用變數」邏輯
src/md_word_renderer/config/default_config.yaml # + excel: 區段
config.yaml                                    # 範例同步
requirements.txt                               # + openpyxl>=3.1,<4
md_word_renderer.spec                          # hiddenimports + openpyxl
md_word_renderer_cli.spec                      # hiddenimports + openpyxl
test_e2e.py                                    # + 2 個 Excel 案例
.gitignore                                     # !templates/excel/*.xlsx 白名單
README.md                                      # + Excel 區塊 + v2.2 開發日誌
AGENTS.md                                      # 測試段落補 test_excel_*.py
doc/README.md                                  # 索引更新（指向新計畫與新增 CHANGELOG）
```

### 不動
- `src/md_word_renderer/parser/*`（輸出 dict 與渲染端無關）
- `src/md_word_renderer/validator/schemas/markdown_schema.json`
- 既有 CLI/GUI 對外命令與介面（向後相容）

## 四、模組設計

### `excel_layout.py`
- 提供 `LayoutConfig` dataclass：包含 `default_template`、`list_sheet_naming`、`header_sheet`、`extra_columns`、`image`、`auto_fit_columns` 等欄位。
- 提供 `LayoutConfig.from_mapping(dict)` / `from_yaml(path)` 載入設定。
- 提供 `sanitize_sheet_name(name: str)`：去除 Excel sheet 名禁用字元 `/ \ ? * [ ] :`。

### `excel_image_handler.py`
- `ExcelImageHandler.embed(worksheet, cell_ref, image_path, max_width, max_height)`：
  - 使用 `openpyxl.drawing.image.Image` 載入圖片
  - 等比縮放後寫入指定儲存格 anchor。
- 若圖片檔不存在則拋 `ImageNotFoundError`，呼叫端決定是否略過。

### `excel_renderer.py`
- `ExcelRenderer` 對齊 `WordRenderer` 介面：`load_template() / render() / save() / render_to_file()`。
- `render(data)` 流程：
  1. 載入樣板（沒有 LAYOUT metadata sheet 時使用 `LayoutConfig` 預設）。
  2. 取得 header sheet：寫入所有「純量頂層欄位」，兩欄 `欄位 | 值`。
  3. 對每個 `type == list` 的頂層欄位：
     - 樣板有同名 sheet 時當作骨架（保留樣式）；否則新建。
     - 標題列：`項目編號 | 值 | 類型` + `extra_columns`
     - DFS 攤平 `children`：每個葉節點寫入一列。
     - 葉節點型別為 `image` → 走 `ExcelImageHandler.embed`。
  4. 自動欄寬（若 `auto_fit_columns=True`）。
- `_prepare_context(data)`：與 WordRenderer 相同方式注入 `data` 鍵，方便樣板用 `{{data["key"]}}` 形式。

### `factory.py`
```python
def build_renderer(template_path: str, format_hint: str = "auto"):
    # auto：依副檔名；docx→WordRenderer；xlsx→ExcelRenderer
    # docx/xlsx：明確指定
```

## 五、CLI 與 GUI

### CLI 旗標
- `render <in> <tpl> <out> [--format {docx,xlsx,auto}] [--no-validate]`
- `batch <in_dir> <tpl> <out_dir> [--format ...] [--continue-on-error]`
- `batch-templates <in> <tpl_dir> <out_dir> [--format ...] [--prefix ...] [--suffix ...]`
- 預設 `auto`：用 `Path(template).suffix.lower()` 推斷。
- 共用函式 `process_one(parser, renderer_factory, in_path, out_path, ...)`。

### GUI
- 三個 window 的「選擇樣板」篩選改為 `Word 模板 (*.docx) ;; Excel 樣板 (*.xlsx)`。
- 「輸出格式」label 由 sample 樣板路徑的副檔名動態決定（不在 UI 暴露切換）。
- `multi_template_window` 支援同資料夾內混雜 `.docx/.xlsx`，逐檔 dispatch。

## 六、設定

### `src/md_word_renderer/config/default_config.yaml` 新增
```yaml
excel:
  default_template: templates/excel/sample_template.xlsx
  list_sheet_naming: "{key}"
  header_sheet:
    enabled: true
    name: "基本資訊"
  extra_columns: [source_field, path, depth]
  image:
    max_width_px: 480
    max_height_px: 360
  auto_fit_columns: true
```

## 七、樣板產生腳本

### `scripts/create_excel_templates.py`
- 用 `openpyxl` 產：
  - `templates/excel/sample_template.xlsx`：基本欄位列（`基本資訊`） + 標題樣式。
  - `templates/excel/with_lists_template.xlsx`：含 `LAYOUT` metadata sheet，定義 header 與 list sheet 配置。
- 命令：`python scripts/create_excel_templates.py`，互動式 overwrite 確認。

## 八、測試計劃

| 檔案 | 案例數 | 重點 |
|---|---|---|
| `test/test_excel_layout.py`（可選） | ≥ 3 | sanitize sheet name、預設值、from_mapping |
| `test/test_excel_renderer.py` | ≥ 10 | 載入樣板、文字替換、children 攤平、nested ≥ 3 層、圖片嵌入、空值、超長字串、auto-fit、缺欄位走 `error_format`、重複 sheet 名容錯 |
| `test/test_excel_factory.py` | ≥ 2 | `xlsx` / `docx` / 不支援副檔名拋錯 |
| `test/test_cli_excel.py` | ≥ 3 | `--format xlsx` / `--format auto`（混合副檔名樣板）/ exit code |
| `test_e2e.py` 加 2 案 | 2 | round-trip 跑完產出 `.xlsx` 至 `output/` |

**驗收目標**：
- 既有 47 個測試 100% 保留綠。
- 新增 ≥ 15 個 Excel 測試全綠。
- `python test_e2e.py` 全綠，產出 `.xlsx` 可用 Excel/LibreOffice 開啟。

## 九、執行順序

1. `pip install openpyxl` + 更新 `requirements.txt`
2. `excel_layout.py`（先寫最小單元測試）
3. `excel_image_handler.py`
4. `excel_renderer.py` + 單元測試
5. `factory.py` + 單元測試
6. CLI 重構：抽 `process_one`、加 `--format`；跑既有 47 測試不退步
7. GUI 微調：樣板 filter、格式 label 從樣板路徑算
8. `scripts/create_excel_templates.py` → 產 `templates/excel/*.xlsx`
9. 補測試至 ≥ 15 案 + `test_e2e.py` 加 2 案
10. 更新 `.spec` 的 `hiddenimports`、驗 `build_exe.py --version both`
11. 文件：`README.md`、`doc/template_syntax_reference.md`、`CHANGELOG.md`、`doc/README.md`、`AGENTS.md`、`.gitignore`
12. 版號 bump → `2.2.0`；驗收：`pytest test/ -v` + `test_e2e.py` + GUI round-trip

## 十、相依與打包

- 新增依賴 `openpyxl>=3.1,<4`（read + write + image）。
- `Pillow` 已被 `requirements-gui.txt` 拉入；GUI 版打包尺寸已含，CLI 版打包需確認 `Pillow` 是否要隨 openpyxl 進入（一般無）。
- `md_word_renderer.spec`、`md_word_renderer_cli.spec` 的 `hiddenimports` 補：
  - `md_word_renderer.renderer.excel_renderer`
  - `md_word_renderer.renderer.excel_image_handler`
  - `md_word_renderer.renderer.excel_layout`
  - `md_word_renderer.renderer.factory`
  - `openpyxl`、`openpyxl.drawing.image`
- `excludes` 不應擋 `openpyxl` / `Pillow`。

## 十一、回歸驗證 checklist

- [ ] `python -m pytest test/ -v` 全綠（既有 47 + 新增 ≥ 15）
- [ ] `python test_e2e.py` 全綠（Word 5 案 + Excel 2 案）
- [ ] `python md2word.py --format auto ...` 行為正確
- [ ] CLI exe 內含 `excel` 等命令：`dist\md2word.exe --help`
- [ ] GUI 三個 window 對 `.xlsx` 樣板 round-trip 成功

## 十二、文件更新

- `README.md`：在「專案狀態」加 `v2.2 ✅`、在「開發日誌」加 `### v2.2.0` 條目、新增「Excel 渲染」章節範例。
- `doc/template_syntax_reference.md`：新增「Excel 樣板語意」章節。
- `CHANGELOG.md`：從無到有；列出 v2.0 / v2.1 / v2.2 重點差異。
- `AGENTS.md`：測試段落增加 Excel 測試檔案清單。
- `doc/README.md`：索引更新（含 `implementation_plan_excel.md` 與 `CHANGELOG.md`）。

---

**維護者**：ytssamuel
**最後更新**：2025-12
