# 實作計劃 - v2.2.1 Excel Jinja2 模板引擎

> 本文件記錄 v2.2.1 的實作規劃。重點在把 Excel renderer 從「auto-append 純量欄位」翻轉為「樣板為主 (template-driven)」，並補上完整的 Jinja2 語意（Phase 1+2+3）。
> 對應決策日期：2025-12；對應版本：`v2.2.1`（patch，向後相容 `auto_flatten_lists` 開關）。

## 一、背景與目標

### 目標
v2.2.0 提供的 `ExcelRenderer` 行為是「auto-append 所有純量欄位到 header sheet」，沒處理 `{{...}}` 樣板文字。v2.2.1 改為「樣板為主」，並補上：
- `{{var}}` 與 `{{data["key"]}}` 儲存格替換
- `{% if %}` 條件渲染
- `{% for %}` 列展開（支援巢狀 children）

### 範圍
- 新增 `excel_template_engine.py`（Jinja2 包裝）
- 重構 `ExcelRenderer.render()` 流程
- 重寫 `create_excel_templates.py` 樣板（含 `{{ }}` / `{% for %}` 示範）
- 既有 16 個 `test_excel_renderer.py` 測試需對齊新語意
- 新增 13+ 個測試（Phase 1: 5、Phase 2: 2、Phase 3: 6）

### 非目標
- 動態巢狀 list（3 層以上 children）
- Excel 公式/樞紐/圖表
- 多樣板合併
- Inline `{% extends %}` / `{% block %}` 等進階 Jinja2 語意

## 二、行為矩陣

| 場景 | v2.2.0 | v2.2.1（auto_flatten_lists=true，預設） | v2.2.1（auto_flatten_lists=false） |
|---|---|---|---|
| 樣板 B2 含 `{{系統名稱}}` | 不處理，append 新資料在後 | **替換** B2 為實際值 | 替換 B2 |
| 樣板沒寫但資料有純量欄位 | append 22 列 | **不 append** | 不 append |
| 樣板有 `異動內容` sheet，內無 `{% for %}` | auto-flatten | auto-flatten | **略過**、不處理 |
| 樣板有 `異動內容` sheet，內含 `{% for %}` | auto-flatten | **走 for 展開** | 走 for 展開 |
| 樣板沒 `異動內容` sheet（資料有 list） | 自動建立 | 自動建立 + auto-flatten | **略過** |

### 向後相容
- `auto_flatten_lists: true`（預設）→ v2.2.0 既有使用 `with_lists_template.xlsx` 的 workflow 不變
- 新行為（template-driven）→ 把 `{{ }}` / `{% for %}` 寫進樣板即可啟用

## 三、模組設計

### `excel_template_engine.py`（新）

```python
class ExcelTemplateEngine:
    def __init__(self, missing_variable: str = "", error_format: str = ""): ...
    def render_cell(self, value: str, context: dict) -> str: ...
    def render_sheet(self, sheet, context: dict) -> int: ...  # 回傳替換 cell 數
    def find_for_markers(self, sheet) -> list[ForMarker]: ...
    def expand_for_loops(self, sheet, context: dict, data: dict) -> int: ...
```

#### Jinja2 環境設定
- `Environment(undefined=ChainableUndefined)` — 缺變數時不拋錯，回傳空字串
- 自定 `finalize`：把所有 `None` 轉 `""`
- 自動注入 `data` 鍵（除頂層外，呼叫端也能用 `data["key"]`）
- 安全模式：禁用 `__` 開頭的 attribute access

#### For marker 偵測
- 掃描 A 欄（或所有 cell），找符合 `{% for VAR in LIST %}` pattern 的 cell
- 同樣找 `{% endfor %}` 標記
- 用 stack 配對（支援巢狀）

#### For 展開演算法
1. 線性掃描所有 row，找 `{% for %}` 與 `{% endfor %}` 位置
2. 用 stack 配對（內層 endfor 先配對內層 for）
3. 對每個 for 區段：
   - 取得 LIST 在 data 中的值（list of dicts）
   - clone body rows 為 N 份
   - 每份綁定 VAR → current item，呼叫 `_apply_template_pass` 於那些 row
   - 刪除 marker rows（含 `{% for %}` 與 `{% endfor %}`）

### `excel_renderer.py`（重構）

```python
def render(self, data):
    processed = self._prepare_context(data)
    existing_sheets = set(self.workbook.sheetnames)
    used = set()
    
    # 1. 標頭 sheet：只套模板，不 append
    if self.layout.header_sheet.enabled:
        name = self._resolve_header_sheet_name(used, existing_sheets)
        self._render_header_sheet(processed, name)
        used.add(name)
    
    # 2. 頂層 list
    for key, value in processed.items():
        if key == "data" or not isinstance(value, list):
            continue
        if key in existing_sheets:
            # 樣板已有：先看有沒有 for marker
            sheet = self.workbook[key]
            if self._has_for_marker(sheet):
                self.engine.expand_for_loops(sheet, processed, data)
            elif self.layout.template_engine.auto_flatten_lists:
                self._flatten_into_sheet(sheet, key, value)
            # else：略過
        elif self.layout.template_engine.auto_flatten_lists:
            name = self.layout.resolve_list_sheet_name(key, used)
            self._render_list_sheet(name, key, value)
            used.add(name)
    
    # 3. 對所有 sheet 套模板 pass
    for sheet in self._iter_data_sheets():
        self.engine.render_sheet(sheet, processed)
```

`_render_header_sheet` 簡化為「不寫入新資料，只確保 header 列存在」。所有純量替換交給 `engine.render_sheet` 在最後一步處理。

## 四、Phase 1+2+3 細節

### Phase 1 — Cell substitution
- `{{var}}` / `{{data["key"]}}` / `{{#1.key}}`
- 非字串 cell（數字、bool、None、日期）跳過
- 缺變數 → 空字串（`ChainableUndefined`）
- cell 內含 `{{` 才處理（節省時間）

### Phase 2 — `{% if %}` 條件
- 完整 Jinja2 if 語法
- false 時 cell 變空（自然結果）

### Phase 3 — `{% for %}` 展開
- 單層與巢狀 children 都支援
- 內建變數：`loop.index`, `loop.first`, `loop.last`（Jinja2 預設）
- 列表為空 → 不展開（marker rows 仍刪除）
- 圖片 leaf 節點 → 嵌入處理

### 範例樣板 `with_lists_template.xlsx`
```
基本資訊 sheet：
A1: 欄位       B1: 值
A2: 系統名稱   B2: {{系統名稱}}
A3: 變更單號   B3: {{變更單號}}

異動內容-測試案例 sheet：
A1: 編號       B1: 內容         C1: 路徑
A2: {% for case in test_cases %}
A3: {{case.number}}   B3: {{case.value}}   C3: {{case.path}}
A4:   {% for c in case.children %}
A5:   {{c.number}}   B5: {{c.value}}
A6:   {% endfor %}
A7: {% endfor %}
```

## 五、設定

`default_config.yaml` 新增：
```yaml
excel:
  default_template: templates/excel/sample_template.xlsx
  ...
  template_engine:
    enabled: true
    auto_flatten_lists: true   # 向後相容
    missing_variable: ""      # silent
```

`LayoutConfig` 新增 `template_engine: TemplateEngineConfig` dataclass。

## 六、檔案改動清單

### 新增
- `src/md_word_renderer/renderer/excel_template_engine.py`
- `doc/implementation_plan_excel_v2_2_1.md`（本檔）

### 修改
- `src/md_word_renderer/renderer/excel_layout.py` — 新增 `TemplateEngineConfig`、從 mapping 載入
- `src/md_word_renderer/renderer/excel_renderer.py` — 重構 render() 流程、新增 helper
- `src/md_word_renderer/config/default_config.yaml` — 新增 template_engine 區段
- `config.yaml` — 同步
- `scripts/create_excel_templates.py` — 重寫兩個 template
- `test/test_excel_renderer.py` — 修改既有測試、新增 13 個
- `test_e2e.py` — 調整 Excel 案預期
- `doc/template_syntax_reference.md` — 第 10 章重寫
- `CHANGELOG.md` — v2.2.1 條目
- `AGENTS.md` — 補充
- `README.md` — 開發日誌
- `src/md_word_renderer/__init__.py` — `__version__ = "2.2.1"`

### 不動
- parser / validator（schema 鎖的是 parser 輸出）
- Markdown 格式語法
- Word renderer / CLI 架構
- spec 檔（Jinja2 早就在 hiddenimports）

## 七、測試矩陣

| 檔案 | 案件數 | 內容 |
|---|---|---|
| `test_excel_renderer.py` 修改 | ~6 案 | 改既有測試符合 template-driven 模型 |
| `test_excel_renderer.py` 新增 | 5 案 | Phase 1 cell substitution |
| `test_excel_renderer.py` 新增 | 2 案 | Phase 2 if |
| `test_excel_renderer.py` 新增 | 6 案 | Phase 3 for |
| `test_excel_factory.py` | 0 案 | 不變 |

**總計**：原本 16 案 → 改寫後約 23 案（覆蓋 1+2+3）+ 整體 test/ 維持 ≥ 90 案綠燈。

## 八、執行順序

1. `excel_template_engine.py`（Phase 1+2+3 核心）
2. `LayoutConfig` 加 `template_engine` 設定
3. `ExcelRenderer.render()` 重構 + helper
4. `default_config.yaml` 加區段
5. 跑既有測試，修改預期
6. `create_excel_templates.py` 重寫，產新 sample
7. 新增 Phase 1+2+3 測試
8. `test_e2e.py` 調整
9. 文件更新（CH10、CHANGELOG、AGENTS、README）
10. bump `__version__ = "2.2.1"`
11. `pytest test/ -v` + `test_e2e.py` + `build_exe.py --version cli` 全綠
12. 實際 round-trip 驗證

## 九、回歸驗證 checklist

- [ ] `python -m pytest test/ -v` 全綠（≥ 90 案）
- [ ] `python test_e2e.py` 全綠（5 Word + 2 Excel）
- [ ] `python md2word.py render data.md sample_template.xlsx out.xlsx` 產出 `.xlsx` 內 `{{ }}` 已被替換
- [ ] `python md2word.py render data.md with_lists_template.xlsx out.xlsx` 產出 `.xlsx` 內 `{% for %}` 區段已展開
- [ ] 既有使用者（v2.2.0 用 `with_lists_template`）走 `auto_flatten_lists=true` 預設仍可正常運作
- [ ] CLI exe 重新打包後 `dist\md2word.exe --format xlsx` 正常

## 十、文件更新

- `doc/template_syntax_reference.md` — 第 10 章大改寫
- `CHANGELOG.md` — 新增 v2.2.1 條目
- `AGENTS.md` — 補一句 v2.2.1 起 Excel 走 Jinja2
- `README.md` — 開發日誌 v2.2.1
- `doc/README.md` — 索引（若有需要）

---

**維護者**：ytssamuel
**最後更新**：2025-12
