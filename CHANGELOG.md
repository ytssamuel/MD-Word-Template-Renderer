# Changelog

本檔案記錄 MD-Word/Excel Template Renderer 的版本變更。採用 [Keep a Changelog](https://keepachangelog.com/zh-TW/) 風格。

## [2.2.1] - 2025-12
### 重寫（Breaking Change）
- **Excel 渲染改為「樣板為主」**：不再自動 append 純量欄位；只有 `{{var}}` 出現的 cell 才會被替換
- `ExcelRenderer.render()` 流程重構為「data 寫入 → for 展開 → 模板 pass」

### 新增
- `ExcelTemplateEngine`（Jinja2 包裝）：支援 `{{var}}` 替換、`{% if %}` 條件、`{% for %}` 列展開
- `LayoutConfig.template_engine` 區段：包含 `enabled` / `auto_flatten_lists` / `missing_variable` / `error_format`
- `config.yaml` 新增 `excel.template_engine` 設定
- `{% for %}` 列展開：單層（v2.2.1）；巢狀 children 為 v2.2.2+ 規劃
- `loop.index` / `loop.first` / `loop.last` 在 for body 內可用
- 缺變數策略：`silent`（預設）/ `keep` / `error_format`
- 換行字串在 template pass 後自動設 wrap_text
- 多張 sheet 範本（`templates/excel/with_lists_template.xlsx`）含兩個 list sheet 示範
- 第三個樣板（`with_layout_template.xlsx`）示範 LAYOUT metadata

### 修改
- `sample_template.xlsx` 改為「手動樣板起點」風格（不再預填示範 row）
- `with_lists_template.xlsx` 改為使用 `{% for %}` 語意
- `create_excel_templates.py` 重寫為新模板語意
- `ExcelRenderer._render_header_sheet` 不再 append 純量
- `ExcelRenderer._apply_template_pass_to_all` 新增：對所有 sheet 走 Jinja2 模板 pass
- LAYOUT metadata 支援 `template_engine_enabled` / `auto_flatten_lists` / `missing_variable` / `error_format` 鍵

### 向後相容
- `auto_flatten_lists: true`（預設值）→ 既有 v2.2.0 的「list sheet 自動攤平」行為保留
- 既有 v2.2.0 使用者不需任何改動，pipeline 仍可運作
- 47 → 100 個測試（新增 53 案，含 Phase 1/2/3）

## [2.2.0] - 2025-12
### 新增
- **Excel 渲染支援（md → .xlsx）**
  - `ExcelRenderer`（API 對齊 `WordRenderer`）
  - `ExcelImageHandler`：等比縮放後嵌入工作表
  - `LayoutConfig`：基本資訊工作表、清單展開命名、圖片尺寸限制等設定
  - `build_renderer()` factory：依樣板副檔名自動選擇 `WordRenderer` / `ExcelRenderer`
- CLI：`render / batch / batch-templates` 加 `--format {auto,docx,xlsx}` 旗標（預設 `auto`）
- GUI：三個 window（單檔/批次/多模板）皆可挑選 `.docx` 或 `.xlsx` 樣板，輸出格式自動偵測
- `LAYOUT` 隱藏工作表：樣板內可選 metadata，覆寫 layout 預設
- 範例樣板：`templates/excel/sample_template.xlsx`、`templates/excel/with_lists_template.xlsx`
- 樣板產生腳本：`scripts/create_excel_templates.py`
- 36 個新單元/整合測試（`test_excel_renderer.py`、`test_excel_factory.py`、`test_cli_excel.py`）
- 2 個新端到端測試（`test_e2e.py` 測試 6、7）

### 修改
- `utils/batch_processor.py`：`output_pattern` 拆出 `output_extension`，副檔名由 caller 注入
- `cli/main.py`：抽出 `process_one()` 共用核心流程；`validate / info` 行為不變
- 所有 GUI window：移除硬編碼 `.docx` 篩選，依樣板副檔名動態判斷
- `.gitignore` 補充 `!templates/excel/*.xlsx` 白名單

### 升級注意事項
- **完全向後相容**：原 47 個測試維持綠；既有的 `.docx` 使用者不需任何改動
- 新增依賴 `openpyxl>=3.1,<4`
- `requirements.txt` 與兩個 `.spec` 的 `hiddenimports` 同步更新

## [2.1.0] - 2025-12-14
### 新增
- 圖片插入功能：`![alt](path)` Markdown 語法解析、`image_handler.py`

## [2.0.0] - 2025-12-14
### 新增
- GUI（CustomTkinter）：單檔/批次/多模板三模式
- PyInstaller 打包：`build_exe.py`、`.spec` 設定檔

## [1.0.0] - 2024
### 新增
- Markdown 解析、Word 渲染、Schema 驗證
- CLI、批次處理、Python API
- 43 個單元測試 + 完整文檔
