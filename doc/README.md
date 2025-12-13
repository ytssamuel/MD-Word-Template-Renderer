# 文檔目錄

本目錄包含 MD-Word Template Renderer 專案的完整技術文檔。

## 📚 文檔索引

### 👤 使用者文檔

| 文件 | 說明 | 適用對象 |
|------|------|----------|
| [user_guide.md](user_guide.md) | **使用手冊** - 完整的使用指南，包含安裝、快速入門、CLI/GUI 使用說明 | 📘 終端使用者、初學者 |
| [template_syntax_reference.md](template_syntax_reference.md) | **模板語法參考** - Word 模板中 Jinja2 語法的詳細說明 | 📗 模板設計者 |

### 🛠️ 開發者文檔

| 文件 | 說明 | 適用對象 |
|------|------|----------|
| [DEV_START.md](DEV_START.md) | **開發準備** - 專案啟動清單、技術決策、開發環境設定 | 👨‍💻 新進開發者 |
| [template_renderer_spec.md](template_renderer_spec.md) | **需求規劃** - 系統需求、資料格式分析、模板語法設計 | 🔍 系統架構師、產品經理 |
| [implementation_plan.md](implementation_plan.md) | **實作計劃** - 開發階段、模組設計、時程規劃 | 📋 專案管理者、開發者 |
| [gui_development_plan.md](gui_development_plan.md) | **GUI 開發計劃** - 圖形介面的設計規劃與實作細節 | 🎨 UI/UX 開發者 |

## 📖 快速導航

### 我想要...

#### 💡 **學習如何使用這個工具**
→ 閱讀 [user_guide.md](user_guide.md)

#### 🎯 **設計 Word 模板**
→ 閱讀 [template_syntax_reference.md](template_syntax_reference.md)

#### 🔨 **參與開發或貢獻程式碼**
→ 從 [DEV_START.md](DEV_START.md) 開始，然後閱讀 [implementation_plan.md](implementation_plan.md)

#### 📐 **了解系統設計理念**
→ 閱讀 [template_renderer_spec.md](template_renderer_spec.md)

#### 🖼️ **開發 GUI 功能**
→ 閱讀 [gui_development_plan.md](gui_development_plan.md)

## 📝 文檔版本

| 文件 | 版本 | 最後更新 |
|------|------|----------|
| user_guide.md | v2.0 | 包含 CLI + GUI 完整功能 |
| template_syntax_reference.md | v1.0 | Jinja2 語法完整參考 |
| template_renderer_spec.md | v1.0 | 初始需求規劃 |
| implementation_plan.md | v2.0 | 包含 GUI 開發計劃 |
| gui_development_plan.md | v1.0 | GUI 功能規劃 |
| DEV_START.md | v1.0 | Phase 1-4 開發清單 |

## 🔗 相關資源

### 專案檔案
- [README.md](../README.md) - 專案主頁
- [CHANGELOG.md](../CHANGELOG.md) - 版本變更記錄（如果有）

### 範例檔案
- `examples/` - 範例 Markdown 資料和模板
- `test/` - 單元測試和測試資料

### 打包相關
- [assets/README.md](../assets/README.md) - 應用程式圖標說明
- `build_exe.py` - 打包自動化腳本

## 💻 技術規格摘要

### 核心功能
- ✅ Markdown 資料解析（支援階層結構）
- ✅ Word 模板渲染（基於 docxtpl + Jinja2）
- ✅ 資料驗證（JSON Schema）
- ✅ 批次處理
- ✅ 命令列工具 (CLI)
- ✅ 圖形介面 (GUI)

### 技術棧
- **語言**: Python 3.10+
- **模板引擎**: docxtpl (Jinja2)
- **Word 處理**: python-docx
- **GUI 框架**: CustomTkinter
- **CLI 框架**: Click
- **驗證**: jsonschema

### 專案結構
```
SpeedBOT/
├── doc/                    # 📁 本目錄 - 所有技術文檔
├── src/                    # 源代碼
│   └── md_word_renderer/   # 主要模組
│       ├── parser/         # Markdown 解析器
│       ├── renderer/       # Word 渲染引擎
│       ├── validator/      # 資料驗證
│       ├── cli/            # 命令列工具
│       └── gui/            # 圖形介面
├── test/                   # 測試套件
├── assets/                 # 圖標等資源
└── examples/               # 範例檔案
```

## 🆘 需要協助？

### 使用問題
1. 查看 [user_guide.md](user_guide.md) 的「疑難排解」章節
2. 檢查 [template_syntax_reference.md](template_syntax_reference.md) 的範例

### 開發問題
1. 查看 [DEV_START.md](DEV_START.md) 的開發準備事項
2. 參考 [implementation_plan.md](implementation_plan.md) 的模組設計

### 系統問題
1. 閱讀 [template_renderer_spec.md](template_renderer_spec.md) 了解設計理念
2. 查看原始碼中的註解和文檔字串

---

**文檔維護者**: MD-Word Template Renderer 開發團隊  
**最後更新**: 2025年12月14日
