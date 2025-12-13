# 開發準備 - Phase 1 啟動清單

## ✅ 已完成的準備工作

### 文件
- ✅ 需求規劃文件（template_renderer_spec.md）
- ✅ 實作計劃文件（implementation_plan.md v2.0）
- ✅ README.md 框架
- ✅ 測試資料（example/OH_20251020.md）

### 設計決策（基於討論結果）
- ✅ 模板語法：`{{變數}}`（Jinja2）
- ✅ 模板引擎：docxtpl
- ✅ 點記法存取：支援
- ✅ 存取方式：名稱和編號雙重支援
- ✅ 空值處理：顯示空白
- ✅ 縮排處理：參考 Python PEP 8
- ✅ 特殊字元：反斜線轉義
- ✅ 錯誤處理：顯示 `[ERROR: 變數不存在]`
- ✅ 資料驗證：支援（JSON Schema）
- ✅ 多檔案處理：支援
- ✅ 設定檔：支援（config.yaml）

### 專案規劃
- ✅ 技術棧選擇
- ✅ 系統架構設計
- ✅ 專案結構規劃
- ✅ 模組規格定義
- ✅ 開發階段劃分
- ✅ 時程規劃（3.5-4 週）
- ✅ 測試策略
- ✅ 驗收標準

## 🚀 接下來的步驟

### Step 1: 環境建置（第 1 天）

```bash
# 1. 建立專案目錄結構
mkdir -p md_word_renderer/{parser,renderer,validator/schemas,config,utils,templates,tests/fixtures,examples}
touch md_word_renderer/__init__.py
touch md_word_renderer/parser/__init__.py
touch md_word_renderer/renderer/__init__.py
touch md_word_renderer/validator/__init__.py
touch md_word_renderer/config/__init__.py
touch md_word_renderer/utils/__init__.py

# 2. 建立 requirements.txt
cat > requirements.txt << 'EOF'
python-docx>=0.8.11
docxtpl>=0.16.7
Jinja2>=3.1.2
PyYAML>=6.0
jsonschema>=4.17.0
pytest>=7.0.0
pytest-cov>=4.0.0
black>=22.0.0
flake8>=5.0.0
EOF

# 3. 建立虛擬環境
python -m venv venv
source venv/bin/activate  # macOS/Linux

# 4. 安裝依賴
pip install -r requirements.txt

# 5. 初始化 Git
git init
cat > .gitignore << 'EOF'
# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
venv/
env/
ENV/

# IDE
.vscode/
.idea/
*.swp
*.swo

# 測試
.pytest_cache/
.coverage
htmlcov/

# 輸出
output/
*.docx
!templates/*.docx
!examples/*.docx

# 日誌
logs/
*.log

# 暫存
.DS_Store
Thumbs.db
EOF

git add .
git commit -m "chore: 初始化專案結構"
```

### Step 2: 開始開發（第 2 天起）

依照 implementation_plan.md 中的順序：

1. **Week 1 (5 天)**
   - Day 1: 環境建置
   - Day 2-3: Markdown 解析器（縮排偵測、特殊字元）
   - Day 4-5: Markdown 解析器（階層建立）+ 資料驗證

2. **Week 2 (5 天)**
   - Day 6-7: Word 渲染器（基本功能）
   - Day 8: CLI 基本版
   - Day 9-10: 批次處理功能

3. **Week 3 (3 天)**
   - Day 11: 設定檔支援
   - Day 12-13: Word 模板範例製作

4. **Week 4 (3-5 天)**
   - Day 14-16: 整合測試
   - Day 17-18: 文件撰寫

## 📋 開發檢查清單

### 環境（必須在開始前完成）
- [ ] Python 3.8+ 已安裝並驗證
- [ ] 虛擬環境已建立並啟動
- [ ] 所有依賴套件已安裝
- [ ] Git 已初始化
- [ ] 專案結構已建立
- [ ] .gitignore 已設定

### 準備（建議完成）
- [ ] 已閱讀 implementation_plan.md
- [ ] 已理解 docxtpl 基本用法
- [ ] 已理解 Jinja2 模板語法
- [ ] 已準備 Word 軟體（製作模板用）
- [ ] 已設定 IDE/編輯器
- [ ] 已設定 Black 和 Flake8

### 開發工具
- [ ] VS Code 或其他 Python IDE
- [ ] Python 擴充套件
- [ ] Git 客戶端
- [ ] Microsoft Word 或 LibreOffice

## 🎯 開發重點提醒

### 編碼規範
- 遵循 PEP 8
- 使用 Type Hints
- 撰寫 Docstring（Google Style）
- 保持函數簡潔（< 50 行）
- 單一職責原則

### Git 規範
- Commit 訊息格式：`<type>(<scope>): <subject>`
- 類型：feat, fix, docs, style, refactor, test, chore
- 每個功能一個分支
- 通過測試後才合併

### 測試要求
- 單元測試涵蓋率 > 85%
- 每個模組都要有測試
- 先寫測試再寫實作（TDD）
- 測試要有意義的名稱

### 效能考量
- 大檔案使用流式處理
- 避免重複讀取檔案
- 快取可重用的資料
- 注意記憶體使用

## 📚 參考資源

### 官方文件
- docxtpl: https://docxtpl.readthedocs.io/
- Jinja2: https://jinja.palletsprojects.com/
- python-docx: https://python-docx.readthedocs.io/
- PyYAML: https://pyyaml.org/
- jsonschema: https://python-jsonschema.readthedocs.io/

### 學習資源
- PEP 8: https://www.python.org/dev/peps/pep-0008/
- Git Commit 規範: https://www.conventionalcommits.org/
- Pytest 教學: https://docs.pytest.org/

## 💡 開發建議

### 先易後難
1. 從簡單的單檔案處理開始
2. 先實作基本功能，再加進階功能
3. 每完成一個功能就寫測試
4. 逐步增加複雜度

### 持續整合
- 每天都要 commit
- 功能完成後立即測試
- 發現問題立即修復
- 保持程式碼整潔

### 文件同步
- 程式碼和註解同步更新
- API 變更時更新文件
- 記錄重要的設計決策
- 維護 CHANGELOG

## ✨ 成功標準

Phase 1 完成時應該達成：

- ✅ 可以解析 OH_20251020.md
- ✅ 可以渲染 Word 模板
- ✅ 支援所有定義的語法
- ✅ 支援多檔案批次處理
- ✅ 支援設定檔
- ✅ 支援資料驗證
- ✅ CLI 完整可用
- ✅ 測試涵蓋率 > 85%
- ✅ 有完整的使用文件
- ✅ 有 3 個模板範例

## 🎉 準備開始！

所有準備工作已完成，現在可以：

```bash
# 1. 確認環境
python --version  # 應該 >= 3.8
pip list  # 檢查套件

# 2. 執行第一個測試
pytest tests/  # 應該看到 0 tests

# 3. 開始開發第一個模組
# 建立 parser/indent_detector.py
# 撰寫測試 tests/test_indent_detector.py
# 實作功能
# 執行測試確認通過
```

**預祝開發順利！有任何問題請參考 implementation_plan.md 📖**

---

**文件建立日期：** 2025-12-12  
**狀態：** ✅ 準備就緒，可以開始開發
