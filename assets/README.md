# 應用程式圖標

## 設計概念

MD-Word Template Renderer 的應用程式圖標結合了 Markdown 和 Word 文檔的視覺元素。

### 設計元素

```
┌─────────────────┐
│   MD  →  W     │  
│  (藍色漸層背景)  │
└─────────────────┘
```

- **左側 "MD"**: 代表 Markdown 輸入
- **中間箭頭**: 表示轉換過程
- **右側 "W"**: 代表 Word 文檔輸出
- **藍色漸層背景**: 專業、科技感的配色
- **圓角矩形**: 代表文檔外框

## 檔案說明

| 檔案 | 用途 | 尺寸 |
|------|------|------|
| `app_icon.ico` | Windows 應用程式圖標 | 16×16 到 256×256 多尺寸 |
| `app_icon.png` | 預覽圖、文檔用圖標 | 256×256 |

## 重新生成圖標

如需修改圖標設計，可以編輯並執行專案根目錄的 `create_icon.py`：

```bash
python create_icon.py
```

## 在 PyInstaller 中使用

圖標已經整合到 `md_word_renderer.spec` 打包配置中：

```python
icon=str(ROOT_DIR / 'assets' / 'app_icon.ico')
```

執行 `pyinstaller md_word_renderer.spec` 時會自動使用此圖標。

## 設計規格

- **主色調**: 藍色系 (#2962FF)
- **強調色**: 白色 (#FFFFFF)
- **風格**: 扁平化、現代設計
- **格式**: ICO (多尺寸), PNG (高解析度)
