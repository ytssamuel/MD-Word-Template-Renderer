#!/usr/bin/env python
"""
MD-Word Template Renderer GUI 啟動器

直接執行此腳本啟動 GUI 應用程式
"""

import sys
from pathlib import Path

# 將 src 目錄加入 Python 路徑
src_dir = Path(__file__).parent / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from md_word_renderer.gui import main


if __name__ == "__main__":
    main()
