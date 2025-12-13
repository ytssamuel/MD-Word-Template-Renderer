#!/usr/bin/env python
"""
MD-Word Template Renderer - 命令列執行腳本

Usage:
    python -m md2word render <input.md> <template.docx> <output.docx>
    python -m md2word batch <input_dir> <template.docx> <output_dir>
    python -m md2word validate <input.md>
    python -m md2word info
"""

import sys
from pathlib import Path

# 添加 src 目錄到路徑
src_path = Path(__file__).parent / 'src'
sys.path.insert(0, str(src_path))

from md_word_renderer.cli import main

if __name__ == '__main__':
    main()
