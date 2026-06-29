# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包規格檔

用於將 MD-Word/Excel Template Renderer GUI 打包為獨立執行檔
"""

import sys
from pathlib import Path

# 專案根目錄
ROOT_DIR = Path(SPECPATH)
SRC_DIR = ROOT_DIR / 'src'

block_cipher = None

# 收集所有資料檔案
datas = [
    # 文件
    (str(ROOT_DIR / 'doc'), 'doc'),
    (str(ROOT_DIR / 'README.md'), '.'),
    # 應用程式圖標
    (str(ROOT_DIR / 'assets'), 'assets'),
    # 範例模板（如果有）
    # (str(ROOT_DIR / 'templates'), 'templates'),
]

# 隱藏的匯入
hiddenimports = [
    'customtkinter',
    'darkdetect',
    'docxtpl',
    'jinja2',
    'docx',
    'lxml',
    'lxml._elementpath',
    'PIL',
    'PIL._tkinter_finder',
    'yaml',
    'jsonschema',
    # Excel 渲染支援 (v2.2)
    'openpyxl',
    'openpyxl.cell',
    'openpyxl.drawing',
    'openpyxl.drawing.image',
    'openpyxl.styles',
    'openpyxl.utils',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'et_xmlfile',
    'md_word_renderer',
    'md_word_renderer.parser',
    'md_word_renderer.parser.markdown_parser',
    'md_word_renderer.renderer',
    'md_word_renderer.renderer.word_renderer',
    'md_word_renderer.renderer.error_handler',
    'md_word_renderer.renderer.excel_renderer',
    'md_word_renderer.renderer.excel_image_handler',
    'md_word_renderer.renderer.excel_layout',
    'md_word_renderer.renderer.factory',
    'md_word_renderer.validator',
    'md_word_renderer.validator.schema_validator',
    'md_word_renderer.gui',
    'md_word_renderer.gui.main_window',
    'md_word_renderer.gui.batch_window',
    'md_word_renderer.gui.multi_template_window',
    'md_word_renderer.gui.settings_window',
    'md_word_renderer.gui.config_manager',
    'md_word_renderer.gui.error_handler',
    'md_word_renderer.gui.template_preview',
]

# 排除的模組（減少檔案大小）
excludes = [
    'matplotlib',
    'numpy',
    'pandas',
    'scipy',
    'IPython',
    'jupyter',
    'notebook',
    'pytest',
    'setuptools',
    'wheel',
    'pip',
]

a = Analysis(
    ['run_gui.py'],
    pathex=[str(SRC_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MD-Word-Renderer',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 模式，不顯示控制台
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT_DIR / 'assets' / 'app_icon.ico'),  # 應用程式圖標
)
