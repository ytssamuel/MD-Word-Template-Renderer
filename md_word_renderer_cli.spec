# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包規格檔 - CLI 版本

用於將 MD-Word Template Renderer CLI 打包為命令列執行檔
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
]

# 隱藏的匯入
hiddenimports = [
    'docxtpl',
    'jinja2',
    'docx',
    'lxml',
    'lxml._elementpath',
    'yaml',
    'jsonschema',
    'click',
    'md_word_renderer',
    'md_word_renderer.parser',
    'md_word_renderer.parser.markdown_parser',
    'md_word_renderer.renderer',
    'md_word_renderer.renderer.word_renderer',
    'md_word_renderer.renderer.error_handler',
    'md_word_renderer.validator',
    'md_word_renderer.validator.schema_validator',
    'md_word_renderer.cli',
    'md_word_renderer.cli.main',
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
    'customtkinter',  # CLI 版本不需要 GUI 庫
    'tkinter',
    'PIL',
    'darkdetect',
]

a = Analysis(
    [str('md2word.py')],
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
    name='md2word',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # CLI 模式，顯示控制台
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT_DIR / 'assets' / 'app_icon.ico'),  # 應用程式圖標
)
