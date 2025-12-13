#!/usr/bin/env python
"""
PyInstaller 打包腳本

自動化打包流程，生成獨立執行檔
"""

import subprocess
import sys
import shutil
from pathlib import Path


def clean_build():
    """清理先前的打包產物"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"清理 {dir_name}/ ...")
            shutil.rmtree(dir_path)


def install_pyinstaller():
    """安裝 PyInstaller"""
    print("檢查 PyInstaller...")
    try:
        import PyInstaller
        print(f"PyInstaller 已安裝: v{PyInstaller.__version__}")
    except ImportError:
        print("安裝 PyInstaller...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])


def build_exe():
    """執行打包"""
    print("\n開始打包...")
    print("=" * 50)
    
    # 使用 spec 檔案打包
    spec_file = Path(__file__).parent / 'md_word_renderer.spec'
    
    if spec_file.exists():
        cmd = [sys.executable, '-m', 'PyInstaller', str(spec_file), '--clean']
    else:
        # 如果沒有 spec 檔案，使用命令列參數
        cmd = [
            sys.executable, '-m', 'PyInstaller',
            '--name=MD-Word-Renderer',
            '--onefile',
            '--windowed',
            '--clean',
            f'--paths={Path(__file__).parent / "src"}',
            '--hidden-import=customtkinter',
            '--hidden-import=darkdetect',
            '--hidden-import=docxtpl',
            '--hidden-import=jinja2',
            '--hidden-import=docx',
            '--hidden-import=lxml',
            '--hidden-import=PIL',
            '--hidden-import=yaml',
            '--hidden-import=jsonschema',
            'run_gui.py'
        ]
    
    print(f"執行命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    return result.returncode == 0


def verify_build():
    """驗證打包結果"""
    exe_path = Path('dist/MD-Word-Renderer.exe')
    
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        print("\n" + "=" * 50)
        print("✅ 打包成功!")
        print(f"   執行檔: {exe_path.absolute()}")
        print(f"   檔案大小: {size_mb:.1f} MB")
        return True
    else:
        print("\n❌ 打包失敗: 找不到執行檔")
        return False


def main():
    """主程式"""
    print("=" * 50)
    print("MD-Word Template Renderer - 打包工具")
    print("=" * 50)
    
    # 1. 清理
    clean_build()
    
    # 2. 確保 PyInstaller 已安裝
    install_pyinstaller()
    
    # 3. 打包
    success = build_exe()
    
    # 4. 驗證
    if success:
        verify_build()
    
    print("\n完成!")
    return 0 if success else 1


if __name__ == '__main__':
    sys.exit(main())
