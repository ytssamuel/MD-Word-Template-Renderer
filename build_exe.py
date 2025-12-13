#!/usr/bin/env python
"""
PyInstaller 打包腳本

自動化打包流程，生成獨立執行檔
支援 GUI 和 CLI 兩個版本
"""

import subprocess
import sys
import shutil
from pathlib import Path
import argparse


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


def build_exe(version='gui'):
    """執行打包
    
    Args:
        version: 'gui', 'cli' 或 'both'
    """
    print("\n開始打包...")
    print("=" * 50)
    
    versions_to_build = []
    if version == 'both':
        versions_to_build = ['gui', 'cli']
    else:
        versions_to_build = [version]
    
    all_success = True
    
    for ver in versions_to_build:
        print(f"\n>>> 打包 {ver.upper()} 版本...")
        
        # 選擇對應的 spec 檔案
        if ver == 'gui':
            spec_file = Path(__file__).parent / 'md_word_renderer.spec'
        else:  # cli
            spec_file = Path(__file__).parent / 'md_word_renderer_cli.spec'
        
        if spec_file.exists():
            cmd = [sys.executable, '-m', 'PyInstaller', str(spec_file), '--clean']
        else:
            print(f"❌ 找不到 spec 檔案: {spec_file}")
            all_success = False
            continue
        
        print(f"執行命令: {' '.join(cmd)}")
        result = subprocess.run(cmd, capture_output=False)
        
        if result.returncode != 0:
            all_success = False
    
    return all_success


def verify_build(version='gui'):
    """驗證打包結果
    
    Args:
        version: 'gui', 'cli' 或 'both'
    """
    versions_to_check = []
    if version == 'both':
        versions_to_check = ['gui', 'cli']
    else:
        versions_to_check = [version]
    
    all_success = True
    
    for ver in versions_to_check:
        if ver == 'gui':
            exe_path = Path('dist/MD-Word-Renderer.exe')
            name = "GUI 版本"
        else:  # cli
            exe_path = Path('dist/md2word.exe')
            name = "CLI 版本"
        
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print("\n" + "=" * 50)
            print(f"✅ {name}打包成功!")
            print(f"   執行檔: {exe_path.absolute()}")
            print(f"   檔案大小: {size_mb:.1f} MB")
        else:
            print(f"\n❌ {name}打包失敗: 找不到執行檔")
            all_success = False
    
    return all_success


def main():
    """主程式"""
    parser = argparse.ArgumentParser(
        description='MD-Word Template Renderer 打包工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例:
  python build_exe.py              # 打包 GUI 版本 (預設)
  python build_exe.py --version gui   # 打包 GUI 版本
  python build_exe.py --version cli   # 打包 CLI 版本
  python build_exe.py --version both  # 打包兩個版本
        """
    )
    parser.add_argument(
        '--version', '-v',
        choices=['gui', 'cli', 'both'],
        default='gui',
        help='要打包的版本 (預設: gui)'
    )
    
    args = parser.parse_args()
    
    print("=" * 50)
    print("MD-Word Template Renderer - 打包工具")
    print(f"版本: {args.version.upper()}")
    print("=" * 50)
    
    # 1. 清理
    clean_build()
    
    # 2. 確保 PyInstaller 已安裝
    install_pyinstaller()
    
    # 3. 打包
    success = build_exe(args.version)
    
    # 4. 驗證
    if success:
        verify_build(args.version)
    
    print("\n完成!")
    return 0 if success else 1


if __name__ == '__main__':
    sys.exit(main())
