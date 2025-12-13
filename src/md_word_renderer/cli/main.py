#!/usr/bin/env python
"""
CLI 主程式

提供命令列介面執行 MD → Word 轉換

Usage:
    md2word render <input_md> <template> <output>
    md2word batch <input_dir> <template> <output_dir>
    md2word validate <input_md>
    md2word info
"""

import argparse
import sys
from pathlib import Path
from typing import Optional, List

from ..parser import MarkdownParser
from ..renderer import WordRenderer
from ..validator import SchemaValidator


def create_parser() -> argparse.ArgumentParser:
    """
    建立命令列參數解析器
    
    Returns:
        argparse.ArgumentParser: 設定好的參數解析器
    """
    parser = argparse.ArgumentParser(
        prog='md2word',
        description='MD-Word Template Renderer - Markdown 轉 Word 文件工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
範例：
  # 單一檔案轉換
  md2word render input.md template.docx output.docx
  
  # 批次轉換（多個 MD 檔 + 一個模板）
  md2word batch ./inputs/ template.docx ./outputs/
  
  # 多模板批次轉換（一個 MD 檔 + 多個模板）
  md2word batch-templates data.md ./templates/ ./outputs/
  
  # 驗證 Markdown 格式
  md2word validate input.md
  
  # 顯示版本資訊
  md2word info
        '''
    )
    
    subparsers = parser.add_subparsers(dest='command', help='可用指令')
    
    # render 子指令
    render_parser = subparsers.add_parser(
        'render', 
        help='渲染單一 Markdown 檔案至 Word'
    )
    render_parser.add_argument(
        'input', 
        help='輸入的 Markdown 檔案路徑'
    )
    render_parser.add_argument(
        'template', 
        help='Word 模板檔案路徑 (.docx)'
    )
    render_parser.add_argument(
        'output', 
        help='輸出的 Word 檔案路徑'
    )
    render_parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='顯示詳細資訊'
    )
    render_parser.add_argument(
        '--no-validate',
        action='store_true',
        help='跳過資料驗證'
    )
    
    # batch 子指令
    batch_parser = subparsers.add_parser(
        'batch', 
        help='批次轉換多個 Markdown 檔案'
    )
    batch_parser.add_argument(
        'input_dir', 
        help='輸入目錄（包含 .md 檔案）'
    )
    batch_parser.add_argument(
        'template', 
        help='Word 模板檔案路徑 (.docx)'
    )
    batch_parser.add_argument(
        'output_dir', 
        help='輸出目錄'
    )
    batch_parser.add_argument(
        '-p', '--pattern',
        default='*.md',
        help='檔案搜尋模式 (預設: *.md)'
    )
    batch_parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='顯示詳細資訊'
    )
    batch_parser.add_argument(
        '--continue-on-error',
        action='store_true',
        help='遇到錯誤時繼續處理其他檔案'
    )
    
    # batch-templates 子指令 (1個md + n個模板 → n個word)
    batch_tpl_parser = subparsers.add_parser(
        'batch-templates', 
        help='使用多個模板渲染同一份 Markdown 資料'
    )
    batch_tpl_parser.add_argument(
        'input', 
        help='輸入的 Markdown 檔案路徑'
    )
    batch_tpl_parser.add_argument(
        'template_dir', 
        help='模板目錄（包含 .docx 檔案）'
    )
    batch_tpl_parser.add_argument(
        'output_dir', 
        help='輸出目錄'
    )
    batch_tpl_parser.add_argument(
        '-p', '--pattern',
        default='*.docx',
        help='模板搜尋模式 (預設: *.docx)'
    )
    batch_tpl_parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='顯示詳細資訊'
    )
    batch_tpl_parser.add_argument(
        '--continue-on-error',
        action='store_true',
        help='遇到錯誤時繼續處理其他模板'
    )
    batch_tpl_parser.add_argument(
        '--prefix',
        default='',
        help='輸出檔案名稱前綴'
    )
    batch_tpl_parser.add_argument(
        '--suffix',
        default='',
        help='輸出檔案名稱後綴（在副檔名之前）'
    )
    
    # validate 子指令
    validate_parser = subparsers.add_parser(
        'validate', 
        help='驗證 Markdown 檔案格式'
    )
    validate_parser.add_argument(
        'input', 
        help='要驗證的 Markdown 檔案路徑'
    )
    validate_parser.add_argument(
        '-s', '--schema',
        help='自訂 JSON Schema 檔案路徑'
    )
    
    # info 子指令
    subparsers.add_parser('info', help='顯示工具版本和相關資訊')
    
    return parser


def cmd_render(args: argparse.Namespace) -> int:
    """
    執行單一檔案渲染
    
    Args:
        args: 命令列參數
        
    Returns:
        int: 結束代碼 (0=成功, 1=失敗)
    """
    input_path = Path(args.input)
    template_path = Path(args.template)
    output_path = Path(args.output)
    
    # 檢查輸入檔案
    if not input_path.exists():
        print(f"❌ 錯誤：找不到輸入檔案 {input_path}")
        return 1
    
    if not template_path.exists():
        print(f"❌ 錯誤：找不到模板檔案 {template_path}")
        return 1
    
    try:
        # 解析 Markdown
        if args.verbose:
            print(f"📄 解析 Markdown: {input_path}")
        
        parser = MarkdownParser()
        data = parser.parse(str(input_path))
        
        field_count = len([k for k in data.keys() if not k.startswith('#')])
        if args.verbose:
            print(f"   ✓ 解析完成，共 {field_count} 個欄位")
        
        # 驗證資料（可選）
        if not args.no_validate:
            validator = SchemaValidator()
            is_valid, errors = validator.validate(data)
            
            if not is_valid:
                print(f"⚠ 警告：資料驗證有 {len(errors)} 個問題")
                for error in errors[:5]:  # 最多顯示 5 個
                    print(f"   - {error}")
        
        # 渲染 Word
        if args.verbose:
            print(f"📝 載入模板: {template_path}")
        
        renderer = WordRenderer()
        renderer.load_template(str(template_path))
        renderer.render(data)
        
        # 確保輸出目錄存在
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # 儲存
        renderer.save(str(output_path))
        
        print(f"✅ 成功輸出至: {output_path}")
        return 0
        
    except Exception as e:
        print(f"❌ 錯誤：{e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def cmd_batch(args: argparse.Namespace) -> int:
    """
    執行批次轉換
    
    Args:
        args: 命令列參數
        
    Returns:
        int: 結束代碼 (0=全部成功, 1=部分失敗)
    """
    input_dir = Path(args.input_dir)
    template_path = Path(args.template)
    output_dir = Path(args.output_dir)
    
    # 檢查輸入目錄
    if not input_dir.exists():
        print(f"❌ 錯誤：找不到輸入目錄 {input_dir}")
        return 1
    
    if not template_path.exists():
        print(f"❌ 錯誤：找不到模板檔案 {template_path}")
        return 1
    
    # 搜尋 Markdown 檔案
    md_files = list(input_dir.glob(args.pattern))
    
    if not md_files:
        print(f"⚠ 警告：在 {input_dir} 中找不到符合 {args.pattern} 的檔案")
        return 1
    
    print(f"📂 找到 {len(md_files)} 個檔案待處理")
    
    # 確保輸出目錄存在
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 初始化
    parser = MarkdownParser()
    renderer = WordRenderer()
    
    success_count = 0
    fail_count = 0
    
    # 批次處理
    for md_file in md_files:
        output_file = output_dir / f"{md_file.stem}.docx"
        
        try:
            if args.verbose:
                print(f"\n處理: {md_file.name}")
            
            # 解析
            data = parser.parse(str(md_file))
            
            # 渲染
            renderer.load_template(str(template_path))
            renderer.render(data)
            renderer.save(str(output_file))
            
            if args.verbose:
                print(f"   ✓ 輸出至 {output_file.name}")
            
            success_count += 1
            
        except Exception as e:
            print(f"   ✗ 失敗: {md_file.name} - {e}")
            fail_count += 1
            
            if not args.continue_on_error:
                print("終止批次處理（使用 --continue-on-error 可繼續處理其他檔案）")
                break
    
    # 顯示結果
    print(f"\n📊 批次處理完成")
    print(f"   ✓ 成功: {success_count} 個")
    print(f"   ✗ 失敗: {fail_count} 個")
    
    return 0 if fail_count == 0 else 1


def cmd_batch_templates(args: argparse.Namespace) -> int:
    """
    使用多個模板渲染同一份 Markdown 資料
    
    Args:
        args: 命令列參數
        
    Returns:
        int: 結束代碼 (0=全部成功, 1=部分失敗)
    """
    input_path = Path(args.input)
    template_dir = Path(args.template_dir)
    output_dir = Path(args.output_dir)
    
    # 檢查輸入檔案
    if not input_path.exists():
        print(f"❌ 錯誤：找不到輸入檔案 {input_path}")
        return 1
    
    # 檢查模板目錄
    if not template_dir.exists():
        print(f"❌ 錯誤：找不到模板目錄 {template_dir}")
        return 1
    
    # 搜尋模板檔案
    template_files = list(template_dir.glob(args.pattern))
    
    if not template_files:
        print(f"⚠ 警告：在 {template_dir} 中找不到符合 {args.pattern} 的模板檔案")
        return 1
    
    print(f"📂 找到 {len(template_files)} 個模板待處理")
    
    # 確保輸出目錄存在
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # 解析 Markdown（只需解析一次）
    try:
        print(f"📄 解析 Markdown: {input_path}")
        parser = MarkdownParser()
        data = parser.parse(str(input_path))
        
        field_count = len([k for k in data.keys() if not k.startswith('#')])
        print(f"   ✓ 解析完成，共 {field_count} 個欄位")
    except Exception as e:
        print(f"❌ 解析 Markdown 失敗：{e}")
        return 1
    
    # 初始化
    renderer = WordRenderer()
    
    success_count = 0
    fail_count = 0
    
    # 批次處理各模板
    for template_file in template_files:
        # 組合輸出檔名
        output_name = f"{args.prefix}{template_file.stem}{args.suffix}.docx"
        output_file = output_dir / output_name
        
        try:
            if args.verbose:
                print(f"\n處理模板: {template_file.name}")
            
            # 渲染
            renderer.load_template(str(template_file))
            renderer.render(data)
            renderer.save(str(output_file))
            
            if args.verbose:
                print(f"   ✓ 輸出至 {output_file.name}")
            
            success_count += 1
            
        except Exception as e:
            print(f"   ✗ 失敗: {template_file.name} - {e}")
            fail_count += 1
            
            if not args.continue_on_error:
                print("終止批次處理（使用 --continue-on-error 可繼續處理其他模板）")
                break
    
    # 顯示結果
    print(f"\n📊 多模板批次處理完成")
    print(f"   ✓ 成功: {success_count} 個")
    print(f"   ✗ 失敗: {fail_count} 個")
    
    return 0 if fail_count == 0 else 1


def cmd_validate(args: argparse.Namespace) -> int:
    """
    執行資料驗證
    
    Args:
        args: 命令列參數
        
    Returns:
        int: 結束代碼 (0=通過, 1=失敗)
    """
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ 錯誤：找不到檔案 {input_path}")
        return 1
    
    try:
        # 解析
        print(f"📄 解析 Markdown: {input_path}")
        parser = MarkdownParser()
        data = parser.parse(str(input_path))
        
        field_count = len([k for k in data.keys() if not k.startswith('#')])
        print(f"   ✓ 解析完成，共 {field_count} 個欄位")
        
        # 驗證
        print("\n🔍 執行驗證...")
        validator = SchemaValidator()
        
        # 載入自訂 Schema（如果有）
        if args.schema:
            schema_path = Path(args.schema)
            if not schema_path.exists():
                print(f"❌ 錯誤：找不到 Schema 檔案 {schema_path}")
                return 1
            validator.load_schema(str(schema_path))
        
        is_valid, errors = validator.validate(data)
        
        if is_valid:
            print("✅ 驗證通過！資料格式正確")
            
            # 顯示欄位摘要
            print("\n📋 欄位摘要：")
            for key, value in data.items():
                if key.startswith('#'):
                    continue
                    
                if isinstance(value, list):
                    print(f"   {key}: [列表，{len(value)} 項]")
                elif isinstance(value, str):
                    preview = value[:30] + "..." if len(value) > 30 else value
                    print(f"   {key}: {preview}")
            
            return 0
        else:
            print(f"❌ 驗證失敗，共 {len(errors)} 個問題：")
            for error in errors:
                print(f"   - {error}")
            return 1
            
    except Exception as e:
        print(f"❌ 錯誤：{e}")
        return 1


def cmd_info() -> int:
    """
    顯示工具資訊
    
    Returns:
        int: 結束代碼
    """
    print("""
╔══════════════════════════════════════════════════╗
║     MD-Word Template Renderer                    ║
║     Markdown → Word 文件轉換工具                 ║
╠══════════════════════════════════════════════════╣
║  版本: 1.0.0                                     ║
║  作者: SpeedBOT Team                             ║
║  授權: MIT                                       ║
╚══════════════════════════════════════════════════╝

功能：
  • 解析特定格式 Markdown（編號. 欄位名稱 | 值）
  • 支援階層結構與縮排
  • 使用 Jinja2 模板引擎渲染 Word
  • 支援迴圈、條件等進階語法
  • 批次處理多個檔案

依賴套件：
  • python-docx >= 0.8.11
  • docxtpl >= 0.16.7
  • Jinja2 >= 3.1.2
  • PyYAML >= 6.0
  • jsonschema >= 4.17.0

詳細說明請參閱: README.md
    """)
    return 0


def cli(args: Optional[List[str]] = None) -> int:
    """
    CLI 入口點
    
    Args:
        args: 命令列參數（用於測試）
        
    Returns:
        int: 結束代碼
    """
    parser = create_parser()
    parsed_args = parser.parse_args(args)
    
    if parsed_args.command is None:
        parser.print_help()
        return 0
    
    if parsed_args.command == 'render':
        return cmd_render(parsed_args)
    elif parsed_args.command == 'batch':
        return cmd_batch(parsed_args)
    elif parsed_args.command == 'batch-templates':
        return cmd_batch_templates(parsed_args)
    elif parsed_args.command == 'validate':
        return cmd_validate(parsed_args)
    elif parsed_args.command == 'info':
        return cmd_info()
    else:
        parser.print_help()
        return 1


def main():
    """主程式入口"""
    sys.exit(cli())


if __name__ == '__main__':
    main()
