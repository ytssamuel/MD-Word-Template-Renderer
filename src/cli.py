#!/usr/bin/env python
"""
MD-Word Template Renderer - 命令列介面

將 Markdown 資料渲染至 Word 模板
"""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Optional

from md_word_renderer.parser import MarkdownParser
from md_word_renderer.renderer import WordRenderer
from md_word_renderer.validator import SchemaValidator
from md_word_renderer.config import ConfigLoader
from md_word_renderer.utils import FileUtils, BatchProcessor


__version__ = "1.0.0"


def setup_logging(verbose: bool = False, log_file: Optional[str] = None):
    """設定日誌"""
    level = logging.DEBUG if verbose else logging.INFO
    
    handlers = [logging.StreamHandler(sys.stdout)]
    
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        handlers.append(logging.FileHandler(log_file, encoding='utf-8'))
    
    logging.basicConfig(
        level=level,
        format='[%(levelname)s] %(message)s',
        handlers=handlers
    )


def render_single_file(markdown_path: str, 
                       template_path: str, 
                       output_path: str,
                       validate: bool = False,
                       schema_path: Optional[str] = None) -> None:
    """
    渲染單個檔案
    
    Args:
        markdown_path: Markdown 檔案路徑
        template_path: Word 模板路徑
        output_path: 輸出檔案路徑
        validate: 是否驗證資料
        schema_path: Schema 檔案路徑
    """
    logger = logging.getLogger(__name__)
    
    # 解析 Markdown
    logger.info(f"讀取 Markdown 檔案: {markdown_path}")
    parser = MarkdownParser()
    data = parser.parse(markdown_path)
    logger.info(f"解析完成，共 {len([k for k in data.keys() if not k.startswith('#')])} 個欄位")
    
    # 資料驗證
    if validate:
        logger.info("執行資料驗證...")
        validator = SchemaValidator(schema_path)
        is_valid, errors = validator.validate(data)
        
        if not is_valid:
            for error in errors:
                logger.warning(f"驗證警告: {error}")
        else:
            logger.info("✓ 資料驗證通過")
    
    # 渲染 Word
    logger.info(f"載入 Word 模板: {template_path}")
    renderer = WordRenderer()
    renderer.load_template(template_path)
    
    logger.info("開始渲染...")
    renderer.render(data)
    
    logger.info(f"儲存至: {output_path}")
    renderer.save(output_path)
    
    logger.info("✓ 處理完成！")


def main():
    """主程式進入點"""
    parser = argparse.ArgumentParser(
        prog='md-word-renderer',
        description='Markdown to Word 模板渲染工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
範例:
  # 單檔案處理
  python cli.py -m data.md -t template.docx -o output.docx
  
  # 多檔案批次處理
  python cli.py -m *.md -t template.docx -o output/
  
  # 使用設定檔
  python cli.py --config config.yaml
  
  # 啟用驗證
  python cli.py -m data.md -t template.docx -o output.docx --validate -v
'''
    )
    
    # 必要參數
    parser.add_argument(
        '-m', '--markdown',
        nargs='+',
        help='Markdown 資料檔案路徑（支援多檔案或通配符）'
    )
    
    parser.add_argument(
        '-t', '--template',
        help='Word 模板檔案路徑'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='輸出路徑（單檔案為檔案路徑，多檔案為目錄）'
    )
    
    # 選項參數
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='顯示詳細處理資訊'
    )
    
    parser.add_argument(
        '--validate',
        action='store_true',
        help='執行資料結構驗證'
    )
    
    parser.add_argument(
        '--schema',
        help='JSON Schema 檔案路徑（驗證用）'
    )
    
    parser.add_argument(
        '--config',
        help='設定檔路徑'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__}'
    )
    
    args = parser.parse_args()
    
    # 載入設定檔
    config = ConfigLoader(args.config)
    
    # 合併 CLI 參數
    template_path = args.template or config.get('template')
    output_path = args.output or config.get('output_dir')
    validate = args.validate or config.get('validation.enabled', False)
    schema_path = args.schema or config.get('validation.schema_path')
    verbose = args.verbose
    
    # 設定日誌
    log_file = config.get('logging.log_file') if config.get('logging.file_output') else None
    setup_logging(verbose, log_file)
    
    logger = logging.getLogger(__name__)
    
    # 檢查必要參數
    if not args.markdown:
        logger.error("請指定 Markdown 檔案（-m 參數）")
        parser.print_help()
        sys.exit(1)
    
    if not template_path:
        logger.error("請指定 Word 模板（-t 參數）")
        parser.print_help()
        sys.exit(1)
    
    if not output_path:
        logger.error("請指定輸出路徑（-o 參數）")
        parser.print_help()
        sys.exit(1)
    
    try:
        # 展開檔案 patterns
        processor = BatchProcessor(verbose=verbose)
        input_files = processor.expand_file_patterns(args.markdown)
        
        if not input_files:
            logger.error(f"找不到符合條件的檔案: {args.markdown}")
            sys.exit(1)
        
        if len(input_files) == 1:
            # 單檔案處理
            render_single_file(
                input_files[0],
                template_path,
                output_path,
                validate=validate,
                schema_path=schema_path
            )
        else:
            # 多檔案批次處理
            def process_func(md_path, tpl_path, out_path):
                render_single_file(
                    md_path, tpl_path, out_path,
                    validate=validate,
                    schema_path=schema_path
                )
            
            results = processor.process_files(
                input_files,
                template_path,
                output_path,
                process_func
            )
            
            # 顯示摘要
            print(f"\n批次處理完成！")
            print(f"成功: {results['success_count']} 個檔案")
            print(f"失敗: {results['failed_count']} 個檔案")
            
            if results['failed']:
                print("\n失敗的檔案:")
                for item in results['failed']:
                    print(f"  - {item['input']}: {item['error']}")
    
    except FileNotFoundError as e:
        logger.error(str(e))
        sys.exit(1)
    except Exception as e:
        logger.error(f"處理失敗: {e}")
        if verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
