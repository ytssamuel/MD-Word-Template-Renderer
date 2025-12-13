#!/usr/bin/env python
"""
MD-Word Template Renderer - 主程式

提供 Python API 使用方式
"""

from md_word_renderer import MarkdownParser, WordRenderer
from md_word_renderer.validator import SchemaValidator
from md_word_renderer.config import ConfigLoader


def render_markdown_to_word(markdown_path: str,
                            template_path: str,
                            output_path: str,
                            validate: bool = False) -> dict:
    """
    將 Markdown 資料渲染至 Word 模板
    
    Args:
        markdown_path: Markdown 檔案路徑
        template_path: Word 模板路徑
        output_path: 輸出檔案路徑
        validate: 是否驗證資料
        
    Returns:
        dict: 解析後的資料（供後續使用）
        
    Example:
        >>> data = render_markdown_to_word(
        ...     "data.md",
        ...     "template.docx",
        ...     "output.docx"
        ... )
    """
    # 解析 Markdown
    parser = MarkdownParser()
    data = parser.parse(markdown_path)
    
    # 資料驗證
    if validate:
        validator = SchemaValidator()
        is_valid, errors = validator.validate(data)
        if not is_valid:
            print(f"警告: 資料驗證有 {len(errors)} 個問題")
            for error in errors:
                print(f"  - {error}")
    
    # 渲染 Word
    renderer = WordRenderer()
    renderer.render_to_file(data, template_path, output_path)
    
    return data


def main():
    """範例使用"""
    import sys
    
    if len(sys.argv) < 4:
        print("用法: python main.py <markdown_file> <template_file> <output_file>")
        print("\n範例:")
        print("  python main.py data.md template.docx output.docx")
        sys.exit(1)
    
    markdown_path = sys.argv[1]
    template_path = sys.argv[2]
    output_path = sys.argv[3]
    
    try:
        data = render_markdown_to_word(markdown_path, template_path, output_path)
        print(f"✓ 成功渲染至 {output_path}")
        print(f"  共處理 {len([k for k in data.keys() if not k.startswith('#')])} 個欄位")
    except Exception as e:
        print(f"✗ 錯誤: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
