"""
CLI 主程式

提供命令列介面執行 MD → Word / Excel 轉換

Usage:
    md2word render <input_md> <template> <output> [--format {docx,xlsx,auto}]
    md2word batch <input_dir> <template> <output_dir> [--format ...]
    md2word batch-templates <input_md> <template_dir> <output_dir> [--format ...]
    md2word validate <input_md>
    md2word info
"""

import argparse
import sys
from pathlib import Path
from typing import Optional, List, Union

# 守門員：Windows cp950 上避免輸出 Unicode 符號時噴例外（v2.2 新增）
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (ValueError, AttributeError):
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8")
    except (ValueError, AttributeError):
        pass

from ..parser import MarkdownParser
from ..renderer import WordRenderer
from ..renderer.factory import build_renderer, detect_format, output_extension_for
from ..validator import SchemaValidator


__all__ = ["create_parser", "cli", "main", "process_one", "resolve_format"]


def __getattr__(name):
    if name == "ExcelRenderer":
        from ..renderer.excel_renderer import ExcelRenderer  # noqa: WPS433
        return ExcelRenderer
    raise AttributeError(name)


# ----------------------------------------------------------------- helpers


def resolve_format(template_path: str, format_hint: str) -> str:
    """將 ``--format`` 的值解析為 ``docx`` / ``xlsx``。``auto`` 時從副檔名推。"""
    fmt = (format_hint or "auto").lower()
    if fmt == "auto":
        return detect_format(template_path)
    if fmt not in {"docx", "xlsx"}:
        raise ValueError(f"不支援的格式: {fmt!r}")
    return fmt


def process_one(
    input_path: str,
    template_path: str,
    output_path: str,
    format_hint: str = "auto",
    validate: bool = True,
    verbose: bool = False,
) -> dict:
    """
    處理單一檔案的核心流程；Word / Excel 共用。

    Returns:
        dict: ``{"format": "docx"|"xlsx", "renderer": <instance>, "fields": int, "output": str}``
    """
    fmt = resolve_format(template_path, format_hint)
    renderer = build_renderer(template_path=template_path, format_hint=fmt)

    if verbose:
        print(f"📄 解析 Markdown: {input_path}")

    parser = MarkdownParser()
    data = parser.parse(str(input_path))
    field_count = len([k for k in data.keys() if not k.startswith("#")])

    if verbose:
        print(f"   ✓ 解析完成，共 {field_count} 個欄位")
        print(f"📝 載入樣板: {template_path} (format={fmt})")

    if validate:
        v = SchemaValidator()
        is_valid, errors = v.validate(data)
        if not is_valid and verbose:
            print(f"⚠ 警告：資料驗證有 {len(errors)} 個問題")
            for error in errors[:5]:
                print(f"   - {error}")

    renderer.load_template(str(template_path))
    renderer.render(data)
    renderer.save(str(output_path))

    return {
        "format": fmt,
        "renderer": renderer,
        "fields": field_count,
        "output": output_path,
    }


# ----------------------------------------------------------------- argparse


def _add_format_flag(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--format",
        dest="format",
        choices=["auto", "docx", "xlsx"],
        default="auto",
        help="輸出格式；預設 auto 從樣板副檔名推斷 (.docx→docx, .xlsx→xlsx)",
    )


def _add_render_parser(parser: argparse.ArgumentParser, is_batch: bool = False,
                       is_batch_templates: bool = False) -> None:
    """共用 render/batch/batch-templates 的參數集合。"""
    if is_batch_templates:
        parser.add_argument("input", help="輸入的 Markdown 檔案路徑")
        parser.add_argument(
            "template_dir", help="模板目錄（包含 .docx / .xlsx 檔案）"
        )
        parser.add_argument("output_dir", help="輸出目錄")
        parser.add_argument(
            "-p", "--pattern",
            default=None,
            help="檔案搜尋模式（預設 *.docx / *.xlsx，視 --format 而定）",
        )
    elif is_batch:
        parser.add_argument("input_dir", help="輸入目錄（包含 .md 檔案）")
        parser.add_argument("template", help="樣板檔案路徑 (.docx 或 .xlsx)")
        parser.add_argument("output_dir", help="輸出目錄")
        parser.add_argument(
            "-p", "--pattern",
            default="*.md",
            help="檔案搜尋模式 (預設: *.md)",
        )
    else:
        parser.add_argument("input", help="輸入的 Markdown 檔案路徑")
        parser.add_argument("template", help="樣板檔案路徑 (.docx 或 .xlsx)")
        parser.add_argument("output", help="輸出的 Word/Excel 檔案路徑")

    parser.add_argument(
        "-v", "--verbose", action="store_true", help="顯示詳細資訊"
    )

    parser.add_argument(
        "--no-validate", dest="no_validate", action="store_true",
        help="跳過資料驗證",
    )

    if is_batch or is_batch_templates:
        parser.add_argument(
            "--continue-on-error", action="store_true",
            help="遇到錯誤時繼續處理其他檔案",
        )

    if is_batch_templates:
        parser.add_argument("--prefix", default="", help="輸出檔案名稱前綴")
        parser.add_argument("--suffix", default="", help="輸出檔案名稱後綴")

    _add_format_flag(parser)


def create_parser() -> argparse.ArgumentParser:
    """建立命令列參數解析器"""
    parser = argparse.ArgumentParser(
        prog="md2word",
        description="MD-Word/Excel Template Renderer - Markdown 轉 Word 或 Excel 文件工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
範例：
  # 單一檔案轉換（Word）
  md2word render input.md template.docx output.docx

  # 單一檔案轉換（Excel，自動從樣板副檔名偵測）
  md2word render input.md template.xlsx output.xlsx

  # 明確指定 Excel
  md2word render input.md template.xlsx output.xlsx --format xlsx

  # 批次轉換
  md2word batch ./inputs/ template.docx ./outputs/

  # 多模板批次轉換
  md2word batch-templates data.md ./templates/ ./outputs/

  # 驗證 Markdown 格式
  md2word validate input.md

  # 顯示版本資訊
  md2word info
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="可用指令")

    render_p = subparsers.add_parser(
        "render", help="渲染單一 Markdown 檔案至 Word/Excel"
    )
    _add_render_parser(render_p)

    batch_p = subparsers.add_parser(
        "batch", help="批次轉換多個 Markdown 檔案"
    )
    _add_render_parser(batch_p, is_batch=True)

    btpl_p = subparsers.add_parser(
        "batch-templates", help="使用多個模板渲染同一份 Markdown 資料"
    )
    _add_render_parser(btpl_p, is_batch_templates=True)

    validate_parser = subparsers.add_parser(
        "validate", help="驗證 Markdown 檔案格式"
    )
    validate_parser.add_argument("input", help="要驗證的 Markdown 檔案路徑")
    validate_parser.add_argument("-s", "--schema", help="自訂 JSON Schema 檔案路徑")

    subparsers.add_parser("info", help="顯示工具版本和相關資訊")

    return parser


# ----------------------------------------------------------------- commands


def cmd_render(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    template_path = Path(args.template)
    output_path = Path(args.output)

    if not input_path.exists():
        print(f"❌ 錯誤：找不到輸入檔案 {input_path}")
        return 1
    if not template_path.exists():
        print(f"❌ 錯誤：找不到模板檔案 {template_path}")
        return 1

    try:
        fmt = resolve_format(str(template_path), args.format)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        process_one(
            input_path=str(input_path),
            template_path=str(template_path),
            output_path=str(output_path),
            format_hint=args.format,
            validate=not getattr(args, "no_validate", False),
            verbose=args.verbose,
        )

        if args.verbose:
            print(f"   ✓ 已使用 {fmt} 渲染器")
        print(f"✅ 成功輸出至: {output_path}")
        return 0
    except Exception as e:
        print(f"❌ 錯誤：{e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1


def cmd_batch(args: argparse.Namespace) -> int:
    input_dir = Path(args.input_dir)
    template_path = Path(args.template)
    output_dir = Path(args.output_dir)

    if not input_dir.exists():
        print(f"❌ 錯誤：找不到輸入目錄 {input_dir}")
        return 1
    if not template_path.exists():
        print(f"❌ 錯誤：找不到模板檔案 {template_path}")
        return 1

    md_files = list(input_dir.glob(args.pattern))
    if not md_files:
        print(f"⚠ 警告：在 {input_dir} 中找不到符合 {args.pattern} 的檔案")
        return 1

    try:
        fmt = resolve_format(str(template_path), args.format)
    except ValueError as e:
        print(f"❌ 錯誤：{e}")
        return 1

    output_ext = f".{fmt}"
    print(f"📂 找到 {len(md_files)} 個檔案待處理 (格式: {fmt})")

    output_dir.mkdir(parents=True, exist_ok=True)

    parser = MarkdownParser()
    success_count = 0
    fail_count = 0

    for md_file in md_files:
        output_file = output_dir / f"{md_file.stem}{output_ext}"
        try:
            if args.verbose:
                print(f"\n處理: {md_file.name}")

            data = parser.parse(str(md_file))
            renderer = build_renderer(template_path=str(template_path), format_hint=fmt)
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

    print(f"\n📊 批次處理完成")
    print(f"   ✓ 成功: {success_count} 個")
    print(f"   ✗ 失敗: {fail_count} 個")
    return 0 if fail_count == 0 else 1


def cmd_batch_templates(args: argparse.Namespace) -> int:
    input_path = Path(args.input)
    template_dir = Path(args.template_dir)
    output_dir = Path(args.output_dir)

    if not input_path.exists():
        print(f"❌ 錯誤：找不到輸入檔案 {input_path}")
        return 1
    if not template_dir.exists():
        print(f"❌ 錯誤：找不到模板目錄 {template_dir}")
        return 1

    # 列出樣板檔案
    candidate_files = []
    for ext in ("docx", "xlsx"):
        candidate_files.extend(template_dir.glob(f"*.{ext}"))
    if args.pattern:
        pattern_files = list(template_dir.glob(args.pattern))
        if pattern_files:
            candidate_files = pattern_files

    # 去重
    seen = set()
    template_files = []
    for p in candidate_files:
        if str(p) not in seen:
            seen.add(str(p))
            template_files.append(p)
    candidate_files = template_files

    if not candidate_files:
        print(f"⚠ 警告：在 {template_dir} 中找不到任何 .docx / .xlsx 樣板")
        return 1

    print(f"📂 找到 {len(candidate_files)} 個樣板待處理")

    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        print(f"📄 解析 Markdown: {input_path}")
        parser = MarkdownParser()
        data = parser.parse(str(input_path))
        field_count = len([k for k in data.keys() if not k.startswith("#")])
        print(f"   ✓ 解析完成，共 {field_count} 個欄位")
    except Exception as e:
        print(f"❌ 解析 Markdown 失敗：{e}")
        return 1

    success_count = 0
    fail_count = 0

    for template_file in candidate_files:
        try:
            fmt = resolve_format(str(template_file), args.format)
        except ValueError:
            fmt = detect_format(str(template_file))

        output_ext = f".{fmt}"
        output_name = f"{args.prefix}{template_file.stem}{args.suffix}{output_ext}"
        output_file = output_dir / output_name

        try:
            if args.verbose:
                print(f"\n處理樣板: {template_file.name} (format={fmt})")

            renderer = build_renderer(template_path=str(template_file), format_hint=fmt)
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
                print("終止批次處理（使用 --continue-on-error 可繼續處理其他樣板）")
                break

    print(f"\n📊 多模板批次處理完成")
    print(f"   ✓ 成功: {success_count} 個")
    print(f"   ✗ 失敗: {fail_count} 個")
    return 0 if fail_count == 0 else 1


def cmd_validate(args: argparse.Namespace) -> int:
    input_path = Path(args.input)

    if not input_path.exists():
        print(f"❌ 錯誤：找不到檔案 {input_path}")
        return 1

    try:
        print(f"📄 解析 Markdown: {input_path}")
        parser = MarkdownParser()
        data = parser.parse(str(input_path))

        field_count = len([k for k in data.keys() if not k.startswith("#")])
        print(f"   ✓ 解析完成，共 {field_count} 個欄位")

        print("\n🔍 執行驗證...")
        validator = SchemaValidator()

        if args.schema:
            schema_path = Path(args.schema)
            if not schema_path.exists():
                print(f"❌ 錯誤：找不到 Schema 檔案 {schema_path}")
                return 1
            validator.load_schema(str(schema_path))

        is_valid, errors = validator.validate(data)

        if is_valid:
            print("✅ 驗證通過！資料格式正確")
            print("\n📋 欄位摘要：")
            for key, value in data.items():
                if key.startswith("#"):
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
    print("""
╔══════════════════════════════════════════════════╗
║     MD-Word Template Renderer                    ║
║     Markdown → Word / Excel 文件轉換工具         ║
╠══════════════════════════════════════════════════╣
║  版本: 2.2.1                                     ║
║  作者: ytssamuel                             ║
║  授權: MIT                                       ║
╚══════════════════════════════════════════════════╝

功能：
  • 解析特定格式 Markdown（編號. 欄位名稱 | 值）
  • 支援階層結構與縮排
  • 渲染至 Word（docxtpl + Jinja2）或 Excel（openpyxl）
  • 支援迴圈、條件等進階語法
  • 批次處理多個檔案

依賴套件：
  • python-docx >= 0.8.11
  • docxtpl >= 0.16.7
  • Jinja2 >= 3.1.2
  • openpyxl >= 3.1
  • PyYAML >= 6.0
  • jsonschema >= 4.17.0

詳細說明請參閱: README.md
    """)
    return 0


def cli(args: Optional[List[str]] = None) -> int:
    parser = create_parser()
    parsed_args = parser.parse_args(args)

    if parsed_args.command is None:
        parser.print_help()
        return 0

    if parsed_args.command == "render":
        return cmd_render(parsed_args)
    elif parsed_args.command == "batch":
        return cmd_batch(parsed_args)
    elif parsed_args.command == "batch-templates":
        return cmd_batch_templates(parsed_args)
    elif parsed_args.command == "validate":
        return cmd_validate(parsed_args)
    elif parsed_args.command == "info":
        return cmd_info()
    else:
        parser.print_help()
        return 1


def main():
    sys.exit(cli())


if __name__ == "__main__":
    main()
