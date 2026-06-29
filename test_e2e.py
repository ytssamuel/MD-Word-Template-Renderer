#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
端到端測試

測試完整的 Markdown → Word / Excel 渲染流程
"""

import sys
from pathlib import Path

# 守門員：Windows cp950 上避免輸出 Unicode 符號時噴例外
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

# 添加 src 到路徑
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from md_word_renderer.parser import MarkdownParser
from md_word_renderer.renderer import WordRenderer
from md_word_renderer.validator import SchemaValidator


try:
    from md_word_renderer.renderer.factory import build_renderer
    HAS_EXCEL = True
except Exception:
    HAS_EXCEL = False


def test_simple_render():
    """測試簡單模板渲染"""
    print("=" * 50)
    print("測試 1: 簡單模板渲染")
    print("=" * 50)

    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')

    print(f"✓ 解析完成，共 {len([k for k in data.keys() if not k.startswith('#')])} 個欄位")

    renderer = WordRenderer()
    renderer.load_template('templates/simple_template.docx')
    renderer.render(data)

    output_dir = Path('output')
    output_dir.mkdir(exist_ok=True)

    output_path = 'output/test_simple_output.docx'
    renderer.save(output_path)

    print(f"✓ 渲染完成，輸出至: {output_path}")
    return True


def test_example_render():
    """測試範例模板渲染（含迴圈）"""
    print("\n" + "=" * 50)
    print("測試 2: 範例模板渲染（含迴圈）")
    print("=" * 50)

    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')

    test_cases = data.get('異動內容-測試案例', [])
    print(f"✓ 解析完成，測試案例共 {len(test_cases)} 項")

    renderer = WordRenderer()
    renderer.load_template('templates/example_template.docx')
    renderer.render(data)

    output_path = 'output/test_example_output.docx'
    renderer.save(output_path)

    print(f"✓ 渲染完成，輸出至: {output_path}")
    return True


def test_full_render():
    """測試完整模板渲染"""
    print("\n" + "=" * 50)
    print("測試 3: 完整模板渲染")
    print("=" * 50)

    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')

    validator = SchemaValidator()
    is_valid, errors = validator.validate(data)

    if is_valid:
        print("✓ 資料驗證通過")
    else:
        print(f"⚠ 資料驗證有 {len(errors)} 個警告")

    renderer = WordRenderer()
    renderer.load_template('templates/full_template.docx')
    renderer.render(data)

    output_path = 'output/test_full_output.docx'
    renderer.save(output_path)

    print(f"✓ 渲染完成，輸出至: {output_path}")
    return True


def test_api_usage():
    """測試 Python API 使用"""
    print("\n" + "=" * 50)
    print("測試 4: Python API 使用")
    print("=" * 50)

    from md_word_renderer import MarkdownParser, WordRenderer

    parser = MarkdownParser()
    renderer = WordRenderer()

    data = parser.parse('referance/sample_data.md')

    renderer.render_to_file(
        data,
        'templates/simple_template.docx',
        'output/test_api_output.docx'
    )

    print("✓ API 測試完成")
    return True


def test_data_structure():
    """測試資料結構"""
    print("\n" + "=" * 50)
    print("測試 5: 資料結構驗證")
    print("=" * 50)

    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')

    assert '系統名稱' in data, "缺少 系統名稱"
    assert '變更單號' in data, "缺少 變更單號"
    assert '異動內容-測試案例' in data, "缺少 異動內容-測試案例"

    assert '#1' in data, "缺少 #1 編號索引"
    assert data['#1']['key'] == '系統名稱', "#1 應該是系統名稱"

    test_cases = data['異動內容-測試案例']
    assert isinstance(test_cases, list), "異動內容-測試案例 應該是列表"
    assert len(test_cases) == 5, f"應該有 5 個測試項目，實際有 {len(test_cases)} 個"

    first_case = test_cases[0]
    assert 'children' in first_case, "測試項目應該有 children"
    assert len(first_case['children']) == 4, f"第一個測試項目應該有 4 個子項目"

    assert data.get('中介軟體') == "", "中介軟體 應該是空字串"

    print("✓ 所有資料結構驗證通過")
    return True


# --------------------------------------------------------------- Excel (v2.2)


def test_excel_simple_render():
    """v2.2：Markdown → Excel 簡單渲染（透過工廠）"""
    if not HAS_EXCEL:
        print("(略過：openpyxl 不可用)")
        return True

    print("\n" + "=" * 50)
    print("測試 6 (v2.2): Excel 簡單渲染")
    print("=" * 50)

    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')

    template = 'templates/excel/sample_template.xlsx'
    output_path = 'output/test_excel_simple_output.xlsx'

    Path('output').mkdir(exist_ok=True)
    renderer = build_renderer(template_path=template, format_hint="auto")
    renderer.render_to_file(data, template, output_path)

    print(f"✓ Excel 渲染完成，輸出至: {output_path}")
    return True


def test_excel_with_lists_render():
    """v2.2：含清單展開的 Markdown → Excel 渲染"""
    if not HAS_EXCEL:
        print("(略過：openpyxl 不可用)")
        return True

    print("\n" + "=" * 50)
    print("測試 7 (v2.2): Excel 含清單展開（with_lists_template）")
    print("=" * 50)

    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')

    template = 'templates/excel/with_lists_template.xlsx'
    output_path = 'output/test_excel_with_lists_output.xlsx'

    Path('output').mkdir(exist_ok=True)
    renderer = build_renderer(template_path=template, format_hint="auto")
    renderer.render_to_file(data, template, output_path)

    print(f"✓ Excel 渲染完成，輸出至: {output_path}")
    return True


def main():
    """執行所有測試"""
    print("\n" + "=" * 60)
    print("   MD-Word/Excel Template Renderer - 端到端測試")
    print("=" * 60)

    tests = [
        ("簡單模板渲染 (Word)", test_simple_render),
        ("範例模板渲染 (Word)", test_example_render),
        ("完整模板渲染 (Word)", test_full_render),
        ("Python API (Word)", test_api_usage),
        ("資料結構驗證", test_data_structure),
        ("Excel 簡單渲染 (v2.2)", test_excel_simple_render),
        ("Excel 含清單渲染 (v2.2)", test_excel_with_lists_render),
    ]

    passed = 0
    failed = 0

    for name, test_func in tests:
        try:
            result = test_func()
            if result:
                passed += 1
        except Exception as e:
            print(f"\n✗ 測試失敗: {name}")
            print(f"  錯誤: {e}")
            failed += 1

    print("\n" + "=" * 60)
    print(f"   測試結果: {passed} 通過, {failed} 失敗")
    print("=" * 60)

    if failed == 0:
        print("\n所有測試通過！")
        print("\n輸出檔案：")
        for f in sorted(Path('output').glob('*.docx')):
            print(f"  - {f}")
        for f in sorted(Path('output').glob('*.xlsx')):
            print(f"  - {f}")

    return failed == 0


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
