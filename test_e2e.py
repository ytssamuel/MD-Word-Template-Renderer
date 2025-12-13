#!/usr/bin/env python
"""
端到端測試

測試完整的 Markdown → Word 渲染流程
"""

import sys
from pathlib import Path

# 添加 src 到路徑
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from md_word_renderer.parser import MarkdownParser
from md_word_renderer.renderer import WordRenderer
from md_word_renderer.validator import SchemaValidator


def test_simple_render():
    """測試簡單模板渲染"""
    print("=" * 50)
    print("測試 1: 簡單模板渲染")
    print("=" * 50)
    
    # 解析 Markdown
    parser = MarkdownParser()
    data = parser.parse('referance/OH_20251020.md')
    
    print(f"✓ 解析完成，共 {len([k for k in data.keys() if not k.startswith('#')])} 個欄位")
    
    # 渲染 Word
    renderer = WordRenderer()
    renderer.load_template('templates/simple_template.docx')
    renderer.render(data)
    
    # 確保輸出目錄存在
    output_dir = Path('output')
    output_dir.mkdir(exist_ok=True)
    
    output_path = 'output/test_simple_output.docx'
    renderer.save(output_path)
    
    print(f"✓ 渲染完成，輸出至: {output_path}")
    return True


def test_example_render():
    """測試範例模板渲染"""
    print("\n" + "=" * 50)
    print("測試 2: 範例模板渲染（含迴圈）")
    print("=" * 50)
    
    # 解析 Markdown
    parser = MarkdownParser()
    data = parser.parse('referance/OH_20251020.md')
    
    # 顯示測試案例數量
    test_cases = data.get('異動內容-測試案例', [])
    print(f"✓ 解析完成，測試案例共 {len(test_cases)} 項")
    
    # 渲染 Word
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
    
    # 解析 Markdown
    parser = MarkdownParser()
    data = parser.parse('referance/OH_20251020.md')
    
    # 資料驗證
    validator = SchemaValidator()
    is_valid, errors = validator.validate(data)
    
    if is_valid:
        print("✓ 資料驗證通過")
    else:
        print(f"⚠ 資料驗證有 {len(errors)} 個警告")
    
    # 渲染 Word
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
    
    # 一站式使用
    parser = MarkdownParser()
    renderer = WordRenderer()
    
    data = parser.parse('referance/OH_20251020.md')
    
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
    data = parser.parse('referance/OH_20251020.md')
    
    # 驗證基本欄位
    assert '系統名稱' in data, "缺少 系統名稱"
    assert '變更單號' in data, "缺少 變更單號"
    assert '異動內容-測試案例' in data, "缺少 異動內容-測試案例"
    
    # 驗證編號索引
    assert '#1' in data, "缺少 #1 編號索引"
    assert data['#1']['key'] == '系統名稱', "#1 應該是系統名稱"
    
    # 驗證階層結構
    test_cases = data['異動內容-測試案例']
    assert isinstance(test_cases, list), "異動內容-測試案例 應該是列表"
    assert len(test_cases) == 5, f"應該有 5 個測試項目，實際有 {len(test_cases)} 個"
    
    # 驗證子項目
    first_case = test_cases[0]
    assert 'children' in first_case, "測試項目應該有 children"
    assert len(first_case['children']) == 4, f"第一個測試項目應該有 4 個子項目"
    
    # 驗證空值處理
    assert data.get('中介軟體') == "", "中介軟體 應該是空字串"
    
    print("✓ 所有資料結構驗證通過")
    return True


def main():
    """執行所有測試"""
    print("\n" + "=" * 60)
    print("   MD-Word Template Renderer - 端到端測試")
    print("=" * 60)
    
    tests = [
        ("簡單模板渲染", test_simple_render),
        ("範例模板渲染", test_example_render),
        ("完整模板渲染", test_full_render),
        ("Python API", test_api_usage),
        ("資料結構驗證", test_data_structure),
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
        print("\n🎉 所有測試通過！")
        print("\n輸出檔案：")
        for f in Path('output').glob('*.docx'):
            print(f"  - {f}")
    
    return failed == 0


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
