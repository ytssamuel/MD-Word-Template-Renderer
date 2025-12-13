#!/usr/bin/env python
"""測試解析結果"""
import sys
import json
sys.path.insert(0, 'src')

from md_word_renderer.parser import MarkdownParser

def main():
    parser = MarkdownParser()
    data = parser.parse('referance/sample_data.md')
    
    # 顯示異動內容
    items = data.get('異動內容-測試案例', [])
    print(f"=== 異動內容-測試案例 (共 {len(items)} 項) ===\n")
    
    for item in items:
        print(f"{item['number']}. {item['value']}")
        for child in item.get('children', []):
            print(f"   {child['number']}. {child['value']}")
            for grandchild in child.get('children', []):
                print(f"      {grandchild['number']}. {grandchild['value']}")
    
    print("\n=== 基本欄位 ===")
    for key in ['系統名稱', '變更單號', '測試日期']:
        if key in data:
            print(f"{key}: {data[key]}")

if __name__ == '__main__':
    main()
