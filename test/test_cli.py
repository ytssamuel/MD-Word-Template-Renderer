#!/usr/bin/env python
"""
CLI 測試

測試命令列工具的各種功能
"""

import os
import sys
import unittest
import tempfile
from pathlib import Path

# 添加 src 到路徑
sys.path.insert(0, str(Path(__file__).parent.parent / 'src'))

from md_word_renderer.cli.main import cli, create_parser


class TestCLI(unittest.TestCase):
    """CLI 功能測試"""
    
    @classmethod
    def setUpClass(cls):
        """設定測試環境"""
        cls.test_dir = Path(__file__).parent.parent
        cls.sample_md = cls.test_dir / 'referance' / 'sample_data.md'
        cls.template = cls.test_dir / 'templates' / 'simple_template.docx'
        cls.temp_dir = tempfile.mkdtemp()
    
    def test_parser_creation(self):
        """測試參數解析器建立"""
        parser = create_parser()
        self.assertIsNotNone(parser)
        self.assertEqual(parser.prog, 'md2word')
    
    def test_help_command(self):
        """測試 help 指令不會報錯"""
        result = cli([])
        self.assertEqual(result, 0)
    
    def test_info_command(self):
        """測試 info 指令"""
        result = cli(['info'])
        self.assertEqual(result, 0)
    
    def test_validate_command(self):
        """測試 validate 指令"""
        if not self.sample_md.exists():
            self.skipTest(f"測試檔案不存在: {self.sample_md}")
        
        result = cli(['validate', str(self.sample_md)])
        self.assertEqual(result, 0)
    
    def test_validate_missing_file(self):
        """測試 validate 指令 - 檔案不存在"""
        result = cli(['validate', 'nonexistent_file.md'])
        self.assertEqual(result, 1)
    
    def test_render_command(self):
        """測試 render 指令"""
        if not self.sample_md.exists():
            self.skipTest(f"測試檔案不存在: {self.sample_md}")
        if not self.template.exists():
            self.skipTest(f"模板檔案不存在: {self.template}")
        
        output_path = Path(self.temp_dir) / 'test_output.docx'
        
        result = cli([
            'render',
            str(self.sample_md),
            str(self.template),
            str(output_path)
        ])
        
        self.assertEqual(result, 0)
        self.assertTrue(output_path.exists())
    
    def test_render_missing_input(self):
        """測試 render 指令 - 輸入檔案不存在"""
        output_path = Path(self.temp_dir) / 'test_output2.docx'
        
        result = cli([
            'render',
            'nonexistent.md',
            str(self.template),
            str(output_path)
        ])
        
        self.assertEqual(result, 1)
    
    def test_render_missing_template(self):
        """測試 render 指令 - 模板檔案不存在"""
        if not self.sample_md.exists():
            self.skipTest(f"測試檔案不存在: {self.sample_md}")
        
        output_path = Path(self.temp_dir) / 'test_output3.docx'
        
        result = cli([
            'render',
            str(self.sample_md),
            'nonexistent.docx',
            str(output_path)
        ])
        
        self.assertEqual(result, 1)
    
    def test_render_verbose(self):
        """測試 render 指令 - verbose 模式"""
        if not self.sample_md.exists():
            self.skipTest(f"測試檔案不存在: {self.sample_md}")
        if not self.template.exists():
            self.skipTest(f"模板檔案不存在: {self.template}")
        
        output_path = Path(self.temp_dir) / 'test_output_verbose.docx'
        
        result = cli([
            'render',
            str(self.sample_md),
            str(self.template),
            str(output_path),
            '-v'
        ])
        
        self.assertEqual(result, 0)
    
    def test_render_no_validate(self):
        """測試 render 指令 - 跳過驗證"""
        if not self.sample_md.exists():
            self.skipTest(f"測試檔案不存在: {self.sample_md}")
        if not self.template.exists():
            self.skipTest(f"模板檔案不存在: {self.template}")
        
        output_path = Path(self.temp_dir) / 'test_output_noval.docx'
        
        result = cli([
            'render',
            str(self.sample_md),
            str(self.template),
            str(output_path),
            '--no-validate'
        ])
        
        self.assertEqual(result, 0)


class TestBatchCommand(unittest.TestCase):
    """批次指令測試"""
    
    @classmethod
    def setUpClass(cls):
        """設定測試環境"""
        cls.test_dir = Path(__file__).parent.parent
        cls.sample_dir = cls.test_dir / 'test' / 'sample_inputs'
        cls.template = cls.test_dir / 'templates' / 'simple_template.docx'
        cls.temp_dir = tempfile.mkdtemp()
    
    def test_batch_command(self):
        """測試 batch 批次指令"""
        if not self.sample_dir.exists():
            self.skipTest(f"測試目錄不存在: {self.sample_dir}")
        if not self.template.exists():
            self.skipTest(f"模板檔案不存在: {self.template}")
        
        output_dir = Path(self.temp_dir) / 'batch_output'
        
        result = cli([
            'batch',
            str(self.sample_dir),
            str(self.template),
            str(output_dir)
        ])
        
        self.assertEqual(result, 0)
        
        # 檢查輸出檔案
        output_files = list(output_dir.glob('*.docx'))
        self.assertGreater(len(output_files), 0)
    
    def test_batch_missing_dir(self):
        """測試 batch 指令 - 目錄不存在"""
        result = cli([
            'batch',
            'nonexistent_dir',
            str(self.template),
            self.temp_dir
        ])
        
        self.assertEqual(result, 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
