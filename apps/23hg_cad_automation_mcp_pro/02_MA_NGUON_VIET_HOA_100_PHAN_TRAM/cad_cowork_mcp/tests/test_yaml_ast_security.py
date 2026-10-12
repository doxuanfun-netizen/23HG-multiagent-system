# -*- coding: utf-8 -*-
import unittest
import sys
import os

# Thêm đường dẫn cha để import module
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from command_router.yaml_loader import _compile_trigger

class TestYamlAstSecurity(unittest.TestCase):
    """Kiểm tra việc biên dịch biểu thức điều kiện trigger bằng AST thay cho eval()."""

    def test_safe_in_expression(self):
        trigger = _compile_trigger("'ve_hinh' in intent")
        self.assertTrue(trigger("Toi muon ve_hinh chu nhat"))
        self.assertFalse(trigger("Toi muon sua doi tuong"))

    def test_safe_and_not_expression(self):
        trigger = _compile_trigger("'circle' in intent and 'block' not in intent")
        self.assertTrue(trigger("draw circle please"))
        self.assertFalse(trigger("draw circle block please"))

    def test_reject_dangerous_tokens(self):
        # Biểu thức cố tình gọi hàm hoặc import phải bị AST từ chối
        with self.assertRaises(Exception):
            _compile_trigger("__import__('os').system('calc')")

if __name__ == "__main__":
    unittest.main()
