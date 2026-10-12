# -*- coding: utf-8 -*-
"""
Kiểm tra đảm bảo toàn bộ mã nguồn Python (*.py) không còn bất kỳ ký tự hoặc dấu câu CJK nào.
"""
import unittest
import os
import re

class TestNoCjkInSource(unittest.TestCase):
    """Quét toàn bộ cây thư mục mã nguồn *.py để xác nhận không còn ký tự CJK."""

    def test_no_cjk_characters_in_python_files(self):
        cjk_pattern = re.compile(r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]')
        pkg_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        cjk_violations = {}
        for root, dirs, files in os.walk(pkg_root):
            for f in files:
                if f.endswith('.py'):
                    fpath = os.path.join(root, f)
                    with open(fpath, 'r', encoding='utf-8', errors='ignore') as fp:
                        content = fp.read()
                    matches = cjk_pattern.findall(content)
                    if matches:
                        rel = os.path.relpath(fpath, pkg_root)
                        cjk_violations[rel] = matches

        self.assertEqual(
            len(cjk_violations),
            0,
            f"Vẫn còn file Python chứa ký tự CJK: {list(cjk_violations.keys())}"
        )

if __name__ == "__main__":
    unittest.main()
