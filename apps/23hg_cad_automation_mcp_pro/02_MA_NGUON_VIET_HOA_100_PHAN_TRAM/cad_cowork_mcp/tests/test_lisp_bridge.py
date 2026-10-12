# -*- coding: utf-8 -*-
"""
Kiểm thử tệp AutoLISP Bridge (04_PLUGIN_CAD_NETLOAD/23hg_bridge.lsp)
Kiểm tra tính toàn vẹn cú pháp LISP, cân bằng dấu ngoặc đơn, các lệnh được định nghĩa và tính an toàn mã nguồn.
"""
import unittest
import os
import re


class TestLispBridge(unittest.TestCase):
    """Kiểm tra hợp thức và an toàn của 23hg_bridge.lsp."""

    @classmethod
    def setUpClass(cls):
        # Đường dẫn tới file lsp trong repo
        test_dir = os.path.dirname(__file__)
        repo_root = os.path.abspath(os.path.join(test_dir, "..", "..", ".."))
        cls.lsp_path = os.path.join(repo_root, "04_PLUGIN_CAD_NETLOAD", "23hg_bridge.lsp")

    def test_lsp_file_exists(self):
        """File 23hg_bridge.lsp tồn tại và có dung lượng hợp lệ."""
        self.assertTrue(os.path.isfile(self.lsp_path), f"Không tìm thấy file: {self.lsp_path}")
        self.assertGreater(os.path.getsize(self.lsp_path), 50)

    def test_parentheses_balance(self):
        """Cú pháp AutoLISP phải cân bằng tuyệt đối số lượng dấu mở ngoặc '(' và đóng ngoặc ')'."""
        with open(self.lsp_path, "r", encoding="utf-8") as f:
            content = f.read()

        # Loại bỏ các chuỗi chuỗi và comment trước khi đếm ngoặc
        clean_content = re.sub(r";;;.*", "", content)
        clean_content = re.sub(r'";.*', "", clean_content)
        open_count = clean_content.count("(")
        close_count = clean_content.count(")")
        self.assertEqual(
            open_count, close_count,
            f"Lỗi cú pháp AutoLISP: Số dấu '(' ({open_count}) khác số dấu ')' ({close_count})"
        )

    def test_defined_commands_present(self):
        """Đảm bảo các lệnh hỗ trợ c:AI_DRAW và c:23HG_STATUS được định nghĩa chuẩn xác."""
        with open(self.lsp_path, "r", encoding="utf-8") as f:
            content = f.read()

        self.assertIn("c:AI_DRAW", content)
        self.assertIn("c:23HG_STATUS", content)

    def test_lsp_no_malicious_system_calls(self):
        """Bridge LISP tuyệt đối không chứa các hàm nguy hiểm như startapp, vl-file-delete."""
        with open(self.lsp_path, "r", encoding="utf-8") as f:
            content = f.read().lower()

        forbidden_tokens = ["startapp", "vl-file-delete", "vl-file-rename", "format c:"]
        for token in forbidden_tokens:
            self.assertNotIn(token, content, f"Phát hiện token nguy hiểm trong LISP bridge: {token}")


if __name__ == "__main__":
    unittest.main()
