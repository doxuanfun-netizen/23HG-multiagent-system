# -*- coding: utf-8 -*-
import unittest
import json
import os
import re

class TestManifestSchema(unittest.TestCase):
    """Kiểm tra tính toàn vẹn của danh mục lệnh và loại bỏ hoàn toàn dịch máy lỗi."""

    def setUp(self):
        self.manifest_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "api_commands_manifest.json"
        )
        with open(self.manifest_path, "r", encoding="utf-8") as f:
            self.data = json.load(f)

    def test_categories_name_en_valid(self):
        """Đảm bảo không còn các danh mục dịch máy lỗi như 'Ôn lại', 'Lớp', 'xem', 'tài liệu'."""
        bad_name_en = ["Ôn lại", "Lớp", "xem", "tài liệu"]
        for cat in self.data.get("categories", []):
            name_en = cat.get("nameEn")
            self.assertNotIn(
                name_en,
                bad_name_en,
                f"Danh mục vẫn còn bị dịch máy lỗi: {name_en}"
            )

    def test_no_chinese_characters(self):
        """Đảm bảo không còn ký tự tiếng Trung hoặc dấu câu tiếng Trung trong toàn bộ file manifest."""
        with open(self.manifest_path, "r", encoding="utf-8") as f:
            content = f.read()

        chinese_chars = re.findall(r'[\u4e00-\u9fff\uff00-\uffef]', content)
        self.assertEqual(
            len(chinese_chars),
            0,
            f"Vẫn còn sót {len(chinese_chars)} ký tự/dấu câu tiếng Trung trong manifest!"
        )

if __name__ == "__main__":
    unittest.main()
