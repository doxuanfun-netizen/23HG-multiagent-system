# -*- coding: utf-8 -*-
"""
Kiểm thử bộ xử lý ngôn ngữ tự nhiên NLPProcessor (nlp_processor.py)
Kiểm tra khả năng trích xuất màu sắc, tọa độ, số học và phân loại ý định hình học từ văn bản tự nhiên.
"""
import unittest
from cad_cowork_mcp.nlp_processor import NLPProcessor


class TestNLPProcessor(unittest.TestCase):
    """Kiểm tra các hàm trích xuất dữ liệu của NLPProcessor."""

    def setUp(self):
        self.nlp = NLPProcessor()

    def test_extract_color_vietnamese_and_english(self):
        """Trích xuất đúng mã màu ACI (AutoCAD Color Index) từ tên màu tiếng Việt và tiếng Anh."""
        # Tiếng Việt
        self.assertEqual(self.nlp.extract_color_from_command("vẽ đường thẳng màu đỏ"), 1)
        self.assertEqual(self.nlp.extract_color_from_command("vẽ hình tròn màu vàng"), 2)
        self.assertEqual(self.nlp.extract_color_from_command("vẽ hình chữ nhật màu xanh lá"), 3)
        self.assertEqual(self.nlp.extract_color_from_command("tô màu đen"), 250)

        # Tiếng Anh
        self.assertEqual(self.nlp.extract_color_from_command("draw circle with red color"), 1)
        self.assertEqual(self.nlp.extract_color_from_command("line green"), 3)
        self.assertEqual(self.nlp.extract_color_from_command("color cyan"), 4)

        # Mặc định khi không có màu -> 7 (White)
        self.assertEqual(self.nlp.extract_color_from_command("vẽ hình tròn bán kính 50"), 7)

    def test_extract_coordinates_2d_and_3d(self):
        """Trích xuất danh sách điểm tọa độ (x, y, z) từ nhiều định dạng cú pháp."""
        # Tọa độ 2D trong ngoặc
        coords_2d = self.nlp._extract_coordinates("vẽ đoạn thẳng từ (100, 200) đến (300, 400)")
        self.assertEqual(len(coords_2d), 2)
        self.assertEqual(coords_2d[0], (100.0, 200.0, 0.0))
        self.assertEqual(coords_2d[1], (300.0, 400.0, 0.0))

        # Tọa độ 3D
        coords_3d = self.nlp._extract_coordinates("tại điểm 10.5, 20.5, 50.0")
        self.assertEqual(len(coords_3d), 1)
        self.assertEqual(coords_3d[0], (10.5, 20.5, 50.0))

    def test_extract_numbers(self):
        """Trích xuất chính xác các số nguyên và số thực từ chuỗi văn bản."""
        nums = self.nlp._extract_numbers("bán kính 50.5 tại vị trí -10.2 và 100")
        self.assertEqual(nums, [50.5, -10.2, 100.0])


if __name__ == "__main__":
    unittest.main()
