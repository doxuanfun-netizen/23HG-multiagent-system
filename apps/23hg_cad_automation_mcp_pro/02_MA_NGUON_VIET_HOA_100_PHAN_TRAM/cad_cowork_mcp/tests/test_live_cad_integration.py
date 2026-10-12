# -*- coding: utf-8 -*-
"""
Kiểm thử tích hợp trực tiếp với AutoCAD đang chạy (Live CAD COM Integration Test)
Chạy trực tiếp trên phiên bản AutoCAD thực tế trên máy tính (nếu CAD đang mở).
Nếu AutoCAD không chạy, kiểm thử tự động bỏ qua (Skip) một cách an toàn mà không làm gãy pipeline CI.
"""
import unittest
from cad_cowork_mcp.cad_controller import CADController


def is_live_cad_available() -> bool:
    """Kiểm tra xem có ứng dụng CAD thực tế nào đang mở sẵn hay không."""
    try:
        controller = CADController()
        return controller.is_running()
    except Exception:
        return False


class TestLiveCADIntegration(unittest.TestCase):
    """Kiểm thử tích hợp thực tế với AutoCAD qua Native COM."""

    @classmethod
    def setUpClass(cls):
        cls.controller = CADController()
        cls.cad_running = cls.controller.start_cad()

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "controller") and cls.controller:
            cls.controller.disconnect()

    def test_live_connection_status(self):
        """Xác định trạng thái kết nối trực tiếp với AutoCAD trên máy người dùng."""
        status = self.controller.is_running()
        self.assertIsInstance(status, bool)

    def test_live_document_info_query(self):
        """Đọc thông tin bản vẽ từ AutoCAD thực tế (nếu đang chạy)."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        res = self.controller.get_document_info()
        self.assertTrue(res.get("success"), f"Truy xuất live document info thất bại: {res}")
        self.assertIn("data", res)
        self.assertIn("name", res["data"])
        self.assertEqual(res.get("executionMethod"), "native_com")

    def test_live_get_layers_query(self):
        """Đọc danh sách Layer từ AutoCAD thực tế qua Native COM."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        res = self.controller.get_layers()
        self.assertTrue(res.get("success"), f"Truy xuất live layers thất bại: {res}")
        self.assertIn("layers", res.get("data", {}))
        # Bản vẽ AutoCAD luôn có tối thiểu Layer '0'
        layer_names = [l["name"] for l in res["data"]["layers"]]
        self.assertIn("0", layer_names)

    def test_live_get_active_layer(self):
        """Lấy layer hiện hành từ AutoCAD thực tế."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        res = self.controller.get_active_layer()
        self.assertTrue(res.get("success"))
        self.assertTrue(bool(res["data"]["name"]))

    def test_live_get_units(self):
        """Đọc đơn vị bản vẽ thực tế và hệ số quy đổi m/m2 từ AutoCAD."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        res = self.controller.get_units()
        self.assertTrue(res.get("success"), f"Lỗi đọc units thực tế: {res}")
        self.assertIn("insunits", res["data"])
        self.assertIn("scale_to_m", res["data"])
        self.assertEqual(res.get("executionMethod"), "native_com")

    def test_live_quantity_by_layer(self):
        """Đo bóc hình học theo layer từ bản vẽ thực tế trên AutoCAD."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        res = self.controller.quantity_by_layer()
        self.assertTrue(res.get("success"), f"Lỗi quantity_by_layer thực tế: {res}")
        self.assertIn("layers", res["data"])
        self.assertIn("grand_total_length_m", res["data"])
        self.assertIn("grand_total_area_m2", res["data"])
        self.assertEqual(res.get("executionMethod"), "native_com")

    def test_live_count_blocks(self):
        """Thống kê block từ bản vẽ thực tế trên AutoCAD."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        res = self.controller.count_blocks()
        self.assertTrue(res.get("success"), f"Lỗi count_blocks thực tế: {res}")
        self.assertIn("blocks", res["data"])
        self.assertIn("total_blocks_found", res["data"])
        self.assertEqual(res.get("executionMethod"), "native_com")

    def test_live_export_boq(self):
        """Xuất file Excel BoQ thực tế từ AutoCAD với 100% công thức sống."""
        if not self.cad_running:
            self.skipTest("AutoCAD không mở sẵn trên máy, bỏ qua kiểm thử Live COM.")

        import tempfile, os
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_file = os.path.join(tmp_dir, "LIVE_BOQ_TEST.xlsx")
            res = self.controller.export_boq(file_path=out_file, format="xlsx")
            self.assertTrue(res.get("success"), f"Lỗi export_boq thực tế: {res}")
            self.assertTrue(os.path.exists(out_file))


if __name__ == "__main__":
    unittest.main()
