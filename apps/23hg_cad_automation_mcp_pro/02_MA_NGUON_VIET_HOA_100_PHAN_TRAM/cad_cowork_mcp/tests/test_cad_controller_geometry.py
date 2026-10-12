# -*- coding: utf-8 -*-
"""
Kiểm thử các hàm tính toán hình học và thao tác đối tượng CADController (cad_controller.py)
Kiểm tra lineweight, chuyển đổi tọa độ 2D/3D, gọi COM ModelSpace và cơ chế tự động sao lưu.
"""
import unittest
import math
from unittest.mock import MagicMock, patch
from cad_cowork_mcp.cad_controller import CADController


class TestCADControllerGeometry(unittest.TestCase):
    """Kiểm tra xử lý hình học và kiểm soát dữ liệu trong CADController."""

    def setUp(self):
        self.controller = CADController()
        self.mock_doc = MagicMock()
        self.mock_app = MagicMock()
        self.mock_app.ActiveDocument = self.mock_doc
        self.controller.app = self.mock_app
        self.controller.doc = self.mock_doc

    def test_validate_lineweight(self):
        """Xác thực độ dày nét (lineweight): giá trị hợp lệ được giữ nguyên, giá trị sai lệch quy về 0."""
        # None -> None
        self.assertIsNone(self.controller.validate_lineweight(None))
        # Hợp lệ (chuẩn AutoCAD ACI lineweights)
        self.assertEqual(self.controller.validate_lineweight(25), 25)
        self.assertEqual(self.controller.validate_lineweight(50), 50)
        self.assertEqual(self.controller.validate_lineweight(0), 0)
        # Bất hợp lệ (ví dụ: 17, 999) -> quy về 0
        self.assertEqual(self.controller.validate_lineweight(17), 0)
        self.assertEqual(self.controller.validate_lineweight(999), 0)

    def test_draw_circle_parameters(self):
        """draw_circle chuyển đổi tâm và bán kính chính xác sang COM AddCircle."""
        self.controller.draw_circle(center=(100, 200, 0), radius=50.0)
        self.mock_doc.ModelSpace.AddCircle.assert_called_once()
        args, kwargs = self.mock_doc.ModelSpace.AddCircle.call_args
        self.assertEqual(args[1], 50.0)

    def test_draw_line_2d_to_3d_conversion(self):
        """draw_line tự động nâng cấp tọa độ 2D (x, y) thành 3D (x, y, 0)."""
        self.controller.draw_line(start_point=(10, 20), end_point=(100, 200))
        self.mock_doc.ModelSpace.AddLine.assert_called_once()

    def test_draw_rectangle_corners(self):
        """draw_rectangle tính toán đúng 4 góc khép kín và ủy quyền cho draw_polyline."""
        with patch.object(self.controller, "draw_polyline") as mock_poly:
            self.controller.draw_rectangle(corner1=(0, 0), corner2=(100, 50))
            mock_poly.assert_called_once()
            points_arg = mock_poly.call_args[0][0]
            # 5 điểm khép kín: (0,0), (100,0), (100,50), (0,50), (0,0)
            self.assertEqual(len(points_arg), 5)
            self.assertEqual(points_arg[0], (0, 0, 0))
            self.assertEqual(points_arg[1], (100, 0, 0))
            self.assertEqual(points_arg[2], (100, 50, 0))
            self.assertEqual(points_arg[3], (0, 50, 0))
            self.assertEqual(points_arg[4], (0, 0, 0))
            self.assertTrue(mock_poly.call_args[0][1])  # closed = True

    def test_draw_arc_angles_radians(self):
        """draw_arc tự động đổi độ (degrees) sang radian cho COM AddArc."""
        self.controller.draw_arc(center=(0, 0, 0), radius=25, start_angle=0, end_angle=180)
        self.mock_doc.ModelSpace.AddArc.assert_called_once()
        args = self.mock_doc.ModelSpace.AddArc.call_args[0]
        # Bán kính ở vị trí 1
        self.assertEqual(args[1], 25)
        # start_angle: 0 deg = 0 rad
        self.assertAlmostEqual(args[2], 0.0)
        # end_angle: 180 deg = pi rad
        self.assertAlmostEqual(args[3], math.pi)


if __name__ == "__main__":
    unittest.main()
