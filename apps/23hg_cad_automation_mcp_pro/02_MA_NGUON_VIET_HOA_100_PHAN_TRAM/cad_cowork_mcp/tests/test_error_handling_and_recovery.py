# -*- coding: utf-8 -*-
"""
Kiểm thử các nhánh xử lý lỗi, phục hồi và chịu lỗi (Error Handling, Fault Tolerance & Recovery)
Kiểm tra chi tiết cơ chế phát hiện ngắt kết nối COM RPC_E_DISCONNECTED, chặn hộp thoại Modal Dialog,
bảo vệ timeout, cơ chế gợi ý sửa lỗi fix_hint và khả năng tự phục hồi của CADController và Server.
"""
import unittest
from unittest.mock import MagicMock, patch
from cad_cowork_mcp.cad_controller import (
    CADController,
    _is_com_disconnect
)
from cad_cowork_mcp.server import _get_command_fix_hint


class TestErrorHandlingAndRecovery(unittest.TestCase):
    """Kiểm tra các kịch bản lỗi biên và phục hồi sự cố trong CADController và Server."""

    def setUp(self):
        self.controller = CADController()
        self.mock_doc = MagicMock()
        self.mock_app = MagicMock()
        self.mock_app.ActiveDocument = self.mock_doc
        self.controller.app = self.mock_app
        self.controller.doc = self.mock_doc

    def test_detect_com_disconnection_errors(self):
        """Hàm _is_com_disconnect nhận diện chính xác các mã lỗi mất kết nối COM RPC."""
        self.assertTrue(_is_com_disconnect("Error -2147417848: The object invoked has disconnected"))
        self.assertTrue(_is_com_disconnect("RPC_E_DISCONNECTED"))
        self.assertTrue(_is_com_disconnect("COM ngắt kết nối đột ngột"))
        self.assertTrue(_is_com_disconnect("Đã ngắt kết nối với máy khách của nó"))
        self.assertFalse(_is_com_disconnect("Invalid parameter value"))
        self.assertFalse(_is_com_disconnect(""))

    def test_modal_dialog_blocks_command_to_prevent_hang(self):
        """AutoCAD đang kẹt hộp thoại Modal Dialog (is_quiescent=False) sẽ từ chối gửi lệnh để tránh treo ứng dụng."""
        # Giả lập trạng thái AutoCAD đang bận hộp thoại Modal
        self.controller.is_quiescent = MagicMock(return_value=False)
        res = self.controller.send_command("_.CIRCLE 0,0 50")
        self.assertFalse(res["success"])
        self.assertIn("Modal Dialog", res["message"])

    def test_com_disconnect_auto_recovery_retry(self):
        """Khi gặp lỗi ngắt kết nối COM, Gateway tự động thử kết nối lại và gửi lại lệnh thành công."""
        self.controller.is_quiescent = MagicMock(return_value=True)

        # Lần gửi 1 gặp lỗi ngắt kết nối COM -2147417848, lần gửi 2 sau khi reconnect thành công
        self.mock_doc.SendCommand.side_effect = [
            Exception("(-2147417848, 'The object invoked has disconnected from its clients.', None, None)"),
            None  # Thành công
        ]

        with patch.object(self.controller, "_try_reconnect", return_value=True):
            res = self.controller.send_command("_.LINE 0,0 100,100")
            self.assertTrue(res["success"])
            self.assertIn("sau khi kết nối lại", res["message"])

    def test_server_fix_hint_generation(self):
        """Hàm _get_command_fix_hint tự động sinh hướng dẫn tham số chi tiết khi lệnh API thất bại."""
        hint = _get_command_fix_hint("api_draw_circle")
        self.assertIsNotNone(hint)
        self.assertIn("required_params", hint)
        self.assertIn("optional_params", hint)
        self.assertIn("center", hint["required_params"])
        self.assertIn("radius", hint["required_params"])

    def test_smart_cad_command_when_cad_not_running(self):
        """Khi phần mềm CAD hoàn toàn không chạy và start_cad thất bại, trả về mã lỗi -8 chuẩn hóa."""
        fresh_controller = CADController()
        with patch.object(fresh_controller, "is_running", return_value=False), \
             patch.object(fresh_controller, "start_cad", return_value=False):
            res = fresh_controller.smart_cad_command("_.CIRCLE 0,0 50")
            self.assertFalse(res["success"])
            self.assertEqual(res["errorCode"], -8)
            self.assertIn("Không tìm thấy phần mềm CAD", res["message"])


if __name__ == "__main__":
    unittest.main()
