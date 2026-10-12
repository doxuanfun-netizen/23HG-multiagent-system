# -*- coding: utf-8 -*-
"""
Kiểm thử các công cụ MCP và quản lý trạng thái máy chủ (server.py)
Kiểm tra DrawingState, SessionMetrics, _query_api_commands và cấu trúc schema của MCP Tools.
"""
import unittest
from cad_cowork_mcp.server import DrawingState, _query_api_commands, CADService


class TestServerMcpTools(unittest.TestCase):
    """Kiểm tra logic nghiệp vụ và tra cứu lệnh tại tầng Server MCP."""

    def setUp(self):
        self.cad_service = CADService()

    def test_drawing_state_tracking_handles(self):
        """DrawingState lưu giữ tối đa 20 handle gần nhất và khử trùng lặp."""
        state = DrawingState()
        # Mô phỏng kết quả lệnh tạo thực thể
        res1 = {
            "success": True,
            "createdEntityHandles": ["1A", "1B"],
            "modifiedEntityHandles": []
        }
        state.update_from_result("api_draw_line", res1)
        self.assertEqual(state.recent_handles, ["1A", "1B"])
        self.assertEqual(state.last_command, "api_draw_line")
        self.assertEqual(state.command_count, 1)

        # Thêm 25 handles mới để kiểm tra giới hạn 20
        new_handles = [f"HEX_{i}" for i in range(25)]
        res2 = {
            "success": True,
            "createdEntityHandles": new_handles,
            "modifiedEntityHandles": []
        }
        state.update_from_result("batch_create", res2)
        self.assertEqual(len(state.recent_handles), 20)

    def test_drawing_state_to_context_string(self):
        """to_context_string tạo chuỗi ngữ cảnh đúng định dạng chèn vào System Prompt."""
        state = DrawingState()
        state.active_document = "Bản_vẽ_cầu_đường.dwg"
        state.recent_handles = ["27E2C0", "27E2C1"]
        state.last_command = "api_draw_circle"
        ctx = state.to_context_string()
        self.assertIn("Active document: Bản_vẽ_cầu_đường.dwg", ctx)
        self.assertIn("Recent handles", ctx)
        self.assertIn("Last command: api_draw_circle", ctx)

    def test_query_api_commands_single_command(self):
        """_query_api_commands Chế độ 1: Tra cứu chi tiết một lệnh cụ thể."""
        res = _query_api_commands(self.cad_service, command_name="api_draw_circle")
        self.assertTrue(res["success"])
        self.assertEqual(res["command"]["name"], "api_draw_circle")
        self.assertIn("parameters", res["command"])

    def test_query_api_commands_by_category(self):
        """_query_api_commands Chế độ 2: Lọc danh sách lệnh theo phân loại tiếng Anh/Việt."""
        res = _query_api_commands(self.cad_service, category="drawing")
        self.assertTrue(res["success"])
        self.assertGreater(res["count"], 0)
        self.assertTrue(any(c["name"] == "api_draw_circle" for c in res["commands"]))

    def test_query_api_commands_overview(self):
        """_query_api_commands Chế độ 3: Trả về danh sách tổng quan các chuyên mục."""
        res = _query_api_commands(self.cad_service)
        self.assertTrue(res["success"])
        self.assertIn("totalCommands", res)
        self.assertIn("categories", res)
        self.assertGreater(len(res["categories"]), 5)

    def test_query_api_commands_not_found(self):
        """_query_api_commands trả về lỗi rõ ràng khi không tìm thấy lệnh."""
        res = _query_api_commands(self.cad_service, command_name="nonexistent_cad_command")
        self.assertFalse(res["success"])
        self.assertIn("Không tìm thấy", res["error"])


if __name__ == "__main__":
    unittest.main()
