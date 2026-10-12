# -*- coding: utf-8 -*-
"""
Kiểm thử bộ thực thi hàng loạt BatchExecutor (DAG Dependency & Parameter Resolution)
Kiểm tra chi tiết thuật toán sắp xếp tôpô, phát hiện chu trình, giải mã biến ${id.field}, và dry_run.
"""
import unittest
from unittest.mock import MagicMock
from cad_cowork_mcp.batch_executor import (
    BatchExecutor,
    CyclicDependencyError,
    BatchExecutionError
)


class TestBatchExecutorDAG(unittest.TestCase):
    """Kiểm tra logic xử lý DAG và chuyển tiếp tham số trong BatchExecutor."""

    def setUp(self):
        self.mock_controller = MagicMock()
        self.executor = BatchExecutor(self.mock_controller)

    def test_topological_sort_linear(self):
        """Sắp xếp tôpô chuỗi phụ thuộc tuyến tính: c1 -> c2 -> c3."""
        commands = [
            {"id": "c3", "command_name": "cmd3", "depends_on": ["c2"]},
            {"id": "c1", "command_name": "cmd1"},
            {"id": "c2", "command_name": "cmd2", "depends_on": ["c1"]},
        ]
        order = self.executor._topological_sort(commands)
        self.assertEqual(order, ["c1", "c2", "c3"])

    def test_topological_sort_diamond(self):
        """Sắp xếp tôpô đồ thị kim cương (Diamond DAG): A -> B, A -> C, (B, C) -> D."""
        commands = [
            {"id": "D", "command_name": "cmdD", "depends_on": ["B", "C"]},
            {"id": "B", "command_name": "cmdB", "depends_on": ["A"]},
            {"id": "C", "command_name": "cmdC", "depends_on": ["A"]},
            {"id": "A", "command_name": "cmdA"},
        ]
        order = self.executor._topological_sort(commands)
        self.assertEqual(order[0], "A")
        self.assertEqual(order[-1], "D")
        self.assertTrue(order.index("B") > order.index("A"))
        self.assertTrue(order.index("C") > order.index("A"))

    def test_detect_cyclic_dependency(self):
        """Phát hiện lỗi phụ thuộc vòng tròn (Cyclic Dependency) và ném exception CyclicDependencyError."""
        commands = [
            {"id": "A", "command_name": "cmdA", "depends_on": ["B"]},
            {"id": "B", "command_name": "cmdB", "depends_on": ["A"]},
        ]
        with self.assertRaises(CyclicDependencyError):
            self.executor._topological_sort(commands)

    def test_dry_run_validation_success(self):
        """Chế độ dry_run kiểm tra DAG hợp lệ và xuất kế hoạch thực thi trước khi gửi lệnh vào CAD."""
        commands = [
            {"id": "c1", "command_name": "api_draw_circle", "parameters": {"radius": 50}},
            {"id": "c2", "command_name": "api_move_entity", "parameters": {"entityHandle": "${c1.handle}"}, "depends_on": ["c1"]},
        ]
        result = self.executor.dry_run(commands)
        self.assertTrue(result["valid"])
        self.assertEqual(result["execution_order"], ["c1", "c2"])
        self.assertEqual(len(result["execution_plan"]), 2)
        self.assertEqual(result["total_commands"], 2)

    def test_dry_run_detects_missing_dependency_declaration(self):
        """dry_run cảnh báo khi tham chiếu ${c1.handle} nhưng quên khai báo depends_on: ['c1']."""
        commands = [
            {"id": "c1", "command_name": "api_draw_circle"},
            {"id": "c2", "command_name": "api_move_entity", "parameters": {"handle": "${c1.handle}"}},
        ]
        result = self.executor.dry_run(commands)
        self.assertFalse(result["valid"])
        self.assertTrue(any("chưa khai báo trong depends_on" in issue for issue in result["issues"]))

    def test_parameter_resolution_deep_fields(self):
        """Giải mã thành công các cú pháp tham chiếu biến: ${id.handle}, ${id.data.field}, mảng lồng nhau."""
        self.executor.results = {
            "cmd1": {
                "success": True,
                "handle": "27E2C0",
                "data": {"area": 314.15, "center": [10, 20, 0]}
            }
        }
        self.executor.handle_map = {"cmd1": "27E2C0"}

        # 1. Trực tiếp .handle
        params_1 = {"target": "${cmd1.handle}"}
        resolved_1 = self.executor._resolve_references(params_1)
        self.assertEqual(resolved_1["target"], "27E2C0")

        # 2. Qua .data.area
        params_2 = {"calcArea": "${cmd1.data.area}"}
        resolved_2 = self.executor._resolve_references(params_2)
        self.assertEqual(resolved_2["calcArea"], 314.15)


if __name__ == "__main__":
    unittest.main()
