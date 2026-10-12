# -*- coding: utf-8 -*-
"""
UNIT TESTS — 23HG CAD AUTOMATION MCP PRO INTEGRATION
Kiểm tra cấu trúc và tính sẵn sàng của phân hệ MCP CAD v2.0.0 (Native COM Interop).
"""

import json
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CAD_APP_DIR = os.path.join(ROOT, "apps", "23hg_cad_automation_mcp_pro")
SRC_DIR = os.path.join(CAD_APP_DIR, "02_MA_NGUON_VIET_HOA_100_PHAN_TRAM")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)


class TestCadAutomationMcpPro(unittest.TestCase):

    def setUp(self):
        self.app_dir = CAD_APP_DIR

    def test_directory_structure_and_modules(self):
        """Kiểm tra sự tồn tại của 5 thư mục cốt lõi và README.md."""
        expected_dirs = [
            "01_CAI_DAT_TU_DONG_1_CLICK",
            "02_MA_NGUON_VIET_HOA_100_PHAN_TRAM",
            "03_FILE_CAU_HINH_CAC_PHAN_MEM_AI",
            "04_PLUGIN_CAD_NETLOAD",
            "05_HUONG_DAN_SU_DUNG_VA_DANH_MUC_LENH",
        ]
        for d in expected_dirs:
            p = os.path.join(self.app_dir, d)
            self.assertTrue(os.path.isdir(p), f"Thư mục {d} phải tồn tại trong apps/23hg_cad_automation_mcp_pro")

        readme = os.path.join(self.app_dir, "README.md")
        self.assertTrue(os.path.isfile(readme), "README.md phải tồn tại")
        with open(readme, "r", encoding="utf-8") as f:
            content = f.read()
        self.assertIn("23HG CAD AUTOMATION MCP PRO", content)
        self.assertIn("v2.0.0", content)

    def test_ai_configurations_valid(self):
        """Kiểm tra các file cấu hình MCP cho Claude Desktop, Antigravity, Cursor, Windsurf."""
        cfg_dir = os.path.join(self.app_dir, "03_FILE_CAU_HINH_CAC_PHAN_MEM_AI")
        configs = [
            "claude_desktop_config.json",
            "antigravity_mcp_config.json",
            "cursor_mcp_config.json",
            "windsurf_mcp_config.json",
        ]
        for cfg in configs:
            p = os.path.join(cfg_dir, cfg)
            self.assertTrue(os.path.isfile(p), f"File {cfg} phải tồn tại")
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            self.assertIn("mcpServers", data)
            self.assertIn("cad-mcp", data["mcpServers"])
            server = data["mcpServers"]["cad-mcp"]
            self.assertEqual(server.get("command"), "python")
            self.assertIn("cad_cowork_mcp", server.get("args", []))

    def test_cad_cowork_mcp_config_json(self):
        """Kiểm tra file config.json của gói cad_cowork_mcp."""
        cfg_path = os.path.join(SRC_DIR, "cad_cowork_mcp", "config.json")
        self.assertTrue(os.path.isfile(cfg_path), "cad_cowork_mcp/config.json phải tồn tại")
        with open(cfg_path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        self.assertIn("server", cfg)
        self.assertIn("cad", cfg)
        self.assertEqual(cfg["server"].get("version"), "2.0.0")

    def test_cad_controller_import_and_takeoff_interface(self):
        """Kiểm tra import CADController và giao diện bóc tách khối lượng QS Takeoff."""
        from cad_cowork_mcp.cad_controller import CADController, SAFE_COMMANDS_WHITELIST

        self.assertIn("LINE", SAFE_COMMANDS_WHITELIST)
        self.assertIn("CIRCLE", SAFE_COMMANDS_WHITELIST)
        self.assertIn("PLINE", SAFE_COMMANDS_WHITELIST)

        ctrl = CADController()
        self.assertIsNone(ctrl.app)
        self.assertIsNone(ctrl.doc)
        # Các phương thức API chuyên sâu phải có mặt
        self.assertTrue(hasattr(ctrl, "get_units"))
        self.assertTrue(hasattr(ctrl, "quantity_by_layer"))
        self.assertTrue(hasattr(ctrl, "count_blocks"))
        self.assertTrue(hasattr(ctrl, "export_boq"))
        self.assertTrue(hasattr(ctrl, "smart_cad_command"))
        self.assertTrue(hasattr(ctrl, "execute_api_command"))

        # Kiểm tra lớp BatchExecutor cho chuỗi lệnh phụ thuộc DAG
        from cad_cowork_mcp.batch_executor import BatchExecutor
        executor = BatchExecutor(ctrl)
        self.assertTrue(hasattr(executor, "execute"))


if __name__ == "__main__":
    unittest.main()
