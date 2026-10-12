# -*- coding: utf-8 -*-
"""
Bộ kiểm thử tự động hệ thống 23HG CAD Automation MCP Pro
Kiểm tra toàn diện cơ chế Security Guard (Whitelist + Blacklist + Gateway + Confirm + AutoBackup)
Chạy bằng lệnh: python -m unittest discover tests
"""
import unittest
from cad_cowork_mcp.cad_controller import (
    validate_command_security,
    CADController,
    DEFAULT_BLOCKED_REGEX,
    DESTRUCTIVE_COMMANDS,
    SAFE_COMMANDS_WHITELIST,
    REQUIRE_CONFIRM_PATTERNS
)


class TestSecurityGuard(unittest.TestCase):
    """Kiểm tra lớp bảo mật Security Guard chống lệnh nguy hiểm và kiểm soát Whitelist."""

    def test_global_blacklist_keywords(self):
        """Kiểm tra danh sách đen toàn cục chặn các hàm/lệnh nguy hiểm tiềm ẩn."""
        dangerous_commands = [
            "SHELL calc.exe",
            "DOS dir",
            "SCRIPT evil.scr",
            "(startapp \"calc.exe\")",
            "(vl-file-delete \"C:/important.dwg\")",
            "(vl-file-copy \"a.lsp\" \"b.lsp\")",
            "(vl-file-rename \"a\" \"b\")",
            "(vl-directory-files \"C:/\")",
            "(load \"evil.lsp\")",
            "NETLOAD evil.dll",
            "ARXLOAD evil.arx",
            "ERASE ALL",
            "_ERASE ALL",
            "PURGE ALL",
        ]
        for cmd in dangerous_commands:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Cơ chế bảo mật phải chặn từ khóa nguy hiểm trong lệnh: {cmd}"
                )

    def test_destructive_commands_blocked_without_confirm(self):
        """Kiểm tra chặn các lệnh CAD phá hoại, đóng hoặc thoát CAD."""
        destructive = [
            "QUIT",
            "_QUIT",
            "QUIT\n",
            "EXIT",
            "_EXIT",
            "CLOSEALL",
            "CLOSE",
            "_CLOSE",
            "SAVE",
            "QSAVE",
            "SAVEAS",
            "PURGE",
            "_PURGE",
        ]
        for cmd in destructive:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Cơ chế bảo mật phải chặn lệnh phá hoại: {cmd}"
                )

    def test_destructive_commands_require_confirm(self):
        """Kiểm tra các lệnh xóa/hủy chủ đích (ERASE, EXPLODE, -LAYER D, MOVE ALL, WBLOCK, ssget) bắt buộc có confirm=True hoặc dry_run=True."""
        dangerous_ops = [
            "ERASE",
            "_ERASE",
            "EXPLODE",
            "_EXPLODE",
            "_.EXPLODE (ssget \"X\")",
            "_.ERASE (ssget \"X\")",
            "-LAYER D OldLayer",
            "-LAYER DELETE OldLayer",
            "MOVE ALL",
            "WBLOCK",
            "_WBLOCK",
        ]
        for cmd in dangerous_ops:
            with self.subTest(cmd=cmd):
                # 1. Không có confirm và không có dry_run -> Phải bị chặn
                is_safe_default, msg = validate_command_security(cmd, confirm=False, dry_run=False)
                self.assertFalse(
                    is_safe_default,
                    f"Lệnh phá hoại có chủ đích '{cmd}' phải bị chặn khi không có confirm/dry_run"
                )

                # 2. Chế độ dry_run=True -> Cho phép kiểm thử an toàn
                is_safe_dry, _ = validate_command_security(cmd, confirm=False, dry_run=True)
                self.assertTrue(
                    is_safe_dry,
                    f"Lệnh '{cmd}' phải được cho phép ở chế độ dry_run=True"
                )

                # 3. Khi người dùng xác nhận tường minh confirm=True -> Cho phép thực thi
                is_safe_confirmed, _ = validate_command_security(cmd, confirm=True, dry_run=False)
                self.assertTrue(
                    is_safe_confirmed,
                    f"Lệnh '{cmd}' phải được cho phép khi người dùng xác nhận confirm=True"
                )

    def test_arbitrary_autolisp_rejected(self):
        """Kiểm tra nguyên tắc 1: Chặn tuyệt đối AutoLISP tùy ý gửi từ bên ngoài."""
        lisp_payloads = [
            "(princ \"Hello\")",
            "(setvar \"FILEDIA\" 0)",
            "(defun c:hack () (princ \"pwned\"))",
            "(eval (read \"(startapp \\\"calc.exe\\\")\"))",
            "(vl-load-com)",
            "  (command \"_CIRCLE\" \"0,0\" \"50\")",
        ]
        for cmd in lisp_payloads:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Mã AutoLISP tùy ý phải bị chặn theo nguyên tắc bảo mật cấu trúc: {cmd}"
                )

    def test_concatenation_bypass_blocked(self):
        """Kiểm tra chặn các nỗ lực bypass bằng cách nối lệnh qua dấu chấm phẩy (;) hoặc xuống dòng (\n)."""
        concatenated_payloads = [
            "_.CIRCLE 0,0 50;QUIT",
            "_.LINE 0,0 10,10\n_QUIT",
            "_.ZOOM _E;SHELL calc.exe",
            "_.REGEN\n_ERASE ALL",
            "_.PLINE 0,0 100,0 100,100 C;FORMAT C:",
        ]
        for cmd in concatenated_payloads:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Nỗ lực nối chuỗi lệnh độc hại phải bị phát hiện và chặn đứng: {cmd}"
                )

    def test_mixed_case_bypass_blocked(self):
        """Kiểm tra chặn các chuỗi hoa thường lẫn lộn (mixed case evasion)."""
        evasion_payloads = [
            "sTaRtApP calc.exe",
            "_eRaSe aLl",
            "qUiT",
            "cLoSeAlL",
            "sAvEaS",
        ]
        for cmd in evasion_payloads:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Biến thể hoa thường của từ khóa nguy hiểm phải bị chặn: {cmd}"
                )

    def test_whitelist_safe_commands_allowed(self):
        """Kiểm tra danh mục Whitelist cho phép các lệnh vẽ, hiệu chỉnh, layer và truy vấn hợp lệ."""
        safe_commands = [
            "_.ZOOM _E",
            "_.REGEN",
            "_.CIRCLE 0,0 50",
            "_.PLINE 0,0 100,0 100,100 C",
            "LINE 0,0 50,50",
            "ARC 0,0 10,10 20,0",
            "RECTANG 0,0 100,50",
            "_.-LAYER M NEW_LAYER C 1 NEW_LAYER ",
            "DIMLINEAR 0,0 100,0 50,20",
            "DIMALIGNED 0,0 100,100 50,70",
            "DIMRADIUS 0,0 50",
            "MOVE 10,10 20,20",
            "COPY 0,0 50,50",
            "ROTATE 0,0 45",
            "SCALE 0,0 2",
            "OFFSET 5",
            "FILLET",
            "CHAMFER",
            "TRIM",
            "EXTEND",
            "DIST 0,0 100,100",
            "AREA",
            "ID 50,50",
            "LIST",
            "PAN",
            "_.-INSERT \"SampleBlock\" 0,0 1 1 0",
            "-INSERT",
            "HATCH",
            "TEXT 0,0 2.5 0 SampleText",
            "MTEXT 0,0 100,50 SampleText",
        ]
        for cmd in safe_commands:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertTrue(
                    is_safe,
                    f"Lệnh an toàn nằm trong Whitelist không được bị chặn nhầm: {cmd} (Lý do: {msg})"
                )

    def test_unapproved_commands_blocked(self):
        """Kiểm tra các lệnh không thuộc Whitelist phải bị từ chối mặc định."""
        unapproved = [
            "UNKNOWN_CAD_CMD 1 2 3",
            "FORMAT C:",
            "DEL *.dwg",
            "NETSTAT -ano",
            "POWERSHELL Get-Process",
        ]
        for cmd in unapproved:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Lệnh không thuộc Whitelist phải bị chặn: {cmd}"
                )

    def test_safe_word_in_text_argument_not_false_positive(self):
        """Kiểm tra chuỗi Text chứa từ nhạy cảm như 'close' không bị chặn nhầm."""
        cmd = "_.TEXT 0,0 2.5 0 \"Please close the front door\""
        is_safe, msg = validate_command_security(cmd)
        self.assertTrue(
            is_safe,
            f"Văn bản chứa từ 'close' trong lệnh TEXT không được bị chặn nhầm: {cmd} (Lý do: {msg})"
        )

    def test_native_com_methods_exist(self):
        """Kiểm tra các hàm Native COM Interop thay thế triệt để DLL cũ tồn tại đầy đủ."""
        controller = CADController()
        self.assertTrue(callable(getattr(controller, "get_document_info", None)))
        self.assertTrue(callable(getattr(controller, "get_layers", None)))
        self.assertTrue(callable(getattr(controller, "get_active_layer", None)))
        self.assertTrue(callable(getattr(controller, "get_entities_summary", None)))
        self.assertTrue(callable(getattr(controller, "backup_current_drawing", None)))
        self.assertTrue(callable(getattr(controller, "get_units", None)))
        self.assertTrue(callable(getattr(controller, "quantity_by_layer", None)))
        self.assertTrue(callable(getattr(controller, "count_blocks", None)))
        self.assertTrue(callable(getattr(controller, "export_boq", None)))

    def test_explode_all_requires_confirm_in_gateway(self):
        """Kiểm tra lệnh _.EXPLODE (ssget \"X\") bị chặn ở Gateway khi không có confirm=True."""
        controller = CADController()
        # 1. confirm=False -> Bị chặn ngay từ tầng Gateway an ninh
        res_no_confirm = controller.send_command("_.EXPLODE (ssget \"X\")", confirm=False)
        self.assertFalse(res_no_confirm.get("success"))
        msg = res_no_confirm.get("message", "")
        self.assertTrue("confirm=True" in msg or "nguy cơ phá hoại" in msg)

        # 2. dry_run=True -> Cho phép kiểm thử an toàn
        res_dry = controller.send_command("_.EXPLODE (ssget \"X\")", dry_run=True)
        self.assertTrue(res_dry.get("success"))
        self.assertTrue(res_dry.get("dry_run"))

    def test_native_com_query_methods_without_dll(self):
        """Kiểm tra get_document_info, get_layers, get_active_layer, get_entities_summary trả về dữ liệu chuẩn qua Native COM 100% không cần DLL."""
        from unittest.mock import MagicMock
        controller = CADController()
        mock_doc = MagicMock()
        mock_doc.Name = "Drawing1.dwg"
        mock_doc.FullName = "C:/Drawings/Drawing1.dwg"
        mock_doc.Path = "C:/Drawings"
        mock_doc.Saved = True
        mock_doc.ModelSpace.Count = 42
        mock_doc.ActiveLayer.Name = "0"
        mock_doc.Layers.Count = 2

        layer1 = MagicMock()
        layer1.Name = "0"
        layer1.color = 7
        layer2 = MagicMock()
        layer2.Name = "Defpoints"
        layer2.color = 5
        mock_doc.Layers.Item.side_effect = [layer1, layer2]

        controller.app = MagicMock()
        controller.app.ActiveDocument = mock_doc
        controller.doc = mock_doc

        # 1. get_document_info
        res_doc = controller.get_document_info()
        self.assertTrue(res_doc["success"])
        self.assertEqual(res_doc["data"]["name"], "Drawing1.dwg")
        self.assertEqual(res_doc["data"]["entityCount"], 42)
        self.assertEqual(res_doc["executionMethod"], "native_com")

        # 2. get_layers
        res_layers = controller.get_layers()
        self.assertTrue(res_layers["success"])
        self.assertEqual(res_layers["data"]["count"], 2)
        self.assertEqual(res_layers["executionMethod"], "native_com")

        # 3. get_active_layer
        res_al = controller.get_active_layer()
        self.assertTrue(res_al["success"])
        self.assertEqual(res_al["data"]["name"], "0")

        # 4. get_entities_summary
        mock_ent1 = MagicMock()
        mock_ent1.ObjectName = "AcDbLine"
        mock_ent2 = MagicMock()
        mock_ent2.ObjectName = "AcDbCircle"
        mock_doc.ModelSpace.Count = 2
        mock_doc.ModelSpace.Item.side_effect = [mock_ent1, mock_ent2]
        res_summary = controller.get_entities_summary()
        self.assertTrue(res_summary["success"])
        self.assertEqual(res_summary["data"]["totalCount"], 2)
        self.assertIn("Line", res_summary["data"]["breakdown"])
        self.assertIn("Circle", res_summary["data"]["breakdown"])

    def test_smart_cad_command_confirm_guard(self):
        """Kiểm tra smart_cad_command truyền cờ confirm xuống SendCommand và chặn lệnh phá hoại khi confirm=False."""
        from unittest.mock import MagicMock
        controller = CADController()
        mock_doc = MagicMock()
        controller.app = MagicMock()
        controller.app.ActiveDocument = mock_doc
        controller.doc = mock_doc
        controller.is_quiescent = MagicMock(return_value=True)

        # 1. _.EXPLODE (ssget "X") khi confirm=False -> Bị chặn
        res_fail = controller.smart_cad_command("_.EXPLODE (ssget \"X\")", confirm=False)
        self.assertFalse(res_fail["success"])
        msg = res_fail.get("message", "")
        self.assertTrue("confirm=True" in msg or "nguy cơ phá hoại" in msg)

        # 2. _.EXPLODE (ssget "X") khi confirm=True -> Được phép gửi
        res_ok = controller.smart_cad_command("_.EXPLODE (ssget \"X\")", confirm=True)
        self.assertTrue(res_ok["success"])
        mock_doc.SendCommand.assert_called_with("_.EXPLODE (ssget \"X\")\n")

        # 3. Lệnh thông thường _.-INSERT không cần confirm
        res_insert = controller.smart_cad_command("_.-INSERT \"Door\" 0,0 1 1 0")
        self.assertTrue(res_insert["success"])

    def test_batch_executor_confirm_support(self):
        """Kiểm tra BatchExecutor nhận và truyền tham số confirm xuống từng lệnh."""
        from unittest.mock import MagicMock
        from cad_cowork_mcp.batch_executor import BatchExecutor

        controller = CADController()
        mock_doc = MagicMock()
        controller.app = MagicMock()
        controller.app.ActiveDocument = mock_doc
        controller.doc = mock_doc
        controller.is_quiescent = MagicMock(return_value=True)

        batch_executor = BatchExecutor(controller)

        # Batch chứa lệnh phá hoại nhưng confirm=False
        commands = [
            {"id": "c1", "command_name": "_.EXPLODE (ssget \"X\")"}
        ]
        res = batch_executor.execute(commands, confirm=False)
        self.assertFalse(res["success"])

        # Batch chứa lệnh phá hoại với confirm=True
        res_confirmed = batch_executor.execute(commands, confirm=True)
        self.assertTrue(res_confirmed["success"])


if __name__ == "__main__":
    unittest.main()
