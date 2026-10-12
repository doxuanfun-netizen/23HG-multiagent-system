# -*- coding: utf-8 -*-
"""
Kiểm thử lớp bảo vệ Whitelist nghiêm ngặt (Strict Whitelist Defense)
Chứng minh hệ thống tuân thủ nguyên tắc Defense-in-Depth:
Dù kẻ tấn công cố tình tránh mọi từ khóa trong Blacklist, tầng 4 (Whitelist)
vẫn chặn đứng 100% các lệnh không thuộc danh mục được phép.
"""
import unittest
from cad_cowork_mcp.cad_controller import validate_command_security, SAFE_COMMANDS_WHITELIST


class TestStrictWhitelistDefense(unittest.TestCase):
    """Kiểm tra tường lửa Whitelist nghiêm ngặt không thể bị vượt qua."""

    def test_non_blacklisted_unknown_verbs_blocked_by_whitelist(self):
        """Các lệnh lạ không nằm trong Blacklist vẫn bị Whitelist chặn tuyệt đối."""
        unknown_commands = [
            "INVENTED_COMMAND 1 2 3",
            "CUSTOM_PLUGIN_RUN",
            "EXEC_EXTERNAL_BINARY",
            "RUN_SYS_UTIL",
            "MY_EXPLOIT_PAYLOAD",
            "CURL http://malicious.com",
            "WGET http://malicious.com",
            "NC 192.168.1.1 4444",
            "SSH root@server",
            "FTP put evil.dwg",
            "REG ADD HKLM\\Software",
            "TASKKILL /F /IM acad.exe",
            "CERTUTIL -urlcache -split",
            "BITSADMIN /transfer job",
            "WMIC process call create",
            "MSHTA vbscript:Close()",
        ]
        for cmd in unknown_commands:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(
                    is_safe,
                    f"Lệnh không thuộc Whitelist '{cmd}' phải bị chặn ở Tầng 4 (Whitelist)"
                )
                self.assertIn("Whitelist", msg)

    def test_whitelist_contains_only_geometry_and_inquiry(self):
        """Đảm bảo danh mục Whitelist chỉ gồm các lệnh hình học, hiệu chỉnh, layer và truy vấn."""
        for verb in SAFE_COMMANDS_WHITELIST:
            self.assertTrue(verb.isupper(), f"Whitelist verb '{verb}' phải viết hoa chuẩn hóa.")
            self.assertFalse(
                verb in {"QUIT", "EXIT", "SHELL", "DOS", "SCRIPT", "NETLOAD", "ARXLOAD"},
                f"Whitelist tuyệt đối không được chứa lệnh phá hoại '{verb}'"
            )

    def test_multiline_and_semicolon_whitelist_enforcement(self):
        """Mọi lệnh con trong chuỗi đa lệnh (qua ';' hoặc '\\n') đều phải vượt qua Whitelist."""
        mixed_payloads = [
            "_.LINE 0,0 10,10;EVIL_CMD",
            "_.CIRCLE 0,0 50\nUNKNOWN_VERB 123",
            "_.PLINE 0,0 10,0 10,10 C;_.REGEN;PAYLOAD_X",
        ]
        for cmd in mixed_payloads:
            with self.subTest(cmd=cmd):
                is_safe, msg = validate_command_security(cmd)
                self.assertFalse(is_safe)
                self.assertIn("Whitelist", msg)


if __name__ == "__main__":
    unittest.main()
