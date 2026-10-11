# -*- coding: utf-8 -*-
"""
Giữ CLAUDE.md và hook trước khi push khớp với repo thật:
  - CLAUDE.md dưới 200 dòng (dài hơn thì dễ bị bỏ sót khi đọc);
  - mọi đường dẫn nhắc trong CLAUDE.md đều tồn tại (không để tài liệu trỏ vào file đã đổi tên/xóa);
  - .githooks/pre-push chạy đúng các bước CI đang chạy, cộng bộ test apps.
"""

import glob
import os
import re
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLAUDE_MD = os.path.join(ROOT, "CLAUDE.md")
HOOK = os.path.join(ROOT, ".githooks", "pre-push")
MAX_LINES = 200


def _read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


class ClaudeMdTest(unittest.TestCase):
    def test_under_200_lines(self):
        n = len(_read(CLAUDE_MD).splitlines())
        self.assertLessEqual(n, MAX_LINES, f"CLAUDE.md có {n} dòng")

    def test_referenced_paths_exist(self):
        text = _read(CLAUDE_MD)
        missing = []
        for token in re.findall(r"`([^`\s]+)`", text):
            if "/" not in token or "<" in token or "\\" in token or token.startswith(("http", "~")):
                continue
            path = token.split("#")[0].rstrip("/")
            if not os.path.exists(os.path.join(ROOT, path)):
                missing.append(token)
        self.assertEqual(missing, [], "CLAUDE.md nhắc tới đường dẫn không tồn tại")

    def test_test_command_matches_ci(self):
        self.assertIn("python -m unittest discover -s tests -t .", _read(CLAUDE_MD))


class PrePushHookTest(unittest.TestCase):
    def test_hook_runs_ci_steps_and_apps_suite(self):
        hook = _read(HOOK)
        ci = "\n".join(_read(p) for p in glob.glob(os.path.join(ROOT, ".github", "workflows", "*.yml")))
        for step in ("unittest discover -s tests -t .", "run_state_graph.py --solver-test"):
            self.assertIn(step, ci, f"CI không còn bước '{step}' — cập nhật hook và test này")
            self.assertIn(step, hook)
        self.assertIn("apps/23hg_schedule_assistant_pro", hook)
        self.assertTrue(hook.startswith("#!/bin/sh"))

    @unittest.skipIf(os.name == "nt", "quyền thực thi không áp dụng trên Windows")
    def test_hook_is_executable(self):
        self.assertTrue(os.access(HOOK, os.X_OK))


if __name__ == "__main__":
    unittest.main()
