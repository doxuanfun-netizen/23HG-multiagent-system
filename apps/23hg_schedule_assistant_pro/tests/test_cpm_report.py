# -*- coding: utf-8 -*-
import csv
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "1_Scripts_TuDongHoa"))
import cpm_report  # noqa: E402


@unittest.skipUnless(os.path.exists(os.path.join(ROOT, "23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm")), "Cần file .xlsm")
class TestReport(unittest.TestCase):
    def test_report_has_all_rows_and_known_finish(self):
        with tempfile.TemporaryDirectory() as td:
            out = os.path.join(td, "r.csv")
            n = cpm_report.build_report(out)
            with open(out, newline="", encoding="utf-8-sig") as f:
                rows = list(csv.DictReader(f))
        self.assertEqual(n, 37)
        self.assertEqual(len(rows), 37)
        top = next(r for r in rows if r["WBS"] == "1")
        self.assertEqual(top["EF"], "2025-10-02")
        self.assertEqual(top["Gang"], "GANG")
        self.assertEqual({r["Loai"] for r in rows}, {"Summary", "Task", "Milestone"})


if __name__ == "__main__":
    unittest.main()
