# -*- coding: utf-8 -*-
"""
Test cho chính khung đối chiếu chuẩn (tests/golden/harness/runner.py).

Mục tiêu: đảm bảo khung không bao giờ báo "đạt" cho ca chưa có dữ liệu duyệt,
và ca sai thì báo FAIL với thoát mã khác 0.
"""

import json
import os
import tempfile
import unittest

from tests.golden.harness import runner

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(HERE, "golden", "harness", "manifest.json")


def _write(cases):
    fd, path = tempfile.mkstemp(suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump({"cases": cases}, f, ensure_ascii=False)
    return path


def _active(cid, checks, source_type="hand_formula", **extra):
    base = {"id": cid, "title": cid, "source_type": source_type, "status": "active",
            "producer": "tools.steel_qs:rebar_unit_mass", "checks": checks}
    if source_type == "approved_dossier":
        base["approved_ref"] = "hồ sơ mẫu"
    base.update(extra)
    return base


class ManifestValidationTest(unittest.TestCase):
    def test_real_manifest_loads(self):
        cases = runner.load_cases(MANIFEST)
        self.assertGreaterEqual(len(cases), 1)

    def test_duplicate_id_rejected(self):
        with self.assertRaisesRegex(runner.ManifestError, "trùng id"):
            runner.validate({"cases": [_active("a", [{"args": [16], "expected": 1.58}]),
                                       _active("a", [{"args": [16], "expected": 1.58}])]})

    def test_pending_needs_reason(self):
        case = {"id": "p", "title": "p", "source_type": "approved_dossier", "status": "pending"}
        with self.assertRaisesRegex(runner.ManifestError, "reason"):
            runner.validate({"cases": [case]})

    def test_approved_active_needs_source_reference(self):
        case = _active("x", [{"args": [16], "expected": 1.58}], source_type="approved_dossier")
        del case["approved_ref"]
        with self.assertRaisesRegex(runner.ManifestError, "approved_ref"):
            runner.validate({"cases": [case]})

    def test_active_needs_producer_and_checks(self):
        no_producer = _active("x", [{"args": [16], "expected": 1.58}])
        del no_producer["producer"]
        with self.assertRaisesRegex(runner.ManifestError, "producer"):
            runner.validate({"cases": [no_producer]})
        with self.assertRaisesRegex(runner.ManifestError, "checks"):
            runner.validate({"cases": [_active("y", [])]})

    def test_unknown_source_type_rejected(self):
        case = _active("z", [{"args": [16], "expected": 1.58}], source_type="guess")
        with self.assertRaisesRegex(runner.ManifestError, "source_type"):
            runner.validate({"cases": [case]})


class RunCaseTest(unittest.TestCase):
    def test_pending_is_never_pass(self):
        case = {"id": "p", "title": "p", "source_type": "approved_dossier", "status": "pending",
                "reason": "chưa có file"}
        res = runner.run_case(case)
        self.assertEqual(res["result"], "CHƯA ĐỐI CHIẾU")
        self.assertNotEqual(res["result"], "PASS")

    def test_matching_value_passes(self):
        res = runner.run_case(_active("ok", [{"args": [16], "expected": 1.580247, "tol": 1e-6}]))
        self.assertEqual(res["result"], "PASS")

    def test_value_outside_tolerance_fails_with_detail(self):
        res = runner.run_case(_active("bad", [{"args": [16], "expected": 1.60, "tol": 0.001}]))
        self.assertEqual(res["result"], "FAIL")
        self.assertIn("kỳ vọng 1.6", res["detail"])

    def test_tolerance_boundary_is_inclusive(self):
        # 1.5802469... nằm đúng trong biên 1.58 ± 0.01 → phải đạt
        res = runner.run_case(_active("edge", [{"args": [16], "expected": 1.58, "tol": 0.01}]))
        self.assertEqual(res["result"], "PASS")

    def test_exception_in_producer_is_fail_not_crash(self):
        res = runner.run_case(_active("err", [{"args": [-1], "expected": 1.0}]))
        self.assertEqual(res["result"], "FAIL")
        self.assertIn("ValueError", res["detail"])


class SummaryAndExitCodeTest(unittest.TestCase):
    def test_approved_counts_only_approved_passes(self):
        cases = [
            _active("f", [{"args": [16], "expected": 1.580247, "tol": 1e-6}]),          # hand_formula đạt
            _active("a", [{"args": [16], "expected": 1.580247, "tol": 1e-6}],
                    source_type="approved_dossier"),                                  # approved đạt
            {"id": "p", "title": "p", "source_type": "approved_dossier", "status": "pending",
             "reason": "chưa có"},
        ]
        results = [runner.run_case(c) for c in cases]
        s = runner.summarize(cases, results)
        self.assertEqual(s["approved_total"], 2)
        self.assertEqual(s["approved_pass"], 1)
        self.assertFalse(s["any_fail"])

    def test_main_returns_1_when_an_active_case_fails(self):
        path = _write([_active("bad", [{"args": [16], "expected": 9.9, "tol": 0.001}])])
        try:
            self.assertEqual(runner.main(["--manifest", path]), 1)
        finally:
            os.remove(path)

    def test_main_returns_0_when_only_pending_and_passing(self):
        path = _write([_active("ok", [{"args": [16], "expected": 1.580247, "tol": 1e-6}]),
                       {"id": "p", "title": "p", "source_type": "approved_dossier",
                        "status": "pending", "reason": "chưa có"}])
        try:
            self.assertEqual(runner.main(["--manifest", path]), 0)
        finally:
            os.remove(path)

    def test_render_states_approved_coverage(self):
        cases = [{"id": "p", "title": "p", "source_type": "approved_dossier", "status": "pending",
                  "reason": "chưa có"}]
        results = [runner.run_case(c) for c in cases]
        text = runner.render(results, runner.summarize(cases, results))
        self.assertIn("0/1 ca đạt", text)


def _fake_table():
    return {"10": {"a": 1.0, "b": 2.0}, "12": {"a": 3.0, "b": 4.0}}


def _boom():
    raise RuntimeError("hỏng")


def _write_csv(text):
    fd, path = tempfile.mkstemp(suffix=".csv")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def _table(cid, csv_path, producer="tests.test_golden_harness:_fake_table", tol=None, **extra):
    base = {"id": cid, "title": cid, "source_type": "approved_dossier", "status": "active", "kind": "table",
            "approved_ref": "hồ sơ mẫu", "producer": producer, "expected_csv": csv_path,
            "tol": tol or {"a": 0.01, "b": 0.01}}
    base.update(extra)
    return base


class TableKindTest(unittest.TestCase):
    def setUp(self):
        self.paths = []

    def tearDown(self):
        for p in self.paths:
            os.remove(p)

    def csv(self, text):
        self.paths.append(_write_csv(text))
        return self.paths[-1]

    def test_matching_table_passes(self):
        res = runner.run_case(_table("t", self.csv("k,a,b\n10,1.0,2.0\n12,3.0,4.0\n")))
        self.assertEqual(res["result"], "PASS", res["detail"])
        self.assertIn("4 ô đạt", res["detail"])

    def test_cell_outside_tolerance_names_key_and_column(self):
        res = runner.run_case(_table("t", self.csv("k,a,b\n10,1.0,2.0\n12,3.0,4.5\n")))
        self.assertEqual(res["result"], "FAIL")
        self.assertIn("12/b", res["detail"])
        self.assertIn("1/4 ô lệch", res["detail"])

    def test_missing_and_extra_keys_fail(self):
        res = runner.run_case(_table("t", self.csv("k,a,b\n10,1.0,2.0\n14,5.0,6.0\n")))
        self.assertEqual(res["result"], "FAIL")
        self.assertIn("thiếu khóa ['14']", res["detail"])
        self.assertIn("không có trong hồ sơ ['12']", res["detail"])

    def test_producer_exception_is_fail(self):
        res = runner.run_case(_table("t", self.csv("k,a\n10,1.0\n"), producer="tests.test_golden_harness:_boom",
                                     tol={"a": 0.1}))
        self.assertEqual(res["result"], "FAIL")
        self.assertIn("RuntimeError", res["detail"])

    def test_table_needs_csv_and_tolerances(self):
        no_csv = _table("t", "")
        with self.assertRaisesRegex(runner.ManifestError, "expected_csv"):
            runner.validate({"cases": [no_csv]})
        bad_tol = _table("u", "x.csv")
        bad_tol["tol"] = 0.1
        with self.assertRaisesRegex(runner.ManifestError, "tol"):
            runner.validate({"cases": [bad_tol]})

    def test_unknown_kind_rejected(self):
        case = _active("k", [{"args": [16], "expected": 1.58}], kind="matrix")
        with self.assertRaisesRegex(runner.ManifestError, "kind"):
            runner.validate({"cases": [case]})

    def test_real_bbs_table_case_runs_without_error(self):
        case = next(c for c in runner.load_cases(MANIFEST) if c["id"] == "bbs_tong_hop_theo_duong_kinh")
        res = runner.run_case(case)
        self.assertIn(res["result"], ("PASS", "FAIL"))
        self.assertNotIn("lỗi", res["detail"])


if __name__ == "__main__":
    unittest.main()
