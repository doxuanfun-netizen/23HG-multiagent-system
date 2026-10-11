# -*- coding: utf-8 -*-
"""
Khung đối chiếu chuẩn (golden harness).

Mỗi ca trong manifest.json gọi một hàm của hệ thống với đầu vào đã biết rồi so với giá trị kỳ vọng.

source_type của ca:
  - "approved_dossier": kỳ vọng lấy từ hồ sơ đã được duyệt. Chỉ loại này mới được tính là đối chiếu nghiệm thu.
  - "hand_formula": kỳ vọng tính tay từ công thức. Chỉ kiểm tra công thức, KHÔNG phải bằng chứng hồ sơ.

kind của ca (mặc định "scalar"):
  - "scalar": mỗi phần tử trong 'checks' gọi hàm một lần, so một số với 'expected' ± 'tol'.
  - "table": gọi hàm một lần với 'args'/'kwargs', hàm trả {khóa: {cột: số}}; so từng ô với file
    'expected_csv' (cột đầu là khóa) theo dung sai từng cột trong 'tol'. Khóa thiếu hoặc thừa đều là lệch.

status của ca:
  - "active": đã có kỳ vọng, chạy kiểm tra.
  - "pending": chưa có dữ liệu duyệt. Báo cáo là CHƯA ĐỐI CHIẾU, không bao giờ tính là đạt.

Chạy:  python tests/golden/harness/runner.py [--manifest ...]
Thoát mã 1 nếu có ca active bị FAIL.
"""

import argparse
import csv
import importlib
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

DEFAULT_MANIFEST = os.path.join(HERE, "manifest.json")
SOURCE_TYPES = ("approved_dossier", "hand_formula")
STATUSES = ("active", "pending")
KINDS = ("scalar", "table")
MAX_LISTED = 8                    # số ô lệch liệt kê trong chi tiết; phần còn lại chỉ đếm


class ManifestError(ValueError):
    pass


def validate(data):
    """Kiểm tra cấu trúc manifest; sai thì báo lỗi kèm mã ca để sửa nhanh."""
    if not isinstance(data.get("cases"), list) or not data["cases"]:
        raise ManifestError("manifest phải có danh sách 'cases' không rỗng")
    seen = set()
    for c in data["cases"]:
        cid = c.get("id", "<không có id>")
        for key in ("id", "source_type", "status", "title"):
            if not c.get(key):
                raise ManifestError(f"{cid}: thiếu trường '{key}'")
        if cid in seen:
            raise ManifestError(f"{cid}: trùng id")
        seen.add(cid)
        if c["source_type"] not in SOURCE_TYPES:
            raise ManifestError(f"{cid}: source_type phải là một trong {SOURCE_TYPES}")
        if c["status"] not in STATUSES:
            raise ManifestError(f"{cid}: status phải là một trong {STATUSES}")
        if c["status"] == "pending" and not c.get("reason"):
            raise ManifestError(f"{cid}: ca pending phải ghi 'reason' (đang thiếu dữ liệu gì)")
        kind = c.get("kind", "scalar")
        if kind not in KINDS:
            raise ManifestError(f"{cid}: kind phải là một trong {KINDS}")
        if c["status"] == "active":
            if ":" not in c.get("producer", ""):
                raise ManifestError(f"{cid}: ca active phải có 'producer' dạng 'module:ham'")
            if kind == "scalar" and not c.get("checks"):
                raise ManifestError(f"{cid}: ca active phải có 'checks'")
            if kind == "table":
                if not c.get("expected_csv"):
                    raise ManifestError(f"{cid}: ca table phải có 'expected_csv'")
                if not isinstance(c.get("tol"), dict) or not c["tol"]:
                    raise ManifestError(f"{cid}: ca table phải có 'tol' dạng {{cột: dung sai}}")
            if c["source_type"] == "approved_dossier" and not c.get("approved_ref"):
                raise ManifestError(f"{cid}: ca approved_dossier phải ghi 'approved_ref' (nguồn hồ sơ đã duyệt)")
    return data["cases"]


def load_cases(path=DEFAULT_MANIFEST):
    with open(path, encoding="utf-8") as f:
        return validate(json.load(f))


def resolve(producer):
    module_name, func_name = producer.split(":", 1)
    return getattr(importlib.import_module(module_name), func_name)


def _read_expected_csv(path):
    path = path if os.path.isabs(path) else os.path.join(ROOT, path)
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    header, body = rows[0], [r for r in rows[1:] if r and any(x.strip() for x in r)]
    return {r[0].strip(): {h: float(v) for h, v in zip(header[1:], r[1:]) if v.strip()} for r in body}


def _run_table(case):
    cid, src = case["id"], case["source_type"]
    fail = lambda detail: {"id": cid, "source_type": src, "result": "FAIL", "detail": detail}
    try:
        got = resolve(case["producer"])(*case.get("args", []), **case.get("kwargs", {}))
    except Exception as exc:  # lỗi trong hàm cũng là FAIL, không để harness sập
        return fail(f"lỗi {type(exc).__name__}: {exc}")
    got = {str(k): v for k, v in got.items()}
    expected = _read_expected_csv(case["expected_csv"])
    tol = case["tol"]
    problems = []
    missing = [k for k in expected if k not in got]
    extra = [k for k in got if k not in expected]
    if missing:
        problems.append(f"hệ thống thiếu khóa {missing}")
    if extra:
        problems.append(f"hệ thống có khóa không có trong hồ sơ {extra}")
    cells = []
    n_cells = 0
    for k in expected:
        if k not in got:
            continue
        for col, t in tol.items():
            if col not in expected[k]:
                continue
            n_cells += 1
            g, e = got[k].get(col), expected[k][col]
            if g is None or abs(g - e) > t:
                shown = "không có" if g is None else f"{g:,.3f}"
                diff = "" if g is None else f", lệch {g - e:+,.3f}"
                cells.append(f"{k}/{col}: nhận {shown}, kỳ vọng {e:,.3f}{diff} (± {t})")
    if cells:
        listed = "; ".join(cells[:MAX_LISTED]) + (f"; … và {len(cells) - MAX_LISTED} ô khác" if len(cells) > MAX_LISTED else "")
        problems.append(f"{len(cells)}/{n_cells} ô lệch: {listed}")
    if problems:
        return fail(" | ".join(problems))
    return {"id": cid, "source_type": src, "result": "PASS", "detail": f"{n_cells} ô đạt ({len(expected)} khóa)"}


def run_case(case):
    cid, src = case["id"], case["source_type"]
    if case["status"] == "pending":
        return {"id": cid, "source_type": src, "result": "CHƯA ĐỐI CHIẾU", "detail": case["reason"]}
    if case.get("kind", "scalar") == "table":
        return _run_table(case)
    fn = resolve(case["producer"])
    failures = []
    for chk in case["checks"]:
        args, kwargs = chk.get("args", []), chk.get("kwargs", {})
        expected, tol = chk["expected"], chk.get("tol", 0)
        try:
            got = fn(*args, **kwargs)
        except Exception as exc:  # lỗi trong hàm cũng là FAIL, không để harness sập
            failures.append(f"args={args} kwargs={kwargs}: lỗi {type(exc).__name__}: {exc}")
            continue
        if abs(got - expected) > tol:
            failures.append(f"args={args} kwargs={kwargs}: nhận {got}, kỳ vọng {expected} ± {tol}")
    detail = "; ".join(failures) if failures else f"{len(case['checks'])} điểm kiểm đạt"
    return {"id": cid, "source_type": src, "result": "FAIL" if failures else "PASS", "detail": detail}


def summarize(cases, results):
    """Đếm riêng đối chiếu hồ sơ duyệt; hand_formula không được cộng vào mục này."""
    approved = [c for c in cases if c["source_type"] == "approved_dossier"]
    approved_pass = [r for r in results if r["source_type"] == "approved_dossier" and r["result"] == "PASS"]
    return {
        "approved_total": len(approved),
        "approved_pass": len(approved_pass),
        "any_fail": any(r["result"] == "FAIL" for r in results),
    }


def render(results, summary):
    lines = [f"{'KẾT QUẢ':<16}{'NGUỒN':<20}{'CA':<36}CHI TIẾT"]
    for r in results:
        lines.append(f"{r['result']:<16}{r['source_type']:<20}{r['id']:<36}{r['detail']}")
    lines.append("")
    lines.append(f"Đối chiếu hồ sơ đã duyệt: {summary['approved_pass']}/{summary['approved_total']} ca đạt.")
    lines.append("Ca hand_formula chỉ kiểm tra công thức, không tính vào mục trên.")
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Chạy khung đối chiếu chuẩn")
    ap.add_argument("--manifest", default=DEFAULT_MANIFEST)
    a = ap.parse_args(argv)
    cases = load_cases(a.manifest)
    results = [run_case(c) for c in cases]
    summary = summarize(cases, results)
    print(render(results, summary))
    return 1 if summary["any_fail"] else 0


if __name__ == "__main__":
    sys.exit(main())
