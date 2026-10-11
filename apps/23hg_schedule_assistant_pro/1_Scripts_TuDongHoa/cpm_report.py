# -*- coding: utf-8 -*-
"""Xuất kết quả CPM tham chiếu (bộ tính Python) ra CSV để so với macro Excel trên Windows.
Dùng:  python cpm_report.py [--out FILE.csv]"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import os

import build_cpm_formula as bf
import cpm_engine as ce

DEFAULT_OUT = os.path.join(bf.ROOT, "4_CPM_CONG_THUC_THU_NGHIEM", "KET_QUA_THAM_CHIEU_PYTHON.csv")


def _d(x):
    return x.date() if isinstance(x, dt.datetime) else x


def build_report(out_path: str) -> int:
    rows, start, hol = bf.rows_from_xlsm()
    tasks = [ce.Task(r["wbs"], r["dur"], r["kind"], r["preds"]) for r in rows]
    res = ce.schedule(tasks, _d(start), ce.Calendar(hol))
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["WBS", "Loai", "ThoiLuong", "ES", "EF", "LS", "LF", "TF_ngay_cong", "Gang",
                    "ES_da_luu", "EF_da_luu", "Lech_ES_ngay_lich", "Lech_EF_ngay_lich"])
        for r in rows:
            e = res[r["wbs"]]
            ses, sef = _d(r["stored_es"]), _d(r["stored_ef"])
            w.writerow([r["wbs"], r["kind"], r["dur"], e.es, e.ef, e.ls, e.lf, e.tf, "GANG" if e.critical else "",
                        ses, sef, (e.es - ses).days if ses else "", (e.ef - sef).days if sef else ""])
    return len(rows)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", default=DEFAULT_OUT)
    a = ap.parse_args()
    print("Đã ghi", build_report(a.out), "dòng ->", a.out)


if __name__ == "__main__":
    main()
