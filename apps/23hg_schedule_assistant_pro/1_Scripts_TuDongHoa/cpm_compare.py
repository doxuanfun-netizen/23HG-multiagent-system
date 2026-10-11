# -*- coding: utf-8 -*-
"""So sánh ngày ES/EF đang lưu trong sheet TIEN_DO (ví dụ sau khi bấm 'Cập Nhật Tiến Độ' rồi lưu file)
với kết quả bộ tính CPM chuẩn (ngày công). Thoát với mã 1 nếu có lệch.
Dùng:  python cpm_compare.py [duong_dan_file.xlsm] [--rain]   (--rain: áp hệ số mưa K_tt như VBA)"""
from __future__ import annotations

import datetime as dt
import os
import sys
from typing import List, Tuple

import build_cpm_formula as bf
import cpm_engine as ce


def _d(x):
    return x.date() if isinstance(x, dt.datetime) else x


def compare(path: str = bf.SRC_XLSM, use_rain: bool = False) -> Tuple[List[tuple], int]:
    rows, start, hol = bf.rows_from_xlsm(path)
    tasks = [ce.Task(r["wbs"], r["dur"], r["kind"], r["preds"], rain=r.get("rain", 1.0)) for r in rows]
    res = ce.schedule(tasks, _d(start), ce.Calendar(hol), use_rain=use_rain)
    bad = []
    for r in rows:
        e = res[r["wbs"]]
        ses, sef = _d(r["stored_es"]), _d(r["stored_ef"])
        if (ses, sef) != (e.es, e.ef):
            bad.append((r["wbs"], ses, e.es, sef, e.ef))
    return bad, len(rows)


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    use_rain = "--rain" in argv
    args = [a for a in argv if not a.startswith("--")]
    path = args[0] if args else bf.SRC_XLSM
    bad, n = compare(path, use_rain)
    print(f"File: {os.path.basename(path)} | {n} dòng | lệch: {len(bad)} | hệ số mưa: {'có' if use_rain else 'không'}")
    for w, ses, ces, sef, cef in bad:
        print(f"  {w:<10} ES lưu {ses} vs chuẩn {ces} | EF lưu {sef} vs chuẩn {cef}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
