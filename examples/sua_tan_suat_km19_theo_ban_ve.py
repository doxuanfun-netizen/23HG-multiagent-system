# -*- coding: utf-8 -*-
"""
Sửa file theo dõi tần suất thí nghiệm Cầu Km19 theo bản vẽ thẩm định (logic: tools/tan_suat_theo_ban_ve.py).

Đầu vào : file Excel theo dõi tần suất của dự án (--vao; file làm việc của đơn vị, KHÔNG commit vào repo)
          examples/data/km19_tan_suat_theo_ban_ve.json   quy tắc dòng bê tông → trang + nhãn bản vẽ
          BBS đã sửa theo bản vẽ + du_lieu.json (HSTK đã OCR)
Đầu ra  : bản đã sửa (--ra): bê tông theo bản vẽ, thép dựng lại từ BBS, công thức tần suất sửa, DM sinh lại,
          sheet NGHI_VAN (cột Đã soát / Ghim).

Chạy:  python examples/sua_tan_suat_km19_theo_ban_ve.py --vao <file.xlsx> --ra <file_moi.xlsx>
"""

import argparse
import datetime as dt
import json
import os
import sys

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tools.bbs_reconcile import load_pages, read_bbs  # noqa: E402
from tools.estimate_reconcile import NghiVan, resolve_f, write_flag_sheet  # noqa: E402
from tools.tan_suat_theo_ban_ve import (apply_concrete, apply_dm, apply_steel, apply_th_bqd, apply_tong_hop,  # noqa: E402
                                       concrete_events, steel_rows_from_bbs, update_bang_nhap)

DEFAULT_CONFIG = os.path.join(ROOT, "examples", "data", "km19_tan_suat_theo_ban_ve.json")


def run(vao: str, ra: str, config: str = DEFAULT_CONFIG, ngay: str = None) -> dict:
    with open(config, encoding="utf-8") as f:
        cfg = json.load(f)
    pages = load_pages(os.path.join(ROOT, cfg["du_lieu_ban_ve"]))
    _, sections = read_bbs(os.path.join(ROOT, cfg["bbs"]))
    wb = openpyxl.load_workbook(vao)
    flags = []
    ws_bt = wb[cfg["sheet_be_tong"]]
    bt_rows = apply_concrete(ws_bt, cfg, pages, flags)
    conc_ev = concrete_events(ws_bt, bt_rows, cfg)

    # ngày đổ bê tông (giá trị Excel lưu sẵn trong file — nhiều ô lấy từ sheet tiến độ) để xếp thép theo thời gian
    cached = openpyxl.load_workbook(vao, data_only=True)[cfg["sheet_be_tong"]]
    bt_dates = {r: cached.cell(r, 3).value for r in bt_rows}
    groups = steel_rows_from_bbs(sections, cfg, ws_bt, bt_rows, flags, bt_dates)
    steel_total, rowmap, steel_ev = apply_steel(wb[cfg["sheet_thep"]], groups, cfg, flags)

    if cfg.get("sheet_thep_bqd_glc") in wb.sheetnames:
        th_total, first, last = apply_th_bqd(wb[cfg["sheet_thep_bqd_glc"]], cfg["sheet_thep"], rowmap, flags)
        update_bang_nhap(wb, th_total, first, last)
    cap_bbs = round(sum(s.L * s.n * s.unit for s in rowmap.values() if s.key == "cap") / 1000, 3)
    cap_ban_ve = resolve_f(cfg["cap_dul_ban_ve_tan"], pages)[0]
    apply_tong_hop(wb[cfg["sheet_tong_hop"]], cfg["sheet_thep"], steel_total, cfg, flags, cap_ban_ve, cap_bbs)
    n_dm = apply_dm(wb["DM"], cfg["sheet_be_tong"], cfg["sheet_thep"], conc_ev, steel_ev, flags) if "DM" in wb.sheetnames else 0
    if "BBLM" in wb.sheetnames:
        flags.append(NghiVan("BBLM", 4, "Nút chọn phiếu", "Nút xoay (spinner) chọn số phiếu không giữ được khi công cụ lưu "
                                                          "file — nhập số thứ tự phiếu vào ô N4, hoặc chép sheet BBLM từ "
                                                          "file gốc sang."))
    ngay = ngay or dt.date.today().strftime("%d/%m/%Y")
    write_flag_sheet(wb, flags, [], f"NGHI VẤN TẦN SUẤT THÍ NGHIỆM — sửa theo bản vẽ thẩm định ngày {ngay} (soát và "
                                    f"ghim ở cột I)")
    wb.save(ra)
    return {"nghi_van": len(flags), "dong_thep": len(rowmap), "dm": n_dm, "lan_bt": len(conc_ev), "lan_thep": len(steel_ev),
            "cap_bbs": cap_bbs, "cap_ban_ve": cap_ban_ve}


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sửa file tần suất thí nghiệm Km19 theo bản vẽ thẩm định")
    ap.add_argument("--vao", required=True)
    ap.add_argument("--ra", required=True)
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--ngay", default=None)
    a = ap.parse_args(argv)
    st = run(a.vao, a.ra, a.config, a.ngay)
    print(f"-> {a.ra}: {st['dong_thep']} dòng thép, {st['lan_bt']} lần lấy mẫu vật liệu bê tông, {st['lan_thep']} lần "
          f"thép, DM {st['dm']} dòng, {st['nghi_van']} nghi vấn")
    return 0


if __name__ == "__main__":
    sys.exit(main())
