# -*- coding: utf-8 -*-
"""
Sửa dự toán G_XD (08) và Phụ lục 03a (09) Cầu Km19+529.080 theo bản vẽ thẩm định, đánh dấu NGHI VẤN
(logic trong tools/estimate_reconcile.py).

Đầu vào : examples/data/km19_du_toan_goc_truoc_sua.xlsx   dự toán gốc trước khi sửa (giữ để chạy lại được)
          examples/data/km19_03a_goc_truoc_sua.xlsx        03a gốc trước khi sửa
          examples/data/km19_du_toan_theo_ban_ve.json      dòng QS → trang + nhãn trên bản vẽ
          05_DU_LIEU_GOC_SCAN_MARKER/du_lieu.json          HSTK đã OCR
Đầu ra  : 03_KINH_TE_QS_DU_TOAN_THANH_TOAN/08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx
          03_KINH_TE_QS_DU_TOAN_THANH_TOAN/09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx
          (cột M 'NGHI VẤN', cột N khối lượng cũ, sheet NGHI_VAN có cột 'Đã soát / Ghim')

Chạy:  python examples/sua_du_toan_03a_km19_theo_ban_ve.py [--ngay dd/mm/yyyy]
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

from tools.bbs_reconcile import load_pages  # noqa: E402
from tools.estimate_reconcile import apply_03a, apply_qs, write_flag_sheet  # noqa: E402

DATA = os.path.join(ROOT, "examples", "data")
OUT_DIR = os.path.join(ROOT, "examples", "HO_SO_CAU_KM19_529", "03_KINH_TE_QS_DU_TOAN_THANH_TOAN")
DEFAULT_DT = os.path.join(DATA, "km19_du_toan_goc_truoc_sua.xlsx")
DEFAULT_03A = os.path.join(DATA, "km19_03a_goc_truoc_sua.xlsx")
DEFAULT_CONFIG = os.path.join(DATA, "km19_du_toan_theo_ban_ve.json")
OUT_DT = os.path.join(OUT_DIR, "08_Du_Toan_GXD_Thong_Tu_11_2021.xlsx")
OUT_03A = os.path.join(OUT_DIR, "09_Thanh_Toan_Khoi_Luong_Phu_Luc_03a.xlsx")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sửa dự toán và 03a Km19 theo bản vẽ thẩm định, đánh dấu nghi vấn")
    ap.add_argument("--du-toan", default=DEFAULT_DT)
    ap.add_argument("--p03a", default=DEFAULT_03A)
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--out-du-toan", default=OUT_DT)
    ap.add_argument("--out-03a", default=OUT_03A)
    ap.add_argument("--ngay", default=dt.date.today().strftime("%d/%m/%Y"), help="ngày ghi trên file")
    a = ap.parse_args(argv)

    with open(a.config, encoding="utf-8") as f:
        cfg = json.load(f)
    pages = load_pages(os.path.join(ROOT, cfg["du_lieu_ban_ve"]))
    note = (f"ĐÃ SỬA THEO BẢN VẼ THẨM ĐỊNH ngày {a.ngay} — ô vàng: số theo bản vẽ; cột M: NGHI VẤN; cột N: khối lượng "
            f"cũ. Xem sheet NGHI_VAN để soát và ghim.")
    qs = cfg["qs_sheet"]

    wb = openpyxl.load_workbook(a.du_toan)
    flags = apply_qs(wb[qs], cfg, pages, qs, note)
    write_flag_sheet(wb, flags, cfg.get("thieu_trong_du_toan", []), "NGHI VẤN DỰ TOÁN — soát và ghim (cột I)")
    wb.save(a.out_du_toan)
    print(f"-> Dự toán: {a.out_du_toan} ({len(flags)} nghi vấn + {len(cfg.get('thieu_trong_du_toan', []))} hạng mục thiếu)")

    wb = openpyxl.load_workbook(a.p03a)
    flags03 = apply_qs(wb[qs], cfg, pages, qs, note)
    apply_03a(wb, cfg["phu_luc_03a"], qs, flags03)
    write_flag_sheet(wb, flags03, [], "NGHI VẤN PHỤ LỤC 03a — soát và ghim (cột I)")
    wb.save(a.out_03a)
    print(f"-> Phụ lục 03a: {a.out_03a} ({len(flags03)} nghi vấn)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
