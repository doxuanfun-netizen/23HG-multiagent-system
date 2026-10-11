# -*- coding: utf-8 -*-
"""
Sửa BBS Cầu Km19+529.080 theo bản vẽ thẩm định và đánh dấu NGHI VẤN (logic trong tools/bbs_reconcile.py).

Đầu vào : examples/data/km19_bbs_goc_truoc_sua.xlsx      BBS gốc trước khi sửa (giữ để chạy lại được)
          examples/data/km19_doi_chieu_ban_ve.json        ánh xạ hạng mục → trang bản vẽ, số cấu kiện
          05_DU_LIEU_GOC_SCAN_MARKER/du_lieu.json         HSTK đã OCR
Đầu ra  : 02_XUONG_TIEN_CHE_COT_THEP/04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx  (giữ tên để các script khác dùng)
          gồm sheet THONG_KE (cột NGHI VẤN), TONG_HOP theo Ø, DOI_CHIEU_BAN_VE, NGHI_VAN (cột Đã soát / Ghim).

Chạy:  python examples/sua_bbs_km19_theo_ban_ve.py [--bbs ...] [--config ...] [--out ...]
"""

import argparse
import datetime as dt
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from tools.bbs_reconcile import load_config, load_pages, read_bbs, reconcile, unit_changes, write_bbs  # noqa: E402

DEFAULT_BBS = os.path.join(ROOT, "examples", "data", "km19_bbs_goc_truoc_sua.xlsx")
DEFAULT_CONFIG = os.path.join(ROOT, "examples", "data", "km19_doi_chieu_ban_ve.json")
DEFAULT_OUT = os.path.join(ROOT, "examples", "HO_SO_CAU_KM19_529", "02_XUONG_TIEN_CHE_COT_THEP",
                           "04_Thong_Ke_Thep_Chi_Tiet_BBS_396_Dong.xlsx")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Sửa BBS Km19 theo bản vẽ thẩm định, đánh dấu nghi vấn")
    ap.add_argument("--bbs", default=DEFAULT_BBS)
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--out", default=DEFAULT_OUT)
    ap.add_argument("--ngay", default=dt.date.today().strftime("%d/%m/%Y"), help="ngày ghi trên file")
    a = ap.parse_args(argv)

    cfg = load_config(a.config)
    titles, sections = read_bbs(a.bbs)
    pages = load_pages(os.path.join(ROOT, cfg["du_lieu_ban_ve"]))
    new_sections, reports, flags = reconcile(sections, cfg, pages)
    note = (f"ĐÃ SỬA THEO BẢN VẼ THẨM ĐỊNH ngày {a.ngay} — dòng tô vàng/cam có NGHI VẤN (cột R), xem sheet NGHI_VAN "
            f"và DOI_CHIEU_BAN_VE. kg/m theo quy ước hồ sơ (bảng solver cắt thép).")
    stats = write_bbs(a.out, titles, new_sections, reports, flags, unit_changes(sections), note)
    print(f"-> Đã ghi BBS đã sửa: {a.out}")
    print(f"   {stats['dong_thep']} dòng thép, {stats['dong_nghi_van']} dòng có nghi vấn, "
          f"{stats['muc_nghi_van']} mục trong sheet NGHI_VAN")
    for r in reports:
        print(f"   {r.hanh_dong:<20} {r.hang_muc}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
