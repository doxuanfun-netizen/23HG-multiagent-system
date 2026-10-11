# -*- coding: utf-8 -*-
"""
Hàm "producer" cho khung đối chiếu: gọi đúng code hệ thống đang dùng và trả kết quả dạng bảng để so với hồ sơ.
Không tự tính lại theo cách riêng — mục đích là đo chính hệ thống.
"""

import os
from collections import defaultdict

from tools.bbs_loader import load_bbs
from tools.steel_qs import rebar_unit_mass

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


def _key(diameter_mm):
    return f"{diameter_mm:g}"          # 10 → "10", 15.2 → "15.2"


def bbs_theo_duong_kinh(path, sheet=None):
    """
    Tổng chiều dài và khối lượng thép thanh theo Ø mà hệ thống tính từ một BBS:
    đọc bằng tools.bbs_loader (như bộ cắt thép), khối lượng theo tools.steel_qs.rebar_unit_mass (như bóc tách QS).
    Dòng bị bộ đọc loại (sai đường kính, chiều dài...) và cáp DƯL không có trong kết quả.
    """
    res = load_bbs(path if os.path.isabs(path) else os.path.join(ROOT, path), sheet=sheet)
    out = defaultdict(lambda: {"tong_dai_m": 0.0, "tong_kg": 0.0})
    for d in res.demands:
        length_m = d.length_mm * d.quantity / 1000.0
        row = out[_key(d.diameter_mm)]
        row["tong_dai_m"] += length_m
        row["tong_kg"] += length_m * rebar_unit_mass(d.diameter_mm)
    return dict(out)
