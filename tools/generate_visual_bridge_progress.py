# -*- coding: utf-8 -*-
"""
Hệ thống Tự động Xuất Sơ đồ Tiến độ Trực quan & Hồ sơ Hoàn công Cầu (As-Built Visual Progress Dashboard)
Bố cục theo bảng theo dõi kỹ thuật công trường:
- Bảng thông số nhịp & khối lượng đỉnh
- Sơ đồ trắc dọc cầu (Elevation view) với phân tầng cấu kiện và mã màu trực quan:
  + Đã thi công (Xanh lá)
  + Đang triển khai (Vàng / Cam)
  + Chưa thi công (Trắng)
- Mặt bằng bố trí cọc khoan nhồi (Pile Layout Matrix) cho từng mố trụ M1, T1, T2, M2
- Bảng tiến độ kết cấu nhịp (Superstructure Beam & Slab Progress)
- Phân tích chi tiết hiện trạng đường găng công trường theo dữ liệu thực tế:
  * 26/26 Cọc khoan nhồi: Hoàn thành 100%
  * Bệ mố M1, M2 & Bệ trụ T1, T2: Hoàn thành 100%
  * Thân trụ T1: Đốt 1 Đã xong, Đốt 2 Đang triển khai
  * Thân trụ T2: Đốt 1 Đã xong, Đốt 2 Đang triển khai
  * Mố M1: Bệ đã xong, Thân đang triển khai cốt thép / ván khuôn
  * Mố M2: Đã hoàn thành thân mố, đang triển khai đỉnh mố / tường cánh

Dữ liệu nằm trong examples/data/asbuilt_km19.json (không gõ cứng trong script):
  - hang_muc: các hạng mục hiện trường (ngày, tình trạng, KCS);
  - ke_hoach_duc_dam: ngày bắt đầu, chu kỳ, số nhịp, số phiến mỗi nhịp, số phiến đã đúc → lịch đúc được TÍNH.
Giới hạn: sơ đồ HTML vẽ cố định 3 nhịp; màu khối dầm trên sơ đồ SVG chưa đổi theo số phiến đã đúc.

Chạy:  python tools/generate_visual_bridge_progress.py [--config ...] [--html ...] [--xlsx ...]
"""

import argparse
import datetime as dt
import math
import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ARTIFACT_DIR = os.environ.get("AEC_ARTIFACT_DIR", os.path.join(ROOT, "output"))
CONFIG_PATH = os.path.join(ROOT, "examples", "data", "asbuilt_km19.json")
DEFAULT_XLSX = os.path.join(ROOT, "examples", "HO_SO_CAU_KM19_529", "01_HIEN_TRUONG_QLCL_KCS",
                            "So_Do_Tien_Do_Truc_Quan_AsBuilt_Cau_Km19.xlsx")
TRANG_THAI = ("ĐÃ THI CÔNG", "ĐANG TRIỂN KHAI", "CHƯA THI CÔNG")
ROW_KEYS = ("hang_muc", "vi_tri", "quy_cach", "mac_bt", "khoi_luong_m3", "tien_do_thiet_ke",
            "ngay_khoi_cong", "ngay_hoan_thanh", "tinh_trang", "danh_gia_kcs", "ghi_chu")
HTML_NHIP = 3                     # template HTML vẽ cố định 3 nhịp


class AsBuiltConfigError(ValueError):
    pass


def load_config(path=CONFIG_PATH):
    """Đọc và kiểm tra cấu hình As-Built; sai thì báo lỗi kèm vị trí."""
    with open(path, encoding="utf-8") as f:
        cfg = json.load(f)
    kh = cfg.get("ke_hoach_duc_dam") or {}
    try:
        dt.date.fromisoformat(kh["ngay_bat_dau"])
    except (KeyError, TypeError, ValueError):
        raise AsBuiltConfigError("ke_hoach_duc_dam.ngay_bat_dau phải có dạng YYYY-MM-DD")
    for key in ("chu_ky_ngay", "so_nhip", "phien_moi_nhip"):
        if not isinstance(kh.get(key), int) or kh[key] <= 0:
            raise AsBuiltConfigError(f"ke_hoach_duc_dam.{key} phải là số nguyên dương")
    tong = kh["so_nhip"] * kh["phien_moi_nhip"]
    if not isinstance(kh.get("da_duc"), int) or not 0 <= kh["da_duc"] <= tong:
        raise AsBuiltConfigError(f"ke_hoach_duc_dam.da_duc phải trong khoảng 0..{tong}")
    markers = 0
    for i, item in enumerate(cfg.get("hang_muc") or [], start=1):
        if item.get("chen_ke_hoach_dam"):
            markers += 1
            continue
        missing = [k for k in ROW_KEYS if k not in item]
        if missing:
            raise AsBuiltConfigError(f"hang_muc[{i}] thiếu {missing}")
        if item["tinh_trang"] not in TRANG_THAI:
            raise AsBuiltConfigError(f"hang_muc[{i}] '{item['hang_muc']}': tinh_trang phải là một trong {TRANG_THAI}")
    if markers != 1:
        raise AsBuiltConfigError("hang_muc phải có đúng một mục {\"chen_ke_hoach_dam\": true}")
    return cfg


def _d(d):
    return d.strftime("%d/%m/%Y")


def _dm(d):
    return d.strftime("%d/%m")


def ke_hoach_dam(kh):
    """Lịch đúc dầm tính từ cấu hình: đúc lần lượt từng phiến, mỗi phiến một chu kỳ, theo thứ tự nhịp."""
    start = dt.date.fromisoformat(kh["ngay_bat_dau"])
    cycle, n, so_nhip, da_duc = kh["chu_ky_ngay"], kh["phien_moi_nhip"], kh["so_nhip"], kh["da_duc"]
    nhip = []
    for i in range(so_nhip):
        bd = start + dt.timedelta(days=i * n * cycle)
        kt = start + dt.timedelta(days=(i + 1) * n * cycle)
        done = min(max(da_duc - i * n, 0), n)
        trang_thai = TRANG_THAI[2] if done == 0 else (TRANG_THAI[0] if done == n else TRANG_THAI[1])
        lich = [(bd + dt.timedelta(days=k * cycle), bd + dt.timedelta(days=(k + 1) * cycle)) for k in range(n)]
        nhip.append({"so": i + 1, "bat_dau": bd, "ket_thuc": kt, "da_duc": done, "trang_thai": trang_thai, "lich": lich})
    return {"tong": so_nhip * n, "da_duc": da_duc, "chu_ky": cycle, "phien_moi_nhip": n, "nhip": nhip}


def _tinh_trang_nhip(nh, n):
    if nh["da_duc"] == 0:
        return f"Chưa thi công (0/{n} phiến)"
    if nh["da_duc"] == n:
        return f"Đã đúc xong ({n}/{n} phiến)"
    return f"Đang đúc ({nh['da_duc']}/{n} phiến)"


def html_tokens(cfg):
    """Giá trị điền vào các token @@...@@ của template HTML."""
    kh = cfg["ke_hoach_duc_dam"]
    plan = ke_hoach_dam(kh)
    if len(plan["nhip"]) != HTML_NHIP:
        raise AsBuiltConfigError(f"Template HTML vẽ cố định {HTML_NHIP} nhịp; cấu hình có {len(plan['nhip'])} nhịp")
    n, tong, cycle = plan["phien_moi_nhip"], plan["tong"], plan["chu_ky"]
    t = {
        "TONG_PHIEN": str(tong), "DA_DUC": str(plan["da_duc"]), "CHU_KY": str(cycle),
        "PHIEN_MOI_NHIP": str(n), "NGAY_MOI_NHIP": str(n * cycle), "BT_NHIP": f"{kh['bt_moi_nhip_m3']:g}",
        "NGAY_1_BE": str(tong * cycle), "NGAY_2_BE": str(math.ceil(tong * cycle / 2)),
        "TT_TONG": ("Chưa thi công" if plan["da_duc"] == 0 else
                    "Đã đúc xong" if plan["da_duc"] == tong else "Đang đúc"),
        "BAI_DUC_HIEN_TRANG": ("Hiện chưa đúc phiến nào." if plan["da_duc"] == 0 else
                               f"Đã đúc {plan['da_duc']}/{tong} phiến."),
    }
    for nh in plan["nhip"]:
        k = f"N{nh['so']}"
        pct = round(100 * nh["da_duc"] / n)
        t[f"{k}_BD"], t[f"{k}_KT"], t[f"{k}_BD_NGAN"] = _d(nh["bat_dau"]), _d(nh["ket_thuc"]), _dm(nh["bat_dau"])
        t[f"{k}_TIEN_DO"] = f"{nh['da_duc']}/{n} phiến ({pct}%)"
        t[f"{k}_TT"] = _tinh_trang_nhip(nh, n)
        t[f"{k}_TT_HOA"] = nh["trang_thai"]
        t[f"{k}_LICH"] = " → ".join(f"D{j} ({_dm(a)}-{_dm(b)})" for j, (a, b) in enumerate(nh["lich"], start=1))
    return t


def fill_tokens(template, tokens):
    out = template
    for k, v in tokens.items():
        out = out.replace(f"@@{k}@@", v)
    if "@@" in out:
        i = out.index("@@")
        raise AsBuiltConfigError(f"Template còn token chưa điền: {out[i:i + 40]!r}")
    return out


def beam_rows(cfg):
    """Các dòng Excel 'Đúc dầm Super-T Nhịp i' tính từ ke_hoach_duc_dam."""
    kh = cfg["ke_hoach_duc_dam"]
    plan = ke_hoach_dam(kh)
    n, cycle = plan["phien_moi_nhip"], plan["chu_ky"]
    rows = []
    for nh in plan["nhip"]:
        note = (f"KH bắt đầu đúc {_d(nh['bat_dau'])}, chu kỳ {cycle} ngày/phiến ({n * cycle} ngày)" if nh["so"] == 1
                else f"KH đúc sau Nhịp {nh['so'] - 1}, chu kỳ {cycle} ngày/phiến ({n * cycle} ngày)")
        kcs = f"Chưa đúc (0/{n})" if nh["da_duc"] == 0 else f"Đã đúc ({nh['da_duc']}/{n})"
        rows.append((f"Đúc dầm Super-T Nhịp {nh['so']}", "Bãi đúc", f"{n} Phiến L={kh['chieu_dai_dam_m']:g}m",
                     kh["mac_bt"], kh["bt_moi_nhip_m3"], f"{_dm(nh['bat_dau'])} - {_d(nh['ket_thuc'])}",
                     f"{_d(nh['bat_dau'])} (KH)", f"{_d(nh['ket_thuc'])} (KH)", nh["trang_thai"], kcs, note))
    return rows


def excel_rows(cfg):
    rows = []
    for item in cfg["hang_muc"]:
        if item.get("chen_ke_hoach_dam"):
            rows.extend(beam_rows(cfg))
        else:
            rows.append(tuple(item[k] for k in ROW_KEYS))
    return rows


def tom_tat(rows, cfg):
    """Dòng tóm tắt hiện trạng, đếm từ dữ liệu thay vì gõ tay."""
    count = {t: sum(1 for r in rows if r[8] == t) for t in TRANG_THAI}
    plan = ke_hoach_dam(cfg["ke_hoach_duc_dam"])
    return (f"Hiện trạng thi công As-Built: {count[TRANG_THAI[0]]} hạng mục đã thi công • "
            f"{count[TRANG_THAI[1]]} đang triển khai • {count[TRANG_THAI[2]]} chưa thi công • "
            f"Dầm Super-T đã đúc {plan['da_duc']}/{plan['tong']} phiến")


def generate_html_dashboard(output_path: str, cfg=None):
    """Tạo tệp HTML trực quan; lịch đúc dầm điền từ cấu hình (xem html_tokens)."""
    cfg = cfg or load_config()
    html_content = """<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Sơ Đồ Tiến Độ Trực Quan & Hoàn Công Cầu Km19+529.080</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    @media print {
      body { background: white !important; color: black !important; padding: 0 !important; }
      .no-print { display: none !important; }
      .print-shadow-none { box-shadow: none !important; border: 1px solid #ccc !important; }
    }
    .svg-node { transition: all 0.25s ease; cursor: pointer; }
    .svg-node:hover { filter: brightness(1.15) drop-shadow(0 4px 6px rgba(0,0,0,0.15)); }
    .pulse-amber { animation: pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite; }
    @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: .75; } }
  </style>
</head>
<body class="bg-slate-900 text-slate-100 antialiased p-4 md:p-6 min-h-screen">
  <div class="max-w-7xl mx-auto space-y-6">

    <!-- HEADER TITLE & CONTROLS -->
    <div class="bg-slate-800/90 border border-slate-700/80 rounded-2xl p-6 shadow-xl backdrop-blur-sm">
      <div class="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
        <div>
          <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-emerald-500/20 text-emerald-400 text-xs font-semibold uppercase tracking-wider mb-2 border border-emerald-500/30">
            <span class="w-2 h-2 rounded-full bg-emerald-400 animate-ping"></span>
            Báo Cáo Tiến Độ Thực Tế Hiện Trường (As-Built Live Tracking)
          </div>
          <h1 class="text-2xl md:text-3xl font-bold tracking-tight text-white flex items-center gap-3">
            <span>🌉</span> SƠ ĐỒ TIẾN ĐỘ THI CÔNG TRỰC QUAN CẦU KM19+529.080
          </h1>
          <p class="text-slate-400 text-sm mt-1">
            Quy mô: Cầu 3 Nhịp dầm Super-T $L=3\times 38.2\text{m} = 114.6\text{m}$ • Kết cấu móng 26 Cọc khoan nhồi $\varnothing 1200 / \varnothing 1000$ • Tải trọng HL93
          </p>
        </div>

        <!-- KPI SUMMARY CARDS -->
        <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div class="bg-slate-900/80 border border-slate-700 rounded-xl p-3 text-center">
            <div class="text-xs text-slate-400 font-medium">Tiến độ chung</div>
            <div class="text-xl font-bold text-emerald-400 mt-1">62.1%</div>
            <div class="text-[11px] text-slate-500">1.731 / 2.789 m³ BT</div>
          </div>
          <div class="bg-slate-900/80 border border-slate-700 rounded-xl p-3 text-center">
            <div class="text-xs text-slate-400 font-medium">Cọc khoan nhồi</div>
            <div class="text-xl font-bold text-emerald-400 mt-1">26/26 cọc</div>
            <div class="text-[11px] text-emerald-500 font-medium">Hoàn thành 100%</div>
          </div>
          <div class="bg-slate-900/80 border border-slate-700 rounded-xl p-3 text-center">
            <div class="text-xs text-slate-400 font-medium">Bệ mố & bệ trụ</div>
            <div class="text-xl font-bold text-emerald-400 mt-1">4/4 bệ</div>
            <div class="text-[11px] text-emerald-500 font-medium">Hoàn thành 100%</div>
          </div>
          <div class="bg-slate-900/80 border border-slate-700 rounded-xl p-3 text-center">
            <div class="text-xs text-slate-400 font-medium">Dầm Super-T (@@TONG_PHIEN@@ phiến)</div>
            <div class="text-xl font-bold text-slate-300 mt-1">@@DA_DUC@@/@@TONG_PHIEN@@ phiến</div>
            <div class="text-[11px] text-cyan-400 font-semibold">KH đúc: @@N1_BD_NGAN@@ (@@CHU_KY@@ ngày/phiến)</div>
          </div>
        </div>
      </div>

      <!-- LEGEND / CHÚ THÍCH MÀU SẮC -->
      <div class="mt-6 pt-4 border-t border-slate-700/60 flex flex-wrap items-center justify-between gap-4">
        <div class="flex flex-wrap items-center gap-6 text-sm">
          <span class="font-semibold text-slate-300">Chú thích phân vùng trạng thái:</span>
          <div class="flex items-center gap-2">
            <span class="w-4 h-4 rounded bg-emerald-500 border border-emerald-400 shadow-sm shadow-emerald-500/50"></span>
            <span class="text-slate-200">Đã thi công hoàn thành (Completed)</span>
          </div>
          <div class="flex items-center gap-2">
            <span class="w-4 h-4 rounded bg-amber-500 border border-amber-400 shadow-sm shadow-amber-500/50 animate-pulse"></span>
            <span class="text-slate-200">Đang triển khai thi công (In Progress)</span>
          </div>
          <div class="flex items-center gap-2">
            <span class="w-4 h-4 rounded bg-slate-700 border border-slate-500"></span>
            <span class="text-slate-300">Chưa thi công (Not Started)</span>
          </div>
        </div>

        <div class="flex items-center gap-2 no-print">
          <button onclick="window.print()" class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-white rounded-lg text-xs font-medium transition flex items-center gap-1.5 border border-slate-600">
            <span>🖨️</span> In Bản Vẽ A3/A4
          </button>
          <button onclick="resetInspect()" class="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-medium transition flex items-center gap-1.5">
            <span>🔍</span> Xem Toàn Thể Cầu
          </button>
        </div>
      </div>
    </div>

    <!-- MAIN INTERACTIVE BRIDGE CANVAS -->
    <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">

      <!-- LEFT: VISUAL DRAWING (3 COLS) -->
      <div class="lg:col-span-3 bg-slate-800/90 border border-slate-700 rounded-2xl p-5 shadow-xl flex flex-col justify-between overflow-x-auto">

        <!-- 1. TOP SPAN & CONCRETE QUANTITY TABLE -->
        <div class="mb-4">
          <div class="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2 flex items-center gap-2">
            <span>📋</span> Bảng Thông Số Chiều Dài Nhịp & Khối Lượng Bê Tông Thiết Kế
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-xs text-center border-collapse border border-slate-700 rounded-lg overflow-hidden">
              <thead>
                <tr class="bg-slate-700/80 text-slate-200">
                  <th class="p-2 border border-slate-600 font-semibold w-24">Phân đoạn</th>
                  <th class="p-2 border border-slate-600 font-bold text-amber-300">MỐ M1</th>
                  <th class="p-2 border border-slate-600 font-bold text-cyan-300">NHỊP 1 (M1 - T1)</th>
                  <th class="p-2 border border-slate-600 font-bold text-amber-300">TRỤ T1</th>
                  <th class="p-2 border border-slate-600 font-bold text-cyan-300">NHỊP 2 (T1 - T2)</th>
                  <th class="p-2 border border-slate-600 font-bold text-amber-300">TRỤ T2</th>
                  <th class="p-2 border border-slate-600 font-bold text-cyan-300">NHỊP 3 (T2 - M2)</th>
                  <th class="p-2 border border-slate-600 font-bold text-amber-300">MỐ M2</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-700 bg-slate-800/50">
                <tr>
                  <td class="p-2 border border-slate-700 font-medium text-slate-400 bg-slate-800">Khối lượng BT</td>
                  <td class="p-2 border border-slate-700 text-emerald-400 font-semibold">183.4 m³</td>
                  <td class="p-2 border border-slate-700 text-amber-400 font-semibold">229.9 m³</td>
                  <td class="p-2 border border-slate-700 text-emerald-400 font-semibold">543.1 m³</td>
                  <td class="p-2 border border-slate-700 text-amber-400 font-semibold">229.9 m³</td>
                  <td class="p-2 border border-slate-700 text-emerald-400 font-semibold">463.9 m³</td>
                  <td class="p-2 border border-slate-700 text-amber-400 font-semibold">229.9 m³</td>
                  <td class="p-2 border border-slate-700 text-emerald-400 font-semibold">408.7 m³</td>
                </tr>
                <tr>
                  <td class="p-2 border border-slate-700 font-medium text-slate-400 bg-slate-800">Chiều dài đúc</td>
                  <td class="p-2 border border-slate-700 text-slate-300">Bệ: 9.8m</td>
                  <td class="p-2 border border-slate-700 text-slate-300 font-bold">L = 38.20 m</td>
                  <td class="p-2 border border-slate-700 text-slate-300">Bệ: 12.6m</td>
                  <td class="p-2 border border-slate-700 text-slate-300 font-bold">L = 38.20 m</td>
                  <td class="p-2 border border-slate-700 text-slate-300">Bệ: 12.6m</td>
                  <td class="p-2 border border-slate-700 text-slate-300 font-bold">L = 38.20 m</td>
                  <td class="p-2 border border-slate-700 text-slate-300">Bệ: 9.8m</td>
                </tr>
                <tr class="bg-slate-750">
                  <td class="p-2 border border-slate-700 font-medium text-slate-400 bg-slate-800">Trạng thái</td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-amber-500/20 text-amber-300 border border-amber-500/30 text-[10px] font-bold">LÊN BỆ MỐ</span>
                  </td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-slate-700 text-slate-300 border border-slate-600 text-[10px] font-bold">KH: @@N1_BD@@</span>
                  </td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 text-[10px] font-bold">XONG ĐỐT 1</span>
                  </td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-slate-700 text-slate-300 border border-slate-600 text-[10px] font-bold">KH: @@N2_BD@@</span>
                  </td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 text-[10px] font-bold">XONG ĐỐT 1</span>
                  </td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-slate-700 text-slate-300 border border-slate-600 text-[10px] font-bold">KH: @@N3_BD@@</span>
                  </td>
                  <td class="p-2 border border-slate-700">
                    <span class="px-2 py-0.5 rounded bg-emerald-500/20 text-emerald-300 border border-emerald-500/30 text-[10px] font-bold">ĐÃ LÊN THÂN</span>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        <!-- 2. HIGH PRECISION SVG ELEVATION SCHEMATIC -->
        <div class="bg-slate-900/90 border border-slate-700/80 rounded-xl p-3 relative overflow-hidden">
          <div class="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2 flex items-center justify-between">
            <span class="flex items-center gap-2"><span>📐</span> Sơ Đồ Trắc Dọc Cầu Km19+529.080 & Mã Màu Tiến Độ As-Built</span>
            <span class="text-[11px] text-slate-500 italic">Click vào từng cấu kiện để xem chi tiết nghiệm thu KCS</span>
          </div>

          <!-- SVG Canvas (Width 1000, Height 500) -->
          <svg viewBox="0 0 1000 480" class="w-full h-auto select-none" id="bridgeSvg">
            <defs>
              <!-- Gradients for realistic civil engineering feel -->
              <linearGradient id="gradGround" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stop-color="#334155" stop-opacity="0.8"/>
                <stop offset="100%" stop-color="#1e293b" stop-opacity="0.9"/>
              </linearGradient>
              <linearGradient id="gradDone" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stop-color="#22c55e"/>
                <stop offset="100%" stop-color="#15803d"/>
              </linearGradient>
              <linearGradient id="gradProgress" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stop-color="#f59e0b"/>
                <stop offset="100%" stop-color="#b45309"/>
              </linearGradient>
              <linearGradient id="gradNotDone" x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stop-color="#475569"/>
                <stop offset="100%" stop-color="#334155"/>
              </linearGradient>
              <pattern id="diagHatch" width="10" height="10" patternTransform="rotate(45 0 0)" patternUnits="userSpaceOnUse">
                <line x1="0" y1="0" x2="0" y2="10" stroke="#f59e0b" stroke-width="2" />
              </pattern>
            </defs>

            <!-- TẦNG ĐỊA CHẤT / GROUND LEVEL -->
            <path d="M 0 320 Q 200 310, 500 330 T 1000 320 L 1000 480 L 0 480 Z" fill="url(#gradGround)" opacity="0.35"/>
            <line x1="10" y1="320" x2="990" y2="320" stroke="#64748b" stroke-dasharray="6,4" stroke-width="1.5"/>
            <text x="20" y="335" fill="#94a3b8" font-size="11" font-weight="600">ĐƯỜNG MẶT ĐẤT TỰ NHIÊN / LÒNG SUỐI</text>

            <!-- ======================================================== -->
            <!-- MỐ M1 (X = 90) -->
            <!-- ======================================================== -->
            <g id="grp_m1" class="svg-node" onclick="inspectElem('M1')">
              <!-- Cọc khoan nhồi M1 (3 cọc L=20m) -->
              <rect x="75" y="320" width="16" height="130" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="97" y="320" width="16" height="130" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="119" y="320" width="16" height="130" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <text x="105" y="440" fill="#ffffff" font-size="10" text-anchor="middle" font-weight="bold">3 CỌC D1000 (XONG)</text>

              <!-- Bê tông lót mố M1 -->
              <rect x="65" y="312" width="80" height="8" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1"/>

              <!-- Bệ mố M1 (Đã hoàn thành) -->
              <rect x="68" y="275" width="74" height="37" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="2"/>
              <text x="105" y="298" fill="#ffffff" font-size="11" text-anchor="middle" font-weight="bold">BỆ MỐ M1</text>

              <!-- Thân mố M1 (Đang triển khai - Vàng Cam) -->
              <rect x="78" y="210" width="54" height="65" rx="3" fill="url(#gradProgress)" stroke="#f59e0b" stroke-width="2" class="pulse-amber"/>
              <text x="105" y="245" fill="#ffffff" font-size="11" text-anchor="middle" font-weight="bold">THÂN M1</text>
              <text x="105" y="260" fill="#fef08a" font-size="9" text-anchor="middle">(GIA CÔNG THÉP)</text>

              <!-- Đỉnh mố & Tường cánh M1 (Chưa thi công) -->
              <polygon points="78,210 132,210 138,175 78,175" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5"/>
              <text x="105" y="195" fill="#cbd5e1" font-size="10" text-anchor="middle">ĐỈNH & CÁNH</text>

              <!-- Label Flag Mố M1 -->
              <rect x="65" y="140" width="80" height="26" rx="4" fill="#0f172a" stroke="#f59e0b" stroke-width="1.5"/>
              <text x="105" y="152" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">MỐ M1</text>
              <text x="105" y="163" fill="#f59e0b" font-size="8.5" text-anchor="middle" font-weight="bold">LÊN BỆ, ĐANG LÀM THÂN</text>
            </g>


            <!-- ======================================================== -->
            <!-- TRỤ T1 (X = 360) -->
            <!-- ======================================================== -->
            <g id="grp_t1" class="svg-node" onclick="inspectElem('T1')">
              <!-- Cọc khoan nhồi T1 (8 cọc D1200 L=40m) -->
              <rect x="330" y="320" width="14" height="150" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="348" y="320" width="14" height="150" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="366" y="320" width="14" height="150" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="384" y="320" width="14" height="150" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <text x="364" y="445" fill="#ffffff" font-size="10" text-anchor="middle" font-weight="bold">8 CỌC D1200 (XONG)</text>

              <!-- Bê tông lót T1 -->
              <rect x="320" y="312" width="88" height="8" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1"/>

              <!-- Bệ trụ T1 (Đã hoàn thành Đợt 1 + 2) -->
              <rect x="323" y="270" width="82" height="42" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="2"/>
              <text x="364" y="295" fill="#ffffff" font-size="11" text-anchor="middle" font-weight="bold">BỆ TRỤ T1</text>

              <!-- Thân trụ T1 - Đốt 1 (ĐÃ THI CÔNG - Xanh lá) -->
              <rect x="344" y="225" width="40" height="45" rx="2" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <text x="364" y="250" fill="#ffffff" font-size="10" text-anchor="middle" font-weight="bold">ĐỐT 1 (XONG)</text>

              <!-- Thân trụ T1 - Đốt 2 (ĐANG TRIỂN KHAI - Vàng Cam) -->
              <rect x="344" y="180" width="40" height="45" rx="2" fill="url(#gradProgress)" stroke="#f59e0b" stroke-width="1.5" class="pulse-amber"/>
              <text x="364" y="202" fill="#ffffff" font-size="9" text-anchor="middle" font-weight="bold">ĐỐT 2</text>
              <text x="364" y="214" fill="#fef08a" font-size="8" text-anchor="middle">(VÁN KHUÔN)</text>

              <!-- Thân trụ T1 - Đốt 3 (CHƯA THI CÔNG) -->
              <rect x="344" y="140" width="40" height="40" rx="2" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5"/>
              <text x="364" y="163" fill="#cbd5e1" font-size="9" text-anchor="middle">ĐỐT 3</text>

              <!-- Xà mũ trụ T1 & Đá kê gối (CHƯA THI CÔNG) -->
              <polygon points="325,140 403,140 415,115 313,115" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5"/>
              <rect x="330" y="108" width="12" height="7" fill="#64748b"/>
              <rect x="386" y="108" width="12" height="7" fill="#64748b"/>
              <text x="364" y="130" fill="#cbd5e1" font-size="10" text-anchor="middle" font-weight="bold">XÀ MŨ T1</text>

              <!-- Label Flag Trụ T1 -->
              <rect x="324" y="75" width="80" height="26" rx="4" fill="#0f172a" stroke="#22c55e" stroke-width="1.5"/>
              <text x="364" y="87" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">TRỤ T1</text>
              <text x="364" y="98" fill="#22c55e" font-size="8.5" text-anchor="middle" font-weight="bold">ĐÃ XONG ĐỐT 1</text>
            </g>


            <!-- ======================================================== -->
            <!-- TRỤ T2 (X = 640) -->
            <!-- ======================================================== -->
            <g id="grp_t2" class="svg-node" onclick="inspectElem('T2')">
              <!-- Cọc khoan nhồi T2 (8 cọc D1200 L=30m) -->
              <rect x="610" y="320" width="14" height="135" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="628" y="320" width="14" height="135" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="646" y="320" width="14" height="135" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="664" y="320" width="14" height="135" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <text x="644" y="440" fill="#ffffff" font-size="10" text-anchor="middle" font-weight="bold">8 CỌC D1200 (XONG)</text>

              <!-- Bê tông lót T2 -->
              <rect x="600" y="312" width="88" height="8" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1"/>

              <!-- Bệ trụ T2 (Đã hoàn thành Đợt 1 + 2) -->
              <rect x="603" y="270" width="82" height="42" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="2"/>
              <text x="644" y="295" fill="#ffffff" font-size="11" text-anchor="middle" font-weight="bold">BỆ TRỤ T2</text>

              <!-- Thân trụ T2 - Đốt 1 (ĐÃ THI CÔNG - Xanh lá) -->
              <rect x="624" y="225" width="40" height="45" rx="2" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <text x="644" y="250" fill="#ffffff" font-size="10" text-anchor="middle" font-weight="bold">ĐỐT 1 (XONG)</text>

              <!-- Thân trụ T2 - Đốt 2 (ĐANG TRIỂN KHAI - Vàng Cam) -->
              <rect x="624" y="180" width="40" height="45" rx="2" fill="url(#gradProgress)" stroke="#f59e0b" stroke-width="1.5" class="pulse-amber"/>
              <text x="644" y="202" fill="#ffffff" font-size="9" text-anchor="middle" font-weight="bold">ĐỐT 2</text>
              <text x="644" y="214" fill="#fef08a" font-size="8" text-anchor="middle">(VÁN KHUÔN)</text>

              <!-- Thân trụ T2 - Đốt 3 (CHƯA THI CÔNG) -->
              <rect x="624" y="140" width="40" height="40" rx="2" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5"/>
              <text x="644" y="163" fill="#cbd5e1" font-size="9" text-anchor="middle">ĐỐT 3</text>

              <!-- Xà mũ trụ T2 & Đá kê gối (CHƯA THI CÔNG) -->
              <polygon points="605,140 683,140 695,115 593,115" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5"/>
              <rect x="610" y="108" width="12" height="7" fill="#64748b"/>
              <rect x="666" y="108" width="12" height="7" fill="#64748b"/>
              <text x="644" y="130" fill="#cbd5e1" font-size="10" text-anchor="middle" font-weight="bold">XÀ MŨ T2</text>

              <!-- Label Flag Trụ T2 -->
              <rect x="604" y="75" width="80" height="26" rx="4" fill="#0f172a" stroke="#22c55e" stroke-width="1.5"/>
              <text x="644" y="87" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">TRỤ T2</text>
              <text x="644" y="98" fill="#22c55e" font-size="8.5" text-anchor="middle" font-weight="bold">ĐÃ XONG ĐỐT 1</text>
            </g>


            <!-- ======================================================== -->
            <!-- MỐ M2 (X = 900) -->
            <!-- ======================================================== -->
            <g id="grp_m2" class="svg-node" onclick="inspectElem('M2')">
              <!-- Cọc khoan nhồi M2 (7 cọc D1200 L=36m) -->
              <rect x="865" y="320" width="15" height="145" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="886" y="320" width="15" height="145" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="907" y="320" width="15" height="145" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <rect x="928" y="320" width="15" height="145" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1.5"/>
              <text x="905" y="445" fill="#ffffff" font-size="10" text-anchor="middle" font-weight="bold">7 CỌC D1200 (XONG)</text>

              <!-- Bê tông lót mố M2 -->
              <rect x="858" y="312" width="94" height="8" fill="url(#gradDone)" stroke="#16a34a" stroke-width="1"/>

              <!-- Bệ mố M2 (Đã hoàn thành) -->
              <rect x="862" y="275" width="86" height="37" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="2"/>
              <text x="905" y="298" fill="#ffffff" font-size="11" text-anchor="middle" font-weight="bold">BỆ MỐ M2</text>

              <!-- Thân mố M2 (ĐÃ HOÀN THÀNH - Xanh lá theo lệnh) -->
              <rect x="872" y="200" width="66" height="75" rx="3" fill="url(#gradDone)" stroke="#16a34a" stroke-width="2"/>
              <text x="905" y="238" fill="#ffffff" font-size="11" text-anchor="middle" font-weight="bold">THÂN MỐ M2</text>
              <text x="905" y="253" fill="#bbf7d0" font-size="9" text-anchor="middle">(ĐÃ ĐỔ XONG THÂN)</text>

              <!-- Đỉnh mố & Tường cánh M2 (ĐANG TRIỂN KHAI) -->
              <polygon points="872,200 938,200 945,175 872,175" fill="url(#gradProgress)" stroke="#f59e0b" stroke-width="1.5" class="pulse-amber"/>
              <text x="905" y="192" fill="#ffffff" font-size="9.5" text-anchor="middle" font-weight="bold">ĐỈNH & CÁNH</text>

              <!-- Label Flag Mố M2 -->
              <rect x="865" y="140" width="80" height="26" rx="4" fill="#0f172a" stroke="#22c55e" stroke-width="1.5"/>
              <text x="905" y="152" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">MỐ M2</text>
              <text x="905" y="163" fill="#22c55e" font-size="8.5" text-anchor="middle" font-weight="bold">ĐÃ LÊN THÂN MỐ</text>
            </g>


            <!-- ======================================================== -->
            <!-- HỆ DẦM SUPER-T (3 NHỊP L=38.2M) & BẢN MẶT CẦU -->
            <!-- ======================================================== -->
            <!-- NHỊP 1 (M1 - T1) -->
            <g id="grp_span1" class="svg-node" onclick="inspectElem('SPAN1')">
              <rect x="142" y="98" width="171" height="15" rx="2" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="4,2"/>
              <text x="227" y="109" fill="#cbd5e1" font-size="9" text-anchor="middle">DẦM SUPER-T NHỊP 1 (KH ĐÚC: @@N1_BD@@)</text>
              <rect x="138" y="90" width="175" height="7" fill="url(#gradNotDone)" stroke="#64748b" stroke-width="1"/>
            </g>

            <!-- NHỊP 2 (T1 - T2) -->
            <g id="grp_span2" class="svg-node" onclick="inspectElem('SPAN2')">
              <rect x="415" y="98" width="178" height="15" rx="2" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="4,2"/>
              <text x="504" y="109" fill="#cbd5e1" font-size="9" text-anchor="middle">DẦM SUPER-T NHỊP 2 (KH ĐÚC: @@N2_BD@@)</text>
              <rect x="415" y="90" width="178" height="7" fill="url(#gradNotDone)" stroke="#64748b" stroke-width="1"/>
            </g>

            <!-- NHỊP 3 (T2 - M2) -->
            <g id="grp_span3" class="svg-node" onclick="inspectElem('SPAN3')">
              <rect x="695" y="98" width="177" height="15" rx="2" fill="url(#gradNotDone)" stroke="#94a3b8" stroke-width="1.5" stroke-dasharray="4,2"/>
              <text x="783" y="109" fill="#cbd5e1" font-size="9" text-anchor="middle">DẦM SUPER-T NHỊP 3 (KH ĐÚC: @@N3_BD@@)</text>
              <rect x="695" y="90" width="177" height="7" fill="url(#gradNotDone)" stroke="#64748b" stroke-width="1"/>
            </g>

            <!-- Ghi chú nhịp kích thước phía trên -->
            <line x1="140" y1="55" x2="315" y2="55" stroke="#38bdf8" stroke-width="1"/>
            <text x="227" y="50" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">Nhịp 1: L = 38.20 m</text>

            <line x1="415" y1="55" x2="593" y2="55" stroke="#38bdf8" stroke-width="1"/>
            <text x="504" y="50" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">Nhịp 2: L = 38.20 m</text>

            <line x1="695" y1="55" x2="872" y2="55" stroke="#38bdf8" stroke-width="1"/>
            <text x="783" y="50" fill="#38bdf8" font-size="10" text-anchor="middle" font-weight="bold">Nhịp 3: L = 38.20 m</text>
          </svg>
        </div>

        <!-- 3. PILE LAYOUT PLAN (MẶT BẰNG BỐ TRÍ CỌC KHOAN NHỒI FOUNDTECH) -->
        <div class="mt-4 pt-3 border-t border-slate-700/80">
          <div class="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2 flex items-center justify-between">
            <span class="flex items-center gap-2"><span>📍</span> Mặt Bằng Bố Trí Cọc Khoan Nhồi Từng Mố Trụ (FOUNDTECH BORED PILES)</span>
            <span class="text-emerald-400 font-semibold text-xs">Đạt 100% (26/26 Cọc Đã Đổ Bê Tông)</span>
          </div>

          <div class="grid grid-cols-4 gap-3 text-center">

            <!-- M1 PILE PLAN (3 cọc) -->
            <div class="bg-slate-900 border border-slate-700 rounded-xl p-2.5">
              <div class="text-[11px] font-bold text-amber-300">BỆ MỐ M1 (3 CỌC D1000)</div>
              <div class="mt-2 flex justify-center items-center gap-2 bg-slate-950/80 p-2 rounded-lg border border-slate-800">
                <span class="w-6 h-6 rounded-full bg-emerald-500 text-white font-bold text-[10px] flex items-center justify-center border border-emerald-300 shadow-sm" title="Cọc C1 mố M1: 24.71 m3">1</span>
                <span class="w-6 h-6 rounded-full bg-emerald-500 text-white font-bold text-[10px] flex items-center justify-center border border-emerald-300 shadow-sm" title="Cọc C2 mố M1: 24.71 m3">2</span>
                <span class="w-6 h-6 rounded-full bg-emerald-500 text-white font-bold text-[10px] flex items-center justify-center border border-emerald-300 shadow-sm" title="Cọc C3 mố M1: 24.71 m3">3</span>
              </div>
              <div class="text-[10px] text-emerald-400 mt-1 font-semibold">L=20m • Xong 3/3 cọc</div>
            </div>

            <!-- T1 PILE PLAN (8 cọc) -->
            <div class="bg-slate-900 border border-slate-700 rounded-xl p-2.5">
              <div class="text-[11px] font-bold text-cyan-300">BỆ TRỤ T1 (8 CỌC D1200)</div>
              <div class="mt-2 grid grid-cols-4 gap-1.5 bg-slate-950/80 p-1.5 rounded-lg border border-slate-800">
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C1: 47.01 m3">1</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C2: 47.01 m3">2</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C3: 47.01 m3">3</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C4: 47.01 m3">4</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C5: 47.01 m3">5</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C6: 47.01 m3">6</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C7: 47.01 m3">7</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C8: 47.01 m3">8</span>
              </div>
              <div class="text-[10px] text-emerald-400 mt-1 font-semibold">L=40m • Xong 8/8 cọc</div>
            </div>

            <!-- T2 PILE PLAN (8 cọc) -->
            <div class="bg-slate-900 border border-slate-700 rounded-xl p-2.5">
              <div class="text-[11px] font-bold text-cyan-300">BỆ TRỤ T2 (8 CỌC D1200)</div>
              <div class="mt-2 grid grid-cols-4 gap-1.5 bg-slate-950/80 p-1.5 rounded-lg border border-slate-800">
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C1: 35.86 m3">1</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C2: 35.86 m3">2</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C3: 35.86 m3">3</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C4: 35.86 m3">4</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C5: 35.86 m3">5</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C6: 35.86 m3">6</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C7: 35.86 m3">7</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto" title="Cọc C8: 35.86 m3">8</span>
              </div>
              <div class="text-[10px] text-emerald-400 mt-1 font-semibold">L=30m • Xong 8/8 cọc</div>
            </div>

            <!-- M2 PILE PLAN (7 cọc) -->
            <div class="bg-slate-900 border border-slate-700 rounded-xl p-2.5">
              <div class="text-[11px] font-bold text-amber-300">BỆ MỐ M2 (7 CỌC D1200)</div>
              <div class="mt-2 grid grid-cols-4 gap-1.5 bg-slate-950/80 p-1.5 rounded-lg border border-slate-800">
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">1</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">2</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">3</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">4</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">5</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">6</span>
                <span class="w-5 h-5 rounded-full bg-emerald-500 text-white font-bold text-[9px] flex items-center justify-center mx-auto">7</span>
                <span class="w-5 h-5 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center mx-auto text-slate-600 text-[8px]">-</span>
              </div>
              <div class="text-[10px] text-emerald-400 mt-1 font-semibold">L=36m • Xong 7/7 cọc</div>
            </div>

          </div>
        </div>

        <!-- 4. TIẾN ĐỘ DẦM SUPER-T PHÂN NHỊP & CHU KỲ ĐÚC 7 NGÀY/PHIẾN -->
        <div class="mt-4 pt-3 border-t border-slate-700/80">
          <div class="text-xs font-bold uppercase tracking-wider text-slate-400 mb-2 flex items-center justify-between">
            <span class="flex items-center gap-2"><span>🏗️</span> Kế Hoạch Bãi Đúc Dầm Super-T (Khởi Công: @@N1_BD@@ — Chu Kỳ: @@CHU_KY@@ Ngày / 1 Phiến)</span>
            <span class="text-slate-400 font-semibold text-xs">Hiện tại: @@DA_DUC@@/@@TONG_PHIEN@@ phiến (@@TT_TONG@@)</span>
          </div>

          <div class="grid grid-cols-3 gap-3 text-xs text-center">
            <div class="bg-slate-900 border border-slate-700 rounded-xl p-3">
              <div class="font-bold text-slate-300">NHỊP 1 (M1 - T1, L=38.20M)</div>
              <div class="text-[11px] text-amber-400 font-semibold mt-1">Khởi công đúc: @@N1_BD@@</div>
              <div class="w-full bg-slate-800 h-2.5 rounded-full overflow-hidden mt-2">
                <div class="bg-blue-500 h-full rounded-full" style="width: 0%"></div>
              </div>
              <div class="flex justify-between text-[10px] text-slate-400 mt-1">
                <span>@@N1_TIEN_DO@@</span>
                <span class="text-slate-400">Dự kiến xong: @@N1_KT@@</span>
              </div>
            </div>

            <div class="bg-slate-900 border border-slate-700 rounded-xl p-3">
              <div class="font-bold text-slate-300">NHỊP 2 (T1 - T2, L=38.20M)</div>
              <div class="text-[11px] text-slate-400 mt-1">Đúc nối tiếp sau Nhịp 1</div>
              <div class="w-full bg-slate-800 h-2.5 rounded-full overflow-hidden mt-2">
                <div class="bg-blue-500 h-full rounded-full" style="width: 0%"></div>
              </div>
              <div class="flex justify-between text-[10px] text-slate-400 mt-1">
                <span>@@N2_TIEN_DO@@</span>
                <span class="text-slate-400">Dự kiến: @@N2_BD_NGAN@@ - @@N2_KT@@</span>
              </div>
            </div>

            <div class="bg-slate-900 border border-slate-700 rounded-xl p-3">
              <div class="font-bold text-slate-300">NHỊP 3 (T2 - M2, L=38.20M)</div>
              <div class="text-[11px] text-slate-400 mt-1">Đúc nối tiếp sau Nhịp 2</div>
              <div class="w-full bg-slate-800 h-2.5 rounded-full overflow-hidden mt-2">
                <div class="bg-blue-500 h-full rounded-full" style="width: 0%"></div>
              </div>
              <div class="flex justify-between text-[10px] text-slate-400 mt-1">
                <span>@@N3_TIEN_DO@@</span>
                <span class="text-slate-400">Dự kiến: @@N3_BD_NGAN@@ - @@N3_KT@@</span>
              </div>
            </div>
          </div>

          <!-- BẢNG CHI TIẾT CHU KỲ 7 NGÀY / 1 PHIẾN -->
          <div class="mt-3 bg-slate-950/70 p-2.5 rounded-xl border border-slate-800 text-[11px] text-slate-300 flex flex-wrap items-center justify-between gap-2">
            <div>
              ⏱️ <strong>Chu kỳ @@CHU_KY@@ ngày / 1 phiến:</strong> Ngày 1-2 (Lắp cốt thép & ống ghen) → Ngày 3 (Ghép ván khuôn & đổ BT C45) → Ngày 4-5 (Bảo dưỡng đạt 85% R28) → Ngày 6 (Căng kéo cáp DƯL 15.2mm & bơm vữa) → Ngày 7 (Cẩu dầm ra bãi chứa luân chuyển bệ đúc).
            </div>
            <div class="text-cyan-400 font-bold">
              Tổng @@TONG_PHIEN@@ phiến: @@NGAY_1_BE@@ ngày (1 bệ đúc) hoặc @@NGAY_2_BE@@ ngày (2 bệ đúc song song)
            </div>
          </div>
        </div>

      </div>

      <!-- RIGHT: LIVE INSPECTOR & ENGINEERING AUDIT REPORT (1 COL) -->
      <div class="space-y-4">

        <!-- DYNAMIC INSPECTOR PANEL -->
        <div class="bg-slate-800/90 border border-slate-700 rounded-2xl p-5 shadow-xl" id="inspectCard">
          <div class="flex items-center justify-between pb-3 border-b border-slate-700">
            <h3 class="text-sm font-bold text-white uppercase tracking-wider flex items-center gap-2">
              <span id="inspectIcon">🔍</span> <span id="inspectTitle">Chi Tiết Cấu Kiện</span>
            </h3>
            <span id="inspectBadge" class="px-2 py-0.5 rounded text-[11px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
              ĐÃ THI CÔNG
            </span>
          </div>

          <div class="mt-4 space-y-3 text-xs" id="inspectBody">
            <p class="text-slate-400 leading-relaxed">
              Nhấp chuột trực tiếp vào bất kỳ bộ phận nào trên cây cầu (Mố M1, Mố M2, Trụ T1, Trụ T2, Dầm Super-T) để tra cứu thông số kỹ thuật, thể tích bê tông, mác thiết kế và nhật ký KCS hoàn công.
            </p>
            <div class="p-3 bg-slate-900/80 rounded-xl border border-slate-700 space-y-2">
              <div class="flex justify-between"><span class="text-slate-400">Khối lượng lũy kế:</span> <span class="font-bold text-emerald-400">1.731,44 m³</span></div>
              <div class="flex justify-between"><span class="text-slate-400">Hạ bộ hoàn thành:</span> <span class="font-bold text-emerald-400">86.5%</span></div>
              <div class="flex justify-between"><span class="text-slate-400">Số lô thép đã nhập:</span> <span class="font-bold text-slate-200">11 chủng loại $\varnothing$</span></div>
              <div class="flex justify-between"><span class="text-slate-400">Số mẫu bê tông R28:</span> <span class="font-bold text-slate-200">79 tổ mẫu</span></div>
            </div>
          </div>
        </div>

        <!-- SITE DIRECTIVE & ROADMAP ACTIONS -->
        <div class="bg-slate-800/90 border border-slate-700 rounded-2xl p-5 shadow-xl space-y-3">
          <h3 class="text-xs font-bold text-amber-300 uppercase tracking-wider flex items-center gap-2">
            <span>⚡</span> Chỉ Đạo Kỹ Thuật Ban Chỉ Huy Hiện Trường
          </h3>
          <ul class="text-xs space-y-2.5 text-slate-300">
            <li class="flex items-start gap-2 bg-slate-900/60 p-2.5 rounded-lg border border-slate-800">
              <span class="text-amber-400 font-bold mt-0.5">1.</span>
              <div>
                <strong class="text-white">Trụ T1 & T2 (Đường găng CPM):</strong> Đã đổ xong Đốt 1. Tập trung toàn bộ ván khuôn leo và nhân lực thi công Đốt 2 ($V=42\text{m}^3$ & $45\text{m}^3$). Bảo dưỡng ẩm bê tông Đốt 1 tối thiểu 7 ngày.
              </div>
            </li>
            <li class="flex items-start gap-2 bg-slate-900/60 p-2.5 rounded-lg border border-slate-800">
              <span class="text-emerald-400 font-bold mt-0.5">2.</span>
              <div>
                <strong class="text-white">Mố M2:</strong> Đã hoàn thành toàn bộ thân mố! Bàn giao mặt bằng cho tổ đội ván khuôn tường đỉnh và tường cánh để kịp tiến độ lao lắp nhịp 3.
              </div>
            </li>
            <li class="flex items-start gap-2 bg-slate-900/60 p-2.5 rounded-lg border border-slate-800">
              <span class="text-cyan-400 font-bold mt-0.5">3.</span>
              <div>
                <strong class="text-white">Mố M1:</strong> Bệ mố đã nghiệm thu. Đẩy nhanh gia công lắp dựng cốt thép tường thân Đợt 1 để tổ chức đổ bê tông trong 3 ngày tới.
              </div>
            </li>
            <li class="flex items-start gap-2 bg-slate-900/60 p-2.5 rounded-lg border border-slate-800">
              <span class="text-blue-400 font-bold mt-0.5">4.</span>
              <div>
                <strong class="text-white">Bãi đúc dầm Super-T (Kế hoạch @@N1_BD@@):</strong> @@BAI_DUC_HIEN_TRANG@@ Tập trung hoàn thiện bệ đúc, kiểm định kích thủy lực căng kéo cáp DƯL 15.2mm, tập kết cốt thép và nghiệm thu trạm trộn bê tông C45 để sẵn sàng bấm nút khởi công đúc ngày @@N1_BD_NGAN@@ theo chu kỳ @@CHU_KY@@ ngày/phiến.
              </div>
            </li>
          </ul>
        </div>

      </div>

    </div>

  </div>

  <!-- JAVASCRIPT FOR INTERACTIVITY -->
  <script>
    const ElemData = {
      'M1': {
        icon: '🧱',
        title: 'HẠNG MỤC MỐ M1 (KM19+471.78)',
        badge: 'LÊN BỆ MỐ - ĐANG LÀM THÂN',
        badgeClass: 'bg-amber-500/20 text-amber-300 border-amber-500/30',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Cọc khoan nhồi:</span> <span class="text-emerald-400 font-bold">3 cọc D1000, L=20m (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Bê tông bệ mố:</span> <span class="text-emerald-400 font-bold">74.50 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân mố Đợt 1:</span> <span class="text-amber-400 font-bold">45.00 m³ (Đang lắp cốt thép)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân mố Đợt 2:</span> <span class="text-slate-400 font-bold">39.44 m³ (Chưa thi công)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tường đỉnh & cánh:</span> <span class="text-slate-400 font-bold">24.46 m³ (Chưa thi công)</span></div>
            <div class="p-2.5 bg-slate-900 rounded-lg text-slate-300 text-[11px] leading-relaxed">
              💡 <strong>Hiện trạng:</strong> Bệ mố đã xong đạt cường độ R28. Đang lắp đặt cốt thép thân mố Đợt 1, nghiệm thu ván khuôn dự kiến trong tuần.
            </div>
          </div>
        `
      },
      'T1': {
        icon: '🗼',
        title: 'HẠNG MỤC TRỤ T1 (KM19+509.98)',
        badge: 'ĐÃ XONG ĐỐT 1 - LẮP ĐỐT 2',
        badgeClass: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Cọc khoan nhồi:</span> <span class="text-emerald-400 font-bold">8 cọc D1200, L=40m (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Bệ trụ (2 đợt):</span> <span class="text-emerald-400 font-bold">165.00 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân trụ Đốt 1:</span> <span class="text-emerald-400 font-bold">42.00 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân trụ Đốt 2:</span> <span class="text-amber-400 font-bold">42.00 m³ (Lắp ván khuôn leo)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân Đốt 3 & Xà mũ:</span> <span class="text-slate-400 font-bold">83.00 m³ (Chưa thi công)</span></div>
            <div class="p-2.5 bg-slate-900 rounded-lg text-slate-300 text-[11px] leading-relaxed">
              💡 <strong>Đường găng:</strong> Đốt 1 đã đổ xong ngày 28/09, kết quả R7 đạt 88% thiết kế. Hệ ván khuôn trượt/leo đang được cân chỉnh để đổ Đốt 2.
            </div>
          </div>
        `
      },
      'T2': {
        icon: '🗼',
        title: 'HẠNG MỤC TRỤ T2 (KM19+548.18)',
        badge: 'ĐÃ XONG ĐỐT 1 - LẮP ĐỐT 2',
        badgeClass: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Cọc khoan nhồi:</span> <span class="text-emerald-400 font-bold">8 cọc D1200, L=30m (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Bệ trụ (2 đợt):</span> <span class="text-emerald-400 font-bold">165.00 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân trụ Đốt 1:</span> <span class="text-emerald-400 font-bold">45.00 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân trụ Đốt 2:</span> <span class="text-amber-400 font-bold">45.00 m³ (Gia công cốt thép)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân Đốt 3 & Xà mũ:</span> <span class="text-slate-400 font-bold">87.00 m³ (Chưa thi công)</span></div>
            <div class="p-2.5 bg-slate-900 rounded-lg text-slate-300 text-[11px] leading-relaxed">
              💡 <strong>Kiểm định:</strong> Đốt 1 nghiệm thu đạt độ thẳng đứng dung sai < 5mm. Đang bố trí cẩu tháp hỗ trợ đổ bê tông Đốt 2.
            </div>
          </div>
        `
      },
      'M2': {
        icon: '🧱',
        title: 'HẠNG MỤC MỐ M2 (KM19+586.38)',
        badge: 'ĐÃ LÊN THÂN HOÀN THÀNH',
        badgeClass: 'bg-emerald-500/20 text-emerald-300 border-emerald-500/30',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Cọc khoan nhồi:</span> <span class="text-emerald-400 font-bold">7 cọc D1200, L=36m (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Bê tông bệ mố:</span> <span class="text-emerald-400 font-bold">74.50 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Thân mố Đợt 1 + 2:</span> <span class="text-emerald-400 font-bold">84.44 m³ C30 (ĐÃ XONG)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tường đỉnh & đá gối:</span> <span class="text-amber-400 font-bold">14.33 m³ (Đang ván khuôn)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tường cánh mố:</span> <span class="text-amber-400 font-bold">10.13 m³ (Đang ván khuôn)</span></div>
            <div class="p-2.5 bg-slate-900 rounded-lg text-slate-300 text-[11px] leading-relaxed">
              💡 <strong>Tiến độ vượt mốc:</strong> Thân mố M2 đã hoàn thành 100%! Đang làm thép tường đỉnh mố và đá kê gối để sẵn sàng tiếp nhận dầm nhịp 3.
            </div>
          </div>
        `
      },
      'SPAN1': {
        icon: '🌉',
        title: 'KẾT CẤU NHỊP 1 (M1 - T1, L=38.2M)',
        badge: '@@N1_TT_HOA@@ (KH: @@N1_BD@@)',
        badgeClass: 'bg-slate-700 text-slate-300 border-slate-600',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Quy cách dầm:</span> <span class="text-white font-bold">5 phiến Super-T 38.2m C45</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tình trạng hiện tại:</span> <span class="text-slate-300 font-bold">@@N1_TT@@</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Kế hoạch khởi công:</span> <span class="text-amber-400 font-bold">@@N1_BD@@</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Chu kỳ công nghệ:</span> <span class="text-cyan-400 font-bold">@@CHU_KY@@ ngày / 1 phiến</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Dự kiến hoàn thành:</span> <span class="text-emerald-400 font-bold">@@N1_KT@@ (@@NGAY_MOI_NHIP@@ ngày)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tổng BT @@PHIEN_MOI_NHIP@@ phiến:</span> <span class="text-white font-bold">@@BT_NHIP@@ m³</span></div>
            <div class="p-2.5 bg-slate-900 rounded-lg text-slate-300 text-[11px] leading-relaxed">
              💡 <strong>Kế hoạch @@PHIEN_MOI_NHIP@@ phiến Nhịp 1:</strong> @@N1_LICH@@.
            </div>
          </div>
        `
      },
      'SPAN2': {
        icon: '🌉',
        title: 'KẾT CẤU NHỊP 2 (T1 - T2, L=38.2M)',
        badge: '@@N2_TT_HOA@@ (KH: @@N2_BD@@)',
        badgeClass: 'bg-slate-700 text-slate-300 border-slate-600',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Quy cách dầm:</span> <span class="text-white font-bold">5 phiến Super-T 38.2m C45</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tình trạng hiện tại:</span> <span class="text-slate-300 font-bold">@@N2_TT@@</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Kế hoạch đúc:</span> <span class="text-amber-400 font-bold">Đúc nối tiếp sau Nhịp 1</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Chu kỳ công nghệ:</span> <span class="text-cyan-400 font-bold">@@CHU_KY@@ ngày / 1 phiến</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Dự kiến đúc:</span> <span class="text-emerald-400 font-bold">@@N2_BD@@ - @@N2_KT@@ (@@NGAY_MOI_NHIP@@ ngày)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tổng BT @@PHIEN_MOI_NHIP@@ phiến:</span> <span class="text-white font-bold">@@BT_NHIP@@ m³</span></div>
          </div>
        `
      },
      'SPAN3': {
        icon: '🌉',
        title: 'KẾT CẤU NHỊP 3 (T2 - M2, L=38.2M)',
        badge: '@@N3_TT_HOA@@ (KH: @@N3_BD@@)',
        badgeClass: 'bg-slate-700 text-slate-300 border-slate-600',
        content: `
          <div class="space-y-2">
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Quy cách dầm:</span> <span class="text-white font-bold">5 phiến Super-T 38.2m C45</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tình trạng hiện tại:</span> <span class="text-slate-300 font-bold">@@N3_TT@@</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Kế hoạch đúc:</span> <span class="text-amber-400 font-bold">Đúc nối tiếp sau Nhịp 2</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Chu kỳ công nghệ:</span> <span class="text-cyan-400 font-bold">@@CHU_KY@@ ngày / 1 phiến</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Dự kiến đúc:</span> <span class="text-emerald-400 font-bold">@@N3_BD@@ - @@N3_KT@@ (@@NGAY_MOI_NHIP@@ ngày)</span></div>
            <div class="flex justify-between border-b border-slate-700/60 pb-1.5"><span class="text-slate-400">Tổng BT @@PHIEN_MOI_NHIP@@ phiến:</span> <span class="text-white font-bold">@@BT_NHIP@@ m³</span></div>
          </div>
        `
      }
    };

    function inspectElem(code) {
      const data = ElemData[code];
      if (!data) return;
      document.getElementById('inspectIcon').innerText = data.icon;
      document.getElementById('inspectTitle').innerText = data.title;
      const badge = document.getElementById('inspectBadge');
      badge.innerText = data.badge;
      badge.className = 'px-2 py-0.5 rounded text-[11px] font-bold border ' + data.badgeClass;
      document.getElementById('inspectBody').innerHTML = data.content;
    }

    function resetInspect() {
      inspectElem('T1');
    }
  </script>
</body>
</html>
"""
    html_content = fill_tokens(html_content, html_tokens(cfg))
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] Đã xuất bản tệp HTML Visual Dashboard: {output_path}")


def generate_excel_dashboard(output_path: str, cfg=None):
    """Tạo tệp Excel sơ đồ tiến độ trực quan có thể in ấn A3/A4 ngoài công trường"""
    cfg = cfg or load_config()
    data_rows = excel_rows(cfg)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SO_DO_TIEN_DO_TRUC_QUAN"
    ws.views.sheetView[0].showGridLines = True

    # Styling fonts & fills
    font_title = Font(name="Arial", size=15, bold=True, color="1B365D")
    font_sub = Font(name="Arial", size=10, italic=True, color="555555")
    font_hdr = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_bold = Font(name="Arial", size=9, bold=True)
    font_regular = Font(name="Arial", size=9)

    fill_navy = PatternFill("solid", fgColor="1B365D")
    fill_blue = PatternFill("solid", fgColor="2E75B6")
    fill_green = PatternFill("solid", fgColor="28A745") # Đã hoàn thành
    fill_amber = PatternFill("solid", fgColor="FFC000") # Đang triển khai
    fill_gray = PatternFill("solid", fgColor="F2F2F2")  # Chưa thi công
    fill_lot = PatternFill("solid", fgColor="D9E1F2")

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC")
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. Header rows
    ws.merge_cells("A1:N1")
    ws["A1"] = "BAN CHỈ HUY CÔNG TRƯỜNG — BẢNG THEO DÕI TIẾN ĐỘ THI CÔNG TRỰC QUAN CẦU KM19+529.080"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:N2")
    ws["A2"] = tom_tat(data_rows, cfg)
    ws["A2"].font = font_sub
    ws["A2"].alignment = align_center
    ws.row_dimensions[2].height = 20

    # 2. Legend row
    ws.merge_cells("A4:C4")
    ws["A4"] = "CHÚ THÍCH TRẠNG THÁI:"
    ws["A4"].font = font_bold
    ws["A4"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("D4:F4")
    ws["D4"] = "[XANH LÁ] ĐÃ THI CÔNG HOÀN THÀNH"
    ws["D4"].fill = fill_green
    ws["D4"].font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    ws["D4"].alignment = align_center

    ws.merge_cells("H4:J4")
    ws["H4"] = "[VÀNG CAM] ĐANG TRIỂN KHAI THI CÔNG"
    ws["H4"].fill = fill_amber
    ws["H4"].font = Font(name="Arial", size=9, bold=True, color="000000")
    ws["H4"].alignment = align_center

    ws.merge_cells("L4:N4")
    ws["L4"] = "[TRẮNG XÁM] CHƯA THI CÔNG"
    ws["L4"].fill = fill_gray
    ws["L4"].font = Font(name="Arial", size=9, bold=True, color="555555")
    ws["L4"].alignment = align_center

    # 3. Main Data Table
    headers = [
        "HẠNG MỤC CẤU KIỆN", "VỊ TRÍ", "QUY CÁCH / KÍCH THƯỚC", "MÁC BT",
        "KHỐI LƯỢNG (m³)", "TIẾN ĐỘ THIẾT KẾ", "NGÀY KHỞI CÔNG", "NGÀY HOÀN THÀNH",
        "TÌNH TRẠNG HIỆN TRƯỜNG", "ĐÁNH GIÁ KCS R28", "GHI CHÚ KỸ THUẬT"
    ]

    for col_idx, text in enumerate(headers, 1):
        cell = ws.cell(6, col_idx, text)
        cell.font = font_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[6].height = 26

    for row_idx, rdata in enumerate(data_rows, 7):
        ws.row_dimensions[row_idx].height = 20
        status = rdata[8]
        row_fill = fill_green if status == "ĐÃ THI CÔNG" else (fill_amber if status == "ĐANG TRIỂN KHAI" else fill_gray)

        for col_idx, val in enumerate(rdata, 1):
            cell = ws.cell(row_idx, col_idx, val)
            cell.font = font_regular
            cell.border = thin_border
            if col_idx in [5]:
                cell.number_format = "#,##0.00"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_idx in [1, 3, 11]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = align_center

            # Highlight status column
            if col_idx == 9:
                cell.fill = row_fill
                cell.font = font_bold
                if status == "ĐÃ THI CÔNG":
                    cell.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")

    # Column widths
    col_widths = [26, 12, 28, 14, 16, 16, 15, 15, 20, 18, 28]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_path)
    print(f"[OK] Đã xuất bản tệp Excel Sơ đồ tiến độ trực quan: {output_path}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Xuất sơ đồ tiến độ trực quan As-Built (HTML + Excel) từ cấu hình")
    ap.add_argument("--config", default=CONFIG_PATH)
    ap.add_argument("--html", default=os.path.join(ARTIFACT_DIR, "so_do_tien_do_truc_quan_cau_km19.html"))
    ap.add_argument("--xlsx", default=DEFAULT_XLSX)
    a = ap.parse_args(argv)
    cfg = load_config(a.config)
    generate_html_dashboard(a.html, cfg)
    os.makedirs(os.path.dirname(os.path.abspath(a.xlsx)), exist_ok=True)
    generate_excel_dashboard(a.xlsx, cfg)
    return 0


if __name__ == "__main__":
    sys.exit(main())
