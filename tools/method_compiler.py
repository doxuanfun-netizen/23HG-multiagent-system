# -*- coding: utf-8 -*-
"""
METHOD COMPILER — Bộ biên dịch Thuyết minh Biện pháp thi công (BPTC)
==================================================================

Triết lý deterministic-first: MỌI con số do code Python tính từ State Bus
(QS sống, CPM, Fleet, BBS) — LLM (nếu có) chỉ viết văn, tuyệt đối không tính số.

Input (nguồn số liệu duy nhất — hồ sơ thiết kế được duyệt):
  - QS/BOQ            : tools/qs_loader.load_qs         → QSEstimate
  - Tiến độ CPM       : tools/schedule_loader.load_schedule + CPMCalculator
  - BBS thép          : tools/bbs_loader.load_bbs       → CuttingStockSolver
  - Fleet (máy + ca)  : tools/equipment_fleet_scheduler → EquipmentFleetScheduler
  - project (JSON)    : tham số công trình (tên, loại, địa chất, Karst, thiết bị...)

Output:
  - method_params.json : Single Source of Truth — mọi con số + trace về nguồn
  - Thuyet_Minh_BPTC_<cong_trinh>.md  (+ .docx nếu có python-docx)
  - DOI_SOAT_SO_LIEU_BPTC.md          : bảng đối soát số → dòng QS/CPM/Fleet/BBS

Mọi số trong BPTC phải trace được về method_params.json; số không trace được
= số chết → method_auditor báo lỗi. Không bao giờ tự điền mác bê tông / Ø thép /
tải trọng thiết bị nếu chưa có trong QS/thiết kế — ghi "[CẦN THIẾT KẾ DUYỆT]".
"""

from __future__ import annotations

import datetime
import json
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

# ── Nguồn dữ liệu ─────────────────────────────────────────────────────────────
from tools.qs_loader import QSLoadError, load_qs                      # QS/BOQ
from tools.schedule_loader import ScheduleLoadError, load_schedule  # Tiến độ CPM
from tools.cpm_calculator import CPMCalculator                      # ES/EF + đường găng
from tools.bbs_loader import BBSLoadError, load_bbs                 # BBS thép
from tools.cutting_stock_solver import (                            # Cắt thép
    CuttingStockSolver, CutDemand,
)
from tools.equipment_fleet_scheduler import (                       # Fleet: máy + ca + dầu
    EquipmentFleetScheduler, FleetTask,
)


# ═══════════════════════════════════════════════════════════════════════════
# HẰNG SỐ & QUY ƯỚC
# ═══════════════════════════════════════════════════════════════════════════

# Chỗ thiếu số liệu thiết kế — KHÔNG được đoán, đánh dấu chờ duyệt.
PENDING = "[CẦN THIẾT KẾ DUYỆT]"

# Dung sai đối chiếu số (±): cho phép lệch do làm tròn, không quá ngưỡng này.
TOL_ABS = 0.02          # 2% tương đối cho đại lượng khối lượng / chiều dài
TOL_QTY_ABS = 0.5       # 0.5 đơn vị cho đếm (phiến, cọc, gối...) → làm tròn

# 6 bất biến KCS (Hold Point / trình tự) — dùng chung cho auditor.
# Mỗi rule: (task-key regex cần có trước, task-key regex phải sau, mô tả vi phạm, mức)
DAG_RULES = [
    (r"cọc|khoan nhồi|khoan cọc",            r"bệ móng|bệ mố|bệ trụ|đào hố móng|đập đầu cọc",
     "Đổ bệ móng/mố trụ trước khi thi công xong cọc", "CHẶN"),
    (r"bệ móng|bệ mố|bệ trụ",                 r"thân mố|thân trụ|thân đặc|xà mũ",
     "Thi công thân/xà mũ trước khi xong bệ móng", "CHẶN"),
    (r"la[oô] (lắp|dầm)|lắp dầm|hạ dầm|vận chuyển.*dầm|lao dầm",
     r"đổ bê tông.*dầm|đúc dầm|dầm super|gia công.*dầm",
     "Lao/hạ dầm trước khi đúc dầm", "CHẶN"),
    (r"la[oô] (lắp|dầm)|lắp dầm|hạ dầm",      r"bản mặt cầu|mặt cầu|dầm ngang|liên tục nhiệt",
     "Đổ bản mặt cầu trước khi lao dầm", "CHẶN"),
    (r"bản mặt cầu|mặt cầu",                  r"khe co|thoát nước|lan can|chống thấm|bê tông nhựa|thảm",
     "Hoàn thiện (khe co/lan can/chống thấm) trước khi xong bản mặt cầu", "GIẢI TỎA"),
    (r"đổ bê tông|bê tông",                   r"cốt thép|ván khuôn|cốp pha|cốp pha",
     "Đổ bê tông trước khi lắp cốt thép/ván khuôn", "CHẶN"),
]

# Công tác CHE KHUẤT bắt buộc có Hold Point (điểm dừng nghiệm thu) trước khi che.
HOLD_POINT_RULES = [
    (r"cọc khoan nhồi|khoan cọc|cọc nhồi", "Cọc khoan nhồi — nghiệm thu lỗ khoan, thép cọc, mũi cọc trước khi đổ bê tông (bị che khuất)"),
    (r"cốt thép.*(bệ|móng|mố|trụ)|cốt thép", "Cốt thép bệ/móng/mố trụ — nghiệm thu trước khi đổ bê tông (bị che khuất)"),
    (r"karst|hang động|hang karst|khoan thăm dò|khoan dò", "Xử lý Karst — nghiệm thu lỗ khoan thăm dò + phương án xử lý trước khi khoan cọc"),
    (r"ống vách|hạ ống vách", "Hạ ống vách — nghiệm thu cao độ, thẳng đứng trước khi khoan tiếp"),
    (r"bản mặt cầu|cốt thép bản|cốt thép.*mặt cầu", "Cốt thép bản mặt cầu — nghiệm thu trước khi đổ bê tông (bị che khuất)"),
    (r"cáp|dul|dự ứng|căng kéo|luồn cáp", "Căng kéo cáp DƯL — nghiệm thu luồn cáp, cốt thép, ván khuôn trước khi đổ bê tông dầm"),
]

# Nhóm công tác HSE theo QCVN 18:2021/BXD — mỗi nhóm có checklist bắt buộc.
HSE_RULES = [
    (r"thân trụ|xà mũ|trụ cao|bản mặt cầu|dầm ngang|lan can|gờ|tháp|trên cao",
     "Làm việc trên cao >2m — giàn giáo, dây an toàn, lưới chắn, sàn công tác (QCVN 18:2021 §4)"),
    (r"giá lao|lao dầm|cẩu|cần trục|cầu trục|kích|nâng|hạ dầm|vận chuyển.*dầm|palang",
     "Thiết bị nâng / giá lao / cẩu — kiểm định, thử tải, giám sát (QCVN 18:2021 §6)"),
    (r"điện|máy phát|trạm|chiếu sáng|hàn|cáp điện",
     "Điện thi công — rò điện, cầu dao, đất an toàn (QCVN 18:2021 §8)"),
    (r"khoan|hố móng|đào|móng|thung lũng|suối|sông|mưa|lũ",
     "Môi trường / thiên tai — lũ quét, sạt lở, nước ngầm, giằng chống (QCVN 18:2021 §7, §10)"),
    (r"bê tông|đầm|trạm trộn|bồn|bơm|xi măng",
     "Đổ bê tông / trạm trộn — PPE, đầm rung, tràn bê tông (QCVN 18:2021 §5)"),
]

# Tiêu chuẩn được phép viện dẫn (đọc từ project.chi_dan_ky_thuat + mặc định).
DEFAULT_TCVN = [
    "TCVN 11823:2017", "TCVN 9395:2012", "TCVN 4453:1995", "TCVN 3105:2022",
    "TCVN 5574:2018", "TCVN 1651:2018", "TCVN 8826:2011", "TCVN 9168:2012",
    "QCVN 18:2021", "TCVN 8819:2011", "TCVN 8855:2011", "TCVN 12302:2018",
]


# ═══════════════════════════════════════════════════════════════════════════
# DATA CLASSES
# ═══════════════════════════════════════════════════════════════════════════

@dataclass
class Param:
    """Một tham số số trong BPTC — luôn kèm trace về nguồn."""
    key: str                      # ví dụ "piles.count", "concrete.C30.m3"
    label: str                    # nhãn hiển thị tiếng Việt
    value: Any                    # số / chuỗi / list
    unit: str = ""
    source: str = ""              # "QS AF.21111 row 15" / "CPM T2.3" / "Fleet M1" / "BBS Ø25"
    chapter: int = 0              # chương BPTC tham chiếu (0 = nền)


@dataclass
class MethodParams:
    """Toàn bộ tham số số của BPTC — Single Source of Truth."""
    project: Dict[str, Any] = field(default_factory=dict)
    params: List[Param] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def add(self, key, label, value, unit="", source="", chapter=0):
        self.params.append(Param(key, label, value, unit, source, chapter))

    def get(self, key) -> Optional[Param]:
        for p in self.params:
            if p.key == key:
                return p
        return None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "project": self.project,
            "params": [
                {"key": p.key, "label": p.label, "value": p.value,
                 "unit": p.unit, "source": p.source, "chapter": p.chapter}
                for p in self.params
            ],
            "warnings": self.warnings,
            "errors": self.errors,
        }


@dataclass
class CompileResult:
    """Kết quả biên dịch BPTC."""
    params: MethodParams = field(default_factory=MethodParams)
    md_text: str = ""
    docx_path: str = ""
    params_path: str = ""
    reconciliation_path: str = ""
    ok: bool = False
    message: str = ""


# ═══════════════════════════════════════════════════════════════════════════
# TIỆN ÍCH CHUẨN HOÁ VĂN BẢN & KEYWORD
# ═══════════════════════════════════════════════════════════════════════════

def _norm(text: Any) -> str:
    """Chuẩn hoá chuỗi: bỏ dấu, hạ chữ, gộp khoảng trắng — để match keyword."""
    import unicodedata
    s = unicodedata.normalize("NFKD", str(text or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.replace("đ", "d").replace("Đ", "d")
    return re.sub(r"\s+", " ", s).strip().lower()


def _matches(pattern: str, text: str) -> bool:
    return bool(re.search(pattern, text, flags=re.IGNORECASE))


def _find(items: List[Any], pattern: str, attr: str = "description",
          exclude: str = "") -> List[Any]:
    """Tìm item QS/BBS có thuộc tính `attr` match regex (trên text chuẩn hoá).

    `exclude`: regex loại trừ — item match cả pattern lẫn exclude thì BỎ QUA
    (tránh đếm nhầm: vd "Bê tông cọc khoan nhồi" không được tính vào "khoan cọc").
    """
    out = []
    for it in items:
        text = _norm(getattr(it, attr, "") or "")
        if not _matches(pattern, text):
            continue
        if exclude and _matches(exclude, text):
            continue
        out.append(it)
    return out


def _sum_qty(items: List[Any], qty_attr: str = "quantity") -> float:
    return round(sum(float(getattr(it, qty_attr, 0) or 0) for it in items), 6)


def _fmt_num(v: Any) -> str:
    """Định dạng số: int không có phần thập phân, float 3 chữ số, chuỗi nguyên vẹn."""
    if v is None:
        return PENDING
    if isinstance(v, bool):
        return "Có" if v else "Không"
    if isinstance(v, float):
        if abs(v - round(v)) < 1e-9:
            return f"{int(round(v)):,}".replace(",", " ") if abs(v) >= 10000 else str(int(round(v)))
        return f"{v:,.3f}".rstrip("0").rstrip(".")
    if isinstance(v, int):
        return f"{v:,}".replace(",", " ") if abs(v) >= 10000 else str(v)
    return str(v)


def _src_qs(item) -> str:
    return f"QS {item.code} dòng {item.row} — {item.description[:70]}"


def _src_cpm(task) -> str:
    return f"CPM {task.get('code') or task['id']} — {task['name'][:60]}"


def _src_fleet(task: FleetTask, alloc) -> str:
    return f"Fleet {task.task_id} / {alloc.machine_code} — {alloc.machine_name}"


def _src_bbs(demand: CutDemand) -> str:
    return f"BBS Ø{demand.diameter_mm} {demand.grade} — mark {demand.mark} ×{demand.quantity}"


# ═══════════════════════════════════════════════════════════════════════════
# PHÂN TÍCH QS → KHỐI LƯỢNG WBS
# ═══════════════════════════════════════════════════════════════════════════

def _analyze_qs(qs, params: MethodParams) -> None:
    """Gom QS thành các nhóm khối lượng WBS, ghi trace nguồn."""
    items = qs.items

    # ── PHẦN I: Cọc khoan nhồi ────────────────────────────────────────────────
    khoan = _find(items, r"khoan|tao lo", exclude=r"be tong|cot thep|ong vach|dap dau|thi nghiem|siey am")
    params.add("piles.drill_length_m", "Tổng chiều dài khoan cọc nhồi",
               _sum_qty(khoan), "m", "; ".join(_src_qs(i) for i in khoan), chapter=3)

    vach = _find(items, r"ong vach", exclude=r"be tong|cot thep|thi nghiem")
    params.add("piles.casing_length_m", "Chiều dài ống vách thép",
               _sum_qty(vach), "m", "; ".join(_src_qs(i) for i in vach), chapter=3)

    coc_bt = _find(items, r"be tong.*coc|coc.*be tong|be tong khoan nhoi",
                   exclude=r"cot thep|ong vach|dam ngang|dam chu|ban mat|dau ngang|mố|tru|bệ|hố")
    params.add("piles.concrete_m3", "Bê tông cọc khoan nhồi",
               _sum_qty(coc_bt), "m³", "; ".join(_src_qs(i) for i in coc_bt), chapter=3)

    coc_thep = _find(items, r"cot thep coc|thep coc khoan", exclude=r"dam|mo tru|ban mat|lien tuc")
    params.add("piles.rebar_t", "Cốt thép cọc khoan nhồi",
               _sum_qty(coc_thep), "tấn", "; ".join(_src_qs(i) for i in coc_thep), chapter=3)

    dap_dau = _find(items, r"dap dau coc|dap dau")
    params.add("piles.break_head_m3", "Đập đầu cọc",
               _sum_qty(dap_dau), "m³", "; ".join(_src_qs(i) for i in dap_dau), chapter=3)

    # Đếm số cọc — lấy từ mô tả (vd "26 cọc") nếu có, không thì chờ duyệt.
    so_coc = None
    for it in khoan:
        m = re.search(r"(\d+)\s*coc", _norm(it.description))
        if m:
            so_coc = int(m.group(1))
            break
    params.add("piles.count", "Số cọc khoan nhồi",
               so_coc if so_coc is not None else PENDING, "cọc",
               _src_qs(khoan[0]) if khoan else "", chapter=3)

    # ── PHẦN II: Kết cấu phần dưới ────────────────────────────────────────────
    dao = _find(items, r"dao ho mong|dao dat|dao ho", exclude=r"van chuyen|coc nhoi|cau|duong cong vu")
    params.add("sub.excavation_m3", "Đào hố móng",
               _sum_qty(dao), "m³", "; ".join(_src_qs(i) for i in dao), chapter=4)

    bt_lot = _find(items, r"be tong lot")
    params.add("sub.lean_concrete_m3", "Bê tông lót",
               _sum_qty(bt_lot), "m³", "; ".join(_src_qs(i) for i in bt_lot), chapter=4)

    be_mong = _find(items, r"be tong be|be mong|be mo|be tru", exclude=r"coc nhoi|tham")
    params.add("sub.footing_concrete_m3", "Bê tông bệ móng",
               _sum_qty(be_mong), "m³", "; ".join(_src_qs(i) for i in be_mong), chapter=4)

    than = _find(items, r"be tong than|than mo|than tru|tuong canh|tuong dinh|mo chan de|mo chu u",
                 exclude=r"cot thep|be tong coc|dam")
    params.add("sub.body_concrete_m3", "Bê tông thân mố/trụ",
               _sum_qty(than), "m³", "; ".join(_src_qs(i) for i in than), chapter=4)

    xa_mu = _find(items, r"xa mu")
    params.add("sub.coping_concrete_m3", "Bê tông xà mũ",
               _sum_qty(xa_mu), "m³", "; ".join(_src_qs(i) for i in xa_mu), chapter=4)

    thep_kcd = _find(items, r"cot thep mo|cot thep tru|cot thep mo va tru", exclude=r"dam|coc|ban mat|lien tuc")
    params.add("sub.rebar_t", "Cốt thép mố trụ",
               _sum_qty(thep_kcd), "tấn", "; ".join(_src_qs(i) for i in thep_kcd), chapter=4)

    # ── PHẦN III: Kết cấu phần trên ──────────────────────────────────────────
    dam_bt = _find(items, r"be tong dam chu|super-?t.*be tong|be tong.*dam chu", exclude=r"dam ngang|cot thep")
    params.add("sup.girder_concrete_m3", "Bê tông dầm chủ",
               _sum_qty(dam_bt), "m³", "; ".join(_src_qs(i) for i in dam_bt), chapter=5)

    cap_dul = _find(items, r"cap du|cap dul|tao xoan|du ung", exclude=r"cot thep thuong")
    params.add("sup.tendon_t", "Cáp dự ứng lực",
               _sum_qty(cap_dul), "tấn", "; ".join(_src_qs(i) for i in cap_dul), chapter=5)

    thep_dam = _find(items, r"cot thep.*dam|cot thep thuong", exclude=r"coc|mo |tru|ban mat|lien tuc")
    params.add("sup.girder_rebar_t", "Cốt thép dầm",
               _sum_qty(thep_dam), "tấn", "; ".join(_src_qs(i) for i in thep_dam), chapter=5)

    lao = _find(items, r"lao lap|van chuyen va lao|lao dam")
    params.add("sup.girder_launch_count", "Số phiến dầm lao lắp",
               _sum_qty(lao), "phiến", "; ".join(_src_qs(i) for i in lao), chapter=6)

    goi = _find(items, r"goi chau|goi cau|goi chậu", exclude=r"be tong|cot thep")
    params.add("sup.bearing_count", "Số gối cầu",
               _sum_qty(goi), "cái", "; ".join(_src_qs(i) for i in goi), chapter=6)

    van_khuon_tam = _find(items, r"van khuon duc san|tam van khuon|van khuon tam", exclude=r"be tong")
    params.add("sup.formwork_panels", "Tấm ván khuôn đúc sẵn",
               _sum_qty(van_khuon_tam), "tấm", "; ".join(_src_qs(i) for i in van_khuon_tam), chapter=7)

    dam_ngang = _find(items, r"be tong dam ngang|dam ngang", exclude=r"cot thep|ban mat|mo |tru ")
    params.add("sup.crossbeam_concrete_m3", "Bê tông dầm ngang",
               _sum_qty(dam_ngang), "m³", "; ".join(_src_qs(i) for i in dam_ngang), chapter=7)

    ban_mat = _find(items, r"ban mat cau", exclude=r"cot thep|van khuon|tam ")
    params.add("sup.deck_concrete_m3", "Bê tông bản mặt cầu",
               _sum_qty(ban_mat), "m³", "; ".join(_src_qs(i) for i in ban_mat), chapter=7)

    lien_tuc = _find(items, r"lien tuc nhiet|moi noi lien tuc")
    params.add("sup.joint_concrete_m3", "Bê tông liên tục nhiệt",
               _sum_qty(lien_tuc), "m³", "; ".join(_src_qs(i) for i in lien_tuc), chapter=7)

    thep_bmc = _find(items, r"cot thep ban mat|cot thep.*lien tuc", exclude=r"coc|mo |tru |dam chu")
    params.add("sup.deck_rebar_t", "Cốt thép bản mặt cầu",
               _sum_qty(thep_bmc), "tấn", "; ".join(_src_qs(i) for i in thep_bmc), chapter=7)

    khe_co = _find(items, r"khe co")
    params.add("sup.expansion_joint_m", "Khe co giãn",
               _sum_qty(khe_co), "m", "; ".join(_src_qs(i) for i in khe_co), chapter=7)

    # ── PHẦN IV: Hoàn thiện ──────────────────────────────────────────────────
    ban_qua_do = _find(items, r"ban qua do")
    params.add("fin.approach_slab_m3", "Bê tông bản quá độ",
               _sum_qty(ban_qua_do), "m³", "; ".join(_src_qs(i) for i in ban_qua_do), chapter=7)

    dap_sau_mo = _find(items, r"dap.*sau mo|dap hat chon loc|dap.*dam chat k", exclude=r"cat san lap")
    params.add("fin.backfill_m3", "Đắp sau mố",
               _sum_qty(dap_sau_mo), "m³", "; ".join(_src_qs(i) for i in dap_sau_mo), chapter=7)

    lan_can = _find(items, r"lan can|go lan can|tay vin", exclude=r"coc tru")
    params.add("fin.parapet_m", "Lan can / gờ",
               _sum_qty(lan_can), "m", "; ".join(_src_qs(i) for i in lan_can), chapter=7)

    thoat_nuoc = _find(items, r"thoat nuoc|ho ga|ong pvc")
    params.add("fin.drainage_set", "Bộ thoát nước",
               _sum_qty(thoat_nuoc), "bộ", "; ".join(_src_qs(i) for i in thoat_nuoc), chapter=7)

    chong_tham = _find(items, r"chong tham|tham be tong nhua|mang chong tham|be tong nhua",
                       exclude=r"dam|cau ")
    params.add("fin.waterproofing_m2", "Chống thấm + thảm BTN",
               _sum_qty(chong_tham), "m²", "; ".join(_src_qs(i) for i in chong_tham), chapter=7)

    taluy = _find(items, r"taluy|gia co|da hoc xay")
    params.add("fin.slope_protection_m3", "Gia cố taluy",
               _sum_qty(taluy), "m³", "; ".join(_src_qs(i) for i in taluy), chapter=7)

    # Thử tải
    thu_tai = _find(items, r"thu tai|kiem dinh")
    params.add("fin.load_test", "Thử tải cầu",
               _sum_qty(thu_tai), "gói", "; ".join(_src_qs(i) for i in thu_tai), chapter=8)


# ═══════════════════════════════════════════════════════════════════════════
# PHÂN TÍCH CPM → TRÌNH TỰ + ĐƯỜNG GĂNG
# ═══════════════════════════════════════════════════════════════════════════

def _analyze_cpm(schedule_path: str, project: Dict[str, Any], params: MethodParams) -> Dict[str, Any]:
    """
    Chạy CPMCalculator → ES/EF/LS/LF, đường găng, tổng thời gian.
    Trả dict {result, tasks_by_code, ordered, critical} để dùng tiếp.
    """
    loaded = load_schedule(schedule_path)
    if loaded.errors:
        params.errors.append(
            f"Tiến độ {loaded.source} có {len(loaded.errors)} dòng sai — sửa file trước khi biên dịch BPTC")
        return {}

    start_date = project.get("start_date") or loaded.project_start or ""
    result = CPMCalculator().calculate(
        loaded.tasks,
        start_date_str=start_date,
        non_working_weekdays=project.get("non_working_weekdays", ()),
        holidays=project.get("holidays", ()),
    )
    if result.status == "ERROR":
        params.errors.append("CPM lỗi logic: " + "; ".join(result.warnings[:5]))
        return {}

    by_id = {t.task_id: t for t in result.tasks}
    # Ưu tiên 'code' (1.1, 2.3) nếu file có; không thì dùng id (2,3,7...)
    def _code(t):
        return str(getattr(t, "code", "") or t.task_id)

    code_map = {}   # task_id -> display code
    for raw in loaded.tasks:
        code_map[raw["id"]] = str(raw.get("code") or raw["id"])

    ordered = sorted(result.tasks, key=lambda t: (t.es, t.task_id))
    critical = [t for t in result.tasks if t.is_critical]

    params.add("cpm.total_days", "Tổng thời gian thi công",
               result.total_duration_days, "ngày làm việc",
               f"CPM {schedule_path}", chapter=1)
    params.add("cpm.start", "Ngày khởi công",
               result.project_start or start_date or PENDING, "", f"CPM {schedule_path}", chapter=1)
    params.add("cpm.finish", "Ngày hoàn thành",
               result.project_finish or PENDING, "", f"CPM {schedule_path}", chapter=1)
    params.add("cpm.critical_path", "Đường găng (theo thứ tự)",
               " → ".join(code_map.get(t.task_id, _code(t)) for t in ordered if t.is_critical),
               "", f"CPM {schedule_path}", chapter=1)
    params.add("cpm.task_count", "Số công việc trong tiến độ",
               len(result.tasks), "việc", f"CPM {schedule_path}", chapter=1)

    # Bảng trình tự thi công — mỗi công việc: mã, tên, thời gian, ES, EF, găng.
    rows = []
    for t in ordered:
        rows.append({
            "code": code_map.get(t.task_id, _code(t)),
            "name": t.name,
            "duration_days": t.duration_days,
            "es": t.es, "ef": t.ef,
            "start_date": t.start_date, "finish_date": t.finish_date,
            "critical": bool(t.is_critical),
            "source": f"CPM {_code(t)}",
        })
    params.add("cpm.sequence", "Trình tự thi công (theo ES)",
               rows, "", f"CPM {schedule_path}", chapter=2)

    return {"result": result, "ordered": ordered, "critical": critical,
            "code_map": code_map, "by_id": by_id,
            "dates": {code_map.get(t.task_id, _code(t)): (t.start_date, t.finish_date)
                      for t in result.tasks}}


# ═══════════════════════════════════════════════════════════════════════════
# PHÂN TÍCH FLEET → MÁY + CA + DẦU
# ═══════════════════════════════════════════════════════════════════════════

def _analyze_fleet(schedule_path: str, project: Dict[str, Any],
                   params: MethodParams, fleet_tasks_cfg: Optional[List[Dict[str, Any]]] = None,
                   cpm_dates: Optional[Dict[str, Tuple[str, str]]] = None) -> None:
    """
    Dùng EquipmentFleetScheduler tính ca máy / máy / dầu / nhân lực.

    fleet_tasks_cfg: list {"code"|"task_id", "task_norm_key", "quantity"?, "labor_per_shift"?}
        — ánh xạ công tác CPM → task_norm_key trong data/equipment_norms_vincons.json.
    project.equipment: list {"name", "count", "spec", "code"?}
        — máy đặc thù cầu (máy khoan cọc, giá lao, cẩu...) KHÔNG có trong norms.
        Đây là nguồn duy nhất cho máy cầu; key auto-mỗi tên công tác bằng keyword.

    Nguyên tắc: ca máy/dầu chỉ tính cho công tác ĐÃ ánh xạ được norm (fleet-mapped).
    Công tác cầu đặc thù chưa có norm → máy lấy từ project.equipment, cảnh báo rõ
    "chưa có norm máy" để bổ sung file norms theo dự án.
    """
    loaded = load_schedule(schedule_path)
    if not loaded.tasks:
        params.warnings.append("Fleet: không có công việc CPM để phân bổ máy")
        return

    scheduler = EquipmentFleetScheduler()
    norms_tasks = scheduler.norms_data.get("tasks", {})

    # Tra cứu cfg theo mã công tác (code ưu tiên, fallback sang id).
    cfg_by_code: Dict[str, Dict[str, Any]] = {}
    if fleet_tasks_cfg:
        for cfg in fleet_tasks_cfg:
            key = str(cfg.get("code") or cfg.get("task_id") or "")
            if key:
                cfg_by_code[key] = cfg

    # Ngày bắt đầu/kết thúc mỗi công tác = từ CPM (file_start/file_finish),
    # không phải cùng 1 ngày khởi công.
    def _task_dates(t) -> Tuple[datetime.date, datetime.date]:
        dur = max(1, int(round(float(t["duration"]))))
        raw_start = ""
        if cpm_dates and str(t.get("code") or t["id"]) in cpm_dates:
            raw_start = cpm_dates[str(t.get("code") or t["id"])][0]
        raw_start = raw_start or t.get("file_start", "")
        try:
            sd = datetime.date.fromisoformat(raw_start) if raw_start else loaded.project_start
            if isinstance(sd, str):
                sd = datetime.date.fromisoformat(sd) if sd else datetime.date.today()
        except (ValueError, TypeError):
            sd = datetime.date.today()
        return sd, sd + datetime.timedelta(days=dur - 1)

    tasks: List[FleetTask] = []
    unmapped: List[str] = []
    for t in loaded.tasks:
        code = str(t.get("code") or t["id"])
        cfg = cfg_by_code.get(code) or cfg_by_code.get(str(t["id"]))
        norm_key = (cfg or {}).get("task_norm_key")
        if norm_key and norm_key not in norms_tasks:
            params.warnings.append(
                f"Fleet: công tác {code} chỉ norm '{norm_key}' không có trong data — bỏ qua ánh xạ")
            norm_key = None

        if norm_key is None:
            unmapped.append(f"{code} {t['name'][:50]}")
            continue  # KHÔNG đoán norm cho công tác cầu đặc thù

        qty = (cfg or {}).get("quantity")
        if qty is None:
            params.warnings.append(
                f"Fleet: công tác {code} có norm '{norm_key}' nhưng thiếu quantity — lấy thời gian làm KL quy ước")
            qty = float(t["duration"])

        sd, ed = _task_dates(t)
        ft = FleetTask(
            task_id=code,
            code=code,
            name=t["name"],
            unit=str(norms_tasks[norm_key].get("unit", "")) or "ca",
            quantity=float(qty),
            start_date=sd,
            end_date=ed,
            shifts_per_day=int((cfg or {}).get("shifts_per_day", project.get("shifts_per_day", 2))),
            task_norm_key=norm_key,
            labor_per_shift=float((cfg or {}).get("labor_per_shift", norms_tasks[norm_key].get("labor_per_shift", 4.0))),
        )
        tasks.append(ft)

    if not tasks:
        params.warnings.append("Fleet: không có công tác nào ánh xạ được norm — ca máy/dầu dùng PENDING")
        params.add("fleet.total_shifts", "Tổng ca máy (công tác đã ánh xạ norm)", PENDING, "ca",
                   "", chapter=2)
        params.add("fleet.total_fuel_l", "Tổng dầu diesel (công tác đã ánh xạ norm)", PENDING, "lít",
                   "", chapter=2)
        params.add("fleet.peak_labor", "Quân số đỉnh (người/ngày, Fleet norm)", PENDING, "người",
                   "", chapter=2)
        params.add("fleet.machines", "Danh mục máy từ Fleet norms", [], "", "", chapter=2)
        params.add("fleet.unmapped_tasks", "Công tác chưa có norm máy (cần khai báo equipment)",
                   unmapped, "", "—", chapter=2)
    else:
        for t in tasks:
            scheduler.compute_task_allocations(t)
        dmin = min(t.start_date for t in tasks)
        dmax = max(t.end_date for t in tasks)
        matrix = scheduler.build_daily_matrix(tasks, dmin, dmax)

        machine_summary: Dict[str, Dict[str, Any]] = {}
        for t in tasks:
            for m in t.machine_allocations:
                rec = machine_summary.setdefault(m.machine_code, {
                    "name": m.machine_name, "total_shifts": 0.0, "total_fuel_l": 0.0, "tasks": []})
                rec["total_shifts"] += m.total_shifts_required
                rec["total_fuel_l"] += m.total_fuel_liters
                rec["tasks"].append(t.task_id)

        total_shifts = round(matrix.total_shifts, 1)
        total_fuel = round(matrix.total_fuel_liters, 1)
        peak_labor = max(matrix.daily_labor.values()) if matrix.daily_labor else 0

        scope = f"Fleet norms — {len(tasks)}/{len(loaded.tasks)} công tác đã ánh xạ"
        params.add("fleet.total_shifts", "Tổng ca máy (công tác đã ánh xạ norm)", total_shifts, "ca",
                   scope, chapter=2)
        params.add("fleet.total_fuel_l", "Tổng dầu diesel (công tác đã ánh xạ norm)", total_fuel, "lít",
                   scope, chapter=2)
        params.add("fleet.peak_labor", "Quân số đỉnh (người/ngày, Fleet norm)", round(peak_labor, 1), "người",
                   scope, chapter=2)
        params.add("fleet.machines", "Danh mục máy từ Fleet norms",
                   [
                       {"code": code, "name": v["name"],
                        "total_shifts": round(v["total_shifts"], 1),
                        "total_fuel_l": round(v["total_fuel_l"], 1),
                        "peak_machines": matrix.peak_machines.get(code, 0),
                        "tasks": sorted(set(v["tasks"]))}
                       for code, v in sorted(machine_summary.items())
                   ],
                   "", scope, chapter=2)
        params.add("fleet.unmapped_tasks", "Công tác chưa có norm máy (cần khai báo equipment)",
                   unmapped, "", "—", chapter=2)

    # Máy đặc thù cầu — CHỈ lấy từ project.equipment (hồ sơ thiết bị được duyệt).
    equip_cfg = [e for e in (project.get("equipment") or []) if isinstance(e, dict)]
    params.add("fleet.equipment", "Máy đặc thù cầu (theo hồ sơ thiết bị được duyệt)",
               [{"name": str(e.get("name") or PENDING), "count": e.get("count", PENDING),
                 "spec": str(e.get("spec", "")), "source": "project.equipment"}
                for e in equip_cfg],
               "", "project.equipment", chapter=2)
    if equip_cfg:
        # Dầu cho máy khai báo: code M1..Mn tra fuel norms; không có code → cảnh báo, không đoán.
        fuel_total = 0.0
        fuel_missing = []
        for e in equip_cfg:
            code = str(e.get("code") or "")
            rate = scheduler.fuel_data.get("equipment_rates", {}).get(code, {})
            if "fuel_rate_l_per_shift" in rate:
                fuel_total += float(rate["fuel_rate_l_per_shift"]) * float(e.get("shifts", 1)) * float(e.get("count") or 0)
            else:
                fuel_missing.append(str(e.get("name", code)))
        params.add("fleet.equipment_fuel_l", "Dầu máy khai báo (theo shifts đã chốt)",
                   round(fuel_total, 1) if not fuel_missing else PENDING, "lít",
                   "project.equipment + data/fuel_consumption_norms.json", chapter=2)
        if fuel_missing:
            params.warnings.append(
                "Fleet: máy chưa có mã định mức dầu: " + ", ".join(fuel_missing) +
                " — bổ sung code vào project.equipment")


# ═══════════════════════════════════════════════════════════════════════════
# PHÂN TÍCH BBS + CUTTING STOCK → KẾ HOẠCH THÉP
# ═══════════════════════════════════════════════════════════════════════════

def _analyze_bbs(bbs_path: str, params: MethodParams) -> None:
    """Đọc BBS → CuttingStockSolver → số cây 11.7m, hao hụt, phân kỳ cấp thép."""
    try:
        bbs = load_bbs(bbs_path)
    except BBSLoadError as e:
        params.errors.append(str(e))
        return
    if bbs.errors:
        # Dòng lỗi KHÔNG làm hỏng cả kế hoạch: tính từ dòng hợp lệ,
        # giữ lỗi để Human Gate xử lý (exit code ≠ 0, xem errors).
        params.errors.extend(f"BBS {bbs_path}: {e}" for e in bbs.errors)
        # Gợi ý chẩn đoán: Ký hiệu trông như Ø chuẩn nhưng cột Ø thì không → nghi đảo cột.
        suspect = sum(1 for e in bbs.errors if "không phải thép thanh tiêu chuẩn" in e
                      or "quá ngắn" in e)
        if suspect:
            params.warnings.append(
                f"BBS: {suspect} dòng có thể bị đảo cột Ký hiệu/Đường kính hoặc sai đơn vị "
                f"chiều dài — kiểm tra file gốc, KHÔNG tự đoán sửa")

    demands: List[CutDemand] = bbs.demands
    if not demands:
        if not bbs.errors:
            params.warnings.append("BBS không có thanh thép nào — bỏ qua kế hoạch cắt thép")
        return

    solver = CuttingStockSolver()
    sol = solver.solve(demands, allow_splicing=True, split_long_bars=True)

    params.add("steel.total_pieces", "Tổng số thanh thép trong BBS",
               bbs.total_pieces, "thanh", f"BBS {bbs_path}", chapter=3)
    params.add("steel.bar_length_m", "Chiều dài cây thép nguyên",
               sol.bar_length_mm / 1000.0, "m", "Chuẩn thanh 11.7m", chapter=3)
    params.add("steel.total_bars", "Số cây thép cần mua",
               sol.total_bars_needed, "cây", f"CuttingStockSolver {bbs_path}", chapter=3)
    params.add("steel.waste_pct", "Tỷ lệ hao hụt",
               round(sol.waste_ratio_pct, 2), "%", f"CuttingStockSolver {bbs_path}", chapter=3)
    params.add("steel.total_weight_t", "Tổng khối lượng thép mua",
               round(sol.total_weight_kg / 1000.0, 3), "tấn", f"CuttingStockSolver {bbs_path}", chapter=3)
    params.add("steel.splices", "Số mối nối (nếu cho phép nối)",
               sol.total_splices, "mối", f"CuttingStockSolver {bbs_path}", chapter=3)

    # Phân kỳ cấp thép theo đường kính — dùng groups.
    groups = []
    for g in sol.groups:
        groups.append({
            "diameter_mm": g.diameter_mm, "grade": g.grade,
            "pieces": g.total_pieces, "bars": g.total_bars_needed,
            "waste_pct": round(g.waste_ratio_pct, 2),
            "net_weight_t": round(g.net_weight_kg / 1000.0, 3),
            "status": g.status,
        })
    params.add("steel.groups", "Phân kỳ cấp thép theo Ø",
               groups, "", f"CuttingStockSolver {bbs_path}", chapter=3)

    if sol.waste_ratio_pct > 1.5:
        params.warnings.append(
            f"Hao hụt thép {sol.waste_ratio_pct:.2f}% > 1.5% — cần rà soát BBS/phương án cắt")


# ═══════════════════════════════════════════════════════════════════════════
# SINH method_params.json + RECONCILIATION
# ═══════════════════════════════════════════════════════════════════════════

def build_method_params(qs_path: str, schedule_path: str, project: Dict[str, Any],
                        bbs_path: Optional[str] = None,
                        fleet_tasks_cfg: Optional[List[Dict[str, Any]]] = None) -> MethodParams:
    """Gom toàn bộ số liệu từ QS/CPM/Fleet/BBS thành MethodParams."""
    params = MethodParams()
    params.project = dict(project)

    # ── QS ──
    try:
        qs = load_qs(qs_path, sheet=project.get("qs_sheet"))
    except QSLoadError as e:
        params.errors.append(str(e))
        qs = None
    if qs is not None:
        params.add("project.name", "Tên công trình",
                   project.get("name") or qs.hang_muc or PENDING, "",
                   f"QS {qs_path}", chapter=1)
        params.add("qs.total_items", "Số dòng công tác trong QS",
                   len(qs.items), "dòng", f"QS {qs_path}", chapter=1)
        params.add("qs.sections", "Các phần trong QS",
                   sorted({(i.section or "(không section)") for i in qs.items}),
                   "", f"QS {qs_path}", chapter=1)
        if qs.G_XD:
            params.add("qs.total_G_XD", "Giá trị dự toán sau thuế",
                       qs.G_XD, "VNĐ", f"QS {qs_path}", chapter=1)
        _analyze_qs(qs, params)

    # ── CPM ──
    cpm_ctx = _analyze_cpm(schedule_path, project, params)

    # ── Fleet (dùng ngày ES/EF từ CPM) ──
    _analyze_fleet(schedule_path, project, params, fleet_tasks_cfg=fleet_tasks_cfg,
                   cpm_dates=cpm_ctx.get("dates") if cpm_ctx else None)

    # ── BBS + cắt thép ──
    if bbs_path:
        _analyze_bbs(bbs_path, params)
    else:
        params.warnings.append("Chưa có BBS (--bbs) — bỏ qua kế hoạch cắt thép trong BPTC")

    return params


def write_method_params(params: MethodParams, out_dir: str) -> str:
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "method_params.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(params.to_dict(), f, ensure_ascii=False, indent=2)
    return path


def write_reconciliation(params: MethodParams, out_dir: str) -> str:
    """Bảng đối soát: mỗi số → nguồn (dòng QS/CPM/Fleet/BBS)."""
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, "DOI_SOAT_SO_LIEU_BPTC.md")
    lines = [
        "# ĐỐI SOÁT SỐ LIỆU BPTC",
        "",
        "Mỗi con số trong Thuyết minh BPTC phải trỏ về nguồn (dòng QS / công tác CPM / Fleet / BBS).",
        "Số không trace được nguồn = số chết → sửa BPTC.",
        "",
        "| Tham số | Giá trị | Đơn vị | Nguồn | Chương |",
        "|---|---|---|---|---|",
    ]
    for p in params.params:
        v = p.value
        if isinstance(v, (list, dict)):
            v = f"({len(v)} mục)"
        lines.append(
            f"| `{p.key}` | {_fmt_num(v)} | {p.unit or '—'} | {p.source or '—'} | {p.chapter or '—'} |")
    if params.warnings:
        lines += ["", "## Cảnh báo"] + [f"- {w}" for w in params.warnings]
    if params.errors:
        lines += ["", "## Lỗi"] + [f"- {e}" for e in params.errors]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return path


# ═══════════════════════════════════════════════════════════════════════════
# RENDER BPTC — 8 CHƯƠNG THEO MẪU Km19 (template câu chuẩn + số sống)
# ═══════════════════════════════════════════════════════════════════════════

def _render_chapters(params: MethodParams) -> List[Tuple[str, str]]:
    """Trả list (tiêu đề, nội dung) 8 chương. Số lấy từ params — không LLM."""
    P = lambda k: params.get(k)
    V = lambda k: (params.get(k).value if params.get(k) else PENDING)

    chapters = []

    # ── CHƯƠNG 1 ─────────────────────────────────────────────────────────────
    seq = P("cpm.sequence")
    seq_rows = ""
    if seq and isinstance(seq.value, list):
        seq_rows = "\n".join(
            f"| {r['code']} | {r['name']} | {r['duration_days']} | {r['start_date'] or r['es']} | {r['finish_date'] or r['ef']} | {'Găng' if r['critical'] else ''} |"
            for r in seq.value)
    chapters.append((
        "CHƯƠNG 1: GIỚI THIỆU CHUNG & ĐIỀU KIỆN HIỆN TRƯỜNG",
        f"""- Công trình: **{V('project.name')}**
- Số dòng công tác QS: **{V('qs.total_items')}**
- Giá trị dự toán (sau thuế): **{_fmt_num(V('qs.total_G_XD'))} VNĐ**
- Tổng thời gian thi công: **{V('cpm.total_days')} ngày làm việc**
- Khởi công: **{V('cpm.start')}** — Hoàn thành: **{V('cpm.finish')}**
- Đường găng: **{V('cpm.critical_path')}**

**Trình tự thi công (theo ES):**

| Mã | Công việc | Thời gian (ngày) | Bắt đầu | Kết thúc | Ghi chú |
|---|---|---|---|---|---|
{seq_rows}
"""))

    # ── CHƯƠNG 2 ─────────────────────────────────────────────────────────────
    mach = P("fleet.machines")
    mach_rows = ""
    if mach and isinstance(mach.value, list):
        mach_rows = "\n".join(
            f"| {m['code']} | {m['name']} | {m['total_shifts']} | {m['total_fuel_l']} | {m['peak_machines']} |"
            for m in mach.value)
    chapters.append((
        "CHƯƠNG 2: ĐƯỜNG CÔNG VỤ, BÃI ĐÚC & TỔ CHỨC THI CÔNG",
        f"""- Tổng ca máy: **{V('fleet.total_shifts')} ca**
- Tổng dầu diesel: **{_fmt_num(V('fleet.total_fuel_l'))} lít**
- Quân số đỉnh: **{V('fleet.peak_labor')} người/ngày**

**Danh mục máy thi công chính:**

| Mã | Máy | Tổng ca | Dầu (l) | Máy đỉnh |
|---|---|---|---|---|
{mach_rows}
"""))

    # ── CHƯƠNG 3 ─────────────────────────────────────────────────────────────
    karst = params.project.get("karst", False)
    karst_block = (
        "- **Xử lý hang Karst:** khoan thăm dò sâu ≥5.0m vào đá liền khối dưới mũi cọc; "
        "bơm chèn vữa + đá hộc bịt hang rỗng; hạ ống vách ngàm vào đá gốc; "
        "giữ Bentonite cao hơn mực nước ngầm ≥1.5m; đổ bê tông bằng rút ống Tremie."
        if karst else "- Địa chất không ghi Karst — bỏ qua chuyên mục xử lý hang.")
    steel_groups = P("steel.groups")
    steel_rows = ""
    if steel_groups and isinstance(steel_groups.value, list):
        steel_rows = "\n".join(
            f"| Ø{g['diameter_mm']} | {g['grade']} | {g['pieces']} | {g['bars']} | {g['waste_pct']}% | {g['net_weight_t']} |"
            for g in steel_groups.value)
    chapters.append((
        "CHƯƠNG 3: THI CÔNG CỌC KHOAN NHỒI & XỬ LÝ KARST",
        f"""- Số cọc khoan nhồi: **{V('piles.count')} cọc**
- Tổng chiều dài khoan: **{V('piles.drill_length_m')} m**
- Ống vách thép: **{V('piles.casing_length_m')} m**
- Bê tông cọc: **{V('piles.concrete_m3')} m³**
- Cốt thép cọc: **{V('piles.rebar_t')} tấn**
- Đập đầu cọc: **{V('piles.break_head_m3')} m³**

{karst_block}

**Kế hoạch cắt thép (BBS):**
- Tổng thanh: **{V('steel.total_pieces')}**
- Cây 11.7m cần mua: **{V('steel.total_bars')} cây** (hao hụt **{V('steel.waste_pct')}%**)

| Ø | Mác | Số thanh | Cây cần | Hao hụt | Thép ròng (t) |
|---|---|---|---|---|---|
{steel_rows}
"""))

    # ── CHƯƠNG 4 ─────────────────────────────────────────────────────────────
    chapters.append((
        "CHƯƠNG 4: THI CÔNG KẾT CẤU PHẦN DƯỚI (MỐ & TRỤ)",
        f"""- Đào hố móng: **{V('sub.excavation_m3')} m³**
- Bê tông lót: **{V('sub.lean_concrete_m3')} m³**
- Bê tông bệ móng: **{V('sub.footing_concrete_m3')} m³**
- Bê tông thân mố/trụ: **{V('sub.body_concrete_m3')} m³**
- Bê tông xà mũ: **{V('sub.coping_concrete_m3')} m³**
- Cốt thép mố trụ: **{V('sub.rebar_t')} tấn**
"""))

    # ── CHƯƠNG 5 ─────────────────────────────────────────────────────────────
    chapters.append((
        "CHƯƠNG 5: ĐÚC DẦM & CĂNG KÉO CÁP DỰ ỨNG LỰC",
        f"""- Bê tông dầm chủ: **{V('sup.girder_concrete_m3')} m³**
- Cáp dự ứng lực: **{V('sup.tendon_t')} tấn**
- Cốt thép dầm: **{V('sup.girder_rebar_t')} tấn**
- Số phiến dầm lao lắp: **{V('sup.girder_launch_count')} phiến**
"""))

    # ── CHƯƠNG 6 ─────────────────────────────────────────────────────────────
    chapters.append((
        "CHƯƠNG 6: VẬN CHUYỂN & LAO LẮP DẦM",
        f"""- Số phiến lao lắp: **{V('sup.girder_launch_count')} phiến**
- Số gối cầu: **{V('sup.bearing_count')} cái**
- Thiết bị lao dầm: theo Fleet ở Chương 2 (xem danh mục máy).
"""))

    # ── CHƯƠNG 7 ─────────────────────────────────────────────────────────────
    chapters.append((
        "CHƯƠNG 7: BẢN MẶT CẦU, LIÊN TỤC NHIỆT & HOÀN THIỆN",
        f"""- Tấm ván khuôn đúc sẵn: **{V('sup.formwork_panels')} tấm**
- Bê tông dầm ngang: **{V('sup.crossbeam_concrete_m3')} m³**
- Bê tông bản mặt cầu: **{V('sup.deck_concrete_m3')} m³**
- Bê tông liên tục nhiệt: **{V('sup.joint_concrete_m3')} m³**
- Cốt thép bản mặt cầu: **{V('sup.deck_rebar_t')} tấn**
- Khe co giãn: **{V('sup.expansion_joint_m')} m**
- Bê tông bản quá độ: **{V('fin.approach_slab_m3')} m³**
- Đắp sau mố: **{V('fin.backfill_m3')} m³**
- Lan can / gờ: **{V('fin.parapet_m')} m**
- Bộ thoát nước: **{V('fin.drainage_set')} bộ**
- Chống thấm + thảm BTN: **{V('fin.waterproofing_m2')} m²**
- Gia cố taluy: **{V('fin.slope_protection_m3')} m³**
"""))

    # ── CHƯƠNG 8 ─────────────────────────────────────────────────────────────
    hse_items = _hse_checklist(params)
    hse_md = "\n".join(f"- {i}" for i in hse_items) if hse_items else "- (chưa có công tác kích hoạt HSE)"
    chapters.append((
        "CHƯƠNG 8: KẾ HOẠCH QUẢN LÝ AN TOÀN LAO ĐỘNG (HSE) & VSMT",
        f"""Theo QCVN 18:2021/BXD — checklist bắt buộc theo công tác thực tế trong tiến độ:

{hse_md}

- Thử tải / kiểm định trước khai thác: **{V('fin.load_test')}**
"""))

    return chapters


def _hse_checklist(params: MethodParams) -> List[str]:
    """Lập checklist HSE theo công tác thực tế trong CPM (QCVN 18:2021)."""
    seq = params.get("cpm.sequence")
    items = set()
    if seq and isinstance(seq.value, list):
        names = " ".join(_norm(r["name"]) for r in seq.value)
        for pat, label in HSE_RULES:
            if _matches(pat, names):
                items.add(label)
    return sorted(items)


def render_markdown(params: MethodParams) -> str:
    """Sinh văn bản BPTC 8 chương (Markdown)."""
    title = params.project.get("name") or "CÔNG TRÌNH XÂY DỰNG"
    header = f"""# THUYẾT MINH BIỆN PHÁP TỔ CHỨC THI CÔNG
## {title}

**Nguồn số liệu:** QS/BOQ + Tiến độ CPM + Fleet + BBS — mọi con số trace trong `method_params.json`.
**Cơ chế:** deterministic-first — số do code tính, văn do câu chuẩn; số thiếu thiết kế ghi `{PENDING}`.

---

"""
    body = []
    for title_ch, content in _render_chapters(params):
        body.append(f"### {title_ch}\n\n{content}\n")
    footer = """
---

**Ghi chú:** BPTC này được biên dịch tự động từ `method_params.json`. Mọi thay đổi số liệu
phải sửa ở nguồn (QS/CPM/Fleet/BBS) rồi chạy lại compiler — KHÔNG sửa tay số trong văn bản.
"""
    return header + "\n".join(body) + footer


def render_docx(params: MethodParams, out_path: str) -> bool:
    """Xuất .docx nếu có python-docx. Trả True nếu thành công."""
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        return False
    doc = Document()
    doc.add_heading("THUYẾT MINH BIỆN PHÁP TỔ CHỨC THI CÔNG", 0)
    doc.add_heading(params.project.get("name") or "", 1)
    for ch_title, ch_content in _render_chapters(params):
        doc.add_heading(ch_title, 2)
        for line in ch_content.strip().split("\n"):
            line = line.strip()
            if not line or line.startswith("|---"):
                continue
            if line.startswith("|"):
                cells = [c.strip() for c in line.strip("|").split("|")]
                p = doc.add_paragraph()
                run = p.add_run(" | ".join(cells))
                run.font.size = Pt(9)
            elif line.startswith("- "):
                doc.add_paragraph(line[2:], style="List Bullet")
            elif line.startswith("**") and line.endswith("**"):
                doc.add_paragraph(line.strip("*"), style="Intense Quote")
            else:
                doc.add_paragraph(line)
    doc.save(out_path)
    return True


# ═══════════════════════════════════════════════════════════════════════════
# ORCHESTRATOR — compile()
# ═══════════════════════════════════════════════════════════════════════════

def compile_bptc(qs_path: str, schedule_path: str, project: Dict[str, Any],
                 out_dir: str, bbs_path: Optional[str] = None,
                 fleet_tasks_cfg: Optional[List[Dict[str, Any]]] = None,
                 write_docx: bool = True) -> CompileResult:
    """Chạy toàn bộ pipeline compile: QS/CPM/Fleet/BBS → params + .md (+.docx)."""
    res = CompileResult()
    params = build_method_params(qs_path, schedule_path, project,
                                 bbs_path=bbs_path, fleet_tasks_cfg=fleet_tasks_cfg)
    res.params = params

    os.makedirs(out_dir, exist_ok=True)
    res.params_path = write_method_params(params, out_dir)
    res.reconciliation_path = write_reconciliation(params, out_dir)

    md = render_markdown(params)
    res.md_text = md
    safe = re.sub(r"[^\w\-+.]+", "_", params.project.get("name", "Cong_Trinh"))[:60]
    md_path = os.path.join(out_dir, f"Thuyet_Minh_BPTC_{safe}.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md)

    if write_docx:
        docx_path = os.path.join(out_dir, f"Thuyet_Minh_BPTC_{safe}.docx")
        if render_docx(params, docx_path):
            res.docx_path = docx_path

    res.ok = not params.errors
    res.message = ("OK — BPTC đã biên dịch" if res.ok
                   else "Có lỗi — xem method_params.json / errors")
    return res


# ═══════════════════════════════════════════════════════════════════════════
# CLI — python -m tools.method_compiler
# ═══════════════════════════════════════════════════════════════════════════

def _load_project_arg(path_or_json: Optional[str]) -> Dict[str, Any]:
    if not path_or_json:
        return {}
    if os.path.exists(path_or_json):
        with open(path_or_json, "r", encoding="utf-8") as f:
            return json.load(f)
    try:
        return json.loads(path_or_json)
    except json.JSONDecodeError:
        return {"name": path_or_json}


def main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(
        description="Biên dịch Thuyết minh BPTC từ QS/CPM/Fleet/BBS (deterministic)")
    ap.add_argument("--qs", required=True, help="File QS/BOQ (.xlsx/.csv/.json)")
    ap.add_argument("--qs-sheet", default=None)
    ap.add_argument("--schedule", required=True, help="Tiến độ CPM (.xml/.xlsx/.csv/.json)")
    ap.add_argument("--bbs", default=None, help="BBS thép (.xlsx/.csv/.json)")
    ap.add_argument("--project", default=None,
                    help="JSON hoặc tên công trình: {name, loai, karst, start_date, chi_dan_ky_thuat, thiet_bi_san_co}")
    ap.add_argument("--fleet-tasks", default=None,
                    help="JSON ánh xạ công tác → task_norm_key: [{code|task_id, task_norm_key, quantity?}]")
    ap.add_argument("--out-dir", required=True, help="Thư mục xuất method_params.json + BPTC")
    ap.add_argument("--no-docx", action="store_true", help="Bỏ qua xuất .docx")
    args = ap.parse_args(argv)

    project = _load_project_arg(args.project)
    if args.qs_sheet:
        project["qs_sheet"] = args.qs_sheet
    fleet_cfg = None
    if args.fleet_tasks:
        fleet_cfg = _load_project_arg(args.fleet_tasks) or []
        if isinstance(fleet_cfg, dict):
            fleet_cfg = fleet_cfg.get("tasks", fleet_cfg.get("items", []))

    res = compile_bptc(
        qs_path=args.qs, schedule_path=args.schedule, project=project,
        out_dir=args.out_dir, bbs_path=args.bbs,
        fleet_tasks_cfg=fleet_cfg, write_docx=not args.no_docx,
    )
    print(f"[method_compiler] {res.message}")
    print(f"  - method_params : {res.params_path}")
    print(f"  - đối soát      : {res.reconciliation_path}")
    if res.docx_path:
        print(f"  - docx          : {res.docx_path}")
    if res.params.errors:
        for e in res.params.errors:
            print(f"  ✗ {e}")
    if res.params.warnings:
        for w in res.params.warnings:
            print(f"  ⚠ {w}")
    return 0 if res.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
