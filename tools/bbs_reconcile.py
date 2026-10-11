# -*- coding: utf-8 -*-
"""
Sửa BBS theo bản vẽ thẩm định, đánh dấu NGHI VẤN mọi chỗ sai khác để chủ dự án soát và ghim.

Quy tắc (chủ dự án 11/10/2026): có dấu thẩm định thì lấy theo thẩm định; sai khác thì đánh dấu nghi vấn.
Số liệu bản vẽ lấy qua tools/drawing_rebar_tables (đã tự kiểm tra phép tính và dòng tổng của bản vẽ).
Mỗi hạng mục BBS (ánh xạ trang trong file cấu hình) được xử lý một trong bốn cách:

  KHỚP                 tổng BBS = tổng bản vẽ (sai số nhỏ) → giữ dòng BBS.
  THAY THEO BẢN VẼ     bảng bản vẽ đọc được đầy đủ → dòng lấy theo bản vẽ.
  THAY MỘT PHẦN        bảng thiếu vài dòng (≤ 5%) → các dòng đúng + dòng NGHI VẤN ghi phần thiếu (không tính kg).
  GIỮ BBS – NGHI VẤN   không đọc được / chưa biết số cấu kiện → giữ dòng BBS, mỗi dòng ghi nghi vấn.

Dòng BBS sai rõ ràng (Ø không phải thép thanh, dài < 50 mm) bị loại và ghi lại trong danh sách nghi vấn.
Khối lượng đơn vị theo quy ước hồ sơ: bảng tools/cutting_stock_solver (chủ dự án chọn 11/10/2026).
"""

from __future__ import annotations

import json
import os
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from tools.cutting_stock_solver import REBAR_UNIT_WEIGHT_KG_M, rebar_unit_weight_kg_m
from tools.drawing_rebar_tables import (good_rows, numeric_issues, page_rebar_table, page_tables, partial_gap_kg,
                                        table_usable, zero_weight_rows_kg)

CAP_DUL = "Cáp DƯL 15.2mm"
MIN_LENGTH_M = 0.05
PER_DIA_TOL = (5.0, 0.005)          # so theo Ø: ≤ 5 kg hoặc ≤ 0,5%

KHOP, THAY, THAY_MOT_PHAN, GIU = "KHỚP", "THAY THEO BẢN VẼ", "THAY MỘT PHẦN", "GIỮ BBS – NGHI VẤN"


@dataclass
class BbsRow:
    hang_muc: str
    bo_phan: str
    mark: str
    d: Optional[float]
    mac: str = ""
    hinh: str = ""
    L_m: Optional[float] = None
    n_per: Optional[float] = None
    n_ck: Optional[float] = None
    unit: Optional[float] = None
    nhom: str = ""
    ghi_chu: str = ""
    nghi_van: List[str] = field(default_factory=list)
    placeholder: bool = False

    @property
    def kg(self) -> float:
        if self.placeholder or None in (self.L_m, self.n_per, self.n_ck, self.unit):
            return 0.0
        return self.L_m * self.n_per * self.n_ck * self.unit

    @property
    def is_cable(self) -> bool:
        return self.nhom == CAP_DUL

    @property
    def is_broken(self) -> bool:
        """Dòng sai rõ ràng: Ø không phải thép thanh tiêu chuẩn, hoặc dài < 50 mm (lệch cột/đơn vị)."""
        if self.is_cable or self.placeholder:
            return False
        return self.d is None or int(self.d) not in REBAR_UNIT_WEIGHT_KG_M or (self.L_m or 0) < MIN_LENGTH_M


@dataclass
class Section:
    title: str
    items: "OrderedDict[str, List[BbsRow]]"


@dataclass
class ItemReport:
    hang_muc: str
    trang: str
    hanh_dong: str
    bbs_cu_kg: float
    loi_cu_kg: float
    ban_ve_kg: Optional[float]
    bbs_moi_kg: float = 0.0
    ghi_chu: str = ""


@dataclass
class NghiVan:
    hang_muc: str
    ky_hieu: str
    noi_dung: str
    gia_tri_cu: str = ""
    gia_tri_moi: str = ""
    trang: str = ""
    row_ref: Optional[int] = None     # điền khi ghi file


def nhom_duong_kinh(d: float) -> str:
    return "D<=10mm" if d <= 10 else ("10<D<=18mm" if d <= 18 else "D>18mm")


def mac_thep(d: float, tron_tron: bool = False) -> str:
    """Theo quy ước của BBS cũ: Ø ≤ 10 thép tròn trơn CB240-T, còn lại CB400-V. Bản vẽ ghi 'R25', 'R8' (tròn trơn)
    thì CB240-T bất kể Ø."""
    return "CB240-T" if tron_tron or d <= 10 else "CB400-V"


def drawing_table(page: dict, pg: dict):
    """Bảng thép của trang; trang có nhiều bảng (vd. tr.112) thì chọn bảng chứa ký hiệu 'bang_co_ky_hieu'."""
    mark = pg.get("bang_co_ky_hieu")
    if not mark:
        return page_rebar_table(page)
    return next((t for t in page_tables(page) if t.stated_total_kg is not None and any(r.mark == mark for r in t.rows)),
                None)


def read_bbs(path: str, sheet: str = "THONG_KE_THEP_CHI_TIET") -> Tuple[Dict[str, str], List[Section]]:
    import openpyxl
    ws = openpyxl.load_workbook(path)[sheet]
    titles = {"B2": ws["B2"].value, "B3": ws["B3"].value, "B4": ws["B4"].value}
    sections: List[Section] = []
    for r in range(6, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.strip()[:2].isdigit():
            sections.append(Section(a.strip(), OrderedDict()))
            continue
        if isinstance(a, str) and a.strip().upper().startswith("TỔNG CỘNG"):
            break
        if not isinstance(a, int) or not sections:
            continue
        v = lambda c: ws.cell(r, c).value
        num = lambda c: float(v(c)) if isinstance(v(c), (int, float)) else None
        row = BbsRow(hang_muc=str(v(2) or ""), bo_phan=str(v(3) or ""), mark=str(v(4) or ""), d=num(5),
                     mac=str(v(6) or ""), hinh=str(v(7) or ""), L_m=num(8), n_per=num(9), n_ck=num(10),
                     unit=num(13), nhom=str(v(16) or ""), ghi_chu=str(v(17) or ""))
        sections[-1].items.setdefault(row.hang_muc, []).append(row)
    return titles, sections


def _conv_unit(row: BbsRow) -> Optional[float]:
    if row.is_cable or row.d is None or int(row.d) not in REBAR_UNIT_WEIGHT_KG_M:
        return row.unit
    return rebar_unit_weight_kg_m(int(row.d))


def _per_dia(rows: List[BbsRow], use_conv: bool = False) -> Dict[int, float]:
    out: Dict[int, float] = defaultdict(float)
    for r in rows:
        if not r.is_cable and not r.placeholder and r.d is not None:
            unit = _conv_unit(r) if use_conv else r.unit
            out[int(r.d)] += (r.L_m or 0) * (r.n_per or 0) * (r.n_ck or 0) * (unit or 0)
    return out


def _close(a: float, b: float, abs_tol: float, rel_tol: float) -> bool:
    return abs(a - b) <= max(abs_tol, rel_tol * abs(b))


def _fmt(kg: Optional[float]) -> str:
    return "—" if kg is None else f"{kg:,.2f} kg".replace(",", "X").replace(".", ",").replace("X", ".")


def reconcile(sections: List[Section], config: dict, pages: Dict[int, dict]):
    """Trả (sections mới, báo cáo từng hạng mục, danh sách nghi vấn)."""
    abs_tol, rel_tol = config["khop"]["tuyet_doi_kg"], config["khop"]["tuong_doi"]
    reports: List[ItemReport] = []
    flags: List[NghiVan] = []
    by_name = {it["ten"]: it for it in config["hang_muc"]}
    # hạng mục mới chèn sau một hạng mục có sẵn
    insert_after = {it["chen_sau"]: it for it in config["hang_muc"] if it.get("moi")}

    def old_lookup(rows: List[BbsRow]):
        return {(r.mark, int(r.d or 0), round(r.L_m or 0, 3)): r for r in rows}

    def process(name: str, old_rows: List[BbsRow]) -> List[BbsRow]:
        item = by_name.get(name)
        cable = [r for r in old_rows if r.is_cable]
        broken = [r for r in old_rows if r.is_broken]
        valid = [r for r in old_rows if not r.is_cable and not r.is_broken]
        old_kg = sum(r.kg for r in valid)
        broken_kg = sum(r.kg for r in broken)
        for r in broken:
            flags.append(NghiVan(name, r.mark, "Dòng BBS cũ sai rõ ràng (Ø không phải thép thanh hoặc dài < 50 mm) — đã loại",
                                 f"Ø{r.d:g} L={r.L_m} m × {r.n_per:g} × {r.n_ck:g}, {r.unit} kg/m = {_fmt(r.kg)}"
                                 if r.d is not None and None not in (r.L_m, r.n_per, r.n_ck) else _fmt(r.kg), "đã loại"))
        conv = lambda rows: [_with_conv(r) for r in rows]
        if item is None:                                       # hạng mục không có trong cấu hình: giữ nguyên
            out = conv(valid) + cable
            reports.append(ItemReport(name, "", "KHÔNG ĐỐI CHIẾU", old_kg, broken_kg, None, ghi_chu="chưa ánh xạ trang bản vẽ"))
            return out

        # hạng mục chỉ có tổng bản vẽ (không trích được từng thanh)
        if "trang" not in item:
            drawing = item["tong_ban_ve_kg"]
            pages_txt = ", ".join(f"tr.{p}" for p in item.get("trang_tham_chieu", []))
            if valid and _close(old_kg, drawing, abs_tol, rel_tol):
                reports.append(ItemReport(name, pages_txt, KHOP, old_kg, broken_kg, drawing, ghi_chu=item.get("ghi_chu", "")))
                flags.append(NghiVan(name, "", "Khớp tổng bản vẽ nhưng cấu thành suy ra — cần xác nhận: " + item.get("ghi_chu", ""),
                                     _fmt(old_kg), _fmt(drawing), pages_txt))
                return conv(valid) + cable
            ph = BbsRow(name, name, "—", None, placeholder=True,
                        ghi_chu=f"Bản vẽ {pages_txt}: {_fmt(drawing)}",
                        nghi_van=[f"Bản vẽ {pages_txt} ghi {_fmt(drawing)}; BBS cũ {_fmt(old_kg + broken_kg)}. "
                                  f"{item.get('ghi_chu', '')}"])
            reports.append(ItemReport(name, pages_txt, GIU, old_kg, broken_kg, drawing, ghi_chu=item.get("ghi_chu", "")))
            out = conv(valid)
            for r in out:
                r.nghi_van.append(f"Bản vẽ {pages_txt} ghi {_fmt(drawing)} cho cả hạng mục; BBS {_fmt(old_kg)}")
            return out + [ph] + cable

        # hạng mục có bảng thép theo trang
        info = []
        for pg in item["trang"]:
            t = drawing_table(pages[pg["so"]], pg) if pg["so"] in pages else None
            info.append((pg, t))
        pages_txt = ", ".join(f"tr.{pg['so']}" for pg, _ in info)
        mult_known = all(pg["he_so"] is not None for pg, _ in info)
        have_total = all(t is not None and t.stated_total_kg is not None for _, t in info)
        drawing = (sum(t.stated_total_kg * pg["he_so"] for pg, t in info) if mult_known and have_total else None)
        # chủ dự án quyết định tính cả các thanh mà bản vẽ ghi 0 kg (cấu hình 'tinh_dong_ban_ve_ghi_0')
        cong_ghi_0 = (sum(zero_weight_rows_kg(t) * pg["he_so"] for pg, t in info)
                      if drawing is not None and item.get("tinh_dong_ban_ve_ghi_0") else 0.0)
        if cong_ghi_0:
            flags.append(NghiVan(name, "", f"Bản vẽ ghi 0 kg cho một số thanh (dài × SL × kg/m = {_fmt(cong_ghi_0)}); "
                                           f"{item['tinh_dong_ban_ve_ghi_0']} → tổng đối chiếu = tổng bản vẽ + "
                                           f"{_fmt(cong_ghi_0)}", _fmt(drawing), _fmt(drawing + cong_ghi_0), pages_txt))
            drawing += cong_ghi_0
        all_usable = all(t is not None and table_usable(t) for _, t in info)
        gaps = {pg["so"]: (0.0 if t is not None and table_usable(t) else (partial_gap_kg(t) if t is not None else None))
                for pg, t in info}
        lookup = old_lookup(valid)
        _check_summary(name, item, drawing, flags, abs_tol, rel_tol)

        # 1) khớp: giữ BBS (so cả theo Ø nếu bản vẽ đọc đủ)
        if valid and drawing is not None and not broken and _close(old_kg, drawing, abs_tol, rel_tol):
            per_dia_ok = True
            if all_usable:
                dr: Dict[int, float] = defaultdict(float)
                for pg, t in info:
                    for r in t.rows:
                        dr[r.diameter_mm] += r.total_kg * pg["he_so"]
                old_pd = _per_dia(valid)
                per_dia_ok = all(_close(old_pd.get(d, 0.0), dr.get(d, 0.0), *PER_DIA_TOL) for d in set(dr) | set(old_pd))
            if per_dia_ok:
                reports.append(ItemReport(name, pages_txt, KHOP, old_kg, broken_kg, drawing))
                return conv(valid) + cable

        # 2) thay theo bản vẽ (đủ hoặc thiếu ít), khi biết số cấu kiện
        if mult_known and all(g is not None for g in gaps.values()):
            action = THAY if all(g == 0 for g in gaps.values()) else THAY_MOT_PHAN
            new_rows: List[BbsRow] = []
            for pg, t in info:
                src = f"tr.{pg['so']}"
                for r in good_rows(t):
                    old = lookup.get((r.mark, r.diameter_mm, round(r.length_mm / 1000, 3)))
                    row = BbsRow(name, f"{pg['bo_phan']} ({src})", r.mark if r.mark and not r.mark.isdigit() else "(?)",
                                 float(r.diameter_mm), mac=mac_thep(r.diameter_mm, r.tron_tron), hinh=old.hinh if old else "",
                                 L_m=round(r.length_mm / 1000, 3), n_per=r.quantity, n_ck=pg["he_so"] * r.multiplier,
                                 unit=rebar_unit_weight_kg_m(r.diameter_mm), nhom=nhom_duong_kinh(r.diameter_mm),
                                 ghi_chu=f"Theo bản vẽ thẩm định {src}")
                    row.nghi_van.append(f"Thay theo {src} (BBS cũ {_fmt(old_kg + broken_kg)}, bản vẽ {_fmt(drawing)})")
                    if not r.mark or r.mark.isdigit():
                        row.nghi_van.append(f"OCR không đọc được ký hiệu thanh (STT {r.stt or '?'}) — xem PDF {src}")
                    row.nghi_van.extend(f"{n} — xem PDF {src}" for n in r.notes)
                    if r.ban_ve_ghi_0:
                        row.nghi_van.append(item.get("tinh_dong_ban_ve_ghi_0") or
                                            "BBS vẫn giữ thanh này để gia công (có dài và số lượng trên bản vẽ) nên "
                                            "BBS lớn hơn tổng bản vẽ phần này — chủ dự án quyết định và ghim")
                    new_rows.append(row)
                gap = gaps[pg["so"]]
                if gap:
                    bad = [r for r in t.rows if numeric_issues(r)]
                    detail = "; ".join(f"{r.mark or '?'} Ø{r.diameter_mm}: {numeric_issues(r)[0]}" for r in bad[:4])
                    new_rows.append(BbsRow(
                        name, f"{pg['bo_phan']} ({src})", "—", None, placeholder=True,
                        ghi_chu=f"Thiếu {_fmt(gap * pg['he_so'])} so với tổng {src}",
                        nghi_van=[f"Bản vẽ {src} ghi tổng {_fmt(t.stated_total_kg * pg['he_so'])}; các dòng đọc được "
                                  f"{_fmt((t.stated_total_kg - gap) * pg['he_so'])}; THIẾU {_fmt(gap * pg['he_so'])} do OCR "
                                  f"mất/đọc sai dòng — nhập bổ sung theo PDF" + (f". Dòng đọc sai: {detail}" if detail else "")]))
            reports.append(ItemReport(name, pages_txt, action, old_kg, broken_kg, drawing,
                                      ghi_chu=item.get("ghi_chu_he_so", "")))
            # một mục cho cả hạng mục, không lặp ở từng dòng; chủ dự án đã xác nhận ('da_xac_nhan') thì không còn nghi vấn
            if item.get("ghi_chu_he_so") and not item.get("da_xac_nhan"):
                flags.append(NghiVan(name, "", "Số cấu kiện cần xác nhận: " + item["ghi_chu_he_so"], "", "", pages_txt))
            flags.append(NghiVan(name, "", f"{action}: BBS cũ {_fmt(old_kg)}"
                                           + (f" (+ {len(broken)} dòng lỗi {_fmt(broken_kg)})" if broken else "")
                                           + f"; bản vẽ {_fmt(drawing)}", _fmt(old_kg + broken_kg), _fmt(drawing), pages_txt))
            return new_rows + cable

        # 3) giữ BBS, đánh dấu
        details = []
        for pg, t in info:
            src = f"tr.{pg['so']}"
            if t is None:
                details.append(f"{src}: không đọc được bảng")
            elif pg["he_so"] is None:
                details.append(f"{src} ({pg['bo_phan']}): 1 cấu kiện = {_fmt(t.stated_total_kg)}, chưa biết số cấu kiện")
            elif gaps[pg["so"]] is None:
                details.append(f"{src}: bảng OCR không khớp tổng {_fmt(t.stated_total_kg)}")
        note = "; ".join(details) + (f". {item['ghi_chu_he_so']}" if item.get("ghi_chu_he_so") else "")
        out = conv(valid)
        for r in out:
            r.nghi_van.append(f"Chưa đối chiếu được với bản vẽ — {note}")
        reports.append(ItemReport(name, pages_txt, GIU, old_kg, broken_kg, drawing, ghi_chu=note))
        flags.append(NghiVan(name, "", f"{GIU}: {note}", _fmt(old_kg), _fmt(drawing), pages_txt))
        return out + cable

    new_sections: List[Section] = []
    for sec in sections:
        items: "OrderedDict[str, List[BbsRow]]" = OrderedDict()
        for name, rows in sec.items.items():
            items[name] = process(name, rows)
            if name in insert_after:
                extra = insert_after[name]
                items[extra["ten"]] = process(extra["ten"], [])
        new_sections.append(Section(sec.title, items))
    for rep in reports:
        rep.bbs_moi_kg = sum(r.kg for sec in new_sections for r in sec.items.get(rep.hang_muc, []) if not r.is_cable)
    return new_sections, reports, flags


def _check_summary(name: str, item: dict, drawing: Optional[float], flags: List[NghiVan], abs_tol: float,
                   rel_tol: float) -> None:
    """Đối chiếu tổng các trang chi tiết với bảng tổng hợp toàn cầu ('tong_hop': {trang, kg}); lệch → nghi vấn."""
    th = item.get("tong_hop")
    if not th or drawing is None:
        return
    if not _close(drawing, th["kg"], abs_tol, rel_tol):
        flags.append(NghiVan(name, "", f"Tổng các trang chi tiết {_fmt(drawing)} ≠ bảng tổng hợp tr.{th['trang']} "
                                       f"{_fmt(th['kg'])} (lệch {_fmt(th['kg'] - drawing)}). {th.get('ghi_chu', '')}".strip(),
                             _fmt(drawing), _fmt(th["kg"]), f"tr.{th['trang']}"))


def _with_conv(r: BbsRow) -> BbsRow:
    new = BbsRow(**{k: getattr(r, k) for k in r.__dataclass_fields__ if k != "nghi_van"})
    new.nghi_van = list(r.nghi_van)
    new.unit = _conv_unit(r)
    if not new.nhom and new.d is not None:
        new.nhom = nhom_duong_kinh(new.d)
    return new


def unit_changes(sections: List[Section]) -> List[Tuple[int, float, float]]:
    """(Ø, kg/m cũ trong BBS, kg/m theo quy ước) cho các Ø có thay đổi."""
    seen = {}
    for sec in sections:
        for rows in sec.items.values():
            for r in rows:
                if r.d is not None and not r.is_cable and int(r.d) in REBAR_UNIT_WEIGHT_KG_M and r.unit is not None:
                    seen.setdefault(int(r.d), r.unit)
    return [(d, old, rebar_unit_weight_kg_m(d)) for d, old in sorted(seen.items())
            if abs(old - rebar_unit_weight_kg_m(d)) > 1e-9]


def load_pages(path: str) -> Dict[int, dict]:
    with open(path, encoding="utf-8") as f:
        return {p["page"]: p for p in json.load(f)["pages"]}


def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ── GHI FILE ─────────────────────────────────────────────────────────────────────────────────────────────

HEADERS = ["TT", "Hạng mục kết cấu", "Bộ phận cấu kiện", "Ký hiệu thanh", "Đường kính Ø (mm)", "Mác thép",
           "Hình dạng thanh", "Chiều dài 1 thanh (m)", "Số thanh / cấu kiện", "Số cấu kiện", "Tổng số thanh",
           "Tổng chiều dài (m)", "Trọng lượng 1m (kg/m)", "Tổng khối lượng (kg)", "Tổng khối lượng (Tấn)",
           "Phân nhóm Ø", "Ghi chú & Tiêu chuẩn", "NGHI VẤN (soát & ghim)"]
GENERIC_PREFIXES = ("Thay theo tr.", "Chưa đối chiếu được với bản vẽ")


def write_bbs(path: str, titles: Dict[str, str], sections: List[Section], reports: List[ItemReport],
              flags: List[NghiVan], unit_chg: List[Tuple[int, float, float]], note: str) -> Dict[str, int]:
    """Ghi BBS đã sửa (công thức giữ như BBS cũ). Trả số dòng thép, dòng nghi vấn, mục trong sheet NGHI_VAN."""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "THONG_KE_THEP_CHI_TIET"
    bold, title_font = Font(bold=True), Font(bold=True, size=13, color="1F3A5E")
    head_fill, sec_fill = PatternFill("solid", fgColor="1F3A5E"), PatternFill("solid", fgColor="D9E1F2")
    flag_fill, gap_fill = PatternFill("solid", fgColor="FFF2CC"), PatternFill("solid", fgColor="F8CBAD")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ncol = len(HEADERS)

    ws["B2"], ws["B3"] = titles.get("B2"), titles.get("B3")
    ws["B4"] = note
    for c in ("B2", "B3"):
        ws[c].font = title_font
    ws["B4"].font = Font(italic=True, color="C00000")
    for c, h in enumerate(HEADERS, start=1):
        cell = ws.cell(5, c, h)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), head_fill
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    r = 6
    tt = 0
    row_of: Dict[int, int] = {}            # id(BbsRow) → dòng Excel
    n_steel = n_flag = 0
    for sec in sections:
        ws.cell(r, 1, sec.title).font = bold
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        ws.cell(r, 1).fill = sec_fill
        r += 1
        for rows in sec.items.values():
            for row in rows:
                if row.placeholder:
                    ws.cell(r, 1, "⚠")
                    vals = {2: row.hang_muc, 3: row.bo_phan, 4: row.mark, 17: row.ghi_chu}
                else:
                    tt += 1
                    n_steel += 1
                    ws.cell(r, 1, tt)
                    vals = {2: row.hang_muc, 3: row.bo_phan, 4: row.mark, 5: row.d, 6: row.mac, 7: row.hinh,
                            8: row.L_m, 9: row.n_per, 10: row.n_ck, 11: f"=I{r}*J{r}", 12: f"=H{r}*K{r}",
                            13: row.unit, 14: f"=L{r}*M{r}", 15: f"=N{r}/1000", 16: row.nhom, 17: row.ghi_chu}
                for c, v in vals.items():
                    ws.cell(r, c, v)
                if row.nghi_van:
                    n_flag += 1
                    ws.cell(r, ncol, " | ".join(row.nghi_van))
                    fill = gap_fill if row.placeholder else flag_fill
                    for c in range(1, ncol + 1):
                        ws.cell(r, c).fill = fill
                for c in range(1, ncol + 1):
                    ws.cell(r, c).border = border
                row_of[id(row)] = r
                r += 1
    last = r - 1
    ws.cell(r, 1, "TỔNG CỘNG CỐT THÉP VÀ CÁP DƯL TOÀN BỘ CÔNG TRÌNH").font = bold
    ws.cell(r, 14, f"=SUM(N6:N{last})").font = bold
    ws.cell(r, 15, f"=SUM(O6:O{last})").font = bold
    total_row = r
    r += 2
    ws.cell(r, 2, "BẢNG TỔNG HỢP THEO PHÂN NHÓM ĐƯỜNG KÍNH Ø").font = bold
    r += 1
    for c, h in enumerate(["Nhóm đường kính", "Khối lượng (kg)", "Khối lượng (Tấn)", "Tỷ trọng (%)"], start=2):
        ws.cell(r, c, h).font = bold
    first_grp = r + 1
    for g in ("D<=10mm", "10<D<=18mm", "D>18mm", CAP_DUL):
        r += 1
        ws.cell(r, 2, g)
        ws.cell(r, 3, f'=SUMIF($P$6:$P${last}, "{g}", $N$6:$N${last})')
        ws.cell(r, 4, f"=C{r}/1000")
        ws.cell(r, 5, f"=C{r}/$N${total_row}")
    r += 1
    ws.cell(r, 2, "TỔNG CỘNG").font = bold
    ws.cell(r, 3, f"=SUM(C{first_grp}:C{r - 1})")
    ws.cell(r, 4, f"=SUM(D{first_grp}:D{r - 1})")
    widths = [6, 30, 30, 14, 12, 11, 14, 12, 12, 10, 12, 14, 12, 16, 12, 14, 30, 70]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "E6"

    # ── tổng hợp theo Ø (đủ mọi Ø có mặt, kể cả Ø8 mà bản cũ thiếu) ──
    th = wb.create_sheet("TONG_HOP_THEP_THEO_DUONG_KINH")
    th["A1"], th["A2"] = titles.get("B2"), "BẢNG TỔNG HỢP CỐT THÉP THEO ĐƯỜNG KÍNH Ø (kg/m theo quy ước hồ sơ)"
    th["A1"].font = th["A2"].font = title_font
    cols = ["STT", "Đường kính Ø (mm)", "Mác thép", "Trọng lượng đơn vị (kg/m)", "Tổng chiều dài (m)",
            "Tổng trọng lượng (kg)", "Tổng trọng lượng (Tấn)", "Tỷ trọng (%)"]
    for c, h in enumerate(cols, start=1):
        cell = th.cell(4, c, h)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), head_fill
    dias = sorted({int(x.d) for s in sections for rows in s.items.values() for x in rows
                   if x.d is not None and not x.is_cable and not x.placeholder})
    rr = 4
    src = "THONG_KE_THEP_CHI_TIET"
    for i, d in enumerate(dias, start=1):
        rr += 1
        th.append([i, d, mac_thep(d), rebar_unit_weight_kg_m(d),
                   f"=SUMIF({src}!$E$6:$E${last}, {d}, {src}!$L$6:$L${last})",
                   f"=SUMIF({src}!$E$6:$E${last}, {d}, {src}!$N$6:$N${last})", f"=F{rr}/1000", None])
    rr += 1
    th.append([len(dias) + 1, "15.2 (cáp DƯL)", "ASTM A416 Gr270", None,
               f'=SUMIF({src}!$P$6:$P${last}, "{CAP_DUL}", {src}!$L$6:$L${last})',
               f'=SUMIF({src}!$P$6:$P${last}, "{CAP_DUL}", {src}!$N$6:$N${last})', f"=F{rr}/1000", None])
    total_th = rr + 1
    th.cell(total_th, 2, "TỔNG CỘNG").font = bold
    th.cell(total_th, 6, f"=SUM(F5:F{rr})").font = bold
    th.cell(total_th, 7, f"=SUM(G5:G{rr})").font = bold
    for k in range(5, rr + 1):
        th.cell(k, 8, f"=F{k}/$F${total_th}")
    for i, w in enumerate([6, 18, 16, 18, 18, 20, 18, 12], start=1):
        th.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # ── đối chiếu với bản vẽ theo hạng mục ──
    dc = wb.create_sheet("DOI_CHIEU_BAN_VE")
    dc["A1"] = "ĐỐI CHIẾU BBS VỚI BẢN VẼ THẨM ĐỊNH THEO HẠNG MỤC (thép thanh, không gồm cáp DƯL)"
    dc["A1"].font = title_font
    hdr = ["Hạng mục", "Trang bản vẽ", "Xử lý", "BBS cũ — dòng hợp lệ (kg)", "BBS cũ — dòng lỗi đã loại (kg)",
           "Bản vẽ thẩm định (kg)", "BBS mới (kg, kg/m theo quy ước)", "BBS mới − bản vẽ (kg)", "Ghi chú"]
    for c, h in enumerate(hdr, start=1):
        cell = dc.cell(3, c, h)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), head_fill
        cell.alignment = Alignment(wrap_text=True)
    for rep in reports:
        diff = None if rep.ban_ve_kg is None else round(rep.bbs_moi_kg - rep.ban_ve_kg, 2)
        dc.append([rep.hang_muc, rep.trang, rep.hanh_dong, round(rep.bbs_cu_kg, 2), round(rep.loi_cu_kg, 2),
                   None if rep.ban_ve_kg is None else round(rep.ban_ve_kg, 2), round(rep.bbs_moi_kg, 2), diff, rep.ghi_chu])
        if rep.hanh_dong != KHOP:
            for c in range(1, len(hdr) + 1):
                dc.cell(dc.max_row, c).fill = flag_fill
    for i, w in enumerate([40, 18, 22, 18, 18, 18, 20, 16, 80], start=1):
        dc.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # ── danh sách nghi vấn để soát & ghim ──
    nv = wb.create_sheet("NGHI_VAN")
    nv["A1"] = "DANH SÁCH NGHI VẤN — chủ dự án soát, ghi kết luận vào cột 'Đã soát / Ghim'"
    nv["A1"].font = title_font
    hdr = ["STT", "Dòng (sheet THONG_KE)", "Hạng mục", "Ký hiệu", "Nội dung nghi vấn", "Giá trị cũ", "Giá trị mới",
           "Trang bản vẽ", "Đã soát / Ghim"]
    for c, h in enumerate(hdr, start=1):
        cell = nv.cell(3, c, h)
        cell.font, cell.fill = Font(bold=True, color="FFFFFF"), head_fill
    entries: List[Tuple] = []
    if unit_chg:
        entries.append(("", "Toàn bảng", "", "Đổi kg/m theo quy ước hồ sơ (bảng solver cắt thép): "
                        + ", ".join(f"Ø{d}: {o:g} → {n:g}" for d, o, n in unit_chg), "", "", ""))
    entries.append(("", "Dòng thay theo bản vẽ", "", "Mác thép của dòng lấy theo bản vẽ đặt theo quy ước BBS cũ "
                    "(Ø ≤ 10: CB240-T; Ø ≥ 12: CB400-V); bảng bản vẽ không ghi mác", "", "", ""))
    for f in flags:
        entries.append(("", f.hang_muc, f.ky_hieu, f.noi_dung, f.gia_tri_cu, f.gia_tri_moi, f.trang))
    for sec in sections:
        for rows in sec.items.values():
            for row in rows:
                specific = [n for n in row.nghi_van if not n.startswith(GENERIC_PREFIXES)]
                for n in specific:
                    entries.append((row_of.get(id(row)), row.hang_muc, row.mark, n, "", "", ""))
    for i, (ref, hm, mark, text, old, new, pg) in enumerate(entries, start=1):
        nv.append([i, ref, hm, mark, text, old, new, pg, ""])
    for i, w in enumerate([6, 12, 36, 12, 110, 22, 22, 14, 24], start=1):
        nv.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    for row in nv.iter_rows(min_row=4, max_col=5, min_col=5):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(path)
    return {"dong_thep": n_steel, "dong_nghi_van": n_flag, "muc_nghi_van": len(entries)}
