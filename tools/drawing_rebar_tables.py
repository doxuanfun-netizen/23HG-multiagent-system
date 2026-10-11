# -*- coding: utf-8 -*-
"""
Đọc bảng thống kê cốt thép trên bản vẽ (dữ liệu OCR trong du_lieu.json) và tự kiểm tra trước khi dùng.

Không đoán: tiêu đề OCR hay sai dấu, gộp ô, lệch cột, nên cột chiều dài / số lượng / kg/m / khối lượng được
chọn sao cho PHÉP TÍNH CỦA CHÍNH BẢN VẼ (dài × số lượng × kg/m = khối lượng) đúng ở nhiều dòng nhất. Mỗi bảng
phải khớp dòng tổng mà bản vẽ ghi. Bảng không qua kiểm tra trả về kèm lý do để người dùng xem PDF, và không
được dùng để thay số liệu.

Dữ liệu vào: một trang trong du_lieu.json (khóa 'page', 'tables' với 'header', 'rows', 'values').
"""

from __future__ import annotations

import itertools
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

ROW_REL_TOL = 0.005        # dòng: |dài×SL×kg/m − kg| ≤ 0,5% (làm tròn trên bản vẽ)
ROW_ABS_TOL = 0.06         # … hoặc ≤ 0,06 kg cho dòng rất nhỏ
TABLE_REL_TOL = 0.001      # bảng: |Σ dòng − tổng ghi trên bản vẽ| ≤ 0,1%
TABLE_ABS_TOL = 0.1
UNIT_RANGE = (0.2, 10.0)   # kg/m của thép thanh Ø6..Ø40
REBAR_DIAMETERS = {6, 8, 10, 12, 14, 16, 18, 20, 22, 25, 28, 32, 36, 40}


def norm(text) -> str:
    """Bỏ dấu, viết hoa, gộp khoảng trắng — để so tiêu đề OCR sai dấu ('KHÓI', 'TÒNG', 'SÓ')."""
    s = unicodedata.normalize("NFD", str(text or "")).replace("đ", "d").replace("Đ", "D")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s.upper()).strip()


@dataclass
class DrawingRow:
    page: int
    stt: Optional[str]
    mark: str
    diameter_mm: int
    length_mm: float
    quantity: float
    unit_kg_m: float
    kg: float                        # khối lượng 1 cấu kiện theo bản vẽ
    multiplier: float = 1.0          # cột "cấu kiện" / "số đốt" / "số đơn nguyên"
    total_kg: float = 0.0            # khối lượng × số cấu kiện (theo bản vẽ)
    issues: List[str] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)     # ghi chú không phải lỗi (vd. hệ số suy từ cột tổng)

    @property
    def ok(self) -> bool:
        return not self.issues


@dataclass
class DrawingTable:
    page: int
    title: str
    rows: List[DrawingRow]
    stated_total_kg: Optional[float]
    issues: List[str] = field(default_factory=list)

    @property
    def rows_total_kg(self) -> float:
        return round(sum(r.total_kg for r in self.rows), 2)


def _num(v) -> Optional[float]:
    return float(v) if isinstance(v, (int, float)) else None


def _diameter(text, allow_plain: bool = False) -> Optional[int]:
    """'D16', 'Ø16', 'Φ16' → 16. Số trần ('16') chỉ nhận khi cột có tiêu đề đường kính (allow_plain), vì cột số
    lượng cũng hay có giá trị 16, 28..."""
    t = norm(text).replace("Ø", "D").replace("Φ", "D")
    pattern = r"(?:D|PHI)?\s*(\d{1,2})(?:\s*MM)?" if allow_plain else r"(?:D|PHI)\s*(\d{1,2})(?:\s*MM)?"
    m = re.fullmatch(pattern, t)
    if m and int(m.group(1)) in REBAR_DIAMETERS:
        return int(m.group(1))
    return None


def _is_group_subtotal(text: str) -> bool:
    """Dòng tổng theo nhóm đường kính ('10<D<=18', 'D>18MM') — không phải tổng chung của bảng."""
    return bool(re.search(r"D\s*[<>]=?|<\s*D|>\s*D", text))


def _row_ok(length, qty, unit, kg) -> bool:
    if None in (length, qty, unit, kg) or not (UNIT_RANGE[0] <= unit <= UNIT_RANGE[1]) or length <= 0 or qty <= 0:
        return False
    return abs(length / 1000.0 * qty * unit - kg) <= max(ROW_ABS_TOL, ROW_REL_TOL * kg)


def _blocks(header: Sequence[str], rows: Sequence[Sequence]) -> List[range]:
    """Tách bảng ghép cạnh nhau (vd. bản mặt cầu | liên tục nhiệt): mỗi khối bắt đầu ngay trước một cột Ø."""
    dia_cols = [i for i in range(len(header))
                if "KINH" in norm(header[i]) or sum(1 for r in rows if i < len(r) and _diameter(r[i])) >= 2]
    dia_cols = [i for i in dia_cols if "KINH" in norm(header[i])] or dia_cols
    starts = []
    for i in dia_cols:
        s = max(0, i - 1)
        if not starts or s > starts[-1] + 3:
            starts.append(s)
    if len(starts) <= 1:
        return [range(0, len(header))]
    starts[0] = 0
    return [range(a, b) for a, b in zip(starts, starts[1:] + [len(header)])]


def _parse_block(page: int, title: str, header, rows, values, cols: range) -> Optional[DrawingTable]:
    hdr = {i: norm(header[i]) if i < len(header) else "" for i in cols}
    data = [(r, v) for r, v in zip(rows, values)]
    # cột Ø: ưu tiên cột có tiêu đề đường kính; nếu không có, cột có nhiều ô dạng 'D16'
    plain = {i: "KINH" in hdr[i] for i in cols}
    dia_count = lambda i: sum(1 for r, _ in data if i < len(r) and _diameter(r[i], plain[i]))
    dia_col = max(cols, key=lambda i: (plain[i] and dia_count(i) >= 2, dia_count(i)))
    if dia_count(dia_col) < 2:
        return None
    is_dia = lambda r: dia_col < len(r) and _diameter(r[dia_col], plain[dia_col])
    body = [(r, v) for r, v in data if is_dia(r)]
    numeric = [i for i in cols if i != dia_col and sum(1 for _, v in body if i < len(v) and _num(v[i]) is not None)
               >= max(2, len(body) // 2)]
    # dài × SL × kg/m có tính giao hoán: đổi chỗ cột dài và cột số lượng vẫn "đúng phép tính". Vì vậy khi hòa điểm
    # thì ưu tiên cách gán khớp tiêu đề cột, rồi đến cột dài có giá trị lớn hơn cột số lượng.
    hints = (("CHIEU DAI",), ("SO LUONG", "SO THANH"), ("KG/M", "DON V", "1M DAI"), ("KHOI LUONG",))

    def median(c):
        xs = sorted(_num(v[c]) for _, v in body if c < len(v) and _num(v[c]) is not None)
        return xs[len(xs) // 2] if xs else 0.0

    best, best_key = None, None
    for L, q, u, k in itertools.permutations(numeric, 4):
        score = sum(1 for _, v in body if _row_ok(*(_num(v[c]) if c < len(v) else None for c in (L, q, u, k))))
        if score == 0:
            continue
        bonus = sum(1 for c, words in zip((L, q, u), hints) if any(w in hdr[c] for w in words))
        # cột khối lượng: ưu tiên khối lượng 1 cấu kiện ('KHỐI LƯỢNG', '... KG'), không phải cột 'TỔNG'
        if ("KHOI LUONG" in hdr[k] or hdr[k].endswith("KG") or "(KG)" in hdr[k]) and "TONG" not in hdr[k]:
            bonus += 1
        key = (score, bonus, median(L) >= median(q))
        if best_key is None or key > best_key:
            best, best_key = (L, q, u, k), key
    best_score = best_key[0] if best_key else 0
    if best is None or best_score < max(2, len(body) // 2):
        return DrawingTable(page, title, [], None, [f"không tìm được cột dài/SL/kg/m/khối lượng khớp phép tính "
                                                     f"(khớp {best_score}/{len(body)} dòng)"])
    L, q, u, k = best
    rest = [i for i in numeric if i not in best]
    total_col = next((i for i in rest if "TONG" in hdr[i]), None)
    mult_col = next((i for i in rest if any(w in hdr[i] for w in ("CAU KIEN", "DOT", "DON NGUYEN", "SO KHE", "SO LUONG BAN"))
                     and "KG" not in hdr[i] and "KHOI LUONG" not in hdr[i] and i != total_col), None)
    mark_col = next((i for i in cols if any(w in hdr[i] for w in ("TEN THANH", "SO HIEU", "KY HIEU"))), None)
    if mark_col is None:
        cand = [i for i in cols if i < dia_col and i not in numeric]
        mark_col = cand[-1] if cand else None

    out: List[DrawingRow] = []
    stated = None
    for r, v in data:
        text = " ".join(norm(r[i]) for i in cols if i < len(r))
        if "TONG" in text and "THEP" in text:
            nums = [v[i] for i in cols if i < len(v) and isinstance(v[i], (int, float))]
            if stated is None and nums and not _is_group_subtotal(text):
                stated = nums[-1]
            continue
        d = _diameter(r[dia_col], plain[dia_col]) if dia_col < len(r) else None
        vals = [(_num(v[c]) if c < len(v) else None) for c in (L, q, u, k)]
        if d is None or vals[0] is None or vals[1] is None:
            continue
        mult = _num(v[mult_col]) if mult_col is not None and mult_col < len(v) else None
        if mult is not None and not (1 <= mult <= 500 and abs(mult - round(mult)) <= 0.01):
            mult = None                          # hệ số số cấu kiện phải là số nguyên hợp lý
        total = _num(v[total_col]) if total_col is not None and total_col < len(v) else None
        mark = str(r[mark_col]).strip() if mark_col is not None and mark_col < len(r) else ""
        stt = str(r[0]).strip() if r and str(r[0]).strip().isdigit() else None
        row = DrawingRow(page, stt, mark, d, vals[0], vals[1], vals[2] or 0.0, vals[3] or 0.0)
        if mult is None and total is not None and row.kg:
            m = total / row.kg               # cột số cấu kiện/số đốt trống nhưng có cột tổng: suy từ số trên bản vẽ
            if m >= 1 and abs(m - round(m)) <= 0.01:
                mult = float(round(m))
                if mult != 1:
                    row.notes.append(f"hệ số {mult:g} suy từ cột tổng ({total:,.2f} ÷ {row.kg:,.2f})")
        row.multiplier = mult or 1.0
        row.total_kg = round(row.kg * row.multiplier, 2)
        if not mark or mark.isdigit():
            row.issues.append(f"không đọc được ký hiệu thanh (STT {stt or '?'})")
        if not _row_ok(*vals):
            calc = vals[0] / 1000 * vals[1] * (vals[2] or 0)
            row.issues.append(f"dài×SL×kg/m = {calc:,.2f} ≠ khối lượng ghi {vals[3] or 0:,.2f}")
        if total is not None and abs(total - row.total_kg) > max(ROW_ABS_TOL, ROW_REL_TOL * total):
            row.issues.append(f"khối lượng × hệ số = {row.total_kg:,.2f} ≠ tổng ghi {total:,.2f}")
        out.append(row)
    t = DrawingTable(page, title, out, stated)
    if not out:
        t.issues.append("không đọc được dòng thép nào")
    if stated is None:
        t.issues.append("không thấy dòng tổng khối lượng cốt thép")
    elif out and abs(t.rows_total_kg - stated) > max(TABLE_ABS_TOL, TABLE_REL_TOL * stated):
        t.issues.append(f"Σ dòng = {t.rows_total_kg:,.2f} kg ≠ tổng ghi trên bản vẽ {stated:,.2f} kg")
    return t


def parse_table(page: int, table: dict) -> List[DrawingTable]:
    header, rows, values = table.get("header") or [], table.get("rows") or [], table.get("values") or []
    width = max([len(header)] + [len(r) for r in rows] or [0])
    header = list(header) + [""] * (width - len(header))
    out = []
    for cols in _blocks(header, rows):
        t = _parse_block(page, table.get("title") or "", header, rows, values, cols)
        if t is not None:
            out.append(t)
    return out


def page_tables(page: dict) -> List[DrawingTable]:
    out: List[DrawingTable] = []
    for t in page.get("tables") or []:
        out.extend(parse_table(page["page"], t))
    return out


def numeric_issues(row: DrawingRow) -> List[str]:
    """Lỗi làm sai số liệu (khác lỗi chỉ thiếu ký hiệu)."""
    return [i for i in row.issues if not i.startswith("không đọc được ký hiệu")]


def table_usable(t: DrawingTable) -> bool:
    """Dùng được để thay số liệu: có dòng tổng, Σ dòng khớp, không dòng nào sai số (thiếu ký hiệu thì vẫn dùng)."""
    return (t.stated_total_kg is not None and not t.issues and bool(t.rows)
            and all(not numeric_issues(r) for r in t.rows))


PARTIAL_MAX_GAP = 0.05     # thiếu tối đa 5% so với tổng bản vẽ thì vẫn dùng phần đọc được (kèm dòng nghi vấn)


def page_rebar_table(page: dict) -> Optional[DrawingTable]:
    """
    Một bảng thép cho cả trang. OCR hay cắt một bảng thành nhiều mảnh (vd. Mố M2 trang 40): nếu chỉ một mảnh có
    dòng tổng và ghép các mảnh làm Σ dòng gần tổng hơn thì ghép lại. Trang có nhiều bảng đều có tổng (vd. bản
    mặt cầu loại 1 | loại 2) thì không ghép — trả None để người gọi xử lý riêng.
    """
    tables = [t for t in page_tables(page) if t.rows]
    if not tables:
        return None
    with_total = [t for t in tables if t.stated_total_kg is not None]
    if len(tables) == 1:
        return tables[0]
    if len(with_total) != 1:
        return None
    main = with_total[0]
    merged_rows = [r for t in tables for r in t.rows]
    merged_sum = sum(r.total_kg for r in merged_rows)
    if abs(merged_sum - main.stated_total_kg) >= abs(main.rows_total_kg - main.stated_total_kg):
        return main
    t = DrawingTable(page["page"], main.title, merged_rows, main.stated_total_kg)
    if abs(t.rows_total_kg - t.stated_total_kg) > max(TABLE_ABS_TOL, TABLE_REL_TOL * t.stated_total_kg):
        t.issues.append(f"Σ dòng = {t.rows_total_kg:,.2f} kg ≠ tổng ghi trên bản vẽ {t.stated_total_kg:,.2f} kg")
    return t


def good_rows(t: DrawingTable) -> List[DrawingRow]:
    return [r for r in t.rows if not numeric_issues(r)]


def partial_gap_kg(t: DrawingTable) -> Optional[float]:
    """
    Phần thiếu (kg) khi bảng chỉ đọc được một phần: tổng bản vẽ − Σ các dòng tự khớp. None nếu không dùng được
    (không có tổng, hoặc các dòng đúng đã vượt tổng, hoặc thiếu quá PARTIAL_MAX_GAP).
    """
    if t.stated_total_kg is None:
        return None
    gap = round(t.stated_total_kg - sum(r.total_kg for r in good_rows(t)), 2)
    if gap < -max(TABLE_ABS_TOL, TABLE_REL_TOL * t.stated_total_kg) or gap > PARTIAL_MAX_GAP * t.stated_total_kg:
        return None
    return max(gap, 0.0)
