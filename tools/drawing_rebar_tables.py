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
    tron_tron: bool = False          # thép tròn trơn ('R25', 'R8') — mác CB240-T
    ban_ve_ghi_0: bool = False       # chính bản vẽ ghi khối lượng 0 dù có dài × SL × kg/m (sai trên bản vẽ)

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
    pending_fixes: list = field(default_factory=list)    # sửa OCR chờ kiểm khi ghép với mảnh có dòng tổng

    @property
    def rows_total_kg(self) -> float:
        return round(sum(r.total_kg for r in self.rows), 2)


def _num(v) -> Optional[float]:
    return float(v) if isinstance(v, (int, float)) else None


def _diameter(text, allow_plain: bool = False) -> Optional[int]:
    """'D16', 'Ø16', 'Φ16' → 16. Số trần ('16') chỉ nhận khi cột có tiêu đề đường kính (allow_plain), vì cột số
    lượng cũng hay có giá trị 16, 28..."""
    t = norm(text).replace("Ø", "D").replace("Φ", "D").replace(" ", "")      # OCR hay tách 'D18' thành 'D1 8'
    pattern = r"(?:D|PHI|R)?(\d{1,2})(?:MM)?" if allow_plain else r"(?:D|PHI|R)(\d{1,2})(?:MM)?"
    m = re.fullmatch(pattern, t)
    if m and int(m.group(1)) in REBAR_DIAMETERS:
        return int(m.group(1))
    return None


# kg/m ghi trên bản vẽ (làm tròn) và theo bảng quy ước → Ø. Chỉ dùng khi ô Ø OCR đọc hỏng ('DB', 'D1').
UNIT_TO_DIAMETER = {0.222: 6, 0.395: 8, 0.617: 10, 0.888: 12, 1.21: 14, 1.208: 14, 1.58: 16, 1.578: 16, 2.0: 18,
                    1.998: 18, 2.47: 20, 2.466: 20, 2.98: 22, 2.984: 22, 3.85: 25, 3.853: 25, 4.83: 28, 4.834: 28,
                    6.31: 32, 6.313: 32}


def _diameter_from_unit(text, unit: Optional[float]) -> Optional[int]:
    t = norm(text).replace(" ", "")
    if unit is None or not re.fullmatch(r"[DR]\w{0,2}", t):
        return None
    return next((d for u, d in UNIT_TO_DIAMETER.items() if abs(u - unit) <= 0.0005), None)


def _is_group_subtotal(text: str) -> bool:
    """Dòng tổng theo nhóm đường kính ('10<D<=18', 'D>18MM') — không phải tổng chung của bảng."""
    return bool(re.search(r"D\s*[<>]=?|<\s*D|>\s*D", text))


def _row_ok(length, qty, unit, kg) -> bool:
    if None in (length, qty, unit, kg) or not (UNIT_RANGE[0] <= unit <= UNIT_RANGE[1]) or length <= 0 or qty <= 0:
        return False
    return abs(length / 1000.0 * qty * unit - kg) <= max(ROW_ABS_TOL, ROW_REL_TOL * kg)


def _integer_qty(length, unit, kg) -> Optional[float]:
    """Số thanh suy từ khối lượng ghi: kg ÷ (dài × kg/m), chỉ nhận khi ra số nguyên (sai lệch ≤ làm tròn)."""
    if not (UNIT_RANGE[0] <= unit <= UNIT_RANGE[1]) or length <= 0 or kg <= 0:
        return None
    n = round(kg / (length / 1000.0 * unit))
    return float(n) if n >= 1 and _row_ok(length, n, unit, kg) else None


def _apply_fixes(rows: List[DrawingRow], fixes, stated: Optional[float]) -> List[DrawingRow]:
    """
    Sửa lỗi OCR bằng chính phép tính của bản vẽ. fixes = [(dòng, các cách sửa, dòng_suy)]:
    - dòng lỗi phép tính có vài cách sửa (SL đọc sai / khối lượng đọc sai) hoặc để nguyên;
    - "dòng suy" là dòng phải điền ô OCR trống/hỏng (Ø, chiều dài tách ô, kg/m, SL, khối lượng) — chỉ giữ khi
      được xác nhận.
    Chỉ nhận tổ hợp làm Σ dòng KHỚP tổng ghi trên bản vẽ, ưu tiên tổ hợp để lại ít dòng sai nhất. Không có tổ
    hợp nào khớp thì giữ nguyên số OCR (kèm lỗi) và bỏ các dòng suy. Dòng đã sửa có ghi chú để đánh dấu nghi vấn.
    """
    if not fixes or stated is None:          # mảnh bảng không có dòng tổng: để page_rebar_table kiểm khi ghép
        return rows
    derived = {id(row) for row, _, is_derived in fixes if is_derived}
    unfixed = [r for r in rows if id(r) not in derived]
    if len(fixes) > 10:
        return unfixed
    tol = max(TABLE_ABS_TOL, TABLE_REL_TOL * stated)
    fixed_ids = {id(row) for row, _, _ in fixes}
    base = sum(r.total_kg for r in rows if id(r) not in fixed_ids)

    def contrib(row, fix):
        if fix is None or fix[0] == "NOTE":
            return row.total_kg
        kind, val, _ = fix
        return round((val if kind == "KG" else row.kg) * row.multiplier, 2)

    options = [cands if is_derived else [None] + cands for _, cands, is_derived in fixes]
    best = None
    for combo in itertools.product(*options):
        if abs(round(base + sum(contrib(row, f) for (row, _, _), f in zip(fixes, combo)), 2) - stated) <= tol:
            n = sum(f is None for f in combo)          # ưu tiên tổ hợp sửa hết các dòng sai phép tính
            if best is None or n < best[0]:
                best = (n, combo)
    if best is None:
        return unfixed
    for (r, _, _), f in zip(fixes, best[1]):
        if f is None:
            continue
        kind, val, note = f
        if kind == "SL":
            r.quantity = val
        elif kind == "KG":
            r.kg = val
            r.total_kg = round(val * r.multiplier, 2)
        if kind != "NOTE":
            r.issues = [i for i in r.issues if not i.startswith("dài×SL×kg/m")]
        r.notes.append(note + " — Σ bảng sau sửa khớp tổng bản vẽ")
    return rows


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

    # kg/m chính bản vẽ dùng cho từng Ø — để điền ô kg/m mà OCR đọc hỏng
    page_units: Dict[int, float] = {}
    for r, v in body:
        uu = _num(v[u]) if u < len(v) else None
        if uu is not None and UNIT_RANGE[0] <= uu <= UNIT_RANGE[1]:
            page_units.setdefault(_diameter(r[dia_col], plain[dia_col]), uu)

    out: List[DrawingRow] = []
    fixes = []
    stated = None
    for r, v in data:
        text = " ".join(norm(r[i]) for i in cols if i < len(r))
        if "TONG" in text and "THEP" in text:
            nums = [v[i] for i in cols if i < len(v) and isinstance(v[i], (int, float))]
            if stated is None and nums and not _is_group_subtotal(text):
                stated = nums[-1]
            continue
        dtext = str(r[dia_col]) if dia_col < len(r) else ""
        d = _diameter(dtext, plain[dia_col])
        vals = [(_num(v[c]) if c < len(v) else None) for c in (L, q, u, k)]
        notes = []                           # ô phải suy — dòng chỉ giữ khi Σ bảng khớp tổng (xem _apply_fixes)
        if d is None:
            d = _diameter_from_unit(dtext, vals[2])
            if d is None:
                continue
            notes.append(f"OCR ghi Ø '{dtext.strip()}', Ø{d} suy từ kg/m {vals[2]:g}")
        if not vals[0]:
            continue
        if vals[0] < 1000 and L + 1 < len(r) and L + 1 not in (q, u, k) and str(r[L + 1]).strip() == "0":
            notes.append(f"OCR tách chiều dài thành '{vals[0]:g}' và '0' → {vals[0] * 10:g} mm")
            vals[0] *= 10
        if vals[2] is None and d in page_units:
            vals[2] = page_units[d]
            notes.append(f"kg/m không đọc được, lấy {vals[2]:g} như các thanh Ø{d} khác của bảng")
        mult = _num(v[mult_col]) if mult_col is not None and mult_col < len(v) else None
        if mult is not None and not (1 <= mult <= 500 and abs(mult - round(mult)) <= 0.01):
            mult = None                          # hệ số số cấu kiện phải là số nguyên hợp lý
        total = _num(v[total_col]) if total_col is not None and total_col < len(v) else None
        use_total = False                        # tổng dòng lấy đúng số ghi ở cột tổng
        if vals[3] is None and total and vals[1] and vals[2]:
            calc = vals[0] / 1000 * vals[1] * vals[2]
            m = total / calc
            if mult is None and m >= 1 and abs(m - round(m)) <= ROW_REL_TOL * m:
                mult, vals[3], use_total = float(round(m)), round(calc, 2), True
                notes.append(f"khối lượng 1 cấu kiện trống, = dài × SL × kg/m; hệ số {mult:g} = tổng ÷ khối lượng")
            else:
                vals[3] = round(total / (mult or 1.0), 2)
                notes.append(f"khối lượng 1 cấu kiện trống, lấy từ cột tổng ({total:,.2f})")
        elif vals[3] is None and total and vals[1] is None:
            vals[3] = round(total / (mult or 1.0), 2)
            notes.append(f"khối lượng 1 cấu kiện không đọc được, lấy từ cột tổng ({total:,.2f})")
        if vals[1] is None and vals[2] and vals[3]:
            n = _integer_qty(vals[0], vals[2], vals[3])
            if n is None:
                continue
            vals[1] = n
            notes.append(f"số lượng trống/không đọc được, suy = khối lượng ghi ÷ (dài × kg/m) = {n:g}")
        if vals[1] is None:
            continue
        # khối lượng ghi 0 ở cả cột khối lượng và cột tổng: sai của chính bản vẽ, không phải lỗi OCR
        design_zero = vals[3] == 0 and not total and bool(vals[2])
        fix = None
        if not design_zero and not _row_ok(*vals) and vals[2] and vals[3]:
            fix = []                         # các cách sửa có thể: đọc sai số lượng, hoặc đọc sai khối lượng
            n = _integer_qty(vals[0], vals[2], vals[3])
            if n is not None:
                fix.append(("SL", n, f"OCR đọc SL = {vals[1]:g}, phép tính bản vẽ cho {n:g} (khối lượng ghi ÷ dài ÷ kg/m)"))
            kg = round(vals[0] / 1000 * vals[1] * vals[2], 2)
            if UNIT_RANGE[0] <= vals[2] <= UNIT_RANGE[1] and kg > 0:
                fix.append(("KG", kg, f"OCR đọc khối lượng = {vals[3]:g}, dài × SL × kg/m = {kg:,.2f}"))
            fix = fix or None
        mark = str(r[mark_col]).strip() if mark_col is not None and mark_col < len(r) else ""
        stt = str(r[0]).strip() if r and str(r[0]).strip().isdigit() else None
        row = DrawingRow(page, stt, mark, d, vals[0], vals[1], vals[2] or 0.0, vals[3] or 0.0)
        row.tron_tron = norm(dtext).startswith("R")
        if mult is None and total is not None and row.kg:
            m = total / row.kg               # cột số cấu kiện/số đốt trống nhưng có cột tổng: suy từ số trên bản vẽ
            calc = row.length_mm / 1000 * row.quantity * row.unit_kg_m
            m2 = total / calc if calc else 0.0
            if round(m) >= 1 and abs(m - round(m)) <= 0.01:
                mult = float(round(m))
                if mult != 1:
                    row.notes.append(f"hệ số {mult:g} suy từ cột tổng ({total:,.2f} ÷ {row.kg:,.2f})")
            elif round(m2) >= 1 and abs(m2 - round(m2)) <= ROW_REL_TOL * m2:
                # cột khối lượng làm tròn 0,1 kg nên tổng ÷ khối lượng lệch số nguyên; dùng dài × SL × kg/m
                mult, use_total = float(round(m2)), True
                if mult != 1:
                    row.notes.append(f"hệ số {mult:g} suy từ cột tổng ({total:,.2f} ÷ dài × SL × kg/m {calc:,.2f})"
                                     f" — ô số cấu kiện gộp/không đọc được")
        row.multiplier = mult or 1.0
        row.total_kg = round(row.kg * row.multiplier, 2)
        if total and (use_total or abs(total - row.total_kg) <= max(ROW_ABS_TOL, ROW_REL_TOL * total)):
            row.total_kg = round(total, 2)          # số ghi ở cột tổng (khối lượng 1 cấu kiện làm tròn 0,1 kg)
        if not mark or mark.isdigit():
            row.issues.append(f"không đọc được ký hiệu thanh (STT {stt or '?'})")
        if design_zero:
            row.ban_ve_ghi_0 = True
            row.notes.append(f"BẢN VẼ ghi khối lượng 0 dù dài × SL × kg/m = "
                             f"{row.length_mm / 1000 * row.quantity * row.unit_kg_m:,.2f} kg — sai trên bản vẽ")
        elif not _row_ok(*vals):
            calc = vals[0] / 1000 * vals[1] * (vals[2] or 0)
            row.issues.append(f"dài×SL×kg/m = {calc:,.2f} ≠ khối lượng ghi {vals[3] or 0:,.2f}")
        if total is not None and abs(total - row.total_kg) > max(ROW_ABS_TOL, ROW_REL_TOL * total):
            row.issues.append(f"khối lượng × hệ số = {row.total_kg:,.2f} ≠ tổng ghi {total:,.2f}")
        out.append(row)
        if notes:
            note = "; ".join(notes)
            fixes.append((row, [(kind, val, note + "; " + n) for kind, val, n in fix] if fix
                          else [("NOTE", None, note)], True))
        elif fix:
            fixes.append((row, fix, False))
    out = _apply_fixes(out, fixes, stated)
    t = DrawingTable(page, title, out, stated)
    if stated is None:
        t.pending_fixes = fixes
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
        for seg in _segments(rows, cols):
            t = _parse_block(page, table.get("title") or "", header, [rows[i] for i in seg],
                             [values[i] if i < len(values) else [] for i in seg], cols)
            if t is not None:
                out.append(t)
    return out


def _segments(rows, cols: range) -> List[List[int]]:
    """
    Tách các bảng xếp chồng trong cùng khối cột (vd. tr.112: bản mặt cầu ở trên, liên tục nhiệt ở dưới): mỗi
    bảng kết thúc ở dòng tổng khối lượng thép của nó; các dòng phía sau dòng tổng thuộc bảng kế tiếp (đoạn không có
    dòng thép nào thì bị bỏ qua).
    """
    segs, cur = [], []
    for i, r in enumerate(rows):
        text = " ".join(norm(r[c]) for c in cols if c < len(r))
        cur.append(i)
        if "TONG" in text and "THEP" in text and not _is_group_subtotal(text):
            segs.append(cur)
            cur = []
    segs.append(cur)
    return [s for s in segs if s]


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
    pending = [f for t in tables for f in t.pending_fixes]
    merged_rows = _apply_fixes([r for t in tables for r in t.rows], pending, main.stated_total_kg)
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


def zero_weight_rows_kg(t: DrawingTable) -> float:
    """Khối lượng (dài × SL × kg/m × hệ số) của các dòng mà chính bản vẽ ghi 0 kg (vd. tr.40: W6, W7, W9)."""
    return round(sum(r.length_mm / 1000 * r.quantity * r.unit_kg_m * r.multiplier
                     for r in t.rows if r.ban_ve_ghi_0), 2)


def _plain_number(cell) -> Optional[float]:
    """Số ghi bằng dấu chấm thập phân ('2.604' m) — không dùng cột 'values' chuẩn hóa kiểu VN."""
    t = str(cell or "").strip().replace(" ", "")
    return float(t) if re.fullmatch(r"\d+(\.\d+)?", t) else None


def hanging_bar_rows(page: dict, elevation_page: Optional[dict] = None) -> List[DrawingRow]:
    """
    Bảng phụ "ỐNG SIÊU ÂM VÀ THANH THÉP TREO LỒNG P11-D16 (TÍNH CHO 01 CỌC)" trên trang cọc (tr.64/66/68/70):
    thép treo lồng không nằm trong bảng thép chính nên tổng thép của trang không gồm nó. Ký hiệu, Ø và kg/m lấy
    từ tiêu đề cột; chiều dài (m), số thanh, khối lượng từ dòng số liệu. Ô trống: số thanh suy từ khối lượng,
    hoặc đọc '4-P11-D16' trên trang mặt đứng cọc (elevation_page); khối lượng = dài × SL × kg/m. Có ghi chú để soát.
    """
    for t in page.get("tables") or []:
        header = [norm(h) for h in t.get("header") or []]
        if not any("TREO LONG" in h for h in header):
            continue
        joined = " ".join(header)
        m = re.search(r"(P\d+)\s*-\s*D(\d+)", joined)
        u = re.search(r"(\d+\.\d+)\s*KG/?M", joined.replace(" ", ""))
        if not m or not u:
            continue
        mark, d, unit = m.group(1), int(m.group(2)), float(u.group(1))
        cols = lambda word: [i for i, h in enumerate(header) if "TREO LONG" in h and word in h]
        for row in t.get("rows") or []:
            first = lambda idx: next((v for v in (_plain_number(row[i]) for i in idx if i < len(row)) if v), None)
            length_m, qty, kg = first(cols("CHIEU DAI")), first(cols("SO THANH")), first(cols("KHOI LUONG"))
            if not length_m:
                continue
            notes = []
            if qty is None and kg:
                qty = float(round(kg / (length_m * unit)))
                notes.append(f"số thanh trống trên OCR, suy = khối lượng ÷ (dài × kg/m) = {qty:g}")
            if qty is None and elevation_page is not None:
                e = re.search(rf"(\d+)\s*-\s*{mark}\s*-\s*D{d}", elevation_page.get("text") or "")
                if e:
                    qty = float(e.group(1))
                    notes.append(f"số thanh trống trên OCR, lấy '{e.group(0)}' ở mặt đứng cọc tr.{elevation_page['page']}")
            if qty is None:
                continue
            calc = round(length_m * qty * unit, 2)
            if kg is None:
                kg = calc
                notes.append(f"khối lượng trống trên OCR, = dài × SL × kg/m = {kg:,.2f}")
            r = DrawingRow(page["page"], None, mark, d, length_m * 1000, qty, unit, kg, 1.0, kg, notes=notes)
            if not _row_ok(r.length_mm, qty, unit, kg):
                r.issues.append(f"dài×SL×kg/m = {calc:,.2f} ≠ khối lượng ghi {kg:,.2f}")
            r.notes.insert(0, "thép treo lồng — bảng phụ, không nằm trong tổng thép chính của trang")
            return [r]
    return []


def with_extra_rows(t: DrawingTable, extra: List[DrawingRow]) -> DrawingTable:
    """Bảng thép của trang cộng thêm các dòng bảng phụ (vd. thép treo lồng cọc): tổng = tổng ghi + Σ dòng phụ."""
    if not extra or t.stated_total_kg is None:
        return t
    out = DrawingTable(t.page, t.title, list(t.rows) + list(extra),
                       round(t.stated_total_kg + sum(r.total_kg for r in extra), 2), list(t.issues))
    return out
