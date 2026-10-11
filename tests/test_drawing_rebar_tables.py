# -*- coding: utf-8 -*-
"""
Test bộ đọc bảng thống kê thép trên bản vẽ (tools/drawing_rebar_tables.py).

Mỗi test ứng với một kiểu lỗi OCR đã gặp trong HSTK Cầu Km19 (tiêu đề sai dấu, cột lệch, bảng ghép, cột số
lượng đứng trước cột dài...). Nguyên tắc: chỉ dùng số liệu khi phép tính của chính bản vẽ đúng.
"""

import json
import os
import re
import unittest

from tools.drawing_rebar_tables import (good_rows, page_rebar_table, page_tables, parse_table, partial_gap_kg,
                                        table_usable)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DU_LIEU = os.path.join(ROOT, "examples", "HO_SO_CAU_KM19_529", "05_DU_LIEU_GOC_SCAN_MARKER", "du_lieu.json")


def table(header, rows, title="BẢNG THỐNG KÊ"):
    def val(c):
        c = str(c).strip()
        return float(c) if re.fullmatch(r"-?\d+(\.\d+)?", c) else None
    return {"title": title, "header": header, "rows": rows, "values": [[val(c) for c in r] for r in rows]}


class ColumnMappingTest(unittest.TestCase):
    def test_quantity_before_length_is_not_swapped(self):
        """Bảng cọc: cột số lượng đứng trước cột dài. dài×SL giao hoán nên phải dựa vào tiêu đề để không gán ngược."""
        t = table(["CẤU KIỆN", "TÊN THANH", "ĐƯỜNG KÍNH", "SỐ LƯỢNG THANH", "CHIỀU DÀI THANH MM", "KHỐI LƯỢNG Đ.VỊ KG/M",
                   "TỔNG KHỐI LƯỢNG KG"],
                  [["", "P1", "D10", "1", "96604", "0.617", "59.60"],
                   ["", "P2", "D25", "28", "12000", "3.850", "1293.60"],
                   ["", "", "", "", "", "TỔNG KHỐI LƯỢNG THÉP (KG)", "1353.20"]])
        (t1,) = parse_table(1, t)
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual([(r.length_mm, r.quantity) for r in t1.rows], [(96604, 1), (12000, 28)])

    def test_per_unit_column_not_taken_as_multiplier(self):
        """Mố M1 tr.31: '1 ĐƠN NGUYÊN KG' là khối lượng, không phải số cấu kiện (lỗi cũ: hệ số = 892,43)."""
        t = table(["LIS", "TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "SỐ LƯỢNG", "KHÓI LƯỢNG ĐƠN KHỐI LƯỢNG CHO VỊ KG/M",
                   "1 ĐƠN NGUYÊN KG", "CẤU KIỆN", "TÒNG KHÓI LƯỢNG KG"],
                  [["1", "M1", "D28", "14213", "13", "4.830", "892.43", "", "892.43"],
                   ["2", "A1", "D16", "7857", "166", "1.580", "2060.73", "", "2060.73"],
                   ["", "", "TỔNG HỢP KHÓI LƯỢNG CỐT THÉP", "", "", "(KG)", "2953.16", "", "2953.16"]])
        (t1,) = parse_table(31, t)
        self.assertEqual([r.multiplier for r in t1.rows], [1.0, 1.0])
        self.assertAlmostEqual(t1.rows_total_kg, 2953.16, places=2)
        self.assertTrue(table_usable(t1))

    def test_multiplier_derived_from_total_column(self):
        """Gờ lan can tr.119: cột số đốt trống nhưng cột tổng có → hệ số = tổng ÷ khối lượng (số nguyên)."""
        t = table(["CẤU KIỆN", "TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI THANH MM", "SỐ LƯỢNG THANH", "KHỐI LƯỢNG ĐƠN VỊ KG/M",
                   "KHỐI LƯỢNG KG", "SỐ LƯỢNG ĐỐT", "TỔNG KHỐI LƯỢNG KG"],
                  [["", "G1", "D14", "5910", "19", "1.210", "135.87", "", "2581.53"],
                   ["", "G2", "D18", "3265", "4", "2.000", "26.12", "", "26.12"],
                   ["", "", "TỔNG HỢP KHỐI LƯỢNG CỐT THÉP", "", "", "(KG)", "", "", "2607.65"]])
        (t1,) = parse_table(119, t)
        self.assertEqual([r.multiplier for r in t1.rows], [19.0, 1.0])
        self.assertTrue(t1.rows[0].notes and "suy từ cột tổng" in t1.rows[0].notes[0])
        self.assertTrue(table_usable(t1))

    def test_group_subtotal_is_not_the_table_total(self):
        t = table(["TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "SỐ LƯỢNG", "KG/M", "KHỐI LƯỢNG KG"],
                  [["P1", "D10", "1000", "10", "0.617", "6.17"],
                   ["P2", "D16", "2000", "4", "1.580", "12.64"],
                   ["", "TONGKHOILUONGTHEP (KG)", "", "10<D<=18", "", "12.64"],
                   ["", "TỔNG KHỐI LƯỢNG THÉP (KG)", "", "", "", "18.81"]])
        (t1,) = parse_table(64, t)
        self.assertEqual(t1.stated_total_kg, 18.81)
        self.assertTrue(table_usable(t1))

    def test_plain_numbers_count_as_diameter_only_under_diameter_header(self):
        """Mố M2 tr.40: cột số lượng có 16, 28... không được coi là cột đường kính (lỗi cũ: tách nhầm bảng)."""
        t = table(["STT", "TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "SỐ LƯỢNG", "KG/M", "KHỐI LƯỢNG KG"],
                  [["1", "F1", "D28", "6578", "16", "4.830", "508.35"],
                   ["2", "F2", "D18", "13571", "28", "2.000", "760.00"],
                   ["", "TỔNG HỢP KHỐI LƯỢNG CỐT THÉP", "", "", "", "", "1268.35"]])
        tables = parse_table(40, t)
        self.assertEqual(len(tables), 1)
        self.assertEqual([r.diameter_mm for r in tables[0].rows], [28, 18])

    def test_side_by_side_tables_are_split(self):
        """Tr.112: bản mặt cầu và bảng thứ hai nằm cạnh nhau trong cùng một bảng OCR."""
        t = table(["", "ĐƯỜNG KINH (MM)", "SỐ THANH", "CHIỀU DÀI 1 THANH (MM)", "KG/M", "KHỐI LƯỢNG (KG)",
                   "", "ĐƯỜNG KINH (MM)", "SỐ THANH", "CHIỀU DÀI 1 THANH (MM)", "KG/M", "KHỐI LƯỢNG (KG)"],
                  [["S1", "D14", "84", "39850", "1.210", "4050.35", "L1", "D18", "10", "2000", "2.000", "40.00"],
                   ["S2", "D18", "490", "13135", "2.000", "12872.30", "L2", "D10", "20", "1000", "0.617", "12.34"]])
        tables = parse_table(112, t)
        self.assertEqual(len(tables), 2)
        self.assertEqual([r.mark for r in tables[0].rows], ["S1", "S2"])
        self.assertEqual([r.mark for r in tables[1].rows], ["L1", "L2"])

    def test_row_arithmetic_error_is_flagged(self):
        t = table(["TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "SỐ LƯỢNG", "KG/M", "KHỐI LƯỢNG KG"],
                  [["A", "D16", "1000", "10", "1.580", "15.80"],
                   ["B", "D16", "2000", "10", "1.580", "31.60"],
                   ["C", "D16", "1000", "10", "1.580", "99.00"],
                   ["", "TỔNG HỢP KHỐI LƯỢNG CỐT THÉP", "", "", "", "146.40"]])
        (t1,) = parse_table(1, t)
        self.assertFalse(table_usable(t1))
        self.assertEqual([r.mark for r in good_rows(t1)], ["A", "B"])


class OcrRepairTest(unittest.TestCase):
    """Lỗi OCR một ô được sửa bằng phép tính của chính bản vẽ — chỉ khi Σ dòng sau sửa khớp tổng ghi trên bản vẽ."""
    HDR = ["TÊN THANH", "ĐƯỜNG KÍNH", "SỐ THANH", "CHIỀU DÀI MM", "KG/M", "KHỐI LƯỢNG KG"]

    def _t(self, rows, total):
        return table(self.HDR, [["A", "D16", "10", "1000", "1.580", "15.80"], ["B", "D16", "10", "2000", "1.580", "31.60"]]
                     + rows + [["TỔNG KHỐI LƯỢNG THÉP", "", "", "", "", str(total)]])

    def test_misread_quantity_fixed_when_total_matches(self):
        """Dầm S1B1 tr.88: OCR đọc G6a SL = 6, khối lượng ghi 664,44 = 11,7 × 9 × 6,31."""
        (t1,) = parse_table(88, self._t([["G6a", "D32", "6", "11700", "6.310", "664.44"]], 711.84))
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual(t1.rows[-1].quantity, 9)
        self.assertIn("Σ bảng sau sửa khớp", t1.rows[-1].notes[0])

    def test_blank_quantity_derived_when_total_matches(self):
        (t1,) = parse_table(94, self._t([["G6b", "D32", "", "6860", "6.310", "129.86"]], 177.26))
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual([r.quantity for r in t1.rows], [10, 10, 3])

    def test_misread_weight_fixed_when_total_matches(self):
        """Dầm S1B2 tr.91: 'Thanh neo' 8 × 110 mm × 2,47 = 2,17 kg, OCR đọc 217."""
        (t1,) = parse_table(91, self._t([["N", "D20", "8", "110", "2.470", "217"]], 49.57))
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual(t1.rows[-1].kg, 2.17)

    def test_no_repair_when_total_does_not_confirm_it(self):
        (t1,) = parse_table(1, self._t([["G6a", "D32", "6", "11700", "6.310", "664.44"]], 900.00))
        self.assertFalse(table_usable(t1))
        self.assertEqual(t1.rows[-1].quantity, 6)
        (t2,) = parse_table(1, self._t([["G6b", "D32", "", "6860", "6.310", "129.86"]], 900.00))
        self.assertEqual([r.mark for r in t2.rows], ["A", "B"])          # dòng phải suy SL bị bỏ


class OcrLayoutTest(unittest.TestCase):
    """Lỗi bố cục OCR gặp ở Mố M1/M2 (tr.31, 40), gờ lan can (tr.119), liên tục nhiệt (tr.112)."""
    HDR = ["STT", "TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "", "SỐ LƯỢNG", "KG/M", "KHỐI LƯỢNG KG", "TỔNG KHỐI LƯỢNG KG"]

    def _t(self, rows, total):
        base = [["1", "A", "D16", "1000", "", "10", "1.580", "15.80", "15.80"],
                ["2", "B", "D8", "2000", "", "10", "0.395", "7.90", "7.90"]]
        return table(self.HDR, base + rows + [["", "TỔNG HỢP KHỐI LƯỢNG CỐT THÉP", "", "", "", "", "", "", str(total)]])

    def test_plain_round_bar_is_read(self):
        (t1,) = parse_table(31, self._t([["3", "H1", "R25", "1070", "", "23", "3.850", "94.75", "94.75"]], 118.45))
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual((t1.rows[-1].diameter_mm, t1.rows[-1].tron_tron), (25, True))

    def test_unreadable_diameter_from_unit_weight(self):
        """'DB' (OCR của D8) → Ø8 vì kg/m 0,395; chỉ nhận khi Σ bảng khớp tổng."""
        (t1,) = parse_table(31, self._t([["3", "B1", "DB", "640", "", "50", "0.395", "12.64", "12.64"]], 36.34))
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual(t1.rows[-1].diameter_mm, 8)
        (t2,) = parse_table(31, self._t([["3", "B1", "DB", "640", "", "50", "0.395", "12.64", "12.64"]], 99.0))
        self.assertEqual([r.mark for r in t2.rows], ["A", "B"])

    def test_length_split_into_two_cells(self):
        """Tr.40: '1640' bị tách thành '164' và '0'; ô kg/m và SL hỏng — suy từ cột tổng, kg/m của Ø16 trong bảng."""
        (t1,) = parse_table(40, self._t([["3", "T1", "D16", "189", "0", "CONG", "INH", "TU", "20.90"]], 44.60))
        self.assertTrue(table_usable(t1), t1.issues)
        r = t1.rows[-1]
        self.assertEqual((r.length_mm, r.quantity, r.unit_kg_m), (1890, 7, 1.58))

    def test_zero_weight_written_on_drawing_is_flagged_not_dropped(self):
        """Tr.40: W6/W7/W9 có dài và SL nhưng bản vẽ ghi 0 kg ở cả hai cột — sai của bản vẽ, không phải OCR."""
        (t1,) = parse_table(40, self._t([["3", "W6", "D16", "615", "", "53", "1.580", "0.0", "0.00"]], 23.70))
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertTrue(t1.rows[-1].ban_ve_ghi_0)
        self.assertIn("BẢN VẼ ghi khối lượng 0", t1.rows[-1].notes[0])

    def test_merged_multiplier_cell_with_rounded_weights(self):
        """Tr.119: 'số đốt' 19 là ô gộp; khối lượng làm tròn 0,1 kg nên tổng ÷ khối lượng không ra số nguyên."""
        hdr = ["TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "SỐ LƯỢNG", "KG/M", "KHỐI LƯỢNG KG", "SỐ LƯỢNG ĐỐT",
               "TỔNG KHỐI LƯỢNG KG"]
        t = table(hdr, [["G3", "D14", "2344", "41", "1.210", "116.3", "19.0", "2209.4"],
                        ["G4B", "D12", "388", "13", "0.888", "4.5", "", "85.1"],
                        ["G4A", "D12", "342", "13", "0.888", "", "", "75.0"],
                        ["", "TỔNG HỢP KHỐI LƯỢNG CỐT THÉP", "", "", "", "", "", "2369.5"]])
        (t1,) = parse_table(119, t)
        self.assertTrue(table_usable(t1), t1.issues)
        self.assertEqual([r.multiplier for r in t1.rows], [19, 19, 19])
        self.assertEqual([r.total_kg for r in t1.rows], [2209.4, 85.1, 75.0])

    def test_stacked_tables_are_split_at_their_totals(self):
        """Tr.112: bảng bản mặt cầu ở trên, bảng liên tục nhiệt ở dưới trong cùng khối cột."""
        hdr = ["TÊN THANH", "ĐƯỜNG KÍNH", "SỐ THANH", "CHIỀU DÀI MM", "KG/M", "KHỐI LƯỢNG KG"]
        t = table(hdr, [["S1", "D14", "84", "1000", "1.210", "101.64"], ["S2", "D18", "10", "1000", "2.000", "20.00"],
                        ["", "TỔNG KHỐI LƯỢNG THÉP", "", "", "", "121.64"],
                        ["", "KHỐI LƯỢNG BÊ TÔNG 35MPA", "", "", "", "102.53"],
                        ["L1", "D1 8", "54", "1000", "2.000", "108.00"], ["L2", "D25", "10", "1000", "3.850", "38.50"],
                        ["L3", "D10", "10", "1000", "0.617", "6.17"],
                        ["", "TỔNG KHỐI LƯỢNG THÉP", "", "", "", "152.67"]])
        tables = parse_table(112, t)
        self.assertEqual([x.stated_total_kg for x in tables], [121.64, 152.67])
        self.assertEqual([r.mark for r in tables[1].rows], ["L1", "L2", "L3"])
        self.assertTrue(all(table_usable(x) for x in tables))


class PartialTableTest(unittest.TestCase):
    def _t(self, stated):
        return table(["TÊN THANH", "ĐƯỜNG KÍNH", "CHIỀU DÀI MM", "SỐ LƯỢNG", "KG/M", "KHỐI LƯỢNG KG"],
                     [["A", "D16", "1000", "10", "1.580", "15.80"], ["B", "D16", "2000", "10", "1.580", "31.60"],
                      ["", "TỔNG HỢP KHỐI LƯỢNG CỐT THÉP", "", "", "", str(stated)]])

    def test_small_gap_is_reported(self):
        (t1,) = parse_table(1, self._t(48.40))               # thiếu 1,00 kg ≈ 2%
        self.assertFalse(table_usable(t1))
        self.assertAlmostEqual(partial_gap_kg(t1), 1.0, places=2)

    def test_rows_above_total_or_large_gap_not_usable(self):
        (over,) = parse_table(1, self._t(40.00))
        self.assertIsNone(partial_gap_kg(over))
        (big,) = parse_table(1, self._t(60.00))              # thiếu > 5%
        self.assertIsNone(partial_gap_kg(big))


@unittest.skipUnless(os.path.isfile(DU_LIEU), "thiếu dữ liệu OCR HSTK Km19")
class RealDossierPagesTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        with open(DU_LIEU, encoding="utf-8") as f:
            cls.pages = {p["page"]: p for p in json.load(f)["pages"]}

    def test_verified_pages_match_drawing_totals(self):
        expected = {44: 16177.83, 48: 32558.52, 55: 16177.83, 59: 35186.85, 64: 2507.53, 66: 4705.88,
                    68: 4828.66, 70: 3671.07, 96: 2763.12, 126: 3964.44}
        for page, total in expected.items():
            t = page_rebar_table(self.pages[page])
            self.assertTrue(table_usable(t), f"tr.{page}: {t.issues}")
            self.assertAlmostEqual(t.rows_total_kg, total, delta=0.05, msg=f"tr.{page}")

    def test_super_t_beam_pages_repaired_to_drawing_totals(self):
        for page, total in {88: 6071.25, 91: 6065.28, 94: 5920.08}.items():
            t = page_rebar_table(self.pages[page])
            self.assertTrue(table_usable(t), f"tr.{page}: {t.issues}")
            self.assertAlmostEqual(t.rows_total_kg, total, delta=0.05, msg=f"tr.{page}")

    def test_abutments_and_parapet_read_completely(self):
        """Trước đây thiếu 126,74 / 192,83 / 458,42 kg; nay mọi dòng đọc được và khớp tổng bản vẽ."""
        for page, total in {31: 7958.56, 40: 19648.73, 119: 20069.02}.items():
            t = page_rebar_table(self.pages[page])
            self.assertTrue(table_usable(t), f"tr.{page}: {t.issues}")
            self.assertAlmostEqual(t.rows_total_kg, total, delta=0.5, msg=f"tr.{page}")
        zero = [r.mark for r in page_rebar_table(self.pages[40]).rows if r.ban_ve_ghi_0]
        self.assertEqual(zero, ["W6(T)", "W7(T)", "W9(T)"])

    def test_pile_hanging_bars_p11_side_table(self):
        """Thép treo lồng P11-D16 (4 thanh/cọc) ở bảng phụ, không trong tổng thép chính. Có P11, tổng 4 loại
        cọc = 108.884,23 kg = bảng tổng hợp tr.175 (108,884 t); thiếu P11 thì lệch 0,42 t."""
        from tools.drawing_rebar_tables import hanging_bar_rows
        total = 0.0
        for page, elev, n in ((64, 63, 3), (66, 65, 7), (68, 67, 8), (70, 69, 8)):
            (r,) = hanging_bar_rows(self.pages[page], self.pages[elev])
            self.assertEqual((r.mark, r.diameter_mm, r.quantity), ("P11", 16, 4), f"tr.{page}")
            self.assertTrue(r.ok, r.issues)
            total += (page_rebar_table(self.pages[page]).stated_total_kg + r.total_kg) * n
        self.assertAlmostEqual(total, 108884.23, places=1)
        (t2,) = hanging_bar_rows(self.pages[70], self.pages[69])
        self.assertTrue(any("tr.69" in n for n in t2.notes))         # ô trống: lấy '4-P11-D16' ở mặt đứng

    def test_link_slab_table_on_page_112(self):
        from tools.drawing_rebar_tables import page_tables
        ltn = [t for t in page_tables(self.pages[112]) if any(r.mark == "L2" for r in t.rows)]
        self.assertEqual(len(ltn), 1)
        self.assertEqual(ltn[0].stated_total_kg, 8917.94)
        self.assertTrue(table_usable(ltn[0]), ltn[0].issues)
        self.assertEqual(sorted(r.mark for r in ltn[0].rows), ["L1", "L2", "L3", "S4"])

    def test_pile_rows_have_length_not_quantity(self):
        rows = page_tables(self.pages[64])[0].rows
        self.assertTrue(all(r.length_mm >= 100 for r in rows), [(r.mark, r.length_mm) for r in rows])


if __name__ == "__main__":
    unittest.main()
