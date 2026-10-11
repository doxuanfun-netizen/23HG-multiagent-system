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

    def test_abutment_m1_has_known_gap(self):
        t = page_rebar_table(self.pages[31])
        self.assertAlmostEqual(t.stated_total_kg, 7958.56, places=2)
        self.assertAlmostEqual(partial_gap_kg(t), 126.74, places=2)

    def test_pile_rows_have_length_not_quantity(self):
        rows = page_tables(self.pages[64])[0].rows
        self.assertTrue(all(r.length_mm >= 100 for r in rows), [(r.mark, r.length_mm) for r in rows])


if __name__ == "__main__":
    unittest.main()
