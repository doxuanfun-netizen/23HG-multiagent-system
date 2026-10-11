# -*- coding: utf-8 -*-
"""
Test bộ cắt thép theo từng Ø (examples/generate_rebarcut_dedicated_package.py).

Dùng một BBS nhỏ (3 loại thép) để chạy nhanh; bộ thật (141.590 thanh) mất ~30 giây. Kiểm tra:
  - mỗi Ø có một file RebarCut + một lệnh cắt CNC, có Master, bảng tổng hợp, cáp DƯL, hướng dẫn;
  - số cây và số đoạn cắt khớp tính tay; chạy lại cho kết quả giống hệt;
  - truyền out_dir thì KHÔNG ghi ra ngoài (không đụng thư mục dự án);
  - thư mục gói vi mô chưa tồn tại không làm script chết (lỗi cũ: FileNotFoundError ở bước cuối);
  - các số trong hướng dẫn vận hành và bảng cáp DƯL được tính từ dữ liệu, không gõ cứng.
"""

import csv
import glob
import importlib.util
import os
import tempfile
import unittest

import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "examples", "generate_rebarcut_dedicated_package.py")

# Ø10: 3 thanh 3800 mm → 1 cây (3×3800 = 11400 ≤ 11700 − lưỡi cắt); Ø12: 4 thanh 5000 mm → 2 cây (2 thanh/cây);
# Ø32: 2 thanh 11000 mm → 2 cây (mỗi cây 1 thanh).
BBS = ("mark,diameter_mm,grade,length_mm,quantity\n"
       "A1,10,CB300-T,3800,3\n"
       "B1,12,CB400-V,5000,4\n"
       "C1,32,CB500-V,11000,2\n")
EXPECTED_BARS = {10: 1, 12: 2, 32: 2}
EXPECTED_PIECES = {10: 3, 12: 4, 32: 2}


def load_module():
    spec = importlib.util.spec_from_file_location("gen_rebarcut_dedicated", SCRIPT)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class DedicatedPackageTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.mod = load_module()
        cls.bbs = os.path.join(cls.tmp.name, "bbs.csv")
        with open(cls.bbs, "w", encoding="utf-8") as f:
            f.write(BBS)
        cls.out = os.path.join(cls.tmp.name, "out")
        cls.decoy = os.path.join(cls.tmp.name, "du_an")          # "thư mục dự án" mà script không được đụng tới
        os.makedirs(cls.decoy)
        cls.mod.PROJECT_DIR = cls.decoy
        cls.result = cls.mod.run(out_dir=cls.out, bbs_path=cls.bbs, bbs_sheet=None, quiet=True)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_structure_one_file_per_diameter(self):
        dias = sorted(os.listdir(os.path.join(self.out, "THEO_TUNG_DUONG_KINH_PHI")))
        self.assertEqual(len(dias), 4)                                       # 3 Ø + cáp DƯL
        for dia in EXPECTED_BARS:
            self.assertTrue(any(f"_D{dia:02d}_" in f for f in dias), f"thiếu file Ø{dia}: {dias}")
        self.assertEqual(len(glob.glob(os.path.join(self.out, "LENH_CAT_CNC_CSV", "*.csv"))), 3)
        for name in ("00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx", "00_RebarCut_MASTER_TOAN_CAU_11M7.xlsx",
                     "README_QUY_TRINH_VAN_HANH_BAI_THEP.md"):
            self.assertTrue(os.path.isfile(os.path.join(self.out, name)), name)

    def test_bar_counts_match_hand_calculation(self):
        self.assertEqual(self.result["bars_by_dia"], EXPECTED_BARS)
        self.assertEqual(self.result["total_bars_master"], sum(EXPECTED_BARS.values()))
        self.assertEqual(self.result["dias"], [10, 12, 32])

    def test_cnc_csv_has_every_piece_with_right_total_length(self):
        for path in glob.glob(os.path.join(self.out, "LENH_CAT_CNC_CSV", "*.csv")):
            with open(path, encoding="utf-8-sig", newline="") as f:
                rows = list(csv.reader(f))[1:]
            dia = int(rows[0][1])
            self.assertEqual(len(rows), EXPECTED_PIECES[dia], os.path.basename(path))
            self.assertEqual({len({r[0] for r in rows})}, {EXPECTED_BARS[dia]})   # số cây khác nhau = số cây cần mua
            piece_mm = {10: 3800, 12: 5000, 32: 11000}[dia]
            self.assertAlmostEqual(sum(float(r[3]) for r in rows), piece_mm * EXPECTED_PIECES[dia] / 1000, places=6)

    def test_rerun_is_identical(self):
        again = os.path.join(self.tmp.name, "again")
        self.mod.run(out_dir=again, bbs_path=self.bbs, bbs_sheet=None, quiet=True)
        for path in glob.glob(os.path.join(self.out, "LENH_CAT_CNC_CSV", "*.csv")):
            other = os.path.join(again, "LENH_CAT_CNC_CSV", os.path.basename(path))
            with open(path, "rb") as a, open(other, "rb") as b:
                self.assertEqual(a.read(), b.read(), os.path.basename(path))

    def test_out_dir_given_writes_nothing_outside(self):
        self.assertEqual(os.listdir(self.decoy), [])                           # không sync sang gói vi mô

    def test_operating_guide_numbers_come_from_data(self):
        with open(os.path.join(self.out, "README_QUY_TRINH_VAN_HANH_BAI_THEP.md"), encoding="utf-8") as f:
            text = f.read()
        self.assertIn("toàn bộ 5 cây thép 11.7m", text)
        self.assertIn("3 thanh Ø10", text)
        self.assertIn("4 thanh Ø12", text)
        self.assertIn("3 loại đường kính", text)
        self.assertIn("nhỏ nhất (Ø10) đến lớn nhất (Ø32)", text)
        for literal in ("33.212", "12.278", "76.014", "35.366"):               # số của bộ thật, không được gõ cứng
            self.assertNotIn(literal, text)

    def test_pt_strand_derived_numbers(self):
        path = os.path.join(self.out, "THEO_TUNG_DUONG_KINH_PHI", "12_Cap_Du_Ung_Luc_15.2mm_SuperT.xlsx")
        row = [c.value for c in openpyxl.load_workbook(path).active[7]][1:]
        length, per_beam, beams, strands, total_len, unit_w, weight = row[6:13]
        self.assertEqual(strands, per_beam * beams)
        self.assertAlmostEqual(total_len, length * strands, places=3)
        self.assertAlmostEqual(weight, round(total_len * unit_w, 1), places=1)
        self.assertEqual((strands, round(total_len), weight), (660, 25212, 27783.6))      # đúng số cũ của hồ sơ


def dashboard_rows(path):
    """{nhãn Ø: [giá trị cột B..]} và các dòng còn lại của bảng tổng hợp."""
    ws = openpyxl.load_workbook(path).active
    return [[c.value for c in r][1:] for r in ws.iter_rows(min_row=6)]


class EveryBarPlannedTest(unittest.TestCase):
    """Bộ BBS nhỏ không có thanh dài: không được báo thiếu, và mua đủ chiều dài cho mọi Ø."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        mod = load_module()
        bbs = os.path.join(cls.tmp.name, "bbs.csv")
        with open(bbs, "w", encoding="utf-8") as f:
            f.write(BBS)
        cls.out = os.path.join(cls.tmp.name, "out")
        cls.result = mod.run(out_dir=cls.out, bbs_path=bbs, bbs_sheet=None, quiet=True)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_no_unplanned_and_purchase_covers_demand(self):
        self.assertEqual(self.result["unplanned_by_dia"], {})
        for r in self.result["summary_rows"]:
            self.assertGreaterEqual(r["bars_11m7"] * 11.7 + 1e-9, r["total_len_m"], f"Ø{r['dia']}")
            self.assertFalse(r["status"].startswith("CHƯA ĐỦ"))

    def test_total_row_says_done_and_cable_row_is_derived(self):
        rows = dashboard_rows(os.path.join(self.out, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx"))
        total = next(r for r in rows if r[0] == "TỔNG CỘNG")
        self.assertEqual(total[14], "HOÀN TẤT")
        cable = next(r for r in rows if r[1] == "Ø15.2 (Cáp DƯL)")
        self.assertEqual(cable[3], 660)


# Ø25 dài 14 m > cây 11.7 m và không có vùng cho phép nối → solver không lập kế hoạch được.
LONG_BBS = BBS + "D1,25,CB400-V,14000,3\n"


class LongBarNotSilentlyDroppedTest(unittest.TestCase):
    """Lỗi cũ: thanh dài hơn cây bị solver đưa vào `unplanned`, nhưng bảng tổng hợp bỏ qua — số cây cần mua thiếu
    mà vẫn ghi OPTIMAL. Hồ sơ Km19 thiếu 3.546 thanh (~144 tấn) theo cách này. Chạy với noi_trong_chieu_dai=False
    để giữ trường hợp không tách thanh dài: khi đó chúng phải hiện ra, không được biến mất."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        mod = load_module()
        bbs = os.path.join(cls.tmp.name, "bbs.csv")
        with open(bbs, "w", encoding="utf-8") as f:
            f.write(LONG_BBS)
        cls.out = os.path.join(cls.tmp.name, "out")
        cls.result = mod.run(out_dir=cls.out, bbs_path=bbs, bbs_sheet=None, quiet=True, noi_trong_chieu_dai=False)
        cls.row25 = next(r for r in cls.result["summary_rows"] if r["dia"] == 25)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_summary_row_reports_unplanned(self):
        self.assertEqual(self.result["unplanned_by_dia"], {25: 3})
        self.assertEqual(self.row25["unplanned_qty"], 3)
        self.assertAlmostEqual(self.row25["unplanned_len_m"], 42.0)
        self.assertGreater(self.row25["unplanned_kg"], 0)
        self.assertTrue(self.row25["status"].startswith("CHƯA ĐỦ"), self.row25["status"])

    def test_purchase_plus_unplanned_covers_demand(self):
        for r in self.result["summary_rows"]:
            self.assertGreaterEqual(r["bars_11m7"] * 11.7 + r["unplanned_len_m"] + 1e-9, r["total_len_m"],
                                    f"Ø{r['dia']}")

    def test_dashboard_shows_unplanned_and_warns(self):
        rows = dashboard_rows(os.path.join(self.out, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx"))
        r25 = next(r for r in rows if r[1] == "Ø25")
        self.assertEqual(r25[11], 3)                                    # cột "Thanh CHƯA lập kế hoạch"
        self.assertTrue(str(r25[14]).startswith("CHƯA ĐỦ"))
        total = next(r for r in rows if r[0] == "TỔNG CỘNG")
        self.assertEqual(total[11], 3)
        self.assertEqual(total[14], "CHƯA ĐỦ")
        self.assertTrue(any(isinstance(r[0], str) and r[0].startswith("CẢNH BÁO: 3 thanh") for r in rows))

    def test_per_diameter_workbook_lists_unplanned_bars(self):
        path = glob.glob(os.path.join(self.out, "THEO_TUNG_DUONG_KINH_PHI", "*_D25_*.xlsx"))[0]
        wb = openpyxl.load_workbook(path)
        self.assertIn("CHUA_LAP_KE_HOACH", wb.sheetnames)
        data = [[c.value for c in r] for r in wb["CHUA_LAP_KE_HOACH"].iter_rows(min_row=5)]
        self.assertEqual([(r[1], r[2], r[5]) for r in data], [("D1", 25, 3)])
        self.assertIn("CÒN 3 thanh", wb["PA_TOI_UU"]["A3"].value)


class LongBarSplitLapIncludedTest(unittest.TestCase):
    """Mặc định (chủ dự án xác nhận 11/10/2026): chiều dài BBS của thanh dài hơn cây ĐÃ GỒM đoạn nối 40d.
    Thanh 14m Ø25 → 2 đoạn (số đoạn tối thiểu), tổng 2 đoạn = 14m, mỗi đoạn ≥ 40d = 1m và ≤ 11.7m."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        bbs = os.path.join(cls.tmp.name, "bbs.csv")
        with open(bbs, "w", encoding="utf-8") as f:
            f.write(LONG_BBS)
        cls.out = os.path.join(cls.tmp.name, "out")
        cls.result = load_module().run(out_dir=cls.out, bbs_path=bbs, bbs_sheet=None, quiet=True)
        cls.row25 = next(r for r in cls.result["summary_rows"] if r["dia"] == 25)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_long_bars_are_planned_with_one_splice_each(self):
        self.assertEqual(self.result["unplanned_by_dia"], {})
        self.assertEqual(self.row25["splices_count"], 3)                  # 3 thanh × (2 đoạn − 1)
        self.assertFalse(self.row25["status"].startswith("CHƯA ĐỦ"))
        self.assertGreaterEqual(self.row25["bars_11m7"] * 11.7 + 1e-9, self.row25["total_len_m"])

    def test_pieces_sum_exactly_to_bbs_length(self):
        path = glob.glob(os.path.join(self.out, "LENH_CAT_CNC_CSV", "*D25*.csv"))[0]
        with open(path, encoding="utf-8-sig", newline="") as f:
            pieces = [float(r[3]) for r in list(csv.reader(f))[1:]]
        self.assertEqual(len(pieces), 6)                                   # 3 thanh × 2 đoạn
        self.assertAlmostEqual(sum(pieces), 3 * 14.0, places=6)            # không cộng thêm đoạn nối
        for p in pieces:
            self.assertGreaterEqual(p, 40 * 25 / 1000)
            self.assertLessEqual(p, 11.7)

    def test_dashboard_states_the_rule(self):
        ws = openpyxl.load_workbook(os.path.join(self.out, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx")).active
        self.assertIn("ĐÃ GỒM trong chiều dài BBS", ws["B5"].value)


class RejectedBbsRowsAreReportedTest(unittest.TestCase):
    """Dòng BBS sai (Ø không phải thép thanh, chiều dài 2mm...) bị bộ đọc loại; trước đây script tạo bộ cắt không báo gì."""

    def test_rejected_rows_listed_in_result_and_dashboard(self):
        with tempfile.TemporaryDirectory() as tmp:
            bbs = os.path.join(tmp, "bbs.csv")
            with open(bbs, "w", encoding="utf-8") as f:
                f.write(BBS + "X1,128,CB400-V,2,525\n")
            out = os.path.join(tmp, "out")
            result = load_module().run(out_dir=out, bbs_path=bbs, bbs_sheet=None, quiet=True)
            self.assertEqual(len(result["bbs_errors"]), 1)
            self.assertIn("Ø128", result["bbs_errors"][0])
            rows = dashboard_rows(os.path.join(out, "00_BANG_TONG_HOP_CAT_THEP_THEO_PHI.xlsx"))
            texts = [r[0] for r in rows if isinstance(r[0], str)]
            self.assertTrue(any(t.startswith("CẢNH BÁO: 1 dòng BBS bị loại") for t in texts), texts)
            self.assertTrue(any("Ø128" in t for t in texts))


class DestinationFolderRegressionTest(unittest.TestCase):
    def test_missing_micro_dossier_folder_no_longer_crashes(self):
        """Trước đây: shutil.copyfile vào thư mục BO_HO_SO_02... chưa tồn tại → FileNotFoundError ở bước cuối."""
        with tempfile.TemporaryDirectory() as tmp:
            mod = load_module()
            bbs = os.path.join(tmp, "bbs.csv")
            with open(bbs, "w", encoding="utf-8") as f:
                f.write(BBS)
            project = os.path.join(tmp, "du_an")                     # thư mục dự án MỚI, chưa có gì bên trong
            mod.PROJECT_DIR = project
            mod.TARGET_FOLDER = os.path.join(project, "01_HE_THONG_CAT_THEP_REBARCUT")
            mod.run(bbs_path=bbs, bbs_sheet=None, quiet=True)        # hành vi mặc định: đồng bộ sang gói vi mô
            dest = os.path.join(project, "BO_HO_SO_02_VI_MO_CHUYEN_SAU_14_BO", "01_To_Hop_Cat_Thep_11m7_RebarCut.xlsx")
            self.assertTrue(os.path.isfile(dest))
            self.assertTrue(os.path.isfile(os.path.join(mod.TARGET_FOLDER, "README_QUY_TRINH_VAN_HANH_BAI_THEP.md")))


if __name__ == "__main__":
    unittest.main()
