# -*- coding: utf-8 -*-
"""Kiểm thử bộ tính CPM bằng các mạng ví dụ đã tính tay (không lấy từ chính code).
Lịch: Thứ Hai–Thứ Bảy làm việc, Chủ nhật nghỉ. 01/01/2024 là Thứ Hai."""
import datetime as dt
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "1_Scripts_TuDongHoa"))
import cpm_engine as c  # noqa: E402

D = dt.date
START = D(2024, 1, 1)  # Thứ Hai


def cal(*holidays):
    return c.Calendar(holidays)


class TestCalendar(unittest.TestCase):
    def test_sunday_is_off_saturday_is_on(self):
        k = cal()
        self.assertFalse(k.is_workday(D(2024, 1, 7)))   # Chủ nhật
        self.assertTrue(k.is_workday(D(2024, 1, 6)))    # Thứ Bảy

    def test_add_skips_sunday_and_holiday(self):
        k = cal(D(2024, 1, 2))
        self.assertEqual(k.add(D(2024, 1, 1), 2), D(2024, 1, 4))   # bỏ Thứ Ba lễ
        self.assertEqual(k.add(D(2024, 1, 6), 1), D(2024, 1, 8))   # bỏ Chủ nhật
        self.assertEqual(k.add(D(2024, 1, 8), -1), D(2024, 1, 6))
        self.assertEqual(k.add(D(2024, 1, 3), 0), D(2024, 1, 3))

    def test_float_days(self):
        k = cal()
        self.assertEqual(k.float_days(D(2024, 1, 4), D(2024, 1, 8)), 3)  # Thứ Năm→Thứ Hai: 3 ngày công
        self.assertEqual(k.float_days(D(2024, 1, 4), D(2024, 1, 4)), 0)


class TestForwardBackward(unittest.TestCase):
    def network(self):
        return [
            c.Task("A", 3),
            c.Task("B", 2, preds=[("A", "FS", 0)]),
            c.Task("C", 5, preds=[("A", "FS", 0)]),
            c.Task("D", 1, preds=[("B", "FS", 0), ("C", "FS", 0)]),
        ]

    def test_hand_calculated_network(self):
        r = c.schedule(self.network(), START, cal())
        # Tính xuôi
        self.assertEqual((r["A"].es, r["A"].ef), (D(2024, 1, 1), D(2024, 1, 3)))
        self.assertEqual((r["B"].es, r["B"].ef), (D(2024, 1, 4), D(2024, 1, 5)))
        self.assertEqual((r["C"].es, r["C"].ef), (D(2024, 1, 4), D(2024, 1, 9)))
        self.assertEqual((r["D"].es, r["D"].ef), (D(2024, 1, 10), D(2024, 1, 10)))
        # Tính ngược và dự trữ
        self.assertEqual((r["B"].ls, r["B"].lf), (D(2024, 1, 8), D(2024, 1, 9)))
        self.assertEqual(r["B"].tf, 3)
        for w in ("A", "C", "D"):
            self.assertEqual(r[w].tf, 0, w)
        self.assertEqual([w for w in "ABCD" if r[w].critical], ["A", "C", "D"])

    def test_holiday_pushes_finish(self):
        r = c.schedule([c.Task("A", 3)], START, cal(D(2024, 1, 2)))
        self.assertEqual(r["A"].ef, D(2024, 1, 4))  # Thứ Hai, Thứ Tư, Thứ Năm

    def test_start_on_sunday_moves_to_monday(self):
        r = c.schedule([c.Task("A", 1)], D(2024, 1, 7), cal())
        self.assertEqual(r["A"].es, D(2024, 1, 8))

    def test_milestone_has_equal_dates(self):
        r = c.schedule([c.Task("A", 2), c.Task("M", 0, "Milestone", [("A", "FS", 0)])], START, cal())
        self.assertEqual(r["M"].es, r["M"].ef)
        self.assertEqual(r["M"].es, D(2024, 1, 3))
        self.assertEqual(r["M"].tf, 0)

    def test_ss_with_lag(self):
        r = c.schedule([c.Task("A", 10), c.Task("B", 3, preds=[("A", "SS", 2)])], START, cal())
        self.assertEqual(r["B"].es, D(2024, 1, 3))
        self.assertEqual(r["B"].ef, D(2024, 1, 5))
        # A kéo dài 10 ngày công (đến 12/01) nên B (kết thúc 05/01) có dự trữ: 12/01 - 05/01 theo ngày công
        self.assertEqual(r["A"].tf, 0)
        self.assertGreater(r["B"].tf, 0)

    def test_ff_aligns_finish(self):
        r = c.schedule([c.Task("A", 5), c.Task("B", 3, preds=[("A", "FF", 0)])], START, cal())
        self.assertEqual(r["A"].ef, D(2024, 1, 5))
        self.assertEqual(r["B"].ef, D(2024, 1, 5))
        self.assertEqual(r["B"].es, D(2024, 1, 3))

    def test_negative_lag_overlaps(self):
        r = c.schedule([c.Task("A", 4), c.Task("B", 2, preds=[("A", "FS", -1)])], START, cal())
        self.assertEqual(r["A"].ef, D(2024, 1, 4))
        self.assertEqual(r["B"].es, D(2024, 1, 4))  # FS-1: bắt đầu cùng ngày A kết thúc

    def test_snet_delays_start(self):
        r = c.schedule([c.Task("A", 2, snet=D(2024, 1, 10))], START, cal())
        self.assertEqual(r["A"].es, D(2024, 1, 10))

    def test_deadline_gives_negative_float(self):
        r = c.schedule([c.Task("A", 5)], START, cal(), deadline=D(2024, 1, 3))
        self.assertLess(r["A"].tf, 0)
        self.assertTrue(r["A"].critical)

    def test_summary_rollup(self):
        tasks = [
            c.Task("1", 0, "Summary"),
            c.Task("1.1", 3),
            c.Task("1.2", 2, preds=[("1.1", "FS", 0)]),
        ]
        r = c.schedule(tasks, START, cal())
        self.assertEqual((r["1"].es, r["1"].ef), (D(2024, 1, 1), D(2024, 1, 5)))
        self.assertTrue(r["1"].critical)


class TestRain(unittest.TestCase):
    def test_rain_factor_lengthens_duration_only_when_enabled(self):
        t = [c.Task("A", 10, rain=0.5)]
        self.assertEqual(c.schedule(t, START, cal())["A"].ef, D(2024, 1, 11))                 # 10 ngày công
        self.assertEqual(c.schedule(t, START, cal(), use_rain=True)["A"].ef, D(2024, 1, 23))  # 20 ngày công

    def test_rain_rounding_matches_vba_round_half_even(self):
        self.assertEqual(c.Task("A", 25, rain=0.75).effective_duration(True), 33)  # 33.33 -> 33
        self.assertEqual(c.Task("A", 3, rain=0.5).effective_duration(True), 6)
        self.assertEqual(c.Task("A", 5, rain=1.0).effective_duration(True), 5)

    def test_milestone_ignores_rain(self):
        r = c.schedule([c.Task("M", 0, "Milestone", rain=0.5)], START, cal(), use_rain=True)
        self.assertEqual(r["M"].es, r["M"].ef)


class TestValidation(unittest.TestCase):
    def test_cycle_detected(self):
        with self.assertRaises(c.CpmError):
            c.schedule([c.Task("A", 1, preds=[("B", "FS", 0)]), c.Task("B", 1, preds=[("A", "FS", 0)])], START, cal())

    def test_missing_predecessor(self):
        with self.assertRaises(c.CpmError):
            c.schedule([c.Task("A", 1, preds=[("X", "FS", 0)])], START, cal())

    def test_predecessor_cannot_be_summary(self):
        with self.assertRaises(c.CpmError):
            c.schedule([c.Task("1", 0, "Summary"), c.Task("1.1", 1), c.Task("2", 1, preds=[("1", "FS", 0)])], START, cal())

    def test_parse_pred(self):
        self.assertEqual(c.parse_pred("1.2.1.2SS+20; 1.1.1FS-3, 5"),
                         [("1.2.1.2", "SS", 20), ("1.1.1", "FS", -3), ("5", "FS", 0)])
        self.assertEqual(c.parse_pred(None), [])
        with self.assertRaises(c.CpmError):
            c.parse_pred("abc")


if __name__ == "__main__":
    unittest.main()
