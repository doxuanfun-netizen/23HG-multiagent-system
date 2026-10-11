# -*- coding: utf-8 -*-
"""Chạy lõi CPM bằng VBA (1_Scripts_TuDongHoa/vba_cpm_core.bas) trong LibreOffice (VBA mode)
và so sánh với cpm_engine.py trên nhiều mạng ngẫu nhiên.

Phạm vi: chỉ kiểm lõi tính (không có Worksheet/Excel thật). Bộ chuyển sheet <-> lõi và ribbon vẫn phải
thử trên Excel (xem KIEM_THU_TREN_WINDOWS.md)."""
import datetime as dt
import os
import random
import shutil
import subprocess
import sys
import tempfile
import time
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "1_Scripts_TuDongHoa"))
from _libreoffice import calc_available  # noqa: E402
import cpm_engine as E  # noqa: E402

try:
    import uno  # type: ignore
    from com.sun.star.beans import PropertyValue  # type: ignore
except Exception:  # pragma: no cover
    uno = None

SOFFICE = shutil.which("soffice") if calc_available() else None
BASE0 = dt.date(1899, 12, 30)


def ser(d):
    return (d - BASE0).days


def unser(n):
    return BASE0 + dt.timedelta(days=n)


@unittest.skipUnless(uno and SOFFICE, "cần LibreOffice + python3-uno")
class VbaCoreVsPython(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="lo_vba_")
        cls.port = 20000 + random.randint(0, 9000)
        cls.proc = subprocess.Popen(
            [SOFFICE, "--headless", "--norestore", "-env:UserInstallation=file://" + cls.tmp,
             f"--accept=socket,host=localhost,port={cls.port};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
        ctx = None
        for _ in range(90):
            try:
                ctx = resolver.resolve(f"uno:socket,host=localhost,port={cls.port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        if ctx is None:
            cls._shutdown()
            raise unittest.SkipTest("không khởi động được LibreOffice")
        desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        pv = PropertyValue()
        pv.Name, pv.Value = "Hidden", True
        cls.desk = desk
        cls.doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv,))
        libs = cls.doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        lib = libs.getByName("Standard")
        with open(os.path.join(ROOT, "1_Scripts_TuDongHoa", "vba_cpm_core.bas"), encoding="utf-8") as f:
            lib.insertByName("CpmCore", "Option VBASupport 1\n" + f.read())
        with open(os.path.join(ROOT, "tests", "vba_harness.bas"), encoding="utf-8") as f:
            lib.insertByName("Harness", f.read())
        cls.fn = cls.doc.getScriptProvider().getScript(
            "vnd.sun.star.script:Standard.Harness.CpmRunText?language=Basic&location=document")

    @classmethod
    def _shutdown(cls):
        try:
            cls.proc.terminate()
            cls.proc.wait(timeout=20)
        except Exception:
            cls.proc.kill()
            cls.proc.wait()
        shutil.rmtree(cls.tmp, ignore_errors=True)

    @classmethod
    def tearDownClass(cls):
        try:
            cls.doc.close(True)
        except Exception:
            pass
        cls._shutdown()

    # ---- helpers ----
    def run_vba(self, start, hols, rows):
        txt = f"START={ser(start)}\nHOL={','.join(str(ser(h)) for h in hols)}\n"
        txt += "\n".join("|".join([w, k, str(d), p]) for w, k, d, p in rows)
        out = self.fn.invoke((txt,), (), ())[0]
        if out.startswith("ERR:"):
            return out
        res = {}
        for line in out.splitlines():
            w, es, ef, ls, lf, tf, cr = line.split("|")
            res[w] = (unser(int(es)), unser(int(ef)), unser(int(ls)), unser(int(lf)), int(tf), cr == "1")
        return res

    @staticmethod
    def run_py(start, hols, rows):
        cal = E.Calendar(hols)
        tasks = [E.Task(w, d, k, E.parse_pred(p)) for w, k, d, p in rows]
        r = E.schedule(tasks, start, cal)
        return {w: (v.es, v.ef, v.ls, v.lf, v.tf, v.critical) for w, v in r.items()}

    def assert_same(self, start, hols, rows):
        got = self.run_vba(start, hols, rows)
        self.assertIsInstance(got, dict, got)
        self.assertEqual(got, self.run_py(start, hols, rows), rows)

    @staticmethod
    def random_rows(rng, n_groups=3, per_group=8, shuffle=False):
        rows, leaves = [], []
        for g in range(1, n_groups + 1):
            rows.append((str(g), "Summary", 0, ""))
            for i in range(1, per_group + 1):
                w = f"{g}.{i}"
                preds = []
                if leaves:
                    for _ in range(rng.choice([0, 1, 1, 2, 3])):
                        pw = rng.choice(leaves)
                        rel = rng.choice(["FS", "FS", "FS", "SS", "FF", "SF"])
                        lag = rng.choice([0, 0, 0, 1, 3, -1, -2, 7])
                        preds.append(pw + ("" if (rel == "FS" and lag == 0 and rng.random() < .5) else rel)
                                     + (f"{lag:+d}" if lag else ""))
                kind = "Milestone" if rng.random() < .1 else "Task"
                dur = 0 if kind == "Milestone" else rng.randint(1, 12)
                rows.append((w, kind, dur, ",".join(sorted(set(preds)))))
                leaves.append(w)
        if shuffle:
            summ = [r for r in rows if r[1] == "Summary"]
            body = [r for r in rows if r[1] != "Summary"]
            rng.shuffle(body)
            rows = summ + body
        return rows

    # ---- tests ----
    def test_basic_chain_with_holiday_and_lag(self):
        rows = [("1", "Summary", 0, ""), ("1.1", "Task", 3, ""), ("1.2", "Task", 2, "1.1"),
                ("1.3", "Milestone", 0, "1.2FS+1")]
        self.assert_same(dt.date(2025, 1, 6), [dt.date(2025, 1, 9)], rows)

    def test_start_on_sunday_is_snapped(self):
        rows = [("1", "Summary", 0, ""), ("1.1", "Task", 2, "")]
        got = self.run_vba(dt.date(2025, 1, 5), [], rows)
        self.assertEqual(got["1.1"][0], dt.date(2025, 1, 6))

    def test_random_networks_match_python(self):
        for seed in range(40):
            rng = random.Random(seed)
            rows = self.random_rows(rng)
            hols = [dt.date(2025, 1, 1) + dt.timedelta(days=rng.randint(0, 60)) for _ in range(rng.randint(0, 6))]
            with self.subTest(seed=seed):
                self.assert_same(dt.date(2025, 1, 1) + dt.timedelta(days=rng.randint(0, 20)), hols, rows)

    def test_predecessor_row_after_successor(self):
        for seed in range(15):
            rng = random.Random(1000 + seed)
            rows = self.random_rows(rng, shuffle=True)
            with self.subTest(seed=seed):
                self.assert_same(dt.date(2025, 3, 3), [], rows)

    def test_errors_are_reported_not_computed(self):
        base = [("1", "Summary", 0, "")]
        cases = {
            "missing": base + [("1.1", "Task", 1, "9.9")],
            "cycle": base + [("1.1", "Task", 1, "1.2"), ("1.2", "Task", 1, "1.1")],
            "to summary": base + [("1.1", "Task", 1, "1")],
            "duplicate": base + [("1.1", "Task", 1, ""), ("1.1", "Task", 2, "")],
            "bad text": base + [("1.1", "Task", 1, ""), ("1.2", "Task", 1, "1.1XX")],
            "empty summary": [("1", "Summary", 0, ""), ("2", "Task", 1, "")],
        }
        for name, rows in cases.items():
            with self.subTest(name):
                got = self.run_vba(dt.date(2025, 1, 6), [], rows)
                self.assertIsInstance(got, str)
                self.assertTrue(got.startswith("ERR:"), got)

    def test_sheet_adapter_writes_dates_and_reports_errors(self):
        """CalculateCPM (vba_cpm_adapter.bas) đọc/ghi sheet TIEN_DO; Application.* cần tài liệu có view."""
        pv = PropertyValue()
        pv.Name, pv.Value = "Hidden", False
        doc = self.desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv,))
        try:
            sheets = doc.Sheets
            sheets.insertNewByName("TIEN_DO", 0)
            sheets.insertNewByName("NGAY_NGHI_LE", 1)
            ws = sheets.getByName("TIEN_DO")
            ws.getCellByPosition(5, 1).setString("06/01/2025")
            data = [("1", "S", "Summary", "", 0), ("1.1", "T", "Task", "", 3),
                    ("1.2", "T", "Task", "1.1", 2), ("1.3", "M", "Task", "1.2FS+1", 0)]
            for i, (w, nm, kind, pr, d) in enumerate(data):
                r = 5 + i
                ws.getCellByPosition(0, r).setString("x")
                ws.getCellByPosition(1, r).setString(w)
                ws.getCellByPosition(2, r).setString(nm)
                ws.getCellByPosition(3, r).setString(kind)
                ws.getCellByPosition(4, r).setString(pr)
                ws.getCellByPosition(5, r).setValue(d)
                ws.getCellByPosition(6, r).setValue(1)
            sheets.getByName("NGAY_NGHI_LE").getCellByPosition(1, 5).setString("09/01/2025")
            libs = doc.BasicLibraries
            libs.createLibrary("Standard")
            lib = libs.getByName("Standard")
            for name, fn in (("CpmCore", "vba_cpm_core.bas"), ("Adapter", "vba_cpm_adapter.bas")):
                with open(os.path.join(ROOT, "1_Scripts_TuDongHoa", fn), encoding="utf-8") as f:
                    src = f.read()
                if name == "Adapter":
                    src += '\nSub RunIt()\n CalculateCPM ThisWorkbook.Sheets("TIEN_DO"), True\nEnd Sub\n'
                lib.insertByName(name, "Option VBASupport 1\n" + src)
            run = doc.getScriptProvider().getScript(
                "vnd.sun.star.script:Standard.Adapter.RunIt?language=Basic&location=document")
            run.invoke((), (), ())
            cell = lambda c, r: ws.getCellByPosition(c, r).getString()
            # 1.1: 6-8/1 (9/1 nghỉ lễ); 1.2: 10-11/1; mốc 1.3 = 11/1 + 2 ngày công = 14/1
            self.assertEqual((cell(7, 6), cell(8, 6)), ("06/01/2025", "08/01/2025"))
            self.assertEqual((cell(7, 7), cell(8, 7)), ("10/01/2025", "11/01/2025"))
            self.assertEqual((cell(7, 8), cell(8, 8)), ("14/01/2025", "14/01/2025"))
            self.assertEqual(cell(11, 6), "GANG")
        finally:
            doc.close(True)


if __name__ == "__main__":
    unittest.main()
