# -*- coding: utf-8 -*-
"""Kiểm thử tĩnh cho dự án 23HG. Chạy: python -m unittest discover -s tests -v"""
import ast, glob, os, re, unittest, zipfile
import xml.etree.ElementTree as ET

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "1_Scripts_TuDongHoa")
XLAM = os.path.join(ROOT, "23HG_Schedule_Assistant_Pro.xlam")
XLSM = os.path.join(ROOT, "23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm")
BAT = os.path.join(ROOT, "CAI_DAT_23HG_ENTERPRISE_PRO.bat")
UNBAT = os.path.join(ROOT, "GO_CAI_DAT_23HG.bat")
VBA_SOURCE = os.path.join(SCRIPTS, "rebuild_perfect_master_v3_professional.py")
NS = {"u": "http://schemas.microsoft.com/office/2009/07/customui"}


def read_text(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def active_scripts():
    return sorted(glob.glob(os.path.join(SCRIPTS, "*.py")))


def ribbon_xml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("customUI/customUI14.xml")


class TestScripts(unittest.TestCase):
    def test_scripts_parse(self):
        for f in active_scripts():
            with self.subTest(script=os.path.basename(f)):
                ast.parse(read_text(f), f)

    def test_scripts_compile_without_warnings(self):
        import warnings
        for f in active_scripts():
            with self.subTest(script=os.path.basename(f)):
                with warnings.catch_warnings():
                    warnings.simplefilter("error")
                    compile(read_text(f), f, "exec")

    def test_no_hardcoded_user_paths(self):
        for f in active_scripts():
            with self.subTest(script=os.path.basename(f)):
                text = read_text(f)
                self.assertNotIn("C:\\Users\\baotu", text)


class TestPackages(unittest.TestCase):
    def test_zip_integrity(self):
        for p in (XLAM, XLSM):
            with self.subTest(file=os.path.basename(p)):
                with zipfile.ZipFile(p) as z:
                    self.assertIsNone(z.testzip())
                    self.assertIn("xl/vbaProject.bin", z.namelist())
                    self.assertIn("[Content_Types].xml", z.namelist())

    def test_ribbon_relationship_registered(self):
        for p in (XLAM, XLSM):
            with self.subTest(file=os.path.basename(p)):
                with zipfile.ZipFile(p) as z:
                    rels = z.read("_rels/.rels").decode("utf-8")
                self.assertIn("customUI/customUI14.xml", rels)


class TestRibbon(unittest.TestCase):
    def setUp(self):
        self.xml = ribbon_xml(XLAM)
        self.root = ET.fromstring(self.xml)
        import glob
        self.vba = read_text(VBA_SOURCE) + "\n" + "\n".join(
            read_text(f) for f in glob.glob(os.path.join(SCRIPTS, "*.bas")))

    def test_xlam_and_xlsm_same_ribbon(self):
        self.assertEqual(self.xml, ribbon_xml(XLSM))

    def test_ids_unique(self):
        ids = [e.get("id") for e in self.root.iter() if e.get("id")]
        self.assertEqual(len(ids), len(set(ids)))

    def test_every_onaction_macro_exists_in_vba(self):
        defined = set(re.findall(r"\bSub\s+(\w+)", self.vba))
        used = {e.get("onAction") for e in self.root.iter() if e.get("onAction")}
        self.assertTrue(used)
        self.assertEqual(sorted(used - defined), [])

    def test_buttons_have_label_and_screentip(self):
        for b in self.root.iter("{%s}button" % NS["u"]):
            with self.subTest(button=b.get("id")):
                self.assertTrue(b.get("label"))
                self.assertTrue(b.get("screentip"))

    def test_large_buttons_per_group_at_most_three(self):
        for g in self.root.iter("{%s}group" % NS["u"]):
            large = [b for b in g if b.get("size") == "large"]
            with self.subTest(group=g.get("id")):
                self.assertLessEqual(len(large), 3)


class TestInstaller(unittest.TestCase):
    def read(self, p):
        with open(p, "rb") as f:
            return f.read()

    def test_crlf_line_endings(self):
        for p in (BAT, UNBAT):
            with self.subTest(file=os.path.basename(p)):
                data = self.read(p)
                self.assertIn(b"\r\n", data)
                self.assertEqual(data.count(b"\n"), data.count(b"\r\n"))

    def test_installer_does_not_force_kill_excel(self):
        text = self.read(BAT).decode("utf-8").lower()
        self.assertNotIn("taskkill", text)
        self.assertNotIn("stop-process", text)

    def test_installer_safeguards(self):
        text = self.read(BAT).decode("utf-8")
        for needle in ("tasklist", "errorlevel", "reg add", ":fail", "23HG_Backup"):
            with self.subTest(needle=needle):
                self.assertIn(needle, text)

    def test_uninstaller_reverses_registry(self):
        text = self.read(UNBAT).decode("utf-8")
        self.assertIn("reg delete", text)
        self.assertIn("23HGAddins", text)

    def test_every_label_has_matching_goto_target(self):
        text = self.read(BAT).decode("utf-8")
        labels = set(re.findall(r"^:(\w+)", text, re.M))
        gotos = set(re.findall(r"goto\s+:(\w+)", text, re.I))
        calls = set(re.findall(r"call\s+:(\w+)", text, re.I))
        self.assertEqual(sorted((gotos | calls) - labels), [])


class TestCI(unittest.TestCase):
    def test_workflow_runs_the_test_suite(self):
        text = read_text(os.path.join(ROOT, ".github", "workflows", "tests.yml"))
        self.assertIn("unittest discover -s tests", text)
        self.assertIn("libreoffice", text)
        self.assertIn("requirements.txt", text)

    def test_gitattributes_keeps_bat_crlf(self):
        text = read_text(os.path.join(ROOT, ".gitattributes"))
        self.assertIn("*.bat text eol=crlf", text)


class TestRepoHygiene(unittest.TestCase):
    def test_required_files(self):
        for name in ("README.md", "CHANGELOG.md", "ROADMAP.md", "KIEM_THU_TREN_WINDOWS.md", "requirements.txt", ".gitignore"):
            with self.subTest(file=name):
                self.assertTrue(os.path.exists(os.path.join(ROOT, name)))

    def test_requirements_pin_openpyxl(self):
        text = read_text(os.path.join(ROOT, "requirements.txt"))
        self.assertRegex(text, r"openpyxl==\d+\.\d+\.\d+")


if __name__ == "__main__":
    unittest.main()
