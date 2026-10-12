# -*- coding: utf-8 -*-
"""
Kiem thu hinh hoc thuc te voi so lieu da biet tren AutoCAD (Known Geometry Live Test).
Kiem chung triet de:
1. Hinh chu nhat 5000x4000 mm (5x4 m) co chu vi 18.0 m va dien tich 20.0 m2.
2. Hatch SOLID tren cung layer co dien tich 20.0 m2.
3. Bo tach phai phan tach rieng biet:
   - closed_polyline_area_m2 = 20.0 m2
   - hatch_area_m2 = 20.0 m2
   - Khong cong trung (Double-counting) thanh 40.0 m2
   - Co canh bao trung lap ro rang
4. Xuat file Excel kiem tra 100% cong thuc song Zero Dead Numbers.
5. Tu dong don dep (cleanup) thuc the sau khi test.
"""
import os
import tempfile
import unittest
import openpyxl
import pythoncom
import win32com.client

from cad_controller import CADController


class TestLiveKnownGeometry(unittest.TestCase):
    """Kiem thu do dac hinh hoc mau tren AutoCAD truc tiep."""

    @classmethod
    def setUpClass(cls):
        cls.controller = CADController()
        cls.cad_running = cls.controller.start_cad()

    @classmethod
    def tearDownClass(cls):
        if hasattr(cls, "controller") and cls.controller:
            cls.controller.disconnect()

    def test_live_known_rectangle_and_hatch(self):
        """Kiem tra hinh chu nhat 5x4m kem Hatch chong cong trung va kiem chung cong thuc Excel."""
        if not self.cad_running:
            self.skipTest("AutoCAD khong mo san tren may, bo qua kiem thu Live COM.")

        c = self.controller
        doc = c.doc
        ms = doc.ModelSpace

        # Thiet lap don vi mm (INSUNITS = 4)
        doc.SetVariable("INSUNITS", 4)

        test_layer = "TEST_SAN_BETONG_5X4"
        created_handles = []

        try:
            # 1. Ve hinh chu nhat 5000 x 4000 mm bang CADController
            rect = c.draw_rectangle((0.0, 0.0, 0.0), (5000.0, 4000.0, 0.0), layer=test_layer)
            self.assertIsNotNone(rect, "Khong the ve hinh chu nhat tren AutoCAD")
            created_handles.append(rect.Handle)

            # 2. Tao Hatch SOLID phu len tren hinh chu nhat do tren cung layer
            hatch = ms.AddHatch(0, "SOLID", True)
            loop = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [rect])
            hatch.AppendOuterLoop(loop)
            hatch.Evaluate()
            hatch.Layer = test_layer
            created_handles.append(hatch.Handle)

            # 3. Chay quantity_by_layer cho layer thu nghiem
            res = c.quantity_by_layer(layers=[test_layer])
            self.assertTrue(res.get("success"), f"quantity_by_layer that bai: {res}")

            data = res["data"]
            self.assertEqual(data["layers_count"], 1)
            layer_info = data["layers"][0]

            # 4. Kiem tra so lieu chinh xac:
            # - Chu vi: 2 * (5000 + 4000) = 18000 mm = 18.0 m
            self.assertAlmostEqual(layer_info["raw_length"], 18000.0, places=1)
            self.assertAlmostEqual(layer_info["length_m"], 18.0, places=2)

            # - Dien tich Pline kin: 5000 * 4000 = 20,000,000 mm2 = 20.0 m2
            self.assertAlmostEqual(layer_info["closed_polyline_area_m2"], 20.0, places=2)

            # - Dien tich Hatch: 20,000,000 mm2 = 20.0 m2
            self.assertAlmostEqual(layer_info["hatch_area_m2"], 20.0, places=2)

            # - Chuc nang chong cong trung: raw_area khong bi cong don thanh 40.0 m2
            self.assertAlmostEqual(layer_info["area_m2"], 20.0, places=2)
            self.assertIsNotNone(layer_info.get("area_warning"))
            self.assertIn("trùng lặp", layer_info["area_warning"].lower())

            # 5. Xuat file Excel va kiem tra cong thuc song
            with tempfile.TemporaryDirectory() as tmp_dir:
                xlsx_path = os.path.join(tmp_dir, "KNOWN_GEOMETRY_BOQ.xlsx")
                exp_res = c.export_boq(file_path=xlsx_path, format="xlsx")
                self.assertTrue(exp_res.get("success"))
                self.assertTrue(os.path.exists(xlsx_path))

                wb = openpyxl.load_workbook(xlsx_path, data_only=False)
                ws2 = wb["2_Chi_Tiet_Layer"]

                # Tim dong cua layer TEST_SAN_BETONG_5X4
                target_row = None
                for r in range(4, ws2.max_row + 1):
                    if ws2[f"B{r}"].value == test_layer:
                        target_row = r
                        break
                self.assertIsNotNone(target_row, f"Khong tim thay layer {test_layer} trong Sheet 2")

                # Kiem tra cac cong thuc song
                self.assertEqual(ws2[f"E{target_row}"].value, f"=D{target_row}*0.001")
                self.assertEqual(ws2[f"G{target_row}"].value, f"=F{target_row}*1e-06")
                self.assertEqual(ws2[f"I{target_row}"].value, f"=H{target_row}*1e-06")

        finally:
            # Don dep cac thuc the da tao
            for h in created_handles:
                try:
                    ent = doc.HandleToObject(h)
                    ent.Delete()
                except Exception:
                    pass


if __name__ == "__main__":
    unittest.main()
