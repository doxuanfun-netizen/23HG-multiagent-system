# -*- coding: utf-8 -*-
"""
Kiểm thử bộ công cụ chấm điểm và định tuyến lệnh CAD (ScoringEngine & CommandRouter)
Kiểm tra chi tiết scorer.py, base_router.py và domains/gslc/router.py.
"""
import unittest
from cad_cowork_mcp.command_router.scorer import ScoringEngine
from cad_cowork_mcp.domains.gslc.router import GsLcRouter


class TestScorerAndRouter(unittest.TestCase):
    """Kiểm tra ScoringEngine và GsLcRouter."""

    def setUp(self):
        self.scorer = ScoringEngine()
        self.router = GsLcRouter()

    def test_scorer_base_score_and_bounds(self):
        """Điểm số luôn nằm trong khoảng chuẩn hóa [0.0, 1.0]."""
        cmd_def = {"intent": "draw", "keywords": [], "entity": []}
        score = self.scorer.calculate_score("vẽ hình tròn", cmd_def, "draw", [])
        self.assertGreaterEqual(score, 0.0)
        self.assertLessEqual(score, 1.0)
        self.assertAlmostEqual(score, 0.75)  # BASE 0.5 + intent 0.25

    def test_scorer_keyword_matching_capped(self):
        """Khớp từ khóa được cộng điểm và có trần giới hạn tối đa 3 từ."""
        cmd_def = {
            "intent": "other",
            "keywords": ["đường", "ống", "dẫn", "nước"],
            "entity": []
        }
        # Khớp 2 từ khóa: 0.5 + 2 * 0.15 = 0.80
        score_2 = self.scorer.calculate_score("vẽ đường ống", cmd_def, "pipe", [])
        # Khớp 4 từ khóa: capped at 3 -> 0.5 + 3 * 0.15 = 0.95
        score_4 = self.scorer.calculate_score("vẽ đường ống dẫn nước", cmd_def, "pipe", [])
        self.assertAlmostEqual(score_2, 0.80)
        self.assertAlmostEqual(score_4, 0.95)
        self.assertGreater(score_4, score_2)

    def test_scorer_anti_pattern_penalty(self):
        """Phát hiện anti-pattern sẽ áp dụng điểm phạt thích đáng."""
        cmd_def_clean = {
            "intent": "valve",
            "keywords": ["van"],
            "anti_patterns": ["bơm"]
        }
        score_normal = self.scorer.calculate_score("vẽ van một chiều", cmd_def_clean, "valve", [])
        score_penalized = self.scorer.calculate_score("vẽ van và bơm", cmd_def_clean, "valve", [])
        self.assertLess(score_penalized, score_normal)

    def test_gslc_router_intent_routing(self):
        """GsLcRouter phân loại chính xác các ý định quy trình công nghệ P&ID."""
        # 1. Ý định vẽ van
        res_valve = self.router.route("chèn van một chiều trên đường ống")
        self.assertEqual(res_valve.domain, "gslc")
        self.assertTrue(len(res_valve.candidates) > 0)
        self.assertGreater(res_valve.intent_confidence, 0.0)

        # 2. Ý định vẽ thiết bị / bồn bể
        res_tank = self.router.route("vẽ bồn chứa hóa chất thẳng đứng")
        self.assertEqual(res_tank.domain, "gslc")
        self.assertTrue(len(res_tank.candidates) > 0)

    def test_gslc_disambiguation_hint_generation(self):
        """Bộ định tuyến tạo gợi ý phân biệt khi có nhiều lệnh ứng viên tiềm năng."""
        res = self.router.route("vẽ van")
        self.assertIsInstance(res.candidates, list)
        self.assertTrue(hasattr(res, "disambiguation_hint"))


if __name__ == "__main__":
    unittest.main()
