"""
Lớp cơ sở bộ định tuyến lệnh

Tất cả các bộ định tuyến miền doanh nghiệp đều kế thừa lớp này, Thực hiện các phương pháp trừu tượng sau đây: 
- domain_name: tên miền
- _load_taxonomy: Tải phân loại
- _get_intent_categories: Nhận danh mục ý định
- _get_command_mappings: Nhận ánh xạ lệnh
- _get_negative_rules: Nhận quy tắc tiêu cực
- _recognize_entities: Xác định các loại thực thể
- _generate_candidates: Tạo lệnh ứng viên
- _generate_disambiguation_hint: Tạo gợi ý định hướng
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any

from .types import RoutingResult, CommandCandidate
from .scorer import ScoringEngine


class BaseCommandRouter(ABC):
    """Lớp cơ sở bộ định tuyến lệnh"""

    def __init__(self):
        self.scorer = ScoringEngine()
        self._load_taxonomy()

    @property
    @abstractmethod
    def domain_name(self) -> str:
        """tên miền, giống 'gslc', 'gsbz'"""
        pass

    @abstractmethod
    def _load_taxonomy(self):
        """Tải cấu hình phân loại"""
        pass

    @abstractmethod
    def _get_intent_categories(self) -> Dict:
        """Nhận định nghĩa lớp ý định"""
        pass

    @abstractmethod
    def _get_command_mappings(self) -> Dict:
        """Nhận ánh xạ lệnh"""
        pass

    @abstractmethod
    def _get_negative_rules(self) -> List:
        """Nhận quy tắc hướng dẫn tiêu cực"""
        pass

    @abstractmethod
    def _recognize_entities(self, user_intent: str) -> List[str]:
        """Xác định các loại thực thể"""
        pass

    @abstractmethod
    def _generate_candidates(
        self,
        intent: Dict,
        entities: List[str],
        user_intent: str,
        context: Optional[Dict]
    ) -> List[CommandCandidate]:
        """Tạo lệnh ứng viên"""
        pass

    @abstractmethod
    def _generate_disambiguation_hint(
        self,
        candidates: List[CommandCandidate],
        user_intent: str
    ) -> Optional[str]:
        """Tạo gợi ý định hướng"""
        pass

    def route(self, user_intent: str, context: Optional[Dict] = None) -> RoutingResult:
        """
        Thực hiện định tuyến lệnh

        Args:
            user_intent: Mô tả mục đích của người dùng
            context: bối cảnh tùy chọn (has_frame_handle, has_entity_handles, target_entity_type)

        Returns:
            RoutingResult: Kết quả định tuyến
        """
        # 1. Nhận dạng ý định
        intent = self._recognize_intent(user_intent)

        # 2. Nhận dạng thực thể
        entities = self._recognize_entities(user_intent)

        # 3. Tạo và chấm điểm ứng viên
        candidates = self._generate_candidates(intent, entities, user_intent, context)

        # 4. Áp dụng quy tắc phủ định
        candidates = self._apply_negative_rules(candidates, user_intent)

        # 5. loại
        candidates.sort(key=lambda x: x.score, reverse=True)

        # 6. Tạo gợi ý định hướng
        hint = self._generate_disambiguation_hint(candidates, user_intent)

        # 7. Xây dựng kết quả
        recommended = self._format_recommended(candidates[0]) if candidates else None

        return RoutingResult(
            domain=self.domain_name,
            intent_category=intent["category"],
            intent_confidence=intent["confidence"],
            recommended=recommended,
            candidates=candidates[:5],  # Trả lại nhiều nhất5cá nhân
            disambiguation_hint=hint
        )

    def _recognize_intent(self, user_intent: str) -> Dict:
        """
        Nhận dạng ý định

        Args:
            user_intent: Mô tả mục đích của người dùng

        Returns:
            {"category": str, "confidence": float}
        """
        categories = self._get_intent_categories()
        best_match = {"category": "unknown", "confidence": 0.0}

        for cat_name, cat_def in categories.items():
            # Tính toán kết quả phù hợp từ khóa
            keywords = cat_def.get("keywords", [])
            keyword_score = sum(1 for kw in keywords if kw in user_intent)

            # Tính hình phạt từ khóa đảo ngược
            anti_keywords = cat_def.get("anti_keywords", [])
            anti_score = sum(1 for kw in anti_keywords if kw in user_intent)

            # Điểm tổng hợp
            score = keyword_score - anti_score * 0.5
            confidence = min(1.0, score / 3.0)  # bình thường hóa

            if confidence > best_match["confidence"]:
                best_match = {"category": cat_name, "confidence": confidence}

        return best_match

    def _apply_negative_rules(
        self,
        candidates: List[CommandCandidate],
        user_intent: str
    ) -> List[CommandCandidate]:
        """
        Áp dụng quy tắc khởi động tiêu cực

        Args:
            candidates: Danh sách lệnh ứng cử viên
            user_intent: Mô tả mục đích của người dùng

        Returns:
            Danh sách thí sinh sau khi áp dụng quy chế
        """
        rules = self._get_negative_rules()

        for rule in rules:
            trigger = rule.get("trigger")
            if trigger and callable(trigger) and trigger(user_intent):
                wrong_commands = rule.get("wrong", [])
                why_not = rule.get("why_not", "")

                for candidate in candidates:
                    if candidate.command in wrong_commands:
                        # Điểm thấp hơn đáng kể
                        candidate.score *= 0.3
                        # Đặt hướng dẫn khởi động tiêu cực
                        candidate.why_not = why_not

        return candidates

    def _format_recommended(self, candidate: CommandCandidate) -> Dict[str, Any]:
        """
        Các lệnh định dạng được đề xuất

        Args:
            candidate: thí sinh có điểm cao nhất

        Returns:
            Từ điển lệnh được đề xuất
        """
        return {
            "command": candidate.command,
            "score": round(candidate.score, 2),
            "category": candidate.category,
            "reason": candidate.reason,
            "usage_example": candidate.usage_example
        }
