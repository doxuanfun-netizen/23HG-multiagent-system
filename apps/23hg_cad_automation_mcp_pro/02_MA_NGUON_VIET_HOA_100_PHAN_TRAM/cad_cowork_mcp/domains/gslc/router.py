"""
GsLc bộ định tuyến lệnh

kế thừa BaseCommandRouter, Thực hiện định tuyến thông minh các lệnh biểu đồ dòng quy trình. 
"""

from typing import Dict, List, Optional

from ...command_router.base_router import BaseCommandRouter
from ...command_router.types import CommandCandidate

from .taxonomy import INTENT_CATEGORIES, ENTITY_TYPES, COMMAND_MAPPINGS
from .rules import NEGATIVE_RULES, DISAMBIGUATION_HINTS


class GsLcRouter(BaseCommandRouter):
    """GsLc Bộ định tuyến lệnh biểu đồ dòng quy trình"""

    @property
    def domain_name(self) -> str:
        return "gslc"

    def _load_taxonomy(self):
        """Tải phân loại"""
        self.intent_categories = INTENT_CATEGORIES
        self.entity_types = ENTITY_TYPES
        self.command_mappings = COMMAND_MAPPINGS

    def _get_intent_categories(self) -> Dict:
        return self.intent_categories

    def _get_command_mappings(self) -> Dict:
        return self.command_mappings

    def _get_negative_rules(self) -> List:
        return NEGATIVE_RULES

    def _recognize_entities(self, user_intent: str) -> List[str]:
        """
        Xác định các loại thực thể trong mục đích của người dùng

        Args:
            user_intent: Mô tả mục đích của người dùng

        Returns:
            Danh sách các loại thực thể được công nhận
        """
        entities = []
        for entity_name, entity_def in self.entity_types.items():
            keywords = entity_def.get("keywords", [])
            if any(kw in user_intent for kw in keywords):
                entities.append(entity_name)
        return entities if entities else ["unknown"]

    def _generate_candidates(
        self,
        intent: Dict,
        entities: List[str],
        user_intent: str,
        context: Optional[Dict]
    ) -> List[CommandCandidate]:
        """
        Tạo lệnh ứng viên

        Args:
            intent: Có ý định xác định {"category": str, "confidence": float}
            entities: Danh sách các loại thực thể được công nhận
            user_intent: Mô tả mục đích của người dùng
            context: bối cảnh tùy chọn

        Returns:
            Danh sách lệnh ứng cử viên
        """
        candidates = []

        for cmd_name, cmd_def in self.command_mappings.items():
            # Tính điểm bằng công cụ tính điểm
            score = self.scorer.calculate_score(
                user_intent=user_intent,
                command_def=cmd_def,
                intent_category=intent["category"],
                recognized_entities=entities,
                context=context
            )

            # Tạo ứng viên
            candidate = CommandCandidate(
                command=cmd_name,
                score=score,
                category=cmd_def.get("intent", "unknown"),
                reason=cmd_def.get("description", ""),
                usage_example=f'{{"command_name": "{cmd_name}", "parameters": {{}}}}'
            )

            candidates.append(candidate)

        return candidates

    def _generate_disambiguation_hint(
        self,
        candidates: List[CommandCandidate],
        user_intent: str
    ) -> Optional[str]:
        """
        Tạo gợi ý định hướng

        khi top 2 Khi điểm thí sinh gần nhau, Cung cấp hướng dẫn phân biệt. 

        Args:
            candidates: Danh sách ứng viên được sắp xếp
            user_intent: Mô tả mục đích của người dùng

        Returns:
            Gợi ý định hướng, hoặc None
        """
        if len(candidates) < 2:
            return None

        top_two = candidates[:2]

        # nếu như top 2 Cách biệt điểm số ít hơn 0.3, Cung cấp các mẹo giải mã
        if top_two[0].score - top_two[1].score < 0.3:
            pair = (top_two[0].command, top_two[1].command)
            reverse_pair = (top_two[1].command, top_two[0].command)

            # Tìm gợi ý định hướng được xác định trước
            hint = DISAMBIGUATION_HINTS.get(pair) or DISAMBIGUATION_HINTS.get(reverse_pair)

            if hint:
                return hint

            # Nếu không có lời nhắc được xác định trước, Nhưng có why_not, cũng trở lại
            if top_two[1].why_not:
                return f"Hãy chú ý đến sự khác biệt: {top_two[1].why_not}"

        return None


# Chức năng tiện lợi
def route_gslc_command(user_intent: str, context: Optional[Dict] = None):
    """
    lộ trình GsLc Đặt hàng(Chức năng tiện lợi)

    Args:
        user_intent: Mô tả mục đích của người dùng
        context: bối cảnh tùy chọn

    Returns:
        RoutingResult
    """
    router = GsLcRouter()
    return router.route(user_intent, context)
