"""
Công cụ chấm điểm lệnh

Cơ chế chấm điểm: 
- Điểm cơ bản: 0.5
- phù hợp với ý định: +0.25
- So khớp thực thể: +0.20
- kết hợp từ khóa: +0.15 (mỗi, hầu hết3cá nhân)
- hình phạt chống khuôn mẫu: -0.30 (mỗi)
- Đặc điểm khác biệt: +0.15
"""

from typing import Dict, List, Any, Optional


class ScoringEngine:
    """Công cụ chấm điểm"""

    # Điểm cơ bản
    BASE_SCORE = 0.5

    # Cấu hình trọng lượng
    WEIGHTS = {
        "intent_match": 0.25,           # phù hợp với ý định
        "entity_match": 0.20,           # So khớp loại thực thể
        "keyword_match": 0.15,          # kết hợp từ khóa(mỗi)
        "keyword_max": 3,               # Giới hạn đối sánh từ khóa
        "anti_pattern_penalty": -0.30,  # hình phạt chống khuôn mẫu(mỗi)
        "differentiator_match": 0.15,   # Kết hợp tính năng khác biệt
        "context_boost": 0.10,          # phần thưởng theo ngữ cảnh
    }

    def calculate_score(
        self,
        user_intent: str,
        command_def: Dict[str, Any],
        intent_category: str,
        recognized_entities: List[str],
        context: Optional[Dict] = None
    ) -> float:
        """
        Tính điểm lệnh

        Args:
            user_intent: Mô tả mục đích của người dùng
            command_def: Định nghĩa lệnh
            intent_category: Các danh mục mục đích đã xác định
            recognized_entities: Danh sách các loại thực thể được công nhận
            context: bối cảnh tùy chọn

        Returns:
            điểm chuẩn hóa [0, 1]
        """
        score = self.BASE_SCORE

        # 1. phù hợp với ý định
        if command_def.get("intent") == intent_category:
            score += self.WEIGHTS["intent_match"]

        # 2. So khớp thực thể
        cmd_entities = command_def.get("entity", [])
        if isinstance(cmd_entities, str):
            cmd_entities = [cmd_entities]
        if any(e in recognized_entities for e in cmd_entities):
            score += self.WEIGHTS["entity_match"]

        # 3. kết hợp từ khóa
        keywords = command_def.get("keywords", [])
        keyword_matches = sum(1 for kw in keywords if kw in user_intent)
        score += self.WEIGHTS["keyword_match"] * min(keyword_matches, self.WEIGHTS["keyword_max"])

        # 4. hình phạt chống khuôn mẫu
        anti_patterns = command_def.get("anti_patterns", [])
        anti_matches = sum(1 for kw in anti_patterns if kw in user_intent)
        score += self.WEIGHTS["anti_pattern_penalty"] * anti_matches

        # 5. Kết hợp tính năng khác biệt
        differentiator = command_def.get("differentiator")
        if differentiator:
            # Chia bằng dấu phẩy tiếng Trung hoặc dấu phẩy tiếng Anh
            diff_keywords = [k.strip() for k in differentiator.replace(", ", ",").split(",")]
            if any(kw in user_intent for kw in diff_keywords if kw):
                score += self.WEIGHTS["differentiator_match"]

        # 6. phần thưởng theo ngữ cảnh
        if context:
            context_score = self._evaluate_context(context, command_def)
            score += self.WEIGHTS["context_boost"] * context_score

        # bình thường hóa thành [0, 1]
        return max(0.0, min(1.0, score))

    def _evaluate_context(self, context: Dict, command_def: Dict) -> float:
        """
        Đánh giá sự phù hợp theo ngữ cảnh

        Args:
            context: thông tin theo ngữ cảnh (has_frame_handle, has_entity_handles, target_entity_type)
            command_def: Định nghĩa lệnh

        Returns:
            điểm ngữ cảnh [0, 1]
        """
        score = 0.0

        # Nếu có tay cầm khung, Nên ưu tiên cho các lệnh yêu cầu khung ảnh.
        if context.get("has_frame_handle"):
            cmd_name = command_def.get("name", "")
            if "frame" in cmd_name.lower() or "in_frame" in cmd_name.lower():
                score += 0.5

        # Nếu có một xử lý thực thể, Ưu tiên các lệnh yêu cầu thực thể
        if context.get("has_entity_handles"):
            keywords = command_def.get("keywords", [])
            if any(kw in ["lô", "làm cho khỏe lại", "gia hạn"] for kw in keywords):
                score += 0.3

        # khớp loại thực thể mục tiêu
        target_type = context.get("target_entity_type")
        if target_type:
            cmd_entities = command_def.get("entity", [])
            if isinstance(cmd_entities, str):
                cmd_entities = [cmd_entities]
            if target_type in cmd_entities:
                score += 0.5

        return min(1.0, score)
