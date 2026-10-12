"""
Bộ sưu tập số liệu về khả năng quan sát

cung cấp CommandMetrics(Chỉ báo lệnh đơn)Và SessionMetrics(Số liệu tổng hợp phiên), 
Tự động thu thập sau khi thực hiện lệnh, Có thể vượt qua DrawingState hoặc đăng nhập đầu ra. 
"""

import time
import logging
from typing import Any, Dict, List, Optional
from collections import defaultdict

logger = logging.getLogger('cad_metrics')


class CommandMetrics:
    """Số liệu có cấu trúc cho một lần thực thi lệnh"""

    def __init__(self, command_name: str, execution_method: str = "unknown"):
        self.command_name = command_name
        self.execution_method = execution_method
        self.start_time = time.time()
        self.duration_ms: int = 0
        self.success: bool = False
        self.error_code: Optional[int] = None
        self.retry_count: int = 0
        self.entities_created: int = 0

    def finish(self, result: Dict[str, Any]) -> None:
        """Trích xuất số liệu từ kết quả lệnh"""
        self.duration_ms = int((time.time() - self.start_time) * 1000)
        self.success = result.get("success", False)
        self.error_code = result.get("errorCode")
        self.execution_method = result.get("executionMethod", self.execution_method)
        self.entities_created = len(result.get("createdEntityHandles", []))

    def to_dict(self) -> Dict[str, Any]:
        return {
            "command_name": self.command_name,
            "execution_method": self.execution_method,
            "duration_ms": self.duration_ms,
            "success": self.success,
            "error_code": self.error_code,
            "retry_count": self.retry_count,
            "entities_created": self.entities_created,
        }


class SessionMetrics:
    """Chỉ số tổng hợp ở cấp phiên"""

    def __init__(self):
        self.session_start = time.time()
        self.total_commands: int = 0
        self.successful_commands: int = 0
        self.failed_commands: int = 0
        self.total_duration_ms: int = 0
        self.commands_by_category: Dict[str, int] = defaultdict(int)
        self.error_counts: Dict[int, int] = defaultdict(int)  # errorCode -> count
        self.reconnect_count: int = 0
        self._recent_metrics: List[CommandMetrics] = []  # gần đây 50 dải

    def record(self, metrics: CommandMetrics) -> None:
        """Ghi lại số liệu thực hiện lệnh"""
        self.total_commands += 1
        self.total_duration_ms += metrics.duration_ms

        if metrics.success:
            self.successful_commands += 1
        else:
            self.failed_commands += 1
            if metrics.error_code is not None:
                self.error_counts[metrics.error_code] += 1

        # Sắp xếp theo tiền tố lệnh(api_draw → draw, api_query → query)
        category = self._extract_category(metrics.command_name)
        self.commands_by_category[category] += 1

        # giữ gần đây 50 dải
        self._recent_metrics.append(metrics)
        if len(self._recent_metrics) > 50:
            self._recent_metrics = self._recent_metrics[-50:]

        logger.debug(
            f"[Metrics] {metrics.command_name}: "
            f"{'OK' if metrics.success else f'FAIL({metrics.error_code})'} "
            f"in {metrics.duration_ms}ms"
        )

    def record_reconnect(self) -> None:
        """Ghi lại kết nối lại"""
        self.reconnect_count += 1

    @staticmethod
    def _extract_category(command_name: str) -> str:
        """Trích xuất danh mục từ tên lệnh"""
        if command_name.startswith("api_gslc_"):
            return "gslc"
        elif command_name.startswith("api_"):
            parts = command_name[4:].split("_", 1)
            return parts[0] if parts else "other"
        return "other"

    @property
    def success_rate(self) -> float:
        if self.total_commands == 0:
            return 0.0
        return self.successful_commands / self.total_commands

    @property
    def avg_duration_ms(self) -> float:
        if self.total_commands == 0:
            return 0.0
        return self.total_duration_ms / self.total_commands

    def to_dict(self) -> Dict[str, Any]:
        """Tạo tóm tắt số liệu phiên"""
        session_duration = int(time.time() - self.session_start)
        most_common_errors = sorted(
            self.error_counts.items(), key=lambda x: x[1], reverse=True
        )[:5]

        return {
            "session_duration_seconds": session_duration,
            "total_commands": self.total_commands,
            "successful": self.successful_commands,
            "failed": self.failed_commands,
            "success_rate": round(self.success_rate, 3),
            "avg_duration_ms": round(self.avg_duration_ms, 1),
            "reconnect_count": self.reconnect_count,
            "commands_by_category": dict(self.commands_by_category),
            "most_common_errors": [
                {"errorCode": code, "count": count}
                for code, count in most_common_errors
            ],
        }
