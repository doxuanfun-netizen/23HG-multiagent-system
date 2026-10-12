"""
Định nghĩa kiểu định tuyến lệnh
"""

from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


@dataclass
class CommandCandidate:
    """ứng cử viên chỉ huy"""
    command: str                          # Tên lệnh, giống api_gslc_insert_bottom_equip_tags
    score: float                          # tỉ số trận đấu [0, 1]
    category: str                         # Loại, giống insert/query/brush
    reason: str                           # Lý do khuyến nghị(Mô tả lệnh)
    why_not: Optional[str] = None         # Tại sao bạn không nên chọn cái này(hướng dẫn tiêu cực)
    usage_example: Optional[str] = None   # Ví dụ sử dụng


@dataclass
class RoutingResult:
    """Kết quả định tuyến"""
    domain: str                           # Miền mà lệnh thuộc về, giống gslc/gsbz/basic
    intent_category: str                  # Các danh mục mục đích đã xác định, giống insert/query
    intent_confidence: float              # Sự tự tin có ý định [0, 1]
    recommended: Optional[Dict[str, Any]] # Các lệnh được đề xuất(thí sinh có điểm cao nhất)
    candidates: List[CommandCandidate]    # danh sách ứng viên(hầu hết5cá nhân)
    disambiguation_hint: Optional[str] = None  # Gợi ý định hướng(khitop2Khi tỉ số gần kề)


@dataclass
class IntentCategory:
    """Định nghĩa lớp ý định"""
    name: str                             # Tên danh mục
    display_name: str                     # tên hiển thị
    keywords: List[str] = field(default_factory=list)       # Từ khóa kích hoạt
    anti_keywords: List[str] = field(default_factory=list)  # từ khóa đảo ngược


@dataclass
class EntityType:
    """Định nghĩa loại thực thể"""
    name: str                             # Tên thực thể
    display_name: str                     # tên hiển thị
    keywords: List[str] = field(default_factory=list)  # Xác định từ khóa


@dataclass
class CommandMapping:
    """định nghĩa ánh xạ lệnh"""
    name: str                             # Tên lệnh
    intent: str                           # Danh mục ý định
    entity: Any                           # loại thực thể mục tiêu(str hoặc List[str])
    keywords: List[str] = field(default_factory=list)       # Từ khóa kích hoạt
    anti_patterns: List[str] = field(default_factory=list)  # chống mẫu(Quyền đại biểu)
    description: str = ""                 # Mô tả lệnh
    differentiator: Optional[str] = None  # Đặc điểm khác biệt
