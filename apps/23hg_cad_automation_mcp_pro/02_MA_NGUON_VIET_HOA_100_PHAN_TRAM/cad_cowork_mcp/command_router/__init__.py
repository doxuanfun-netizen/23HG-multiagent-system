"""
khung định tuyến lệnh

Cung cấp khả năng định tuyến thông minh cho các lệnh miền doanh nghiệp, gỡ rối agent Khó khăn trong việc lựa chọn khi phải đối mặt với một số lượng lớn các lệnh tương tự. 

Ngành kiến ​​​​trúc: 
    command_router/          # Khung định tuyến được chia sẻ
    ├── types.py            # định nghĩa kiểu
    ├── scorer.py           # Công cụ chấm điểm
    └── base_router.py      # Lớp cơ sở bộ định tuyến

    domains/                 # Triển khai miền doanh nghiệp
    └── gslc/               # Bộ định tuyến sơ đồ quy trình
        ├── router.py
        ├── taxonomy.py
        └── rules.py

Cách sử dụng: 
    from domains.gslc.router import GsLcRouter

    router = GsLcRouter()
    result = router.route("Tạo thẻ thiết bị theo lô")
    print(result.recommended)  # api_gslc_insert_bottom_equip_tags
"""

from .types import CommandCandidate, RoutingResult
from .base_router import BaseCommandRouter
from .scorer import ScoringEngine

__all__ = [
    'CommandCandidate',
    'RoutingResult',
    'BaseCommandRouter',
    'ScoringEngine',
]
