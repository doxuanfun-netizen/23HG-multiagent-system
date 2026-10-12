"""
GsLc Bộ định tuyến lệnh biểu đồ dòng quy trình

cung cấp 36 cá nhân GsLc API Khả năng định tuyến thông minh cho các lệnh. 

Cách sử dụng: 
    from domains.gslc.router import GsLcRouter

    router = GsLcRouter()
    result = router.route("Tạo thẻ thiết bị theo lô")

    print(result.recommended["command"])  # api_gslc_insert_bottom_equip_tags
    print(result.candidates[1].why_not)   # "auto_number Là để đánh số thực thể hiện có, Không tạo khối thẻ mới"
"""

from .router import GsLcRouter

__all__ = ['GsLcRouter']
