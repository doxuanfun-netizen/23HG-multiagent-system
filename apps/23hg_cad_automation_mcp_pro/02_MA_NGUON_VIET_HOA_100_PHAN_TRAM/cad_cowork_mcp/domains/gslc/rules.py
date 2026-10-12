"""
GsLc Quy tắc điều hướng và phân giải lệnh trùng lặp

Giải quyết 8 cặp lệnh dễ gây nhầm lẫn:
1. insert_bottom_equip_tags vs auto_number (Tạo hàng loạt vs Đánh số)
2. insert_valve vs insert_control_valve (Van thông thường vs Van điều khiển)
3. insert_join_draw_arrow vs insert_outer_pipe_arrow (Nối bản vẽ vs Ống ngoại vi)
4. brush_pipe_from_to vs sync_pipe_data (Làm mới liên kết ống vs Đồng bộ ngoài)
5. query_instrument_in_frame vs query_valve_in_frame (Đồng hồ đo gồm van điều khiển vs Van thường)
6. export_all_data vs export_frame_data (Toàn bộ vs Từng khung)
7. insert_equip_tag vs insert_bottom_equip_tags (Đơn lẻ vs Hàng loạt)
8. auto_number vs insert_equipment (Đánh số vs Chèn thiết bị)
"""

NEGATIVE_RULES = [
    # Quy tắc 1: Tạo hàng loạt mã vị trí vs Tự động đánh số
    {
        "name": "batch_generate_vs_auto_number",
        "trigger": lambda inp: any(k in inp.lower() for k in ["hàng loạt", "tạo mã", "tạo tag", "tạo vị trí", "Tạo hàng loạt", "Tạo thẻ", "Tạo thẻ thiết bị"]) and not any(k in inp.lower() for k in ["đánh số", "đánh stt", "số seri"]),
        "wrong": ["api_gslc_auto_number"],
        "right": "api_gslc_insert_bottom_equip_tags",
        "why_not": "auto_number dùng để đánh số các thực thể đã có (sửa thuộc tính), không phải tạo mới block mã vị trí (tag)"
    },

    # Quy tắc 2: Tự động đánh số vs Chèn/Tạo mới
    {
        "name": "auto_number_vs_insert",
        "trigger": lambda inp: any(k in inp.lower() for k in ["đánh số", "đánh stt", "sắp xếp", "số seri"]) and any(k in inp.lower() for k in ["tự động", "lại", "thứ tự", "tự động", "lại", "loại"]),
        "wrong": ["api_gslc_insert_bottom_equip_tags", "api_gslc_insert_equip_tag", "api_gslc_insert_equipment"],
        "right": "api_gslc_auto_number",
        "why_not": "insert_xxx dùng để tạo mới thực thể; để đánh số thứ tự vui lòng dùng auto_number"
    },

    # Quy tắc 3: Mũi tên nối bản vẽ vs Mũi tên ống ngoại vi
    {
        "name": "join_arrow_vs_outer_arrow",
        "trigger": lambda inp: any(k in inp.lower() for k in ["nối bản vẽ", "xuyên bản vẽ", "sang bản vẽ", "tiếp bản vẽ", "Nhận hình ảnh"]) and not any(k in inp.lower() for k in ["ngoại vi", "ngoại tuyến", "ngoài", "Ống ngoài"]),
        "wrong": ["api_gslc_insert_outer_pipe_arrow"],
        "right": "api_gslc_insert_join_draw_arrow",
        "why_not": "outer_pipe_arrow dùng cho đường ống ngoại vi/phụ trợ; liên kết giữa các bản vẽ vui lòng dùng join_draw_arrow"
    },

    # Quy tắc 4: Mũi tên ống ngoại vi vs Mũi tên nối bản vẽ
    {
        "name": "outer_arrow_vs_join_arrow",
        "trigger": lambda inp: any(k in inp.lower() for k in ["ngoại vi", "ngoại tuyến", "công cộng", "hơi nóng", "nước làm mát", "Ống ngoài", "Công trình công cộng", "hơi nước", "nước làm mát"]) and not any(k in inp.lower() for k in ["nối bản vẽ", "tiếp", "Nhận hình ảnh"]),
        "wrong": ["api_gslc_insert_join_draw_arrow"],
        "right": "api_gslc_insert_outer_pipe_arrow",
        "why_not": "join_draw_arrow dùng để liên kết xuyên bản vẽ; hướng ống ngoại vi/phụ trợ vui lòng dùng outer_pipe_arrow"
    },

    # Quy tắc 5: Van điều khiển vs Van thường
    {
        "name": "control_valve_vs_valve",
        "trigger": lambda inp: any(k in inp.lower() for k in ["van điều khiển", "van điều tiết", "van điện", "van khí nén", "van on/off", "van điều khiển", "van điều tiết", "Van bật/tắt", "Van điện", "Van khí nén"]),
        "wrong": ["api_gslc_insert_valve"],
        "right": "api_gslc_insert_control_valve",
        "why_not": "insert_valve là van thông thường (van bi, van cổng...); van điều khiển/điều tiết vui lòng dùng insert_control_valve"
    },

    # Quy tắc 6: Van thường vs Van điều khiển
    {
        "name": "valve_vs_control_valve",
        "trigger": lambda inp: any(k in inp.lower() for k in ["van bi", "van cổng", "van chặn", "van bướm", "van", "van", "van bi", "cổng", "van cầu", "van bướm"]) and not any(k in inp.lower() for k in ["điều khiển", "điều tiết", "điện", "khí nén", "điều khiển", "điều chỉnh", "công tắc"]),
        "wrong": ["api_gslc_insert_control_valve"],
        "right": "api_gslc_insert_valve",
        "why_not": "insert_control_valve là van điều khiển thuộc hệ đo lường; van thông thường (van bi, cổng...) vui lòng dùng insert_valve"
    },

    # Quy tắc 7: Làm mới liên kết vs Đồng bộ dữ liệu
    {
        "name": "brush_vs_sync",
        "trigger": lambda inp: any(k in inp.lower() for k in ["làm mới", "quét liên kết", "cập nhật hướng", "làm cho khỏe lại", "chải"]) and any(k in inp.lower() for k in ["ống", "đường ống", "đồng hồ", "đo", "đường ống", "mét"]) and not any(k in inp.lower() for k in ["excel", "nhập", "import", "đồng bộ", "json", "từ file", "nhập khẩu", "đồng bộ"]),
        "wrong": ["api_gslc_sync_pipe_data", "api_gslc_sync_instrument_data", "api_gslc_sync_equipment_data"],
        "right": "api_gslc_brush_pipe_from_to",
        "why_not": "sync_xxx dùng để nhập dữ liệu từ ngoài (Excel/JSON); làm mới liên kết trong CAD vui lòng dùng brush_xxx"
    },

    # Quy tắc 8: Đồng bộ dữ liệu vs Làm mới liên kết
    {
        "name": "sync_vs_brush",
        "trigger": lambda inp: any(k in inp.lower() for k in ["đồng bộ", "nhập dữ liệu", "excel", "json", "từ file", "import", "đồng bộ", "nhập khẩu", "từ tập tin"]),
        "wrong": ["api_gslc_brush_pipe_from_to", "api_gslc_brush_instrument_location"],
        "right": "api_gslc_sync_pipe_data",
        "why_not": "brush_xxx dùng để tự động tính toán lại liên kết trong CAD; nhập dữ liệu từ ngoài vui lòng dùng sync_xxx"
    }
]

# Ánh xạ gợi ý phân giải lệnh (Disambiguation Hints)
DISAMBIGUATION_HINTS = {
    ("api_gslc_insert_bottom_equip_tags", "api_gslc_auto_number"):
        "Tạo hàng loạt ≠ Tự động đánh số. Tạo mới là sinh block mã vị trí (tag) mới; đánh số là sửa thuộc tính số của thực thể đã có.",

    ("api_gslc_insert_valve", "api_gslc_insert_control_valve"):
        "Van thông thường (van bi/van cổng/van chặn) ≠ Van điều khiển (van tiết lưu/van điện/van khí nén).",

    ("api_gslc_insert_join_draw_arrow", "api_gslc_insert_outer_pipe_arrow"):
        "Mũi tên nối bản vẽ (xuyên bản vẽ) ≠ Mũi tên ống ngoại vi (hướng cấp phụ trợ bên ngoài).",

    ("api_gslc_brush_pipe_from_to", "api_gslc_sync_pipe_data"):
        "Làm mới liên kết ≠ Đồng bộ dữ liệu. Làm mới là CAD tự tính toán topo; đồng bộ là nhập từ Excel/JSON bên ngoài.",

    ("api_gslc_insert_equip_tag", "api_gslc_insert_bottom_equip_tags"):
        "Chèn đơn lẻ ≠ Tạo hàng loạt. insert_equip_tag cần chỉ định tọa độ; insert_bottom_equip_tags tự động sinh ở chân khung bản vẽ.",

    ("api_gslc_query_instrument_in_frame", "api_gslc_query_valve_in_frame"):
        "Truy vấn đồng hồ đo (gồm cả van điều khiển) ≠ Truy vấn van (chỉ van thông thường)."
}
