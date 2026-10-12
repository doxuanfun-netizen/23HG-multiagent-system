"""
GsLc Phân loại sơ đồ quy trình

sự định nghĩa: 
- INTENT_CATEGORIES: Danh mục ý định(query/insert/brush/export/sync/tool)
- ENTITY_TYPES: Loại thực thể(equipment/instrument/pipe/valve/...)
- COMMAND_MAPPINGS: 36 Ánh xạ chi tiết các lệnh
"""

# Định nghĩa lớp ý định
INTENT_CATEGORIES = {
    "query": {
        "display_name": "Truy vấn / Tìm kiếm",
        "keywords": ["Truy vấn", "Lấy", "danh sách", "trình diễn", "Kiểm tra", "thống kê", "Có gì ở đó", "bao nhiêu"],
        "anti_keywords": ["chèn", "tạo nên", "phát ra", "Thêm vào"]
    },
    "insert": {
        "display_name": "Chèn / Tạo mới",
        "keywords": ["chèn", "tạo nên", "phát ra", "Thêm vào", "địa điểm", "Tạo hàng loạt"],
        "anti_keywords": ["Truy vấn", "Lấy", "xóa bỏ"]
    },
    "brush": {
        "display_name": "Làm mới / Gán liên kết",
        "keywords": ["làm cho khỏe lại", "chải", "sự kết hợp", "ràng buộc", "Đặt vị trí", "Đặt nơi để đi"],
        "anti_keywords": ["chèn", "tạo nên", "đồng bộ", "nhập khẩu"]
    },
    "export": {
        "display_name": "Xuất dữ liệu",
        "keywords": ["Xuất khẩu", "đầu ra", "lưu dưới dạng", "Tạo báo cáo", "Tạo bảng", "Tạo bảng kê khai"],
        "anti_keywords": ["nhập khẩu", "đồng bộ"]
    },
    "sync": {
        "display_name": "Đồng bộ / Nhập",
        "keywords": ["đồng bộ", "nhập khẩu", "gia hạn", "từExcel", "từJSON", "từ tập tin"],
        "anti_keywords": ["Xuất khẩu", "làm cho khỏe lại"]
    },
    "tool": {
        "display_name": "Công cụ / Đánh số",
        "keywords": ["số seri", "đánh số tự động", "xác minh", "Thông thoáng", "di chuyển để đóng băng", "hồi phục"],
        "anti_keywords": ["Tạo thẻ", "Tạo thẻ", "Chèn số tham chiếu"]
    }
}

# Định nghĩa loại thực thể
ENTITY_TYPES = {
    "equipment": {
        "display_name": "Thiết bị",
        "keywords": ["thiết bị", "lò phản ứng", "bể chứa", "bơm", "lò sưởi", "máy ly tâm", "Bơm chân không"]
    },
    "equip_tag": {
        "display_name": "Mã vị trí thiết bị (Tag)",
        "keywords": ["số bit", "Thẻ thiết bị", "TAG", "Số dưới cùng", "Khối thẻ"]
    },
    "instrument": {
        "display_name": "Đồng hồ đo / Thiết bị đo",
        "keywords": ["mét", "cảm biến", "Máy phát", "nhiệt độ", "áp lực ", "mức chất lỏng", "chảy"]
    },
    "control_valve": {
        "display_name": "Van điều khiển",
        "keywords": ["van điều khiển", "van điều tiết", "Van bật/tắt", "Van điện", "Van khí nén", "Tự vận hành"]
    },
    "valve": {
        "display_name": "Van thông thường",
        "keywords": ["van", "van bi", "cổng", "van cầu", "Van kiểm tra", "van bướm", "van an toàn"]
    },
    "pipe": {
        "display_name": "Đường ống",
        "keywords": ["đường ống", "đường ống", "đường ống", "Đến và đi"]
    },
    "join_draw_arrow": {
        "display_name": "Mũi tên nối bản vẽ",
        "keywords": ["Nhận hình ảnh", "Trên các biểu đồ", "Số bản vẽ", "Tiếp tục", "mũi tên kết nối"]
    },
    "outer_pipe_arrow": {
        "display_name": "Mũi tên ống ngoại vi",
        "keywords": ["Ống ngoài", "Công trình công cộng", "hơi nước", "nước làm mát", "phong cách nhạc cụ", "Nitơ"]
    },
    "frame": {
        "display_name": "Khung bản vẽ",
        "keywords": ["Khung tranh", "khung tranh", "Phạm vi vẽ", "vẽ"]
    }
}

# 36 cá nhân GsLc ánh xạ lệnh
COMMAND_MAPPINGS = {
    # ========== Chèn lớp (9) ==========
    "api_gslc_insert_bottom_equip_tags": {
        "intent": "insert",
        "entity": "equip_tag",
        "keywords": ["Tạo số sê-ri theo lô", "tự độngTạo thẻ", "Số dưới cùng", "Tạo thẻ thiết bị", "Chèn số sê-ri theo lô"],
        "anti_patterns": ["số seri", "số seri", "Đánh số lại"],
        "description": "Tự động tạo khối thẻ thiết bị theo lô ở cuối khung",
        "differentiator": "Tự động tạo các thực thể mới theo đợt, Không cần xác định vị trí"
    },
    "api_gslc_insert_equip_tag": {
        "intent": "insert",
        "entity": "equip_tag",
        "keywords": ["Chèn số tham chiếu", "Thêm thẻ", "Tạo thẻ", "thẻ đơn"],
        "anti_patterns": ["lô", "tự động", "đáy"],
        "description": "Chèn một khối thẻ thiết bị vào vị trí đã chỉ định",
        "differentiator": "Cần xác định vị trí cụ thể"
    },
    "api_gslc_insert_equipment": {
        "intent": "insert",
        "entity": "equipment",
        "keywords": ["Cắm thiết bị", "Thêm thiết bị", "Tạo thiết bị", "đặt thiết bị"],
        "anti_patterns": ["lô", "tự động", "số bit"],
        "description": "Chèn một khối thiết bị vào vị trí đã chỉ định"
    },
    "api_gslc_insert_instrument": {
        "intent": "insert",
        "entity": "instrument",
        "keywords": ["Chèn mét", "Thêm nhạc cụ", "Tạo thước đo"],
        "description": "Chèn khối thiết bị vào vị trí đã chỉ định"
    },
    "api_gslc_insert_control_valve": {
        "intent": "insert",
        "entity": "control_valve",
        "keywords": ["Chèn van điều khiển", "Chèn van điều chỉnh", "Thêm van điều khiển", "Thêm van điều chỉnh", "Van bật/tắt"],
        "anti_patterns": ["van bi", "cổng", "van cầu"],
        "description": "Lắp các van điều khiển (van điều chỉnh, van chuyển mạch…) tại các vị trí được chỉ định"
    },
    "api_gslc_insert_valve": {
        "intent": "insert",
        "entity": "valve",
        "keywords": ["Chèn van", "Thêm van", "Chèn van bi", "Chèn van cổng", "Chèn van dừng"],
        "anti_patterns": ["van điều khiển", "van điều tiết", "Van bật/tắt"],
        "description": "Chèn một van thông thường vào vị trí được chỉ định"
    },
    "api_gslc_insert_pipe": {
        "intent": "insert",
        "entity": "pipe",
        "keywords": ["Chèn ống", "Thêm đường ống", "Tạo đường dẫn"],
        "description": "Chèn ống vào vị trí được chỉ định"
    },
    "api_gslc_insert_join_draw_arrow": {
        "intent": "insert",
        "entity": "join_draw_arrow",
        "keywords": ["Chèn mũi tên kết nối", "Thêm sơ đồ kết nối", "Kết nối đồ thị chéo", "Nhận hình ảnh"],
        "anti_patterns": ["Ống ngoài", "Công trình công cộng"],
        "description": "Chèn mũi tên kết nối(Được sử dụng cho nhiều kết nối bản vẽ)"
    },
    "api_gslc_insert_outer_pipe_arrow": {
        "intent": "insert",
        "entity": "outer_pipe_arrow",
        "keywords": ["Chèn mũi tên ống ngoài", "Ống ngoài", "mũi tên tiện ích", "mũi tên hơi nước"],
        "anti_patterns": ["Nhận hình ảnh"],
        "description": "Chèn mũi tên ống ngoài(Tiện ích đến và đi)"
    },

    # ========== Lớp truy vấn (8) ==========
    "api_gslc_query_frame_info": {
        "intent": "query",
        "entity": "frame",
        "keywords": ["Khung truy vấn", "Thông tin khung", "Thống kê khung", "dữ liệu khung"],
        "description": "Thông tin khung truy vấn và dữ liệu thống kê"
    },
    "api_gslc_query_equipment_in_frame": {
        "intent": "query",
        "entity": "equipment",
        "keywords": ["Thiết bị truy vấn", "Danh sách thiết bị", "Thiết bị trong khung tranh", "Có những thiết bị gì"],
        "description": "Truy vấn thực thể thiết bị trong khung hình"
    },
    "api_gslc_query_instrument_in_frame": {
        "intent": "query",
        "entity": "instrument",
        "keywords": ["Công cụ truy vấn", "Danh sách nhạc cụ", "Dụng cụ trong khung", "Có những nhạc cụ gì?"],
        "description": "Truy vấn các công cụ trong khung(Chứa van điều khiển)"
    },
    "api_gslc_query_pipe_in_frame": {
        "intent": "query",
        "entity": "pipe",
        "keywords": ["đường dẫn truy vấn", "Danh sách ống", "Ống trong khung", "Có những đường ống nào?"],
        "description": "Truy vấn các đường ống trong khung"
    },
    "api_gslc_query_valve_in_frame": {
        "intent": "query",
        "entity": "valve",
        "keywords": ["Van kiểm tra", "Danh sách van", "van trong khung", "Có những loại van nào?"],
        "description": "Truy vấn các van thông dụng trong khung hình(Không bao gồm van điều khiển)"
    },
    "api_gslc_query_join_draw_arrow_in_frame": {
        "intent": "query",
        "entity": "join_draw_arrow",
        "keywords": ["Mũi tên kết nối biểu đồ truy vấn", "Danh sách hình ảnh", "Có những kết nối nào?"],
        "description": "Truy vấn mũi tên nối bản vẽ trong khung bản vẽ"
    },
    "api_gslc_query_all_entities_in_frame": {
        "intent": "query",
        "entity": "frame",
        "keywords": ["Truy vấn tất cả các thực thể", "Tất cả trong khung", "Tất cả các thực thể", "tất cả các yếu tố"],
        "description": "Truy vấn tất cả trong khung hình GsLc thực thể"
    },
    "api_gslc_query_entities_by_relative_position": {
        "intent": "query",
        "entity": "equipment",
        "keywords": ["vị trí tương đối", "Các đơn vị liên quan", "các thực thể xung quanh", "thực thể lân cận"],
        "description": "Truy vấn thực thể theo vị trí tương đối(Dụng cụ xung quanh một thiết bị/Van, v.v.)"
    },

    # ========== Làm mới lớp học (3) ==========
    "api_gslc_brush_pipe_from_to": {
        "intent": "brush",
        "entity": "pipe",
        "keywords": ["Chải các đường ống đến và đi", "Nơi đường ống đến và đi", "Chải qua lại", "Ống cọfrom/to", "Thiết bị liên quan đến đường ống", "Hiệp hội đường ống", "Ống cọ", "Đường ống xả"],
        "anti_patterns": ["đồng bộ", "nhập khẩu", "Excel", "JSON"],
        "description": "Duyệt các đường ống đến và đi chỉ bằng một cú nhấp chuột: Theo dõi các thiết bị ở cả hai đầu dọc theo cấu trúc liên kết đường ống(số bit)/Mã công khai, Viết hướng của ống/Đi tới và liên kết vị trí tương đối của hai điểm",
        "differentiator": "Thẻ thiết bị được liên kết/Sự ra đời của mật danh công cộng, Không đồng bộ hóa dữ liệu bên ngoài"
    },
    "api_gslc_brush_instrument_location": {
        "intent": "brush",
        "entity": "instrument",
        "keywords": ["Chải nhạc cụ", "Vị trí dụng cụ chải", "vị trí nhạc cụ", "Hiệp hội nhạc cụ", "đồng hồ làm mới", "Vị trí công cụ ràng buộc", "Đánh răng bằng một cú nhấp chuột"],
        "anti_patterns": ["đồng bộ", "nhập khẩu"],
        "description": "Quét vị trí dụng cụ bằng một cú nhấp chuột: nhấn kết nối/Dấu vết cấu trúc liên kết đường ống liên quan đến thiết bị(số bit)hoặc ống(số ống)và ràng buộc vị trí tương đối"
    },
    "api_gslc_brush_join_draw_arrow": {
        "intent": "brush",
        "entity": "join_draw_arrow",
        "keywords": ["Sơ đồ kết nối bàn chải mũi tên", "Kết nối với liên kết đồ thị", "Làm mới sơ đồ kết nối"],
        "description": "Làm mới liên kết mũi tên sơ đồ kết nối"
    },
    # ========== Lớp xuất khẩu (6) ==========
    "api_gslc_export_all_data": {
        "intent": "export",
        "entity": "frame",
        "keywords": ["Xuất tất cả", "Xuất tất cả dữ liệu", "Xuất đầy đủ", "Xuất tất cả", "Xuất dữ liệu sơ đồ", "Xuất khẩuCADdữ liệu",
                     "Xuất thiết bị", "Dữ liệu thiết bị", "đường ống xuất khẩu", "dữ liệu đường ống", "Công cụ xuất khẩu", "Dữ liệu nhạc cụ", "van xuất khẩu", "Dữ liệu van"],
        "description": "Xuất tất cả dữ liệu sơ đồ(thiết bị+mét+đường ống+van), flat JArray Định dạng"
    },
    "api_gslc_export_frame_data": {
        "intent": "export",
        "entity": "frame",
        "keywords": ["Xuất khung", "dữ liệu khung", "Xuất khung đơn"],
        "description": "Xuất dữ liệu cho một khung đồ thị"
    },
    # Đã xóa: api_gslc_export_equipment/pipeline/instrument/valve_fitting_data
    # Các tính năng đã được tích hợp vào api_gslc_export_all_data

    # ========== Lớp đồng bộ hóa (3) ==========
    "api_gslc_sync_pipe_data": {
        "intent": "sync",
        "entity": "pipe",
        "keywords": ["Đường ống đồng bộ", "Đường ống nhập khẩu", "Cập nhật dữ liệu đường ống", "từExcelđường ống", "từJSONđường ống"],
        "anti_patterns": ["làm cho khỏe lại", "chải"],
        "description": "Đồng bộ hóa các thuộc tính đường ống từ nguồn dữ liệu bên ngoài",
        "differentiator": "từJSON/ExcelNhập thuộc tính cập nhật dữ liệu"
    },
    "api_gslc_sync_equipment_data": {
        "intent": "sync",
        "entity": "equipment",
        "keywords": ["Đồng bộ hóa thiết bị", "Nhập thiết bị", "Cập nhật dữ liệu thiết bị", "từExcelthiết bị"],
        "anti_patterns": ["làm cho khỏe lại"],
        "description": "Đồng bộ hóa thuộc tính thiết bị từ nguồn dữ liệu bên ngoài"
    },
    "api_gslc_sync_instrument_data": {
        "intent": "sync",
        "entity": "instrument",
        "keywords": ["Đồng bộ hóa nhạc cụ", "Nhập khẩu dụng cụ", "Cập nhật dữ liệu máy đo", "từExcelmét"],
        "anti_patterns": ["làm cho khỏe lại"],
        "description": "Đồng bộ hóa các thuộc tính của thiết bị từ các nguồn dữ liệu bên ngoài"
    },

    # ========== Công cụ (6) ==========
    "api_gslc_auto_number": {
        "intent": "tool",
        "entity": ["equipment", "instrument", "pipe", "valve"],
        "keywords": ["đánh số tự động", "số seri", "số seri", "sắp xếp số", "Đánh số lại"],
        "anti_patterns": ["Tạo thẻ", "Tạo thẻ", "Chèn số tham chiếu", "Tạo hàng loạt"],
        "description": "Tự động đánh số các thực thể hiện có(Sửa đổi thuộc tính số)",
        "differentiator": "Sửa đổi thuộc tính số của một thực thể hiện có, Không có thực thể mới nào được tạo"
    },
    "api_gslc_generate_equipment_list": {
        "intent": "tool",
        "entity": "equipment",
        "keywords": ["Tạo danh sách thiết bị", "Danh sách thiết bị", "Tóm tắt thiết bị", "Tổng quan về thiết bị"],
        "description": "Tạo danh sách thiết bị"
    },
    "api_gslc_validate_entity_data": {
        "intent": "tool",
        "entity": "frame",
        "keywords": ["Xác thực dữ liệu", "Kiểm tra dữ liệu", "kiểm tra", "Xác thực dữ liệu"],
        "description": "xác minh GsLc Tính toàn vẹn dữ liệu thực thể"
    },
    "api_gslc_clear_relative_position": {
        "intent": "tool",
        "entity": "equipment",
        "keywords": ["vị trí tương đối rõ ràng", "liên kết rõ ràng", "đặt lại vị trí"],
        "description": "Xóa thông tin vị trí tương đối của một thực thể"
    },
    "api_gslc_move_to_freeze_layer": {
        "intent": "tool",
        "entity": "frame",
        "keywords": ["di chuyển để đóng băng", "lớp đóng băng", "Thực thể ẩn", "đông cứng"],
        "description": "Di chuyển các thực thể sang lớp cố định"
    },
    "api_gslc_restore_from_freeze_layer": {
        "intent": "tool",
        "entity": "frame",
        "keywords": ["Phục hồi từ đóng băng", "khôi phục thực thể", "giải phóng", "tan băng"],
        "description": "Khôi phục các thực thể từ các lớp bị đóng băng"
    }
}
