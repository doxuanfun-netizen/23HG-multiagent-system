#!/usr/bin/env python
"""
Kịch bản thực thi lệnh trực tiếp

dùng cho Electron IPC gọi trực tiếp CAD Đặt hàng(đường vòng MCP giao thức)
vượt qua subprocess Gọi script này để thực thi lệnh và trả về kết quả

Để ý: Kịch bản này tạo ra một độc lập COM kết nối, Và MCP Server Các kết nối không ảnh hưởng lẫn nhau. 
Nếu bạn gặp phải MCP Ngoại lệ cuộc gọi, Có thể cần phải được kiểm tra CAD trạng thái lệnh. 

Usage:
    python direct_command.py <command_name>
    # thông số được thông qua stdin qua JSON Chuyển định dạng

Example:
    echo '{"entityTypes": ["LINE", "CIRCLE"]}' | python direct_command.py api_window_select_with_filter
"""

import sys
import json
import os

# nhập khẩu pythoncom dùng cho COM dọn dẹp
try:
    import pythoncom
    HAS_PYTHONCOM = True
except ImportError:
    HAS_PYTHONCOM = False

try:
    from .cad_controller import CADController
except ImportError:
    # Cho phép chạy trực tiếp: python direct_command.py ...
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from cad_controller import CADController


def main():
    """chức năng chính: Phân tích các tham số dòng lệnh và thực thi CAD Đặt hàng"""
    if len(sys.argv) < 2:
        print(json.dumps({
            "success": False,
            "errorCode": -2,
            "message": "Missing command name",
            "data": None
        }))
        return

    command_name = sys.argv[1]
    parameters = {}

    # từ stdin Đọc thông số JSON(gỡ rối ENAMETOOLONG câu hỏi)
    # Windows Giới hạn độ dài dòng lệnh là khoảng. 32767 tính cách, Tràn ngập tập hợp lựa chọn lớn
    stdin_data = sys.stdin.read().strip()
    if stdin_data:
        try:
            parameters = json.loads(stdin_data)
        except json.JSONDecodeError as e:
            print(json.dumps({
                "success": False,
                "errorCode": -3,
                "message": f"Invalid parameters JSON from stdin: {str(e)}",
                "data": None
            }))
            return

    controller = None
    try:
        # tạo nên CAD bộ điều khiển
        controller = CADController()

        # kết nối CAD(Phải được kết nối trước khi thực hiện lệnh)
        if not controller.start_cad():
            print(json.dumps({
                "success": False,
                "errorCode": -7,
                "message": "Không thể kết nối với CAD, xin hãy chắc chắn CAD Đã bắt đầu",
                "data": None
            }, ensure_ascii=False))
            return

        # thực hiện lệnh(Sử dụng thời gian chờ lâu hơn, Vì chụp ảnh màn hình có thể mất nhiều thời gian)
        timeout = 60.0 if "preview" in command_name.lower() else 30.0
        result = controller.execute_api_command(
            command_name=command_name,
            parameters=parameters,
            timeout=timeout
        )

        # Kết quả đầu ra(Để người gọi nắm bắt)
        print(json.dumps(result, ensure_ascii=False))

    except Exception as e:
        print(json.dumps({
            "success": False,
            "errorCode": -5,
            "message": f"Execution failed: {str(e)}",
            "data": None
        }))
    finally:
        # dọn dẹp COM kết nối(Giúp giải phóng tài nguyên)
        if controller is not None:
            controller.app = None
            controller.doc = None
        if HAS_PYTHONCOM:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass  # Bỏ qua lỗi dọn dẹp


if __name__ == "__main__":
    main()
