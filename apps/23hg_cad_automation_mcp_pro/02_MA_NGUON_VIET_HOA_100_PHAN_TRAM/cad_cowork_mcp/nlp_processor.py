"""
NLP Processor — DEPRECATED

Mô-đun này là trình phân tích cú pháp ngôn ngữ tự nhiên của Trung Quốc dựa trên các biểu thức chính quy, để trích xuất tọa độ từ ngôn ngữ tự nhiên, màu sắc, Hành động, v.v.. 

hiện hữu MCP Dưới cấu trúc, Người gọi là LLM(Claude Chờ đợi), LLM Sự hiểu biết về ý định và cấu trúc tham số đã được hoàn thành, 
Vì vậy MCP Toolchain không còn yêu cầu bổ sung NLP lớp. Mô-đun này được dành riêng cho người không LLM khách hàng(giống Electron IPC)sử dụng, 
nhưng không có ở đó MCP server Được tham chiếu trong đường dẫn chính. 

Để phân biệt lệnh, Vui lòng sử dụng command_router/ khung(Định tuyến tên miền dựa trên quy tắc, còn hơn là NLP). 
"""

import warnings
warnings.warn(
    "nlp_processor Không được dùng nữa, MCP Chuỗi công cụ không còn được sử dụng. Xem tài liệu mô-đun để biết chi tiết. ",
    DeprecationWarning,
    stacklevel=2,
)

import logging
import re
import math
from typing import Any, Dict, List, Optional, Tuple
import json
import os
import os.path

# đọc trực tiếpconfig.jsontài liệu
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

logger = logging.getLogger('nlp_processor')

class NLPProcessor:
    """Lớp xử lý ngôn ngữ tự nhiên, Chịu trách nhiệm phân tích hướng dẫn người dùng và chuyển đổi chúng thànhCADvận hành"""
    
    def __init__(self):
        """
        khởi tạoNLPbộ xử lý
        """
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("Đã khởi tạo bộ xử lý ngôn ngữ tự nhiên")

        # Đối với các phiên bản cũ hơnCADCài đặt không được hỗ trợRGBmàu sắc, Vì vậy để tương thích, Đặt màu theo chỉ mục
        
        # Tên màu chi tiếtRGBÁnh xạ các giá trị
        self.color_rgb_map = {
            # màu cơ bản
            "màu đỏ": 1,
            "màu vàng": 2,
            "màu xanh lá": 3,
            "màu xanh da trời": 4,
            "màu xanh da trời": 5,
            "màu đỏ tươi": 6,
            "Trắng": 7,
            "xám": 8,
            "màu xám nhạt": 9,
            "đen": 250,
            "màu nâu": 251,
            "màu cam": 30,
            "Màu tím": 200,
            "hồng": 221,
            

            # màu cơ bản - phím tiếng anh
            "Red": 1,
            "Yellow": 2,
            "Green": 3,
            "Cyan": 4,
            "Blue": 5,
            "Magenta": 6,
            "White": 7,
            "Gray": 8,
            "Light Gray": 9,
            "Black": 250,
            "Brown": 251,
            "Orange": 30,
            "Purple": 200,
            "Pink": 221,
                       
        }
        
        # Ánh xạ từ khóa mở rộng
        self.shape_keywords = {
            # hình dạng cơ bản
            "đường thẳng": "line", "Dây điện": "line",
            "tròn": "circle", "tròn": "circle",
            "vòng cung": "arc", "vòng cung": "arc",
            "hình chữ nhật": "rectangle", "quảng trường": "rectangle", "quảng trường": "square",
            "đường đa tuyến": "polyline", "Đa tuyến": "polyline",
            "chữ": "text", "Từ": "text",
            "đánh dấu": "dimension", "Kích thước": "dimension",
            
        }
        
        # Ánh xạ từ khóa hành động
        self.action_keywords = {
            # chuyển động cơ bản
            "bức vẽ": "draw", "vẽ tranh": "draw", "tạo nên": "draw", "Thêm vào": "draw",
            "Ôn lại": "modify", "Điều chỉnh": "modify", "Thay đổi": "modify",
            "di chuyển": "move", "quay": "rotate", "Phóng": "scale",
            "phóng to": "scale_up", "thu nhỏ": "scale_down",
            "xóa bỏ": "erase", "Xóa": "erase", "Di dời": "erase",
            "cứu": "save",
            
            # Vận hành chuyên nghiệp
            "đánh dấu": "dimension",
            "điền vào": "hatch",
            "Tạo lớp": "create_layer",
            "Chuyển lớp": "change_layer"
        }
    
      
    def extract_color_from_command(self, command: str) -> Optional[int]:
       
        if command is None:
            return 7

        try:
            num = int(command)
            if num >= 1 and num <= 255:
                return num
        except:
            pass

        # Chuyển đổi lệnh thành chữ thường
        command = command.lower()
        
        # Cố gắng khớp tên màu
        for color_name in self.color_rgb_map.keys():
            if color_name.lower() in command:
                return self.color_rgb_map[color_name]
        
        # Cố gắng khớp với mô tả màu sắc(giống"xanh nhạt")
        color_pattern = r'([Tối hay sáng]?[a-zA-Z\u4e00-\u9fa5]+màu sắc)'
        color_matches = re.findall(color_pattern, command)
        
        for color_match in color_matches:
            # Kiểm tra xem tên màu có được biết không
            if color_match in self.color_rgb_map:
                return self.color_rgb_map[color_match]
                    
        # Nếu không tìm thấy thông tin màu sắc, trở lại7 Màu trắng mặc định
        return 7
    
    def process_command(self, command: str) -> Dict[str, Any]:
        """Xử lý các lệnh ngôn ngữ tự nhiên và trả về kết quả"""
        self.logger.info(f"lệnh xử lý: {command}")
        
        # lệnh phân tích cú pháp
        parsed_command = self.parse_command(command)
      
        # Trả về kết quả phân tích cú pháp
        return parsed_command
    
    def parse_command(self, command: str) -> Dict[str, Any]:
        """Phân tích các lệnh ngôn ngữ tự nhiên và chuyển đổi sangCADThông số vận hành"""
        self.logger.info(f"lệnh phân tích cú pháp: {command}")
        
        # Chuyển đổi lệnh thành chữ thường và loại bỏ khoảng trắng thừa
        command = command.lower().strip()
        
        # Cố gắng xác định loại lệnh
        command_type = self._identify_command_type(command)
        self.logger.debug(f"Loại lệnh được công nhận: {command_type}")
        
        # Phân phối cho các chức năng xử lý khác nhau dựa trên loại lệnh
        if command_type == "draw_line":
            return self._parse_draw_line(command)
        elif command_type == "draw_circle":
            return self._parse_draw_circle(command)
        elif command_type == "draw_arc":
            return self._parse_draw_arc(command)
        elif command_type == "draw_rectangle":
            return self._parse_draw_rectangle(command)
        elif command_type == "draw_polyline":
            return self._parse_draw_polyline(command)
        elif command_type == "draw_text":
            return self._parse_draw_text(command)
        elif command_type == "draw_hatch":
            return self._parse_draw_hatch(command)
        elif command_type == "save":
            return self._parse_save(command)
        else:
            # Trả về kết quả lỗi theo mặc định
            return {
                "type": "unknown",
                "error": "Loại lệnh không được nhận dạng",
                "original_command": command
            }
    
    def _identify_command_type(self, command: str) -> str:
        """Xác định loại lệnh"""
        # Kiểm tra loại hoạt động
        for action, action_type in self.action_keywords.items():
            if action in command:
                # Xử lý hình cơ bản
                for shape, shape_type in self.shape_keywords.items():
                    if shape in command:
                        if action_type == "draw":
                            if shape_type == "line":
                                return "draw_line"
                            elif shape_type == "circle":
                                return "draw_circle"
                            elif shape_type == "arc":
                                return "draw_arc"
                            elif shape_type in ["rectangle", "square"]:
                                return "draw_rectangle"
                            elif shape_type == "polyline":
                                return "draw_polyline"
                            elif shape_type == "text":
                                return "draw_text"
                            elif shape_type == "dimension":
                                return "add_dimension"
        
        # Kiểm tra xem đó có phải là lệnh tạo lớp không
        if "Lớp" in command and any(action in command for action in ["tạo nên", "Mới", "Thêm vào"]):
            return "create_layer"
        
        # Kiểm tra xem đó có phải là lệnh nhãn không
        if "đánh dấu" in command:
            return "add_dimension"
        
        if "cứu" in command:
            return "save"
        
        # Nếu không được công nhận, Trả về loại không xác định
        return "unknown"
    
    def _extract_coordinates(self, text: str) -> List[Tuple[float, float, float]]:
        """Trích xuất điểm tọa độ từ văn bản"""
        # Định dạng tọa độ khớp: (x,y,z) hoặc (x,y) hoặc x,y,z hoặc x,y
        pattern = r'\(?\s*(-?\d+\.?\d*)\s*,\s*(-?\d+\.?\d*)(?:\s*,\s*(-?\d+\.?\d*))?\s*\)?'
        matches = re.finditer(pattern, text)
        
        coordinates = []
        for match in matches:
            x = float(match.group(1))
            y = float(match.group(2))
            z = float(match.group(3)) if match.group(3) else 0.0
            coordinates.append((x, y, z))
        
        return coordinates
    
    def _extract_numbers(self, text: str) -> List[float]:
        """Trích xuất số từ văn bản"""
        pattern = r'(-?\d+\.?\d*)'
        matches = re.findall(pattern, text)
        return [float(match) for match in matches]
    
    def _parse_draw_line(self, command: str) -> Dict[str, Any]:
        """Phân tích lệnh vẽ đường thẳng"""
        # Cố gắng trích xuất tọa độ
        coordinates = self._extract_coordinates(command)
        
        if len(coordinates) >= 2:
            # Nếu tìm thấy ít nhất hai điểm tọa độ, Sử dụng hai điểm đầu tiên làm điểm bắt đầu và điểm kết thúc
            return {
                "type": "draw_line",
                "start_point": coordinates[0],
                "end_point": coordinates[1]
            }
        else:
            # Nếu không tìm thấy đủ điểm tọa độ, Hãy thử sử dụng giá trị mặc định
            # Tại đây bạn có thể đặt điểm bắt đầu và điểm kết thúc mặc định nếu cần
            return {
                "type": "draw_line",
                "start_point": (0, 0, 0),
                "end_point": (100, 100, 0),
                "note": "Sử dụng tọa độ mặc định, Bởi vì thông tin tọa độ không được cung cấp đầy đủ trong lệnh"
            }
    
    def _parse_draw_circle(self, command: str) -> Dict[str, Any]:
        """Phân tích lệnh vẽ vòng tròn"""
        # Cố gắng trích xuất tọa độ và bán kính
        coordinates = self._extract_coordinates(command)
        numbers = self._extract_numbers(command)
        
        # Bán kính trích xuất
        radius = None
        radius_pattern = r'(?:bán kính|r|radius)[^\d]*?(-?\d+\.?\d*)'
        radius_match = re.search(radius_pattern, command, re.IGNORECASE)
        if radius_match:
            radius = float(radius_match.group(1))
        elif len(numbers) > 0:
            # Nếu bán kính không được chỉ định rõ ràng, Sử dụng số đầu tiên được tìm thấy làm bán kính
            radius = numbers[0]
        else:
            # Bán kính mặc định
            radius = 50.0
        
        # Trích xuất điểm trung tâm
        center = None
        if len(coordinates) > 0:
            center = coordinates[0]
        else:
            # Điểm trung tâm mặc định
            center = (0, 0, 0)
        
        return {
            "type": "draw_circle",
            "center": center,
            "radius": radius
        }
    
    def _parse_draw_arc(self, command: str) -> Dict[str, Any]:
        """Phân tích lệnh vẽ vòng cung"""
        # Cố gắng trích xuất tọa độ, bán kính và góc
        coordinates = self._extract_coordinates(command)
        numbers = self._extract_numbers(command)
        
        # Trích xuất điểm trung tâm
        center = None
        if len(coordinates) > 0:
            center = coordinates[0]
        else:
            # Điểm trung tâm mặc định
            center = (0, 0, 0)
        
        # Bán kính trích xuất
        radius = None
        radius_pattern = r'(?:bán kính|r|radius)[^\d]*?(-?\d+\.?\d*)'
        radius_match = re.search(radius_pattern, command, re.IGNORECASE)
        if radius_match:
            radius = float(radius_match.group(1))
        elif len(numbers) > 0:
            # Nếu bán kính không được chỉ định rõ ràng, Sử dụng số đầu tiên được tìm thấy làm bán kính
            radius = numbers[0]
        else:
            # Bán kính mặc định
            radius = 50.0
        
        # Trích xuất góc bắt đầu và góc kết thúc
        start_angle = 0.0
        end_angle = 90.0
        
        start_angle_pattern = r'(?:góc bắt đầu|start angle)[^\d]*?(-?\d+\.?\d*)'
        start_angle_match = re.search(start_angle_pattern, command, re.IGNORECASE)
        if start_angle_match:
            start_angle = float(start_angle_match.group(1))
        
        end_angle_pattern = r'(?:góc cuối|end angle)[^\d]*?(-?\d+\.?\d*)'
        end_angle_match = re.search(end_angle_pattern, command, re.IGNORECASE)
        if end_angle_match:
            end_angle = float(end_angle_match.group(1))
        
        # Đảm bảo tất cả các tham số cần thiết đều có giá trị
        if center is None:
            center = (0, 0, 0)
        if radius is None:
            radius = 50.0
        if start_angle is None:
            start_angle = 0.0
        if end_angle is None:
            end_angle = 90.0
            
        self.logger.debug(f"Phân tích kết quả lệnh cung: điểm trung tâm={center}, bán kính={radius}, góc bắt đầu={start_angle}, góc cuối={end_angle}")
        
        return {
            "type": "draw_arc",
            "center": center,
            "radius": radius,
            "start_angle": start_angle,
            "end_angle": end_angle
        }
    
    def _parse_draw_rectangle(self, command: str) -> Dict[str, Any]:
        """Phân tích lệnh vẽ hình chữ nhật"""
        # Cố gắng trích xuất tọa độ
        coordinates = self._extract_coordinates(command)
        
        if len(coordinates) >= 2:
            # Nếu tìm thấy ít nhất hai điểm tọa độ, Sử dụng hai điểm đầu tiên làm điểm chéo
            return {
                "type": "draw_rectangle",
                "corner1": coordinates[0],
                "corner2": coordinates[1]
            }
        else:
            # Nếu không tìm thấy đủ điểm tọa độ, Cố gắng trích xuất chiều rộng và chiều cao
            width = 100.0
            height = 100.0
            
            width_pattern = r'(?:chiều rộng|width)[^\d]*?(-?\d+\.?\d*)'
            width_match = re.search(width_pattern, command, re.IGNORECASE)
            if width_match:
                width = float(width_match.group(1))
            
            height_pattern = r'(?:cao|height)[^\d]*?(-?\d+\.?\d*)'
            height_match = re.search(height_pattern, command, re.IGNORECASE)
            if height_match:
                height = float(height_match.group(1))
            
            # Nếu tìm thấy điểm tọa độ, sử dụng điều này như một điểm khởi đầu
            if len(coordinates) == 1:
                corner1 = coordinates[0]
                corner2 = (corner1[0] + width, corner1[1] + height, corner1[2])
            else:
                # Điểm bắt đầu và điểm kết thúc mặc định
                corner1 = (0, 0, 0)
                corner2 = (width, height, 0)
            
            return {
                "type": "draw_rectangle",
                "corner1": corner1,
                "corner2": corner2
            }
    
    def _parse_draw_polyline(self, command: str) -> Dict[str, Any]:
        """Phân tích cú pháp lệnh vẽ đa tuyến"""
        # Cố gắng trích xuất tọa độ
        coordinates = self._extract_coordinates(command)
        
        # Kiểm tra xem có cần đóng cửa không
        closed = "đóng cửa" in command or "đóng cửa" in command
        
        if len(coordinates) >= 2:
            # Nếu tìm thấy ít nhất hai điểm tọa độ, Sử dụng chúng làm điểm cho polylines
            return {
                "type": "draw_polyline",
                "points": coordinates,
                "closed": closed
            }
        else:
            # Nếu không tìm thấy đủ điểm tọa độ, lỗi trả về
            return {
                "type": "error",
                "message": "Vẽ một đường đa tuyến cần ít nhất hai điểm tọa độ"
            }
    
    def _parse_draw_text(self, command: str) -> Dict[str, Any]:
        """Phân tích lệnh vẽ văn bản"""
        # Cố gắng trích xuất tọa độ
        coordinates = self._extract_coordinates(command)
        
        # Trích xuất nội dung văn bản
        text_pattern = r'[nội dung văn bản|text|nội dung][: :]\s*[\"\'](.*?)[\"\']'
        text_match = re.search(text_pattern, command)
        
        text = ""
        if text_match:
            text = text_match.group(1)
        else:
            # Cố gắng trích xuất nội dung trong dấu ngoặc kép dưới dạng văn bản
            quote_pattern = r'[\"\'](.*?)[\"\']'
            quote_match = re.search(quote_pattern, command)
            if quote_match:
                text = quote_match.group(1)
            else:
                # Nếu không tìm thấy nội dung trong dấu ngoặc kép, Sử dụng văn bản mặc định
                text = "văn bản mẫu"
        
        # Trích xuất chiều cao văn bản
        height = 2.5  # Chiều cao mặc định
        height_pattern = r'(?:cao|height)[^\d]*?(-?\d+\.?\d*)'
        height_match = re.search(height_pattern, command, re.IGNORECASE)
        if height_match:
            height = float(height_match.group(1))
        
        # Trích xuất góc quay
        rotation = 0.0  # Góc mặc định
        rotation_pattern = r'(?:quay|góc|rotation)[^\d]*?(-?\d+\.?\d*)'
        rotation_match = re.search(rotation_pattern, command, re.IGNORECASE)
        if rotation_match:
            rotation = float(rotation_match.group(1))
        
        # Trích xuất điểm chèn
        position = None
        if len(coordinates) > 0:
            position = coordinates[0]
        else:
            # điểm chèn mặc định
            position = (0, 0, 0)
        
        return {
            "type": "draw_text",
            "position": position,
            "text": text,
            "height": height,
            "rotation": rotation
        }
    
    def _parse_draw_hatch(self, command: str) -> Dict[str, Any]:
        """Phân tích các lệnh điền bản vẽ"""
        # Cố gắng trích xuất một tập hợp các điểm tọa độ
        coordinates = self._extract_coordinates(command)
        
        # Trích xuất tên mẫu nở
        pattern_name = "SOLID"  # Mặc định là tô đậm
        pattern_patterns = [
            r'(?:mẫu|pattern)[^\w]*?["\'](.*?)["\']\'',
            r'(?:mẫu|pattern)[^\w]*?(\w+)'
        ]
        
        for pattern in pattern_patterns:
            pattern_match = re.search(pattern, command, re.IGNORECASE)
            if pattern_match:
                pattern_name = pattern_match.group(1).upper()
                break
        
        # Trích xuất tỷ lệ lấp đầy
        scale = 1.0  # Tỷ lệ mặc định
        scale_pattern = r'(?:Tỷ lệ|Phóng|scale)[^\d]*?(\d+\.?\d*)'
        scale_match = re.search(scale_pattern, command, re.IGNORECASE)
        if scale_match:
            scale = float(scale_match.group(1))
        
        # Kiểm tra xem có đủ điểm để xác định ranh giới điền không
        if len(coordinates) >= 3:
            self.logger.debug(f"Phân tích kết quả lệnh điền: Tập hợp điểm={coordinates}, mẫu={pattern_name}, Tỷ lệ={scale}")
            return {
                "type": "draw_hatch",
                "points": coordinates,
                "pattern_name": pattern_name,
                "scale": scale
            }
        else:
            # nếu không đủ điểm, Trả về thông báo lỗi
            self.logger.warning("Phân tích lệnh điền không thành công: Yêu cầu ít nhất3điểm để xác định ranh giới điền")
            return {
                "type": "error",
                "message": "Việc vẽ phần tô đòi hỏi ít nhất3điểm để xác định ranh giới"
            }


    def _parse_save(self, command: str) -> Dict[str, Any]:
        """Phân tích lệnh lưu"""
        # Cố gắng trích xuất đường dẫn tập tin
        path_pattern = r'(?:con đường|lưu vào|path)[^\w]*?[\"\'](.*?)[\"\']'
        path_match = re.search(path_pattern, command, re.IGNORECASE)
        
        if path_match:
            file_path = path_match.group(1)
        else:
            # Đường dẫn tệp mặc định
            file_path = os.path.join(config["output"]["directory"], config["output"]["default_filename"])
        
        return {
            "type": "save",
            "file_path": file_path
        }