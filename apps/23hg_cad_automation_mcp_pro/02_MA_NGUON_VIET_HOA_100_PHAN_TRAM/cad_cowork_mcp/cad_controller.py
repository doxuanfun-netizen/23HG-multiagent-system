import logging
import math
import time
import os
import re
import json
import tempfile
from typing import Any, Dict, List, Optional, Tuple, Union
import csv
from datetime import datetime
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ===== HỆ THỐNG BẢO MẬT 23HG CAD SECURITY GUARD (WHITELIST + BLACKLIST) =====

# Danh sách regex các hàm / từ khóa nguy hiểm cấm tuyệt đối (Blacklist toàn cục)
DEFAULT_BLOCKED_REGEX = re.compile(
    r"(?i)\b("
    r"startapp"
    r"|vl-file-[a-z0-9_-]+"
    r"|vl-directory-[a-z0-9_-]+"
    r"|\b(?:arxload|netload|vbastmt|vbaload|shell|dos|sh|script)\b"
    r"|\bload\b"
    r"|_?ERASE\s+ALL"
    r"|PURGE\s+ALL"
    r")\b"
)

# Danh sách các lệnh CAD có tính chất phá hoại hoặc thoát ứng dụng
DESTRUCTIVE_COMMANDS = {
    "QUIT", "EXIT", "CLOSEALL", "CLOSE", "SAVE", "QSAVE", "SAVEAS",
    "ERASE", "PURGE", "SHELL", "DOS", "SH", "SCRIPT", "NETLOAD",
    "ARXLOAD", "VBASTMT", "VBALOAD"
}

# Danh mục Whitelist các lệnh CAD an toàn được phép thực thi
SAFE_COMMANDS_WHITELIST = {
    # Nhóm vẽ / hình học (Draw)
    "LINE", "PLINE", "CIRCLE", "ARC", "RECTANG", "RECTANGLE", "POLYGON",
    "ELLIPSE", "SPLINE", "POINT", "HATCH", "BHATCH", "-HATCH", "SOLID",
    "DONUT", "RAY", "XLINE", "REVCLOUD", "BOUNDARY", "REGION", "WIPEOUT",

    # Nhóm hiệu chỉnh (Modify)
    "MOVE", "COPY", "ROTATE", "SCALE", "OFFSET", "MIRROR", "FILLET",
    "CHAMFER", "TRIM", "EXTEND", "STRETCH", "ARRAY", "ARRAYRECT",
    "ARRAYPOLAR", "ARRAYPATH", "BREAK", "JOIN", "EXPLODE", "PEDIT",
    "ALIGN", "LENGTHEN", "ERASE",

    # Nhóm thuộc tính & Layer
    "LAYER", "COLOR", "LINETYPE", "LWEIGHT", "PROPERTIES", "CHPROP",
    "CHANGE", "MATCHPROP",

    # Nhóm Text & Kích thước (Annotation)
    "TEXT", "MTEXT", "STYLE",
    "DIM", "DIMLINEAR", "DIMALIGNED", "DIMANGULAR", "DIMRADIUS",
    "DIMDIAMETER", "DIMARC", "DIMORDINATE", "DIMCONTINUE", "DIMBASELINE",
    "DIMCENTER", "DIMREASSOCIATE", "DIMSTYLE", "DIMOVERRIDE",
    "LEADER", "QLEADER", "MLEADER", "MLEADERSTYLE",

    # Nhóm hiển thị & View
    "ZOOM", "PAN", "REGEN", "REGENALL", "VIEW", "PLAN",
    "VPOINT", "DVIEW", "UCS", "UCSICON",

    # Nhóm truy vấn & Đo đạc (Inquiry)
    "DIST", "AREA", "ID", "LIST", "MEASURE", "DIVIDE", "MASSPROP",

    # Nhóm Block & Bảng
    "BLOCK", "INSERT", "BMAKE", "TABLE", "ATTDEF", "ATTEDIT",
    "BATTMAN", "EATTEDIT", "WBLOCK",

    # Lệnh hệ thống an toàn
    "SETVAR", "GETVAR", "REDRAW", "REDRAWALL"
}

# Tương thích ngược với các module tham chiếu BLOCKED_COMMANDS_REGEX
BLOCKED_COMMANDS_REGEX = DEFAULT_BLOCKED_REGEX


# Danh sách các lệnh có tính phá hoại cao yêu cầu confirm=True hoặc dry_run=True
REQUIRE_CONFIRM_PATTERNS = re.compile(
    r"(?i)(?:_?ERASE\b|_?EXPLODE\b|(?:\b|[_.\-])LAYER\s+D(?:EL(?:ETE)?)?\b|\bMOVE\s+ALL\b|_?WBLOCK\b|\(ssget\b)"
)

# Phát hiện mã AutoLISP tiêm nhiễm (code injection) trong chuỗi lệnh
LISP_INJECTION_PATTERN = re.compile(
    r"(?i)\(\s*(?:ssget|defun|eval|read|apply|setq|startapp|vl-|command)\b"
)


def validate_command_security(command: str, confirm: bool = False, dry_run: bool = False) -> Tuple[bool, str]:
    """Kiểm tra an toàn cho chuỗi lệnh CAD trước khi gửi vào AutoCAD.

    Tuân thủ 4 nguyên tắc bảo mật tối thượng:
    1. Không cho phép gửi mã AutoLISP tùy ý từ nguồn không tin cậy. Chỉ nhận lệnh có cấu trúc.
    2. Các lệnh phá hoại có chủ đích (ERASE, EXPLODE, -LAYER D, MOVE ALL, WBLOCK, ssget) bắt buộc phải có confirm=True hoặc dry_run=True.
    3. Blacklist tuyệt đối các từ khóa nguy hiểm (startapp, shell, vl-file-*, v.v.).
    4. Whitelist nghiêm ngặt cho các nhóm lệnh hình học, hiệu chỉnh, layer và truy vấn.
    """
    if not command or not command.strip():
        return False, "Chuỗi lệnh rỗng."

    # 1. Chặn tuyệt đối mã AutoLISP tùy ý từ chuỗi lệnh bên ngoài
    stripped = command.strip()
    if stripped.startswith("(") or re.search(r"^\s*\(", command, re.MULTILINE):
        return False, (
            "Lệnh AutoLISP tùy ý bị từ chối bởi cơ chế bảo vệ an toàn 23HG Pro. "
            "Vui lòng sử dụng các API lệnh có cấu trúc (structured commands) với tham số đã kiểm tra kiểu và phạm vi."
        )

    # Chặn AutoLISP nhúng (embedded lisp injection như (ssget "X"), (eval ...))
    if LISP_INJECTION_PATTERN.search(command):
        if not confirm and not dry_run:
            return False, (
                "Phát hiện biểu thức AutoLISP nhúng trong chuỗi lệnh (ví dụ: ssget, defun, eval...). "
                "Hệ thống từ chối thực thi trừ khi có xác nhận tường minh 'confirm=True' hoặc chạy ở chế độ 'dry_run=True'."
            )

    # 2. Kiểm tra Blacklist toàn cục
    match = DEFAULT_BLOCKED_REGEX.search(command)
    if match:
        return False, f"Lệnh bị chặn bởi cơ chế bảo vệ an toàn 23HG: Phát hiện từ khóa nguy hiểm '{match.group(0)}'."

    # 3. Kiểm tra các lệnh phá hoại có chủ đích (Yêu cầu confirm=True hoặc dry_run=True)
    confirm_match = REQUIRE_CONFIRM_PATTERNS.search(command)
    if confirm_match:
        if not confirm and not dry_run:
            return False, (
                f"Cảnh báo bảo vệ dữ liệu: Lệnh có nguy cơ phá hoại cao '{confirm_match.group(0)}' "
                f"bắt buộc phải có tham số xác nhận 'confirm=True' hoặc chạy ở chế độ kiểm thử 'dry_run=True' "
                f"để tránh mất mát dữ liệu bản vẽ ngoài ý muốn."
            )

    # 4. Kiểm tra từng dòng lệnh (hỗ trợ lệnh đa dòng và phân cách bằng dấu chấm phẩy)
    raw_sub_commands = []
    for line in command.splitlines():
        for sub in line.split(";"):
            if sub.strip():
                raw_sub_commands.append(sub.strip())

    for line in raw_sub_commands:
        token_match = re.match(r"^[^\s\(\)]+", line)
        if not token_match:
            return False, f"Không xác định được lệnh CAD: {line}"
        raw_token = token_match.group(0)
        clean_verb = re.sub(r"^[_.\-']+", "", raw_token).upper()

        if clean_verb in DESTRUCTIVE_COMMANDS:
            if clean_verb in {"ERASE", "WBLOCK", "EXPLODE"} and (confirm or dry_run):
                pass
            else:
                return False, f"Lệnh bị chặn bởi cơ chế bảo vệ an toàn 23HG: Phát hiện lệnh nguy hiểm '{clean_verb}'."

        if clean_verb in SAFE_COMMANDS_WHITELIST or clean_verb.startswith("DIM"):
            continue

        return False, f"Lệnh '{clean_verb}' không nằm trong danh mục lệnh CAD an toàn (Whitelist)."

    return True, "Passed"

# đọc trực tiếpconfig.jsontài liệu
config_path = os.path.join(os.path.dirname(__file__), 'config.json')
with open(config_path, 'r', encoding='utf-8') as f:
    config = json.load(f)

try:
    import win32com.client
    # pythoncomĐúngpywin32một phần của, Không cần cài đặt riêng
    import pythoncom
except ImportError:
    logging.error("Không thể nhậpwin32com.clienthoặcpythoncom, Hãy chắc chắn rằng nó đã được cài đặtpywin32Thư viện")
    raise

logger = logging.getLogger('cad_controller')

# CAD danh sách ứng viên: (type_key, COM ProgID, tên thân thiện)
# Phát hiện các phiên bản đang chạy theo thứ tự này trong quá trình phát hiện tự động; Cấu hình rõ ràng type Ưu tiên mục này rồi roll back các mục còn lại. 
# Thứ tự là"tất cả trong thời gian chạy"mức độ ưu tiên mặc định của(AutoCAD > Hạo Thần > Trung Vương). 
CAD_CANDIDATES = [
    ("autocad", "AutoCAD.Application", "AutoCAD"),
    ("gstarcad", "GCAD.Application", "GstarCAD"),
    ("zwcad", "ZWCAD.Application", "ZWCAD"),
]


def _normalize_cad_type(cad_type: str) -> str:
    """bình thường hóa config của cad.type cho nội bộ type_key. 

    trở lại "autocad" / "gstarcad" / "zwcad"(chỉ định rõ ràng một CAD), 
    hoặc "auto"(Tự động phát hiện hoạt động CAD —— Mặc định và được đề xuất). 
    """
    t = (cad_type or "").strip().lower()
    if t in ("gcad", "gstarcad"):
        return "gstarcad"
    if t in ("autocad", "acad"):
        return "autocad"
    if t == "zwcad":
        return "zwcad"
    # "auto" / vô giá trị / không rõ → Tự động phát hiện
    return "auto"


# nhu cầu[bối cảnh ứng dụng](CommandFlags.Session)của API Lệnh: Mở/Đóng các hoạt động ở cấp độ thu thập tài liệu như bản vẽ. 
# Các lệnh này phải được CsTool của DLApiExecuteSession phân bổ, chứ không phải bình thường DLApiExecute(bối cảnh tài liệu). 
# Và CsTool kết thúc ApiCommandHandlerBase.RequiresApplicationContext=true Tương ứng một-một với các lệnh. 
APPLICATION_CONTEXT_COMMANDS = {"api_open_document", "api_close_document"}


def _is_com_disconnect(error_str: str) -> bool:
    """Xác định xem thông tin ngoại lệ có COM Kênh bị ngắt kết nối(RPC ngắt kết nối)."""
    if not error_str:
        return False
    lower = error_str.lower()
    return (
        "-2147417848" in error_str
        or "rpc_e_disconnected" in lower
        or "ngắt kết nối" in lower
        or "đã ngắt kết nối với máy khách của nó" in lower
        or "disconnected" in lower
    )


def classify_layer_trade(layer_name: str) -> tuple:
    """Phân loại Layer CAD vào các nhóm Hạng mục Tiên lượng (AEC Trades) chuẩn Việt Nam.
    Trả về: (trade_code, trade_name)
    """
    if not layer_name:
        return ("KHAC", "Khác / Chưa phân loại")

    l = layer_name.upper()

    # 1. Bê tông Móng / Cột
    if any(k in l for k in ("COT", "COL", "COLUMN", "MONG", "FOUNDATION", "FOOTING", "COC", "PILE")):
        return ("COT_MONG", "Cột, Móng & Cọc")

    # 2. Bê tông Dầm / Giằng
    if any(k in l for k in ("DAM", "BEAM", "GIANG", "LINTEL", "DAU_COT")):
        return ("BE_TONG_DAM", "Bê tông Dầm & Giằng")

    # 3. Bê tông Sàn
    if any(k in l for k in ("SAN", "SLAB", "FLOOR", "BAN_THANG", "STAIR")):
        return ("SAN_KET_CAU", "Bê tông Bản sàn & Cầu thang")

    # 4. Cốt thép
    if any(k in l for k in ("THEP", "REBAR", "BBS", "STEEL", "FE_")):
        return ("COT_THEP", "Cốt thép Xây dựng (BBS)")

    # 5. Xây tường
    if any(k in l for k in ("TUONG", "WALL", "XAY", "BRICK", "BLOCK_WALL")):
        return ("XAY_TUONG", "Xây gạch & Tường ngăn")

    # 6. Cửa & Vách kính
    if any(k in l for k in ("CUA", "DOOR", "WINDOW", "VACH", "GLAZ", "GLASS")):
        return ("CUA_VACH", "Cửa đi, Cửa sổ & Vách kính")

    # 7. Hoàn thiện lát nền, trát, ốp, sơn
    if any(k in l for k in ("HOAN_THIEN", "LAT", "OP", "TILE", "FINISH", "SON", "PAINT", "PLASTER", "TRAT")):
        return ("HOAN_THIEN", "Hoàn thiện lát nền, trát & ốp")

    # 8. Cơ điện MEP
    if any(k in l for k in ("MEP", "DIEN", "ELEC", "NUOC", "PLUMB", "HVAC", "PCCC", "FIRE")):
        return ("CO_DIEN_MEP", "Cơ điện & PCCC (MEP)")

    # 9. Kết cấu chung
    if any(k in l for k in ("KET_CAU", "STR", "STRUCT", "BE_TONG", "CONC")):
        return ("KET_CAU_CHUNG", "Kết cấu bê tông chung")

    return ("KHAC", "Khác / Chưa phân loại")


class CADController:
    """CADLớp điều khiển, chịu trách nhiệm vềCADtương tác ứng dụng"""
    
    def __init__(self):
        """khởi tạoCADbộ điều khiển"""
        self.app = None
        self.doc = None
        self.entities = {}  # Lưu trữ các tham chiếu thực thể cho đồ họa được tạo, cho những sửa đổi tiếp theo
        # Tải tham số từ tập tin cấu hình
        self.startup_wait_time = config["cad"]["startup_wait_time"]
        self.command_delay = config["cad"]["command_delay"]
        # LấyCADkiểu(mặc định AUTO: Tự động phát hiện hoạt động CAD, Tương thích với nhiều thiết bị CAD cùng tồn tại)
        # sự ưu tiên: biến môi trường CAD_COWORK_TYPE > config.json > "AUTO". 
        # Các biến môi trường cho phép người dùng ~/.claude.json của env Đặc tả bắt buộc trong khối CAD, 
        # Không cần chỉnh sửa uvx Được đóng gói trong bộ đệm config.json(Chuyển từ phiên bản cũ CAD điểm đau). 
        self.cad_type = os.environ.get("CAD_COWORK_TYPE") or config.get("cad", {}).get("type", "AUTO")
        # Loại ưa thích sau khi chuẩn hóa; "auto" Cho biết phát hiện tự động thuần túy
        self.preferred_cad_type = _normalize_cad_type(self.cad_type)
        # Kết nối thực sự thành công CAD(Điền sau khi thăm dò, Lời nhắc/Chẩn đoán hiển thị các kết nối thực sự CAD)
        self.active_cad_type: Optional[str] = None
        self.active_cad_name: Optional[str] = None
        # Danh sách các giá trị băng thông hợp lệ
        self.valid_lineweights = [0, 5, 9, 13, 15, 18, 20, 25, 30, 35, 40, 50, 53, 60, 70, 80, 90, 100, 106, 120, 140, 158, 200, 211]
        # Bộ đệm dữ liệu truy vấn bóc tách khối lượng gần nhất (tránh quét lại khi xuất BoQ)
        self.last_quantities_data: Optional[Dict[str, Any]] = None
        self.last_blocks_data: Optional[Dict[str, Any]] = None
        logger.info("CAD Controller đã khởi tạo")
    
    def _candidate_order(self) -> List[Tuple[str, str, str]]:
        """Trả về theo thứ tự ưu tiên CAD danh sách ứng viên (type_key, ProgID, tên thân thiện). 

        - preferred_cad_type == "auto": Thăm dò tất cả các ứng cử viên theo thứ tự mặc định. 
        - loại rõ ràng: Đặt loại này đầu tiên, Phần còn lại là dự phòng. 
          Vì vậy ngay cả khi config được cấu hình CAD với hoạt động thực tế CAD không nhất quán(Nhiều loại sản phẩm được cài đặt trên cùng một máy CAD), 
          Nó cũng có thể tự động kết nối với cái đang chạy, Không còn những thất bại nặng nề —— Lần này còn hơn lần này CAD Sửa lỗi cốt lõi cùng tồn tại. 
        """
        if self.preferred_cad_type == "auto":
            return list(CAD_CANDIDATES)
        preferred = [c for c in CAD_CANDIDATES if c[0] == self.preferred_cad_type]
        rest = [c for c in CAD_CANDIDATES if c[0] != self.preferred_cad_type]
        return preferred + rest

    def _connect_running_cad(self):
        """Phát hiện và kết nối để chạy CAD Ví dụ(COM GetActiveObject). 

        theo _candidate_order() Hãy thử từng cái một theo thứ tự, Kết nối với lần chạy đầu tiên CAD, 
        Ghi lại sau thành công active_cad_type / active_cad_name Đối với các từ nhắc nhở và sử dụng chẩn đoán. 

        Returns:
            COM Application sự vật; Trả về khi tất cả các ứng cử viên chưa được chạy None. 
        """
        for type_key, app_id, app_name in self._candidate_order():
            try:
                app = win32com.client.GetActiveObject(app_id)
                self.active_cad_type = type_key
                self.active_cad_name = app_name
                logger.info(f"Đã kết nối tới phần mềm CAD đang chạy: {app_name} ({app_id})")
                return app
            except Exception as e:
                logger.debug(f"Không phát hiện {app_name} ({app_id}) đang chạy: {e}")
                continue
        return None

    def disconnect(self) -> dict:
        """Giải phóng tường minh các tham chiếu COM, hủy đối tượng và gọi CoUninitialize để dọn dẹp phiên kết nối.
        Giúp triệt tiêu việc Autodesk Licensing Service tích tụ các tiến trình con ngầm (AdskLicensingAgent.exe).
        """
        try:
            self.doc = None
            if hasattr(self, "entities") and self.entities:
                self.entities.clear()
            if getattr(self, "app", None) is not None:
                try:
                    del self.app
                except Exception:
                    pass
                self.app = None
            self.active_cad_type = None
            self.active_cad_name = None
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            logger.info("Đã ngắt kết nối COM và giải phóng tài nguyên AutoCAD thành công.")
            return {"success": True, "message": "Đã ngắt kết nối COM và dọn dẹp tài nguyên AutoCAD"}
        except Exception as e:
            logger.debug(f"Lỗi khi dọn dẹp kết nối COM: {e}")
            return {"success": False, "message": f"Lỗi dọn dẹp COM: {str(e)}"}

    def __del__(self):
        try:
            self.disconnect()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()

    def start_cad(self) -> bool:
        """Kết nối tới CAD đang chạy; nếu cần sẽ khởi chạy instance mới."""
        old_app = None
        try:
            # Khởi tạo COM
            pythoncom.CoInitialize()

            # Lưu tham chiếu cũ để dọn dẹp nếu có
            if self.app is not None:
                old_app = self.app
                self.app = None
                self.doc = None

            # 1) Ưu tiên kết nối CAD đang chạy (tự động dò tìm AutoCAD / GstarCAD / ZWCAD)
            self.app = self._connect_running_cad()

            # 2) Không có CAD nào đang chạy -> khởi động theo cấu hình
            if self.app is None:
                if self.preferred_cad_type == "auto":
                    logger.error(
                        "Không phát hiện CAD đang chạy (đã thử AutoCAD/GstarCAD/ZWCAD); "
                        "Chế độ AUTO không tự khởi chạy, vui lòng mở CAD và mở sẵn bản vẽ."
                    )
                    return False
                # Khởi động loại CAD chỉ định
                type_key, app_id, app_name = next(
                    (c for c in CAD_CANDIDATES if c[0] == self.preferred_cad_type),
                    CAD_CANDIDATES[0],
                )
                logger.info(f"Không phát hiện CAD đang chạy, đang khởi động {app_name} ({app_id})...")
                try:
                    self.app = win32com.client.Dispatch(app_id)
                    self.app.Visible = True
                    self.active_cad_type = type_key
                    self.active_cad_name = app_name
                    # Đợi CAD khởi động
                    time.sleep(self.startup_wait_time)
                except Exception as new_app_ex:
                    logger.error(f"Khởi động {app_name} thất bại: {new_app_ex}")
                    return False

            # 3) Lấy hoặc tạo bản vẽ hiện hành
            try:
                if self.app.Documents.Count == 0:
                    logger.info("Đang tạo bản vẽ mới...")
                    self.doc = self.app.Documents.Add()
                else:
                    logger.info("Đang lấy bản vẽ hiện hành...")
                    self.doc = self.app.ActiveDocument
            except Exception as doc_ex:
                logger.warning(f"Lấy bản vẽ thất bại, thử tạo bản vẽ mới: {str(doc_ex)}")
                try:
                    self.doc = self.app.Documents.Add()
                except Exception as new_doc_ex:
                    logger.error(f"Tạo bản vẽ mới thất bại: {str(new_doc_ex)}")
                    return False

            # Đợi CAD hoàn tất xử lý
            time.sleep(2)

            if self.doc is None:
                raise Exception("Không thể lấy đối tượng ActiveDocument hợp lệ")

            # Cố gắng đọc thuộc tính tài liệu để xác minh tính hợp lệ của chúng
            try:
                name = self.doc.Name
                logger.info(f"Tên tài liệu: {name}")
            except Exception as name_ex:
                logger.error(f"Không thể đọc tên tài liệu: {str(name_ex)}")
                raise Exception("Đối tượng tài liệu không hợp lệ")

            logger.info(f"CAD đã kết nối thành công và sẵn sàng ({self.active_cad_name or 'CAD'})")
            return True

        except Exception as e:
            logger.error(f"Khởi động CAD thất bại: {str(e)}")
            return False
        finally:
            # Dọn dẹp các phiên bản cũ
            if old_app is not None:
                try:
                    del old_app
                except Exception:
                    pass


    def is_running(self) -> bool:
        """nghiên cứuCADNó đang chạy à(xác minh COM giá trị đối tượng)"""
        if self.app is None or self.doc is None:
            return False

        # Xác minh đang hoạt động COM Đối tượng vẫn còn giá trị?
        try:
            _ = self.doc.Name  # Truy cập tài sản nhanh chóng không có tác dụng phụ, sẽ kích hoạt RPC gọi 
            return True
        except Exception as e:
            logger.warning(f"Kết nối COM không hợp lệ: {e}")
            return False

    def _refresh_active_document(self) -> bool:
        """Làm mới các tham chiếu tài liệu đang hoạt động, Hãy chắc chắn rằng bạn đang trỏ đến bản vẽ hiện đang mở

        Khi người dùng CAD Khi chuyển đổi bản vẽ trong, self.doc Có thể trỏ đến một tài liệu không còn hoạt động. 
        Phương pháp này phát hiện và cập nhật self.doc Trỏ tới tài liệu hiện đang hoạt động. 

        Returns:
            bool: Làm mới trả về thành công True, Trở lại khi thất bại False
        """
        if self.app is None:
            logger.error("Đối tượng Application rỗng, không thể làm mới tài liệu")
            return False

        try:
            # Xác minh đầu tiên Application Đối tượng có hợp lệ không?
            try:
                _ = self.app.Documents.Count
            except Exception as app_e:
                # Application Đối tượng đã hết hạn, Cần kết nối lại
                logger.warning(f"Đối tượng Application đã mất hiệu lực: {app_e}")
                return False  # Hãy để người gọi kích hoạt kết nối lại

            # Lấy tài liệu hiện đang hoạt động
            new_doc = self.app.ActiveDocument

            if new_doc is None:
                # Kiểm tra xem có tài liệu đang mở không
                if self.app.Documents.Count > 0:
                    new_doc = self.app.Documents.Item(0)
                    logger.info("ActiveDocument rỗng, sử dụng tài liệu mở đầu tiên")
                else:
                    logger.warning("Hiện tại không có tài liệu nào đang mở")
                    return False

            # Kiểm tra xem có cần cập nhật không(Xác định theo tên tài liệu, Tránh so sánh trực tiếp COM sự vật)
            try:
                old_name = self.doc.Name if self.doc else None
            except Exception:
                old_name = None  # Giấy tờ cũ không còn giá trị

            try:
                new_name = new_doc.Name
            except Exception as e:
                logger.error(f"Không thể lấy tên tài liệu mới: {e}")
                return False

            # Cập nhật tài liệu tham khảo
            if old_name != new_name:
                logger.info(f"Phát hiện chuyển đổi bản vẽ: '{old_name}' -> '{new_name}'")
                self.doc = new_doc
                # Làm sạch bộ đệm thực thể cũ(bởi vì thực thể Handle Có thể khác nhau trong các bản vẽ khác nhau)
                self.entities = {}
                logger.info("Đã cập nhật tham chiếu Document, làm sạch cache")
            elif self.doc is None:
                # nếu như self.doc Trống nhưng có cùng tên(Thiết lập lần đầu)
                self.doc = new_doc
                logger.info(f"Thiết lập tham chiếu Document: {new_name}")

            return True

        except Exception as e:
            error_str = str(e)
            # Kiểm tra xem có phải không COM Lỗi ngắt kết nối
            if "-2147417848" in error_str or "RPC_E_DISCONNECTED" in error_str.upper():
                logger.warning(f"Kết nối COM đã ngắt: {e}")
                return False  # Hãy để người gọi kích hoạt kết nối lại

            logger.error(f"Làm mới tài liệu hiện hành thất bại: {e}")
            return False

    def _try_reconnect(self) -> bool:
        """Hãy thử kết nối lại với CAD

        khi COM Khi mất kết nối, Hãy thử thiết lập lại kết nối. 
        Sẽ dọn dẹp cái cũ trước COM sự vật, và sau đó lấy lại. 

        Returns:
            bool: Kết nối lại thành công và quay trở lại True, Trở lại khi thất bại False
        """
        logger.info("Đang thử kết nối lại CAD...")

        try:
            # Làm sạch cái cũ trước COM tham chiếu đối tượng
            self.app = None
            self.doc = None
            self.entities = {}

            # Khởi tạo lại COM(Đảm bảo khởi tạo trong chuỗi hiện tại)
            try:
                pythoncom.CoInitialize()
            except Exception as com_init_e:
                logger.debug(f"COM đã khởi tạo: {com_init_e}")

            # Tự động phát hiện hoạt động CAD(sang AutoCAD/Hạo Thần/Trung Vương)
            self.app = self._connect_running_cad()

            # Không có phiên bản nào đang chạy: Các loại rõ ràng có thể Dispatch khởi động; AUTO chế độ từ bỏ kết nối lại
            if self.app is None:
                if self.preferred_cad_type == "auto":
                    logger.warning("Kết nối lại thất bại: Không phát hiện CAD đang chạy (AutoCAD/GstarCAD/ZWCAD)")
                    return False
                type_key, app_id, app_name = next(
                    (c for c in CAD_CANDIDATES if c[0] == self.preferred_cad_type),
                    CAD_CANDIDATES[0],
                )
                logger.warning(f"Không thấy CAD đang chạy, thử Dispatch khởi động {app_name}...")
                try:
                    self.app = win32com.client.Dispatch(app_id)
                    self.active_cad_type = type_key
                    self.active_cad_name = app_name
                    logger.info("Dispatch thành công")
                except Exception as disp_e:
                    logger.error(f"Khởi động Dispatch thất bại: {disp_e}")
                    return False

            # Đảm bảo ứng dụng hiển thị
            try:
                self.app.Visible = True
            except Exception:
                pass

            # Nhận tài liệu hoạt động
            self.doc = self.app.ActiveDocument

            if self.doc is None:
                # Kiểm tra xem có tài liệu đang mở không
                if self.app.Documents.Count == 0:
                    logger.warning("Không có tài liệu nào đang mở trong CAD")
                    return False
                else:
                    self.doc = self.app.Documents.Item(0)

            if self.doc is None:
                logger.warning("Kết nối lại thành công nhưng không lấy được tài liệu")
                return False

            # Xác minh kết nối
            doc_name = self.doc.Name
            logger.info(f"Kết nối lại thành công ({self.active_cad_name or 'CAD'}), tài liệu hiện hành: {doc_name}")

            return True

        except Exception as e:
            logger.error(f"Kết nối lại thất bại: {e}")
            self.app = None
            self.doc = None
            return False

    def _ensure_valid_connection(self) -> bool:
        """Đảm bảo có giá trị hợp lệ CAD kết nối

        Gọi phương thức này trước khi thực hiện thao tác, sẽ được xử lý tự động: 
        1. Kiểm tra xem kết nối có hợp lệ không
        2. Làm mới tài liệu tham khảo tài liệu(Xử lý chuyển đổi bản vẽ)
        3. Cố gắng kết nối lại nếu cần thiết

        Returns:
            bool: Kết nối hợp lệ và trả về True, Nếu không thì quay lại False
        """
        # Đầu tiên hãy kiểm tra kết nối cơ bản
        if self.app is None:
            return self._try_reconnect()

        # Làm mới tài liệu tham khảo tài liệu(Xử lý các tình huống chuyển đổi bản vẽ)
        if not self._refresh_active_document():
            return self._try_reconnect()

        return True

    def is_quiescent(self) -> bool:
        """Kiểm tra xem AutoCAD có đang ở trạng thái rảnh rỗi (Quiescent) hay đang kẹt hộp thoại Modal Dialog."""
        if self.app is None:
            return False
        try:
            state = self.app.GetAcadState()
            return bool(state.IsQuiescent)
        except Exception:
            return True

    def _raw_send_command_gateway(self, command: str, dry_run: bool = False, confirm: bool = False) -> dict:
        """Cổng bảo mật duy nhất thực thi SendCommand tới AutoCAD.

        Mọi thao tác gửi lệnh cấp thấp bắt buộc phải đi qua cổng này để:
        1. Kiểm tra Whitelist + Blacklist bảo vệ an toàn (Security Guard).
        2. Kiểm tra xác nhận cho các lệnh phá hoại (confirm=True).
        3. Kiểm tra kết nối COM và trạng thái Quiescent (tránh treo do Modal Dialog).
        4. Tự động phục hồi kết nối COM khi bị ngắt.

        Args:
            command: chuỗi lệnh để thực thi
            dry_run: nếu True chỉ kiểm tra an toàn và cú pháp
            confirm: nếu True cho phép các lệnh phá hoại được phê duyệt chạy

        Returns:
            dict: Kết quả thực thi {success, message}
        """
        # 1. Kiểm tra an toàn bảo mật (Security Guard)
        is_safe, sec_msg = validate_command_security(command, confirm=confirm, dry_run=dry_run)
        if not is_safe:
            logger.warning(sec_msg)
            return {"success": False, "message": sec_msg}

        if dry_run:
            return {"success": True, "dry_run": True, "message": f"Kiểm tra an toàn hợp lệ (Passed): {command.strip()}"}

        # 2. Đảm bảo có kết nối hợp lệ
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD không chạy hoặc bị ngắt kết nối. Vui lòng mở AutoCAD và mở một bản vẽ."}

        # 3. Kiểm tra trạng thái Modal Dialog / Quiescent
        if not self.is_quiescent():
            return {"success": False, "message": "AutoCAD đang bận hoặc đang hiển thị hộp thoại cảnh báo (Modal Dialog). Vui lòng đóng hộp thoại trước khi gửi lệnh."}

        try:
            # Đảm bảo lệnh kết thúc bằng dấu xuống dòng
            if not command.endswith("\n"):
                command = command + "\n"

            # Gửi lệnh trực tiếp qua COM
            self.doc.SendCommand(command)

            # Chờ quá trình thực thi lệnh hoàn tất
            time.sleep(self.command_delay)

            logger.info(f"Lệnh đã gửi qua Gateway: {command.strip()}")
            return {"success": True, "message": f"Lệnh '{command.strip()}' đã thực thi thành công"}

        except Exception as e:
            error_msg = str(e)
            # Kiểm tra xem có phải lỗi ngắt kết nối COM không
            if "-2147417848" in error_msg or "RPC_E_DISCONNECTED" in error_msg.upper():
                logger.warning("Phát hiện mất kết nối COM trong Gateway, đang thử kết nối lại...")
                if self._try_reconnect():
                    try:
                        self.doc.SendCommand(command)
                        time.sleep(self.command_delay)
                        logger.info(f"Lệnh đã gửi lại sau khi phục hồi COM: {command.strip()}")
                        return {"success": True, "message": f"Lệnh '{command.strip()}' đã thực thi thành công (sau khi kết nối lại)"}
                    except Exception as retry_e:
                        error_msg = f"Gửi lệnh không thành công sau khi kết nối lại: {str(retry_e)}"

            logger.error(f"Gửi lệnh thất bại tại Gateway: {error_msg}")
            return {"success": False, "message": f"Không gửi được lệnh: {error_msg}"}

    def send_command(self, command: str, dry_run: bool = False, confirm: bool = False) -> dict:
        """Gửi lệnh tới CAD qua SendCommand có cơ chế bảo vệ an toàn Security Guard.

        Ủy quyền toàn bộ cho cổng kiểm soát an ninh trung tâm _raw_send_command_gateway.
        """
        return self._raw_send_command_gateway(command, dry_run=dry_run, confirm=confirm)

    def save_drawing(self, file_path: str) -> bool:
        """Lưu bản vẽ hiện tại vào đường dẫn đã chỉ định"""
        if not self.is_running():
            logger.error("CAD chưa chạy, không thể lưu bản vẽ")
            return False

        try:
            # Đảm bảo thư mục tồn tại (bỏ qua khi chỉ là tên file, không có thư mục)
            parent = os.path.dirname(os.path.abspath(file_path))
            if parent:
                os.makedirs(parent, exist_ok=True)

            # lưu tập tin
            self.doc.SaveAs(file_path)
            logger.info(f"Bản vẽ đã lưu tại: {file_path}")

            return True
        except Exception as e:
            logger.error(f"Lưu bản vẽ thất bại: {str(e)}")
            return False

    def open_document(self, file_path: str) -> dict:
        """Mở bản vẽ DWG/DXF bằng Native COM (Documents.Open)."""
        if not file_path or not os.path.isfile(file_path):
            return {"success": False, "errorCode": -2,
                    "message": f"Không tìm thấy file bản vẽ: {file_path}"}
        if not self.is_running():
            started = self.start_cad()
            if not started:
                return {"success": False, "errorCode": -8,
                        "message": "Không tìm thấy CAD đang chạy. Vui lòng mở CAD rồi thử lại."}
        try:
            abs_path = os.path.abspath(file_path)
            doc = self.app.Documents.Open(abs_path)
            self.doc = doc
            try:
                name = doc.Name
            except Exception:
                name = os.path.basename(abs_path)
            logger.info(f"Đã mở bản vẽ: {abs_path}")
            return {"success": True, "message": f"Đã mở bản vẽ: {name}",
                    "data": {"name": name, "path": abs_path},
                    "executionMethod": "direct_com"}
        except Exception as e:
            logger.error(f"Mở bản vẽ thất bại: {e}")
            return {"success": False, "errorCode": -5,
                    "message": f"Mở bản vẽ thất bại: {e}"}

    def close_document(self, save_changes: bool = False) -> dict:
        """Đóng bản vẽ hiện hành bằng Native COM (Document.Close)."""
        if not self.is_running() or self.doc is None:
            return {"success": False, "errorCode": -8,
                    "message": "Chưa có bản vẽ nào đang mở."}
        try:
            self.doc.Close(save_changes)
            logger.info(f"Đã đóng bản vẽ (save={save_changes}).")
            # Trỏ lại bản vẽ còn mở (nếu còn), nếu không thì None
            try:
                if self.app.Documents.Count > 0:
                    self.doc = self.app.ActiveDocument
                else:
                    self.doc = None
            except Exception:
                self.doc = None
            return {"success": True,
                    "message": f"Đã đóng bản vẽ (save={save_changes}).",
                    "executionMethod": "direct_com"}
        except Exception as e:
            logger.error(f"Đóng bản vẽ thất bại: {e}")
            return {"success": False, "errorCode": -5,
                    "message": f"Đóng bản vẽ thất bại: {e}"}

    def backup_current_drawing(self, backup_dir: Optional[str] = None) -> Optional[str]:
        """Tự động tạo bản sao lưu (.bak) của bản vẽ hiện tại trước các thao tác hàng loạt hoặc nguy hiểm."""
        if not self.is_running() or not self.doc:
            return None
        try:
            fullname = getattr(self.doc, "FullName", None)
            if fullname and os.path.isfile(fullname):
                import shutil
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                dirname = backup_dir or os.path.join(os.path.dirname(fullname), ".23hg_backup")
                os.makedirs(dirname, exist_ok=True)
                filename = os.path.basename(fullname)
                base, ext = os.path.splitext(filename)
                backup_path = os.path.join(dirname, f"{base}_backup_{timestamp}{ext}")
                shutil.copy2(fullname, backup_path)
                logger.info(f"[AUTO-BACKUP] Đã tự động sao lưu bản vẽ tới: {backup_path}")
                return backup_path
            else:
                import tempfile
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                temp_backup = os.path.join(tempfile.gettempdir(), f"23hg_cad_autobackup_{timestamp}.dwg")
                try:
                    self.doc.SaveAs(temp_backup)
                    logger.info(f"[AUTO-BACKUP] Đã sao lưu bản vẽ chưa lưu tới file tạm: {temp_backup}")
                    return temp_backup
                except Exception:
                    return None
        except Exception as e:
            logger.warning(f"[AUTO-BACKUP] Không thể tự động sao lưu bản vẽ: {e}")
        return None
    
    def refresh_view(self) -> None:
        """làm cho khỏe lạiCADxem"""
        if self.is_running():
            try:
                self.doc.Regen(1)  # acAllViewports = 1
            except Exception as e:
                logger.error(f"Làm mới khung nhìn thất bại: {str(e)}")
    
    def validate_lineweight(self, lineweight) -> int:
        """Xác thực và trả về giá trị độ rộng dòng hợp lệ
        
        Nếu giá trị độ rộng dòng được cung cấp không có trong danh sách các giá trị hợp lệ, sau đó trả về giá trị mặc định0
        
        Args:
            lineweight: Giá trị độ rộng dòng để xác minh
            
        Returns:
            Giá trị độ rộng dòng hợp lệ
        """
        if lineweight is None:
            return None
            
        # Kiểm tra xem độ rộng dòng có nằm trong danh sách các giá trị hợp lệ không
        if lineweight in self.valid_lineweights:
            return lineweight
        else:
            logger.warning(f"Giá trị độ rộng dòng {lineweight} không hợp lệ, Giá trị mặc định sẽ được sử dụng 0")
            return 0
    
    def draw_line(self, start_point: Tuple[float, float, float], 
                 end_point: Tuple[float, float, float], layer: str = None, color: int = None, lineweight=None) -> bool:
        """Vẽ một đường thẳng"""
        if not self.is_running():
            return False
            
        try:
            # Đảm bảo các điểm là ba chiều
            if len(start_point) == 2:
                start_point = (start_point[0], start_point[1], 0)
            if len(end_point) == 2:
                end_point = (end_point[0], end_point[1], 0)
      
            # sử dụngVARIANTGói dữ liệu điểm tọa độ
            start_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [start_point[0], start_point[1], start_point[2]])
            end_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                             [end_point[0], end_point[1], end_point[2]])
            
            # Thêm đường thẳng
            line = self.doc.ModelSpace.AddLine(start_array, end_array)
            
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                line.Layer = layer
         
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                line.Color = color

            if lineweight is not None:
                line.LineWeight = self.validate_lineweight(lineweight)
            
            # làm mới chế độ xem
            self.refresh_view()

            logger.debug(f"Đường thẳng được vẽ: điểm bắt đầu{start_point}, kết thúc{end_point}, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return line
            
        except Exception as e:
            logger.error(f"Lỗi khi vẽ đường thẳng: {str(e)}")
            return None
    
    def draw_circle(self, center: Tuple[float, float, float], 
                   radius: float, layer: str = None, color: int = None, lineweight=None) -> Any:
        """Vẽ một vòng tròn"""
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo các điểm là ba chiều
            if len(center) == 2:
                center = (center[0], center[1], 0)
            
            # sử dụngVARIANTGói dữ liệu điểm tọa độ
            center_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [center[0], center[1], center[2]])
            
            # thêm vòng tròn
            circle = self.doc.ModelSpace.AddCircle(center_array, radius)
            
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                circle.Layer = layer
            
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                circle.Color = color

            if lineweight is not None:
                circle.LineWeight = self.validate_lineweight(lineweight)
            
            # làm mới chế độ xem
            self.refresh_view()
            
            logger.debug(f"Vẽ vòng tròn: trung tâm{center}, bán kính{radius}, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return circle
            
        except Exception as e:
            logger.error(f"Lỗi khi vẽ hình tròn: {str(e)}")
            return None
    
    def draw_arc(self, center: Tuple[float, float, float], 
                radius: float, start_angle: float, end_angle: float, layer: str = None, color: int = None, lineweight=None) -> Any:
        """Vẽ vòng cung"""
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo các điểm là ba chiều
            if len(center) == 2:
                center = (center[0], center[1], 0)
                
            # Chuyển đổi góc sang radian
            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)
            
            # sử dụngVARIANTGói dữ liệu điểm tọa độ
            center_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [center[0], center[1], center[2]])
            
            # Thêm vòng cung
            arc = self.doc.ModelSpace.AddArc(center_array, radius, start_rad, end_rad)
            
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                arc.Layer = layer
            
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                arc.Color = color

            if lineweight is not None:
                arc.LineWeight = self.validate_lineweight(lineweight)
            
            # làm mới chế độ xem
            self.refresh_view()
            
            logger.debug(f"Vẽ vòng cung: trung tâm{center}, bán kính{radius}, góc bắt đầu{start_angle}, góc cuối{end_angle}, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return arc
        except Exception as e:
            logger.error(f"Vẽ cung tròn thất bại: {str(e)}")
            return None
 
    def draw_ellipse(self, center: Tuple[float, float, float], 
                    major_axis: float, minor_axis: float, rotation: float = 0, 
                    layer: str = None, color: int = None, lineweight=None) -> Any:
        """Vẽ một hình elip"""
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo các điểm là ba chiều
            if len(center) == 2:
                center = (center[0], center[1], 0)
            
            if rotation is None:
                rotation = 0

            # Chuyển đổi góc quay sang radian
            rotation_rad = math.radians(rotation)
            
            # sử dụngVARIANTGói dữ liệu điểm tọa độ
            center_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [center[0], center[1], center[2]])
            
            # Tính vectơ trục chính của hình elip
            major_x = major_axis * math.cos(rotation_rad)
            major_y = major_axis * math.sin(rotation_rad)
            major_vector = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [major_x, major_y, 0])
            
            # Thêm hình elip
            ellipse = self.doc.ModelSpace.AddEllipse(center_array, major_vector, minor_axis / major_axis)
            
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                ellipse.Layer = layer
            
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                ellipse.Color = color

            if lineweight is not None:
                ellipse.LineWeight = self.validate_lineweight(lineweight)
            
            # làm mới chế độ xem
            self.refresh_view()
            
            logger.debug(f"Vẽ hình elip: trung tâm{center}, trục dài{major_axis}, trục ngắn{minor_axis}, góc quay{rotation}, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return ellipse
        except Exception as e:
            logger.error(f"Vẽ elip thất bại: {str(e)}")
            return None
    
    def draw_polyline(self, points: List[Tuple[float, float, float]], closed: bool = False, layer: str = None, color: int = None, lineweight=None) -> Any:
        """Vẽ đa tuyến"""
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo tất cả các điểm đều là 3D
            processed_points = []
            for point in points:
                if len(point) == 2:
                    processed_points.append((point[0], point[1], 0))
                else:
                    processed_points.append(point)
            
            # Tạo mảng điểm
            point_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                [coord for point in processed_points for coord in point])
            
            # Thêm đường đa tuyến
            polyline = self.doc.ModelSpace.AddPolyline(point_array)
            
            # Nếu cần đóng
            if closed and len(processed_points) > 2:
                polyline.Closed = True
            
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                polyline.Layer = layer
            
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                polyline.Color = color

            if lineweight is not None:
                polyline.LineWeight = self.validate_lineweight(lineweight)
            
            # làm mới chế độ xem
            self.refresh_view()

            logger.debug(f"Vẽ đa tuyến: {len(points)}điểm, {'đóng cửa' if closed else 'không đóng cửa'}, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return polyline
        except Exception as e:
            logger.error(f"Lỗi khi vẽ đường đa tuyến: {str(e)}")
            return None
    
    def draw_rectangle(self, corner1: Tuple[float, float, float], 
                      corner2: Tuple[float, float, float], layer: str = None, color: int = None, lineweight=None) -> Any:
        """Vẽ một hình chữ nhật"""
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo các điểm là ba chiều
            if len(corner1) == 2:
                corner1 = (corner1[0], corner1[1], 0)
            if len(corner2) == 2:
                corner2 = (corner2[0], corner2[1], 0)
                
            # Tính bốn điểm góc của hình chữ nhật
            x1, y1, z1 = corner1
            x2, y2, z2 = corner2
            
            # Tạo bốn điểm của hình chữ nhật
            points = [
                (x1, y1, z1),
                (x2, y1, z1),
                (x2, y2, z1),
                (x1, y2, z1),
                (x1, y1, z1)  # hình chữ nhật khép kín
            ]
            
            # Vẽ hình chữ nhật bằng polylines
            return self.draw_polyline(points, True, layer, color, lineweight)
        except Exception as e:
            logger.error(f"Lỗi khi vẽ hình chữ nhật: {str(e)}")
            return None
    
    def draw_text(self, position: Tuple[float, float, float], 
                 text: str, height: float = 2.5, rotation: float = 0, layer: str = None, color: int = None) -> Any:
        """Thêm văn bản"""
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo các điểm là ba chiều
            if len(position) == 2:
                position = (position[0], position[1], 0)
            
            # sử dụngVARIANTGói dữ liệu điểm tọa độ
            position_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                 [position[0], position[1], position[2]])
                
            # Thêm văn bản
            text_obj = self.doc.ModelSpace.AddText(text, position_array, height)
            
            # Đặt góc quay
            if rotation != 0:
                text_obj.Rotation = math.radians(rotation)
            
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                text_obj.Layer = layer
            
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                text_obj.Color = color
            
            # làm mới chế độ xem
            self.refresh_view()
                                
            logger.debug(f"Đã thêm văn bản: '{text}', Vị trí{position}, cao{height}, quay{rotation}Tiêu, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return text_obj
        except Exception as e:
            logger.error(f"Lỗi khi thêm văn bản: {str(e)}")
            return None
    
    def draw_hatch(self, points: List[Tuple[float, float, float]], 
                  pattern_name: str = "SOLID", scale: float = 1.0, layer: str = None, color: int = None) -> Any:
        """Vẽ một mẫu điền
        
        Args:
            points: Tập hợp các điểm lấp đầy ranh giới, Mỗi điểm là một bộ tọa độ hai chiều hoặc ba chiều
            pattern_name: Điền tên mẫu, Mặc định là"SOLID"(Điền thực thể)
            scale: Tỷ lệ điền vào mẫu, Mặc định là1.0
            layer: Tên lớp, nếu vìNonesau đó sử dụng lớp hiện tại
            color: chỉ số màu, nếu vìNonesau đó sử dụng màu mặc định
            
        Returns:
            Trả về thành công đối tượng đã điền, Trở lại khi thất bạiNone
        """
        if not self.is_running():
            return None
            
        try:
            # Đảm bảo tất cả các điểm đều hợp lệ
            if not points or len(points) < 3:
                logger.error("Tạo mặt cắt thất bại: Cần tối thiểu 3 điểm định nghĩa ranh giới")
                return None
                
            # Tạo các đường polyline khép kín làm ranh giới
            closed_polyline = self.draw_polyline(points, closed=True, layer=layer)
            if not closed_polyline:
                logger.error("Tạo mặt cắt thất bại: Không thể tạo polyline ranh giới")
                return None
                
            # Tạo một đối tượng điền (0Cho biết làm đầy bình thường, TrueThể hiện ranh giới liên kết)
            hatch = self.doc.ModelSpace.AddHatch(0, pattern_name, True)
                
            # Thêm vòng ranh giới bên ngoài
            # sử dụngVARIANTmảng các đối tượng bao bọc
            object_ids = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [closed_polyline])
            hatch.AppendOuterLoop(object_ids)
                
            # Đặt tỷ lệ mẫu tô màu
            hatch.PatternScale = scale
                
            # Nếu một lớp được chỉ định, Đặt lớp
            if layer:
                # Đảm bảo lớp tồn tại
                self.create_layer(layer)
                # Đặt lớp thực thể
                hatch.Layer = layer
                
            # Nếu màu sắc được chỉ định, Đặt màu
            if color is not None:
                hatch.Color = color
                                
            # cập nhật điền (Tính diện tích lấp đầy)
            hatch.Evaluate()
            
            # làm mới chế độ xem
            self.refresh_view()
                
            logger.debug(f"Đã tạo điền: mẫu {pattern_name}, Tỷ lệ {scale}, Lớp{layer if layer else 'mặc định'}, màu sắc{color if color is not None else 'mặc định'}")
            return hatch
        except Exception as e:
            logger.error(f"Lỗi khi tạo mặt cắt: {str(e)}")
            return None
    
    def zoom_extents(self) -> bool:
        """Thu phóng chế độ xem để hiển thị tất cả các đối tượng"""
        if not self.is_running():
            return False
            
        try:
            self.doc.ActiveViewport.ZoomExtents()
            logger.info("Đã zoom extents toàn bộ đối tượng")
            return True
        except Exception as e:
            logger.error(f"Lỗi khi zoom extents: {str(e)}")
            return False
    
    def close(self) -> None:
        """đóng cửaCADbộ điều khiển"""
        try:
            # giải phóngCOMtài nguyên
            if self.app is not None:
                del self.app
            pythoncom.CoUninitialize()
        except:
            pass

    
    def create_layer(self, layer_name: str) -> bool:    # , color: Union[int, Tuple[int, int, int]] = 7
        """Tạo lớp mới
        
        Args:
            layer_name: Tên lớp
            color: giá trị màu, có thểCADchỉ số màu(int)hoặcRGBgiá trị màu(tuple)
            
        Returns:
            Ca phẫu thuật có thành công không?
        """
        if not self.is_running():
            return False
        
        try:
            # Kiểm tra xem lớp đã tồn tại chưa
            for i in range(self.doc.Layers.Count):
                if self.doc.Layers.Item(i).Name == layer_name:
                    # Lớp đã tồn tại, kích hoạt nó
                    self.doc.ActiveLayer = self.doc.Layers.Item(i)
                    return True
                
            # Tạo lớp mới
            new_layer = self.doc.Layers.Add(layer_name)
            
            # Lớp không đặt màu, Đặt màu của thực thể bên trong
            # # Đặt màu
            # if isinstance(color, int):
            #     # Sử dụng chỉ số màu
            #     new_layer.Color = color
            # elif isinstance(color, tuple) and len(color) == 3:
            #     # sử dụngRGBgiá trị
            #     r, g, b = color
            #     # cài đặtTrueColor
            #     new_layer.TrueColor = self._create_true_color(r, g, b)
            
            # Đặt làm lớp hiện tại
            self.doc.ActiveLayer = new_layer
            logger.info(f"Đã tạo Layer mới: {layer_name}")  #, màu sắc: {color}
            return True
        except Exception as e:
            logger.error(f"Lỗi khi tạo Layer: {str(e)}")
            return False

    def add_dimension(self, start_point: Tuple[float, float, float], 
                     end_point: Tuple[float, float, float],
                     text_position: Tuple[float, float, float] = None, textheight: float = 5,layer: str = None, color: int=None) -> Any:
            """Thêm kích thước tuyến tính"""
            if not self.is_running():
                return None
            
            try:
                # Đảm bảo các điểm là ba chiều
                if len(start_point) == 2:
                    start_point = (start_point[0], start_point[1], 0)
                if len(end_point) == 2:
                    end_point = (end_point[0], end_point[1], 0)
                
                # Nếu không có vị trí văn bản được cung cấp, Tính toán tự động
                if text_position is None:
                    # phía trên điểm giữa giữa điểm đầu và điểm cuối
                    mid_x = (start_point[0] + end_point[0]) / 2
                    mid_y = (start_point[1] + end_point[1]) / 2
                    text_position = (mid_x, mid_y + 5, 0)
                elif len(text_position) == 2:
                    text_position = (text_position[0], text_position[1], 0)
                
                # sử dụngVARIANTGói dữ liệu điểm tọa độ
                start_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                 [start_point[0], start_point[1], start_point[2]])
                end_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                               [end_point[0], end_point[1], end_point[2]])
                text_pos_array = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, 
                                                     [text_position[0], text_position[1], text_position[2]])
                
                # Thêm chú thích căn chỉnh
                dimension = self.doc.ModelSpace.AddDimAligned(start_array, end_array, text_pos_array)
                
                # Đặt chiều cao văn bản
                if textheight is not None:
                    dimension.TextHeight = textheight
                
                # Nếu một lớp được chỉ định, Đặt lớp
                if layer:
                    # Đảm bảo lớp tồn tại
                    self.create_layer(layer)
                    # Đặt lớp thực thể
                    dimension.Layer = layer

                # Nếu màu sắc được chỉ định, Đặt màu
                if color is not None:
                    dimension.Color = color
                
                # làm mới chế độ xem
                self.refresh_view()

                logger.info(f"Đã thêm nhãn: từ {start_point} đến {end_point}, Lớp{layer if layer else 'mặc định'}")
                return dimension
            except Exception as e:
                logger.error(f"Lỗi khi thêm kích thước: {str(e)}")
                return None

    # ==================== TRUY VẤN DỮ LIỆU NATIVE COM INTEROP (100% KHÔNG CẦN DLL) ====================

    def get_document_info(self) -> dict:
        """Lấy thông tin chi tiết của bản vẽ hiện tại qua Native COM Interop 100% không cần DLL."""
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            name = getattr(self.doc, "Name", "Unknown")
            fullname = getattr(self.doc, "FullName", "")
            path = getattr(self.doc, "Path", "")
            saved = bool(getattr(self.doc, "Saved", True))

            entity_count = 0
            try:
                entity_count = self.doc.ModelSpace.Count
            except Exception:
                pass

            active_layer = ""
            try:
                active_layer = self.doc.ActiveLayer.Name
            except Exception:
                pass

            layer_count = 0
            try:
                layer_count = self.doc.Layers.Count
            except Exception:
                pass

            return {
                "success": True,
                "data": {
                    "name": name,
                    "fullName": fullname,
                    "path": path,
                    "saved": saved,
                    "entityCount": entity_count,
                    "activeLayer": active_layer,
                    "layerCount": layer_count,
                },
                "executionMethod": "native_com",
                "message": f"Thông tin bản vẽ '{name}' (truy xuất trực tiếp qua Native COM Interop)"
            }
        except Exception as e:
            return {"success": False, "message": f"Lỗi truy xuất thông tin bản vẽ qua COM: {str(e)}"}

    def get_layers(self) -> dict:
        """Lấy danh sách tất cả các layer trong bản vẽ qua Native COM Interop 100% không cần DLL."""
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            layers = []
            for i in range(self.doc.Layers.Count):
                try:
                    layer = self.doc.Layers.Item(i)
                    layers.append({
                        "name": layer.Name,
                        "color": layer.color,
                        "frozen": getattr(layer, "Freeze", False),
                        "locked": getattr(layer, "Lock", False),
                        "on": not getattr(layer, "LayerOn", True) if hasattr(layer, "LayerOn") else True
                    })
                except Exception:
                    continue
            return {
                "success": True,
                "data": {
                    "layers": layers,
                    "count": len(layers),
                    "activeLayer": self.doc.ActiveLayer.Name
                },
                "executionMethod": "native_com",
                "message": f"Tìm thấy {len(layers)} layer (truy xuất trực tiếp qua Native COM Interop)"
            }
        except Exception as e:
            return {"success": False, "message": f"Lỗi truy xuất danh sách layer qua COM: {str(e)}"}

    def get_active_layer(self) -> dict:
        """Lấy tên và thông tin Layer đang kích hoạt qua Native COM Interop."""
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            al = self.doc.ActiveLayer
            return {
                "success": True,
                "data": {
                    "name": al.Name,
                    "color": al.color,
                    "frozen": getattr(al, "Freeze", False),
                    "locked": getattr(al, "Lock", False)
                },
                "executionMethod": "native_com",
                "message": f"Layer hiện hành: {al.Name}"
            }
        except Exception as e:
            return {"success": False, "message": f"Lỗi truy xuất active layer qua COM: {str(e)}"}

    def get_entities_summary(self) -> dict:
        """Thống kê các đối tượng hình học trong ModelSpace qua Native COM Interop."""
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            ms = self.doc.ModelSpace
            count = ms.Count
            type_counts = {}
            sample_size = min(count, 1000)
            for i in range(sample_size):
                try:
                    ent = ms.Item(i)
                    obj_name = getattr(ent, "ObjectName", "Unknown")
                    clean_type = obj_name.replace("AcDb", "").replace("Entity", "")
                    type_counts[clean_type] = type_counts.get(clean_type, 0) + 1
                except Exception:
                    continue
            return {
                "success": True,
                "data": {
                    "totalCount": count,
                    "breakdown": type_counts,
                    "sampled": sample_size
                },
                "executionMethod": "native_com",
                "message": f"Tổng cộng {count} thực thể trong ModelSpace"
            }
        except Exception as e:
            return {"success": False, "message": f"Lỗi thống kê thực thể qua COM: {str(e)}"}

    # ==================== ĐỘNG CƠ BÓC TÁCH KHỐI LƯỢNG (QS TAKEOFF ENGINE) ====================
    # 100% Native COM Interop - Không phụ thuộc DLL - Zero Dead Numbers
    # Phân loại Hạng mục AEC - Chống cộng trùng diện tích - Phát hiện lỗi số 0 âm thầm

    def get_units(self) -> dict:
        """Đọc đơn vị bản vẽ CAD (INSUNITS, MEASUREMENT, DIMLFAC, LUNITS) và tính hệ số quy đổi chuẩn sang mét (m) và mét vuông (m2).
        Tự động cảnh báo đỏ khi INSUNITS = 0 (tránh sai lệch 1000 lần giữa mm và m) hoặc DIMLFAC != 1.0.
        """
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            doc = self.doc
            try:
                insunits = int(doc.GetVariable("INSUNITS"))
            except Exception:
                insunits = 0

            try:
                measurement = int(doc.GetVariable("MEASUREMENT"))
            except Exception:
                measurement = 1

            try:
                dimlfac = float(doc.GetVariable("DIMLFAC"))
            except Exception:
                dimlfac = 1.0

            try:
                lunits = int(doc.GetVariable("LUNITS"))
            except Exception:
                lunits = 2

            # Bản đồ 25 mã đơn vị INSUNITS của AutoCAD sang mét (m):
            unit_map = {
                0: ("Unspecified (Chưa chỉ định đơn vị)", 0.001),
                1: ("Inches", 0.0254),
                2: ("Feet", 0.3048),
                3: ("Miles", 1609.344),
                4: ("Millimeters", 0.001),
                5: ("Centimeters", 0.01),
                6: ("Meters", 1.0),
                7: ("Kilometers", 1000.0),
                8: ("Microinches", 2.54e-8),
                9: ("Mils", 2.54e-5),
                10: ("Yards", 0.9144),
                11: ("Angstroms", 1e-10),
                12: ("Nanometers", 1e-9),
                13: ("Microns", 1e-6),
                14: ("Decimeters", 0.1),
                15: ("Decameters", 10.0),
                16: ("Hectometers", 100.0),
                17: ("Gigameters", 1e9),
                18: ("Astronomical Units", 1.495978707e11),
                19: ("Light Years", 9.460730472e15),
                20: ("Parsecs", 3.085677581e16),
                21: ("US Survey Feet", 0.3048006096),
                22: ("US Survey Inch", 0.0254000508),
                23: ("US Survey Yard", 0.9144018288),
                24: ("US Survey Mile", 1609.3472187),
            }

            unit_name, scale_to_m = unit_map.get(insunits, (f"Unit Code {insunits}", 0.001))
            scale_to_m2 = scale_to_m * scale_to_m

            warnings = []
            if insunits == 0:
                warnings.append(
                    "CẢNH BÁO ĐƠN VỊ (INSUNITS = 0): Bản vẽ chưa khai báo đơn vị đo lường. Hệ thống tạm thời giả định là Millimeters (mm). "
                    "NẾU BẢN VẼ ĐƯỢC THIẾT KẾ THEO MÉT (m) (thường gặp ở bản vẽ hạ tầng, san nền, quy hoạch), SỐ LIỆU SẼ BỊ SAI LỆCH 1000 LẦN! "
                    "Vui lòng kiểm tra kích thước thực tế hoặc gõ lệnh INSUNITS=6 (Meters) hoặc INSUNITS=4 (Millimeters) trong CAD."
                )

            if abs(dimlfac - 1.0) > 1e-4:
                warnings.append(
                    f"CẢNH BÁO DIMLFAC={dimlfac} khác 1.0: Kích thước hiển thị đã bị co giãn tỷ lệ so với hình học CAD thực tế."
                )

            warning_str = "\n".join(warnings) if warnings else None

            return {
                "success": True,
                "data": {
                    "insunits": insunits,
                    "unit_name": unit_name,
                    "scale_to_m": scale_to_m,
                    "scale_to_m2": scale_to_m2,
                    "measurement": "Metric" if measurement == 1 else "Imperial",
                    "dimlfac": dimlfac,
                    "lunits": lunits,
                    "warnings": warnings,
                    "warning": warning_str
                },
                "executionMethod": "native_com",
                "message": f"Đơn vị bản vẽ: {unit_name} (Hệ số quy đổi sang m: {scale_to_m}, m2: {scale_to_m2})"
            }
        except Exception as e:
            logger.error(f"Lỗi đọc đơn vị bản vẽ: {e}")
            return {"success": False, "message": f"Lỗi đọc đơn vị qua COM: {str(e)}"}

    def quantity_by_layer(self, layers: list = None, area_mode: str = "separate") -> dict:
        """Duyệt ModelSpace qua COM, tổng hợp Chiều dài (Length) và Diện tích (Area) theo từng Layer và phân loại hình học.
        Tách riêng diện tích Polyline kín và Hatch để triệt tiêu lỗi cộng trùng số liệu (Double-counting).
        Tự động phân loại Hạng mục AEC và bắt lỗi các đối tượng số 0 âm thầm (Skipped entities).
        """
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            # Lấy đơn vị và hệ số quy đổi
            units_info = self.get_units()
            scale_to_m = 0.001
            scale_to_m2 = 1e-6
            unit_name = "Millimeters"
            if units_info.get("success"):
                scale_to_m = units_info["data"]["scale_to_m"]
                scale_to_m2 = units_info["data"]["scale_to_m2"]
                unit_name = units_info["data"]["unit_name"]

            layer_filter = set(l.lower() for l in layers) if layers else None
            ms = self.doc.ModelSpace
            count = ms.Count

            layers_dict = {}
            skipped_entities = []

            for i in range(count):
                try:
                    ent = ms.Item(i)
                    layer = str(getattr(ent, "Layer", "0"))
                    if layer_filter and layer.lower() not in layer_filter:
                        continue

                    if layer not in layers_dict:
                        t_code, t_name = classify_layer_trade(layer)
                        layers_dict[layer] = {
                            "layer": layer,
                            "trade_code": t_code,
                            "trade_name": t_name,
                            "total_entities": 0,
                            "raw_length": 0.0,
                            "raw_closed_polyline_area": 0.0,
                            "raw_hatch_area": 0.0,
                            "raw_circle_area": 0.0,
                            "raw_ellipse_area": 0.0,
                            "breakdown": {
                                "lines": {"count": 0, "length": 0.0},
                                "polylines": {"count": 0, "length": 0.0, "area": 0.0, "closed_count": 0, "open_count": 0},
                                "circles": {"count": 0, "length": 0.0, "area": 0.0},
                                "arcs": {"count": 0, "length": 0.0},
                                "hatches": {"count": 0, "area": 0.0},
                                "ellipses": {"count": 0, "length": 0.0, "area": 0.0},
                                "others": {"count": 0}
                            }
                        }

                    entry = layers_dict[layer]
                    entry["total_entities"] += 1
                    obj_name = getattr(ent, "ObjectName", "")

                    if obj_name == "AcDbLine":
                        length = float(getattr(ent, "Length", 0.0))
                        entry["raw_length"] += length
                        b = entry["breakdown"]["lines"]
                        b["count"] += 1
                        b["length"] += length

                    elif obj_name in ("AcDbPolyline", "AcDb2dPolyline", "AcDb3dPolyline"):
                        length = 0.0
                        try:
                            length = float(getattr(ent, "Length", 0.0))
                        except Exception:
                            length = 0.0

                        # Xử lý trường hợp số 0 âm thầm: Tự động tính khoảng cách các đỉnh nếu Length = 0
                        if length <= 0.0:
                            try:
                                coords = getattr(ent, "Coordinates", None)
                                if coords and len(coords) >= 4:
                                    step = 2 if obj_name == "AcDbPolyline" else 3
                                    pts = [coords[j:j+step] for j in range(0, len(coords), step)]
                                    for j in range(len(pts) - 1):
                                        dx = pts[j+1][0] - pts[j][0]
                                        dy = pts[j+1][1] - pts[j][1]
                                        dz = (pts[j+1][2] - pts[j][2]) if (step == 3 and len(pts[j+1]) > 2) else 0.0
                                        length += math.sqrt(dx*dx + dy*dy + dz*dz)
                                    if getattr(ent, "Closed", False) and len(pts) > 2:
                                        dx = pts[0][0] - pts[-1][0]
                                        dy = pts[0][1] - pts[-1][1]
                                        dz = (pts[0][2] - pts[-1][2]) if (step == 3 and len(pts[0]) > 2) else 0.0
                                        length += math.sqrt(dx*dx + dy*dy + dz*dz)
                            except Exception as c_err:
                                skipped_entities.append({
                                    "handle": str(getattr(ent, "Handle", "")),
                                    "layer": layer,
                                    "object_name": obj_name,
                                    "reason": f"Không thể tính chiều dài polyline: {c_err}"
                                })

                        entry["raw_length"] += length
                        b = entry["breakdown"]["polylines"]
                        b["count"] += 1
                        b["length"] += length
                        is_closed = bool(getattr(ent, "Closed", False))
                        if is_closed:
                            b["closed_count"] += 1
                            try:
                                area = float(getattr(ent, "Area", 0.0))
                                entry["raw_closed_polyline_area"] += area
                                b["area"] += area
                            except Exception as a_err:
                                skipped_entities.append({
                                    "handle": str(getattr(ent, "Handle", "")),
                                    "layer": layer,
                                    "object_name": obj_name,
                                    "reason": f"Không thể đọc diện tích polyline kín: {a_err}"
                                })
                        else:
                            b["open_count"] += 1

                    elif obj_name == "AcDbCircle":
                        radius = float(getattr(ent, "Radius", 0.0))
                        circ = 2.0 * math.pi * radius
                        entry["raw_length"] += circ
                        area = math.pi * radius * radius
                        try:
                            area = float(getattr(ent, "Area", area))
                        except Exception:
                            pass
                        entry["raw_circle_area"] += area
                        b = entry["breakdown"]["circles"]
                        b["count"] += 1
                        b["length"] += circ
                        b["area"] += area

                    elif obj_name == "AcDbArc":
                        length = float(getattr(ent, "ArcLength", 0.0))
                        entry["raw_length"] += length
                        b = entry["breakdown"]["arcs"]
                        b["count"] += 1
                        b["length"] += length

                    elif obj_name == "AcDbHatch":
                        try:
                            area = float(getattr(ent, "Area", 0.0))
                            entry["raw_hatch_area"] += area
                            b = entry["breakdown"]["hatches"]
                            b["count"] += 1
                            b["area"] += area
                        except Exception as h_err:
                            skipped_entities.append({
                                "handle": str(getattr(ent, "Handle", "")),
                                "layer": layer,
                                "object_name": "AcDbHatch",
                                "reason": f"Không thể đọc diện tích Hatch: {h_err}"
                            })

                    elif obj_name == "AcDbEllipse":
                        area = 0.0
                        try:
                            area = float(getattr(ent, "Area", 0.0))
                            entry["raw_ellipse_area"] += area
                        except Exception:
                            pass
                        length = 0.0
                        try:
                            length = float(getattr(ent, "ArcLength", 0.0))
                            entry["raw_length"] += length
                        except Exception:
                            pass
                        b = entry["breakdown"]["ellipses"]
                        b["count"] += 1
                        b["length"] += length
                        b["area"] += area

                    else:
                        entry["breakdown"]["others"]["count"] += 1

                except Exception as ent_err:
                    skipped_entities.append({
                        "handle": "unknown",
                        "layer": "unknown",
                        "object_name": "unknown",
                        "reason": str(ent_err)
                    })
                    continue

            # Quy đổi đơn vị, phát hiện trùng lặp và phân nhóm Hạng mục (Trade)
            layers_result = []
            potential_double_counting_layers = []
            trade_summary = {}

            grand_total_length_m = 0.0
            grand_total_closed_polyline_area_m2 = 0.0
            grand_total_hatch_area_m2 = 0.0
            grand_total_circle_area_m2 = 0.0
            grand_total_area_m2 = 0.0
            grand_total_entities = 0

            for layer_name in sorted(layers_dict.keys()):
                d = layers_dict[layer_name]
                length_m = round(d["raw_length"] * scale_to_m, 3)
                closed_polyline_area_m2 = round(d["raw_closed_polyline_area"] * scale_to_m2, 3)
                hatch_area_m2 = round(d["raw_hatch_area"] * scale_to_m2, 3)
                circle_area_m2 = round(d["raw_circle_area"] * scale_to_m2, 3)

                # Cảnh báo nếu layer có cả Pline kín và Hatch (nguy cơ cộng trùng)
                area_warning = None
                if closed_polyline_area_m2 > 0 and hatch_area_m2 > 0:
                    area_warning = "Cảnh báo trùng lặp: Layer có cả Polyline kín và Hatch. Đã tách riêng 2 cột để không cộng dồn trùng lặp."
                    potential_double_counting_layers.append(layer_name)

                # Tính raw_area và area_m2 tương thích ngược
                if d["raw_closed_polyline_area"] > 0 or d["raw_circle_area"] > 0:
                    raw_area = d["raw_closed_polyline_area"] + d["raw_circle_area"]
                else:
                    raw_area = d["raw_hatch_area"] + d["raw_ellipse_area"]
                area_m2 = round(raw_area * scale_to_m2, 3)

                grand_total_length_m += length_m
                grand_total_closed_polyline_area_m2 += closed_polyline_area_m2
                grand_total_hatch_area_m2 += hatch_area_m2
                grand_total_circle_area_m2 += circle_area_m2
                grand_total_area_m2 += area_m2
                grand_total_entities += d["total_entities"]

                t_code = d["trade_code"]
                t_name = d["trade_name"]
                if t_code not in trade_summary:
                    trade_summary[t_code] = {
                        "trade_name": t_name,
                        "length_m": 0.0,
                        "closed_polyline_area_m2": 0.0,
                        "hatch_area_m2": 0.0,
                        "circle_area_m2": 0.0,
                        "total_entities": 0,
                        "layer_count": 0
                    }
                ts = trade_summary[t_code]
                ts["length_m"] = round(ts["length_m"] + length_m, 3)
                ts["closed_polyline_area_m2"] = round(ts["closed_polyline_area_m2"] + closed_polyline_area_m2, 3)
                ts["hatch_area_m2"] = round(ts["hatch_area_m2"] + hatch_area_m2, 3)
                ts["circle_area_m2"] = round(ts["circle_area_m2"] + circle_area_m2, 3)
                ts["total_entities"] += d["total_entities"]
                ts["layer_count"] += 1

                layers_result.append({
                    "layer": layer_name,
                    "trade_code": t_code,
                    "trade_name": t_name,
                    "total_entities": d["total_entities"],
                    "raw_length": round(d["raw_length"], 3),
                    "length_m": length_m,
                    "raw_closed_polyline_area": round(d["raw_closed_polyline_area"], 3),
                    "closed_polyline_area_m2": closed_polyline_area_m2,
                    "raw_hatch_area": round(d["raw_hatch_area"], 3),
                    "hatch_area_m2": hatch_area_m2,
                    "raw_circle_area": round(d["raw_circle_area"], 3),
                    "circle_area_m2": circle_area_m2,
                    "raw_area": round(raw_area, 3),
                    "area_m2": area_m2,
                    "area_warning": area_warning,
                    "breakdown": d["breakdown"]
                })

            result_data = {
                "units": {
                    "unit_name": unit_name,
                    "scale_to_m": scale_to_m,
                    "scale_to_m2": scale_to_m2
                },
                "total_entities_scanned": count,
                "layers_count": len(layers_result),
                "grand_total_length_m": round(grand_total_length_m, 3),
                "grand_total_closed_polyline_area_m2": round(grand_total_closed_polyline_area_m2, 3),
                "grand_total_hatch_area_m2": round(grand_total_hatch_area_m2, 3),
                "grand_total_circle_area_m2": round(grand_total_circle_area_m2, 3),
                "grand_total_area_m2": round(grand_total_area_m2, 3),
                "grand_total_entities": grand_total_entities,
                "by_trade": trade_summary,
                "layers": layers_result,
                "diagnostics": {
                    "skipped_entities_count": len(skipped_entities),
                    "skipped_entities": skipped_entities[:20],
                    "potential_double_counting_layers": potential_double_counting_layers
                }
            }

            # Lưu vào bộ nhớ đệm
            self.last_quantities_data = result_data

            msg = (
                f"Bóc tách thành công {len(layers_result)} layer: "
                f"Dài = {grand_total_length_m:.2f} m | "
                f"DT Pline kín = {grand_total_closed_polyline_area_m2:.2f} m2 | "
                f"DT Hatch = {grand_total_hatch_area_m2:.2f} m2 (Đã tách riêng chống trùng lặp)"
            )
            if potential_double_counting_layers:
                msg += f" [Lưu ý: Có {len(potential_double_counting_layers)} layer chứa cả Pline kín và Hatch]"

            return {
                "success": True,
                "data": result_data,
                "executionMethod": "native_com",
                "message": msg
            }
        except Exception as e:
            logger.error(f"Lỗi bóc tách khối lượng layer: {e}")
            return {"success": False, "message": f"Lỗi bóc tách khối lượng qua COM: {str(e)}"}

    def count_blocks(self, block_names: list = None, include_attributes: bool = True, scan_nested: bool = False) -> dict:
        """Thống kê chi tiết Block và thiết bị trong ModelSpace, tự động xử lý Dynamic Block, trích xuất thuộc tính Attribute, phát hiện Xref và Block lồng nhau."""
        if not self._ensure_valid_connection():
            return {"success": False, "message": "CAD chưa chạy hoặc kết nối đã bị ngắt"}
        try:
            ms = self.doc.ModelSpace
            count = ms.Count

            name_filter = set(b.lower() for b in block_names) if block_names else None
            blocks_dict = {}
            total_blocks_found = 0

            for i in range(count):
                try:
                    ent = ms.Item(i)
                    obj_name = getattr(ent, "ObjectName", "")
                    if obj_name not in ("AcDbBlockReference", "AcDbMInsertBlock"):
                        continue

                    # Phân giải tên Dynamic Block (EffectiveName ưu tiên trước)
                    eff_name = getattr(ent, "EffectiveName", None)
                    raw_name = getattr(ent, "Name", "UnknownBlock")
                    block_name = eff_name if eff_name else raw_name

                    if name_filter and block_name.lower() not in name_filter:
                        continue

                    layer = str(getattr(ent, "Layer", "0"))
                    handle = str(getattr(ent, "Handle", ""))

                    # Kiểm tra XRef và Tỷ lệ Scale
                    is_xref = False
                    try:
                        blk_def = self.doc.Blocks.Item(raw_name)
                        is_xref = bool(getattr(blk_def, "IsXRef", False))
                    except Exception:
                        pass

                    x_scale = round(float(getattr(ent, "XScaleFactor", 1.0)), 2)
                    y_scale = round(float(getattr(ent, "YScaleFactor", 1.0)), 2)
                    z_scale = round(float(getattr(ent, "ZScaleFactor", 1.0)), 2)
                    is_scaled = (abs(x_scale - 1.0) > 1e-3 or abs(y_scale - 1.0) > 1e-3 or abs(z_scale - 1.0) > 1e-3)

                    # Điểm chèn tọa độ
                    ins_pt = [0.0, 0.0, 0.0]
                    try:
                        pt_var = getattr(ent, "InsertionPoint", None)
                        if pt_var is not None and len(pt_var) >= 2:
                            ins_pt = [round(float(pt_var[0]), 2), round(float(pt_var[1]), 2), round(float(pt_var[2]) if len(pt_var) > 2 else 0.0, 2)]
                    except Exception:
                        pass

                    # Trích xuất thuộc tính (Attributes)
                    attributes = {}
                    if include_attributes and getattr(ent, "HasAttributes", False):
                        try:
                            attribs = ent.GetAttributes()
                            for att in attribs:
                                tag = str(getattr(att, "TagString", "")).strip()
                                val = str(getattr(att, "TextString", "")).strip()
                                if tag:
                                    attributes[tag] = val
                        except Exception:
                            pass

                    total_blocks_found += 1
                    if block_name not in blocks_dict:
                        t_code, t_name = classify_layer_trade(layer)
                        blocks_dict[block_name] = {
                            "block_name": block_name,
                            "trade_code": t_code,
                            "trade_name": t_name,
                            "is_xref": is_xref,
                            "count": 0,
                            "layers": set(),
                            "attributes_summary": {},
                            "is_scaled": is_scaled,
                            "scales_seen": set(),
                            "instances_handles": [],
                            "sample_instances": []
                        }

                    b_entry = blocks_dict[block_name]
                    b_entry["count"] += 1
                    b_entry["layers"].add(layer)
                    b_entry["scales_seen"].add(f"({x_scale}, {y_scale}, {z_scale})")
                    b_entry["instances_handles"].append(handle)

                    if len(b_entry["sample_instances"]) < 20:
                        b_entry["sample_instances"].append({
                            "handle": handle,
                            "layer": layer,
                            "insertion_point": ins_pt,
                            "attributes": attributes,
                            "scale": [x_scale, y_scale, z_scale]
                        })

                    # Tổng hợp thống kê attribute
                    for k, v in attributes.items():
                        if k not in b_entry["attributes_summary"]:
                            b_entry["attributes_summary"][k] = {}
                        val_str = str(v)
                        b_entry["attributes_summary"][k][val_str] = b_entry["attributes_summary"][k].get(val_str, 0) + 1

                except Exception:
                    continue

            blocks_list = []
            for bname in sorted(blocks_dict.keys()):
                item = blocks_dict[bname]
                blocks_list.append({
                    "block_name": item["block_name"],
                    "trade_code": item["trade_code"],
                    "trade_name": item["trade_name"],
                    "is_xref": item["is_xref"],
                    "count": item["count"],
                    "layers": sorted(list(item["layers"])),
                    "attributes_summary": item["attributes_summary"],
                    "scales_seen": sorted(list(item["scales_seen"])),
                    "total_instances_recorded": len(item["instances_handles"]),
                    "sample_instances": item["sample_instances"][:5]
                })

            result_data = {
                "total_blocks_found": total_blocks_found,
                "unique_blocks_count": len(blocks_list),
                "blocks": blocks_list
            }

            # Lưu vào bộ nhớ đệm
            self.last_blocks_data = result_data

            return {
                "success": True,
                "data": result_data,
                "executionMethod": "native_com",
                "message": f"Thống kê được {total_blocks_found} thực thể block thuộc {len(blocks_list)} loại block khác nhau"
            }
        except Exception as e:
            logger.error(f"Lỗi thống kê block: {e}")
            return {"success": False, "message": f"Lỗi thống kê block qua COM: {str(e)}"}

    def export_boq(self, file_path: str = None, format: str = "xlsx", quantities_data: dict = None, blocks_data: dict = None, title: str = "BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ)", auto_cleanup: bool = False) -> dict:
        """Xuất bảng tiên lượng khối lượng BoQ ra file Excel (.xlsx) với 100% công thức sống (Zero Dead Numbers) hoặc CSV.
        Tự động tách riêng diện tích Polyline kín và Hatch, phân loại Hạng mục AEC và tái sử dụng bộ đệm (Cache).
        Hỗ trợ auto_cleanup=True để tự động ngắt kết nối COM giải phóng RAM và triệt tiêu các tiến trình con ngầm.
        """
        try:
            # 1. Thu thập dữ liệu: ưu tiên tái sử dụng bộ đệm đã truy vấn gần nhất
            if quantities_data is None:
                if self.last_quantities_data is not None:
                    quantities_data = self.last_quantities_data
                else:
                    q_res = self.quantity_by_layer()
                    if q_res.get("success"):
                        quantities_data = q_res.get("data")
                    else:
                        quantities_data = {"layers": [], "units": {"unit_name": "Millimeters", "scale_to_m": 0.001, "scale_to_m2": 1e-6}}

            if blocks_data is None:
                if self.last_blocks_data is not None:
                    blocks_data = self.last_blocks_data
                else:
                    b_res = self.count_blocks()
                    if b_res.get("success"):
                        blocks_data = b_res.get("data")
                    else:
                        blocks_data = {"blocks": []}

            # 2. Xác định đường dẫn file đầu ra
            if not file_path:
                now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                ext = "xlsx" if (format.lower() == "xlsx" and HAS_OPENPYXL) else "csv"
                file_path = os.path.abspath(f"BANG_TIEN_LUONG_BOQ_{now_str}.{ext}")

            out_dir = os.path.dirname(os.path.abspath(file_path))
            if out_dir:
                os.makedirs(out_dir, exist_ok=True)

            units = quantities_data.get("units", {})
            scale_to_m = units.get("scale_to_m", 0.001)
            scale_to_m2 = units.get("scale_to_m2", 1e-6)
            unit_name = units.get("unit_name", "Millimeters")

            layers_list = quantities_data.get("layers", [])
            blocks_list = blocks_data.get("blocks", [])

            # Tên bản vẽ hiện hành
            doc_name = "Drawing"
            try:
                if self.doc:
                    doc_name = self.doc.Name
            except Exception:
                pass

            # 3. Xuất file theo định dạng
            if format.lower() == "xlsx" and HAS_OPENPYXL:
                wb = openpyxl.Workbook()

                # Định nghĩa màu sắc và kiểu dáng chuyên nghiệp chuẩn AEC
                font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
                font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
                font_bold = Font(name="Segoe UI", size=10, bold=True, color="1B365D")
                font_normal = Font(name="Segoe UI", size=10)
                font_small = Font(name="Segoe UI", size=9, italic=True, color="555555")

                fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
                fill_slate = PatternFill(start_color="2B4C7E", end_color="2B4C7E", fill_type="solid")
                fill_total = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
                fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

                thin_border_side = Side(border_style="thin", color="CCCCCC")
                double_border_side = Side(border_style="double", color="1B365D")
                border_data = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
                border_total = Border(top=thin_border_side, bottom=double_border_side, left=thin_border_side, right=thin_border_side)

                align_center = Alignment(horizontal="center", vertical="center")
                align_left = Alignment(horizontal="left", vertical="center")
                align_right = Alignment(horizontal="right", vertical="center")

                # ================= SHEET 2: CHI TIẾT LAYER (tạo trước để Sheet 1 link công thức) =================
                ws2 = wb.create_sheet(title="2_Chi_Tiet_Layer")
                ws2.views.sheetView[0].showGridLines = True

                # Banner Sheet 2
                ws2.merge_cells("A1:L1")
                ws2["A1"] = "CHI TIẾT ĐO BÓC HÌNH HỌC THEO LAYER (TÁCH RIÊNG PLINE KÍN VÀ HATCH)"
                ws2["A1"].font = font_title
                ws2["A1"].fill = fill_navy
                ws2["A1"].alignment = align_center
                ws2.row_dimensions[1].height = 28

                # Header Sheet 2 (Giữ cột E cho Dài m, G cho DT Pline kín m2 để tương thích 100% test)
                headers_s2 = [
                    "STT", "TÊN LAYER", "SỐ ĐỐI TƯỢNG",
                    "CHIỀU DÀI THÔ (CAD)", "CHIỀU DÀI (m)",
                    "DT PLINE KÍN THÔ (CAD)", "DT PLINE KÍN (m2)",
                    "DT HATCH THÔ (CAD)", "DT HATCH (m2)",
                    "HẠNG MỤC (TRADE)", "PHÂN LOẠI THỰC THỂ", "CẢNH BÁO TRÙNG LẶP"
                ]
                for col_idx, h in enumerate(headers_s2, start=1):
                    cell = ws2.cell(row=3, column=col_idx, value=h)
                    cell.font = font_header
                    cell.fill = fill_slate
                    cell.alignment = align_center
                    cell.border = border_data
                ws2.row_dimensions[3].height = 24

                row_idx = 4
                start_row_s2 = 4
                for idx, lyr in enumerate(layers_list, start=1):
                    bk = lyr.get("breakdown", {})
                    bk_parts = []
                    if bk.get("lines", {}).get("count", 0) > 0:
                        bk_parts.append(f"Line: {bk['lines']['count']}")
                    if bk.get("polylines", {}).get("count", 0) > 0:
                        p_cnt = bk["polylines"]["count"]
                        c_cnt = bk["polylines"].get("closed_count", 0)
                        bk_parts.append(f"Pline: {p_cnt} ({c_cnt} kín)")
                    if bk.get("circles", {}).get("count", 0) > 0:
                        bk_parts.append(f"Circle: {bk['circles']['count']}")
                    if bk.get("arcs", {}).get("count", 0) > 0:
                        bk_parts.append(f"Arc: {bk['arcs']['count']}")
                    if bk.get("hatches", {}).get("count", 0) > 0:
                        bk_parts.append(f"Hatch: {bk['hatches']['count']}")
                    if bk.get("ellipses", {}).get("count", 0) > 0:
                        bk_parts.append(f"Ellipse: {bk['ellipses']['count']}")
                    bk_desc = ", ".join(bk_parts) if bk_parts else "Thực thể khác"

                    ws2.cell(row=row_idx, column=1, value=idx).alignment = align_center
                    ws2.cell(row=row_idx, column=2, value=lyr.get("layer", "")).alignment = align_left
                    c_cnt = ws2.cell(row=row_idx, column=3, value=lyr.get("total_entities", 0))
                    c_cnt.alignment = align_right
                    c_cnt.number_format = "#,##0"

                    c_rlen = ws2.cell(row=row_idx, column=4, value=lyr.get("raw_length", 0.0))
                    c_rlen.alignment = align_right
                    c_rlen.number_format = "#,##0.00"

                    # CÔNG THỨC SỐNG: Chiều dài (m) = Cột D * scale_to_m
                    c_len_m = ws2.cell(row=row_idx, column=5, value=f"=D{row_idx}*{scale_to_m}")
                    c_len_m.alignment = align_right
                    c_len_m.number_format = "#,##0.00"

                    # Diện tích Polyline kín
                    raw_closed_area = lyr.get("raw_closed_polyline_area", lyr.get("raw_area", 0.0))
                    c_rarea_p = ws2.cell(row=row_idx, column=6, value=raw_closed_area)
                    c_rarea_p.alignment = align_right
                    c_rarea_p.number_format = "#,##0.00"

                    # CÔNG THỨC SỐNG: Diện tích Pline kín (m2) = Cột F * scale_to_m2
                    c_area_p_m2 = ws2.cell(row=row_idx, column=7, value=f"=F{row_idx}*{scale_to_m2}")
                    c_area_p_m2.alignment = align_right
                    c_area_p_m2.number_format = "#,##0.00"

                    # Diện tích Hatch (TÁCH RIÊNG CHỐNG TRÙNG LẶP)
                    c_rarea_h = ws2.cell(row=row_idx, column=8, value=lyr.get("raw_hatch_area", 0.0))
                    c_rarea_h.alignment = align_right
                    c_rarea_h.number_format = "#,##0.00"

                    # CÔNG THỨC SỐNG: Diện tích Hatch (m2) = Cột H * scale_to_m2
                    c_area_h_m2 = ws2.cell(row=row_idx, column=9, value=f"=H{row_idx}*{scale_to_m2}")
                    c_area_h_m2.alignment = align_right
                    c_area_h_m2.number_format = "#,##0.00"

                    ws2.cell(row=row_idx, column=10, value=lyr.get("trade_name", "")).alignment = align_left
                    ws2.cell(row=row_idx, column=11, value=bk_desc).alignment = align_left
                    ws2.cell(row=row_idx, column=12, value=lyr.get("area_warning") or "").alignment = align_left

                    # Kẻ viền và màu zebra
                    for c_i in range(1, 13):
                        cell = ws2.cell(row=row_idx, column=c_i)
                        cell.font = font_normal
                        cell.border = border_data
                        if row_idx % 2 == 1:
                            cell.fill = fill_zebra
                    row_idx += 1

                end_row_s2 = max(row_idx - 1, start_row_s2)
                total_row_s2 = row_idx

                # Dòng Tổng Cộng Sheet 2
                ws2.cell(row=total_row_s2, column=1, value="TỔNG CỘNG").alignment = align_center
                ws2.cell(row=total_row_s2, column=2, value="").alignment = align_center
                ws2.merge_cells(f"A{total_row_s2}:B{total_row_s2}")

                # SUM formulas
                cell_sum_cnt = ws2.cell(row=total_row_s2, column=3, value=f"=SUM(C{start_row_s2}:C{end_row_s2})")
                cell_sum_cnt.number_format = "#,##0"
                cell_sum_cnt.alignment = align_right

                cell_sum_rlen = ws2.cell(row=total_row_s2, column=4, value=f"=SUM(D{start_row_s2}:D{end_row_s2})")
                cell_sum_rlen.number_format = "#,##0.00"
                cell_sum_rlen.alignment = align_right

                cell_sum_len = ws2.cell(row=total_row_s2, column=5, value=f"=SUM(E{start_row_s2}:E{end_row_s2})")
                cell_sum_len.number_format = "#,##0.00"
                cell_sum_len.alignment = align_right

                cell_sum_rparea = ws2.cell(row=total_row_s2, column=6, value=f"=SUM(F{start_row_s2}:F{end_row_s2})")
                cell_sum_rparea.number_format = "#,##0.00"
                cell_sum_rparea.alignment = align_right

                cell_sum_parea = ws2.cell(row=total_row_s2, column=7, value=f"=SUM(G{start_row_s2}:G{end_row_s2})")
                cell_sum_parea.number_format = "#,##0.00"
                cell_sum_parea.alignment = align_right

                cell_sum_rharea = ws2.cell(row=total_row_s2, column=8, value=f"=SUM(H{start_row_s2}:H{end_row_s2})")
                cell_sum_rharea.number_format = "#,##0.00"
                cell_sum_rharea.alignment = align_right

                cell_sum_harea = ws2.cell(row=total_row_s2, column=9, value=f"=SUM(I{start_row_s2}:I{end_row_s2})")
                cell_sum_harea.number_format = "#,##0.00"
                cell_sum_harea.alignment = align_right

                ws2.cell(row=total_row_s2, column=10, value="Zero Dead Numbers").alignment = align_left
                ws2.cell(row=total_row_s2, column=11, value="100% Live Formulas").alignment = align_left
                ws2.cell(row=total_row_s2, column=12, value="Đã tách riêng Pline kín vs Hatch").alignment = align_left

                for c_i in range(1, 13):
                    cell = ws2.cell(row=total_row_s2, column=c_i)
                    cell.font = font_bold
                    cell.fill = fill_total
                    cell.border = border_total
                ws2.row_dimensions[total_row_s2].height = 22

                # ================= SHEET 3: THỐNG KÊ BLOCK =================
                ws3 = wb.create_sheet(title="3_Thong_Ke_Block")
                ws3.views.sheetView[0].showGridLines = True

                ws3.merge_cells("A1:G1")
                ws3["A1"] = "BẢNG THỐNG KÊ BLOCK & THIẾT BỊ LẮP ĐẶT (DYNAMIC & ATTRIBUTES)"
                ws3["A1"].font = font_title
                ws3["A1"].fill = fill_navy
                ws3["A1"].alignment = align_center
                ws3.row_dimensions[1].height = 28

                headers_s3 = ["STT", "TÊN BLOCK / THIẾT BỊ", "PHÂN BỐ TRÊN LAYER", "SỐ LƯỢNG (Cái/Bộ)", "THUỘC TÍNH NỔI BẬT (ATTRIBUTES)", "HẠNG MỤC (TRADE)", "GHI CHÚ"]
                for col_idx, h in enumerate(headers_s3, start=1):
                    cell = ws3.cell(row=3, column=col_idx, value=h)
                    cell.font = font_header
                    cell.fill = fill_slate
                    cell.alignment = align_center
                    cell.border = border_data
                ws3.row_dimensions[3].height = 24

                b_row_idx = 4
                start_row_s3 = 4
                for idx, blk in enumerate(blocks_list, start=1):
                    ws3.cell(row=b_row_idx, column=1, value=idx).alignment = align_center
                    ws3.cell(row=b_row_idx, column=2, value=blk.get("block_name", "")).alignment = align_left
                    ws3.cell(row=b_row_idx, column=3, value=", ".join(blk.get("layers", []))).alignment = align_left

                    c_bcnt = ws3.cell(row=b_row_idx, column=4, value=blk.get("count", 0))
                    c_bcnt.alignment = align_right
                    c_bcnt.number_format = "#,##0"

                    # Chuỗi tóm tắt attributes
                    att_parts = []
                    for k, vmap in blk.get("attributes_summary", {}).items():
                        v_summary = "/".join([f"{val}({c})" for val, c in list(vmap.items())[:3]])
                        att_parts.append(f"{k}: {v_summary}")
                    att_str = "; ".join(att_parts) if att_parts else "-"

                    ws3.cell(row=b_row_idx, column=5, value=att_str).alignment = align_left
                    ws3.cell(row=b_row_idx, column=6, value=blk.get("trade_name", "")).alignment = align_left
                    ws3.cell(row=b_row_idx, column=7, value="XRef" if blk.get("is_xref") else "").alignment = align_left

                    for c_i in range(1, 8):
                        cell = ws3.cell(row=b_row_idx, column=c_i)
                        cell.font = font_normal
                        cell.border = border_data
                        if b_row_idx % 2 == 1:
                            cell.fill = fill_zebra
                    b_row_idx += 1

                end_row_s3 = max(b_row_idx - 1, start_row_s3)
                total_row_s3 = b_row_idx

                # Dòng Tổng Cộng Sheet 3
                ws3.cell(row=total_row_s3, column=1, value="TỔNG CỘNG").alignment = align_center
                ws3.merge_cells(f"A{total_row_s3}:C{total_row_s3}")

                c_sum_b = ws3.cell(row=total_row_s3, column=4, value=f"=SUM(D{start_row_s3}:D{end_row_s3})")
                c_sum_b.alignment = align_right
                c_sum_b.number_format = "#,##0"

                ws3.cell(row=total_row_s3, column=5, value="Cái / Bộ").alignment = align_center
                ws3.cell(row=total_row_s3, column=6, value="100% Live Formula").alignment = align_left
                ws3.cell(row=total_row_s3, column=7, value="").alignment = align_left

                for c_i in range(1, 8):
                    cell = ws3.cell(row=total_row_s3, column=c_i)
                    cell.font = font_bold
                    cell.fill = fill_total
                    cell.border = border_total
                ws3.row_dimensions[total_row_s3].height = 22

                # ================= SHEET 1: TỔNG HỢP TIÊN LƯỢNG (BẢNG CHÍNH BOQ) =================
                ws1 = wb.active
                ws1.title = "1_Tong_Hop_BoQ"
                ws1.views.sheetView[0].showGridLines = True

                # Banner
                ws1.merge_cells("A1:E1")
                ws1["A1"] = title.upper()
                ws1["A1"].font = font_title
                ws1["A1"].fill = fill_navy
                ws1["A1"].alignment = align_center
                ws1.row_dimensions[1].height = 30

                # Thông tin bản vẽ & đơn vị
                now_str_vn = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                meta_rows = [
                    ("Tên bản vẽ nguồn:", doc_name, "Thời điểm xuất:", now_str_vn),
                    ("Đơn vị gốc CAD:", unit_name, "Hệ số quy đổi sang m:", f"{scale_to_m} (m2: {scale_to_m2})")
                ]
                for m_idx, (k1, v1, k2, v2) in enumerate(meta_rows, start=3):
                    ws1.cell(row=m_idx, column=1, value=k1).font = font_bold
                    ws1.cell(row=m_idx, column=2, value=v1).font = font_normal
                    ws1.cell(row=m_idx, column=3, value=k2).font = font_bold
                    ws1.cell(row=m_idx, column=4, value=v2).font = font_normal

                # Header bảng tổng hợp
                headers_s1 = ["STT", "DANH MỤC TIÊN LƯỢNG / HẠNG MỤC CÔNG TRÌNH", "ĐƠN VỊ", "KHỐI LƯỢNG TIÊN LƯỢNG", "NGUỒN CÔNG THỨC & GHI CHÚ KỸ THUẬT"]
                for col_idx, h in enumerate(headers_s1, start=1):
                    cell = ws1.cell(row=6, column=col_idx, value=h)
                    cell.font = font_header
                    cell.fill = fill_slate
                    cell.alignment = align_center
                    cell.border = border_data
                ws1.row_dimensions[6].height = 24

                # Dòng 1: Chiều dài
                ws1.cell(row=7, column=1, value=1).alignment = align_center
                ws1.cell(row=7, column=2, value="Tổng chiều dài tim tuyến / nét kết cấu hình học (Lines & Polylines)").alignment = align_left
                ws1.cell(row=7, column=3, value="m").alignment = align_center
                c_vol1 = ws1.cell(row=7, column=4, value=f"='2_Chi_Tiet_Layer'!E{total_row_s2}")
                c_vol1.alignment = align_right
                c_vol1.number_format = "#,##0.00"
                c_vol1.font = font_bold
                ws1.cell(row=7, column=5, value="Công thức động liên kết từ sheet 2_Chi_Tiet_Layer").alignment = align_left

                # Dòng 2: Diện tích Polyline kín
                ws1.cell(row=8, column=1, value=2).alignment = align_center
                ws1.cell(row=8, column=2, value="Tổng diện tích hình học khép kín (Closed Polylines, Circles)").alignment = align_left
                ws1.cell(row=8, column=3, value="m2").alignment = align_center
                c_vol2 = ws1.cell(row=8, column=4, value=f"='2_Chi_Tiet_Layer'!G{total_row_s2}")
                c_vol2.alignment = align_right
                c_vol2.number_format = "#,##0.00"
                c_vol2.font = font_bold
                ws1.cell(row=8, column=5, value="Công thức động liên kết từ sheet 2_Chi_Tiet_Layer (DT Pline kín)").alignment = align_left

                # Dòng 3: Khối lượng Block
                ws1.cell(row=9, column=1, value=3).alignment = align_center
                ws1.cell(row=9, column=2, value="Tổng số lượng Block thiết bị và cấu kiện lắp đặt").alignment = align_left
                ws1.cell(row=9, column=3, value="Cái/Bộ").alignment = align_center
                c_vol3 = ws1.cell(row=9, column=4, value=f"='3_Thong_Ke_Block'!D{total_row_s3}")
                c_vol3.alignment = align_right
                c_vol3.number_format = "#,##0"
                c_vol3.font = font_bold
                ws1.cell(row=9, column=5, value="Công thức động liên kết từ sheet 3_Thong_Ke_Block").alignment = align_left

                # Dòng 4: Diện tích Hatch (TÁCH RIÊNG RÕ RÀNG)
                ws1.cell(row=10, column=1, value=4).alignment = align_center
                ws1.cell(row=10, column=2, value="Tổng diện tích Hatch hoàn thiện / vật liệu lát (Hatch Areas)").alignment = align_left
                ws1.cell(row=10, column=3, value="m2").alignment = align_center
                c_vol4 = ws1.cell(row=10, column=4, value=f"='2_Chi_Tiet_Layer'!I{total_row_s2}")
                c_vol4.alignment = align_right
                c_vol4.number_format = "#,##0.00"
                c_vol4.font = font_bold
                ws1.cell(row=10, column=5, value="Công thức động liên kết từ sheet 2_Chi_Tiet_Layer (Đã tách riêng chống trùng)").alignment = align_left

                for r_idx in range(7, 11):
                    for c_i in range(1, 6):
                        c = ws1.cell(row=r_idx, column=c_i)
                        c.border = border_data
                        if r_idx % 2 == 1:
                            c.fill = fill_zebra
                    ws1.row_dimensions[r_idx].height = 20

                # Chân trang chú thích
                ws1.cell(row=12, column=1, value="* Ghi chú: Toàn bộ bảng tính tuân thủ nghiêm ngặt nguyên tắc Zero Dead Numbers. Mọi phép cộng tổng và quy đổi tỷ lệ đều sử dụng công thức sống của Excel.").font = font_small

                # Tự động căn chỉnh độ rộng cột cho tất cả các sheets
                for sheet in (ws1, ws2, ws3):
                    for col in sheet.columns:
                        max_len = 0
                        col_letter = get_column_letter(col[0].column)
                        for cell in col:
                            if cell.row == 1:
                                continue
                            val_str = str(cell.value or "")
                            if len(val_str) > max_len:
                                max_len = len(val_str)
                        sheet.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

                wb.save(file_path)
                logger.info(f"Đã xuất bảng khối lượng BoQ Excel thành công: {file_path}")

            else:
                # Xuất ra định dạng CSV (Hỗ trợ khi format == 'csv' hoặc không có openpyxl)
                csv_file = file_path if file_path.endswith(".csv") else f"{os.path.splitext(file_path)[0]}.csv"
                with open(csv_file, mode="w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow([title])
                    writer.writerow(["Bản vẽ:", doc_name, "Thời điểm:", datetime.now().strftime("%d/%m/%Y %H:%M:%S")])
                    writer.writerow(["Đơn vị CAD:", unit_name, "Hệ số sang m:", scale_to_m, "Hệ số sang m2:", scale_to_m2])
                    writer.writerow([])
                    writer.writerow(["=== CHI TIẾT LAYER ==="])
                    writer.writerow(["STT", "Tên Layer", "Hạng mục", "Số thực thể", "Dài thô (CAD)", f"Dài (m, x{scale_to_m})", "DT Pline kín (CAD)", "DT Pline kín (m2)", "DT Hatch (CAD)", "DT Hatch (m2)", "Cảnh báo"])
                    for idx, lyr in enumerate(layers_list, start=1):
                        writer.writerow([
                            idx,
                            lyr.get("layer"),
                            lyr.get("trade_name"),
                            lyr.get("total_entities"),
                            lyr.get("raw_length"),
                            lyr.get("length_m"),
                            lyr.get("raw_closed_polyline_area"),
                            lyr.get("closed_polyline_area_m2"),
                            lyr.get("raw_hatch_area"),
                            lyr.get("hatch_area_m2"),
                            lyr.get("area_warning") or ""
                        ])
                    writer.writerow([])
                    writer.writerow(["=== THỐNG KÊ BLOCK ==="])
                    writer.writerow(["STT", "Tên Block", "Hạng mục", "Layer", "Số lượng", "Attributes"])
                    for idx, blk in enumerate(blocks_list, start=1):
                        att_summary = json.dumps(blk.get("attributes_summary", {}), ensure_ascii=False)
                        writer.writerow([
                            idx,
                            blk.get("block_name"),
                            blk.get("trade_name"),
                            ", ".join(blk.get("layers", [])),
                            blk.get("count"),
                            att_summary
                        ])
                file_path = csv_file
                logger.info(f"Đã xuất bảng khối lượng BoQ CSV thành công: {file_path}")

            if auto_cleanup:
                self.disconnect()

            return {
                "success": True,
                "file_path": file_path,
                "format": format,
                "layers_count": len(layers_list),
                "blocks_count": len(blocks_list),
                "cleaned_up": bool(auto_cleanup),
                "message": f"Xuất bảng tiên lượng BoQ thành công tới: {file_path} (Zero Dead Numbers, 100% công thức sống, tách riêng Pline kín vs Hatch)"
            }

        except Exception as e:
            logger.error(f"Lỗi xuất bảng tiên lượng BoQ: {e}")
            return {"success": False, "message": f"Lỗi xuất bảng tiên lượng BoQ: {str(e)}"}

    # ==================== API Layer phương thức gọi ====================
    # thường gọi cad-cowork-cstool Tham số hóa của API Lệnh

    def _get_api_base_dir(self) -> str:
        """Lấy API Layer thư mục cơ sở"""
        return os.path.join(tempfile.gettempdir(), "CAD-Cowork")

    def _get_api_request_path(self, request_id: str = None) -> str:
        """Lấy API Đường dẫn tập tin yêu cầu

        Args:
            request_id: hỏi ID, Được sử dụng để cô lập các yêu cầu đồng thời. None sau đó sử dụng đường dẫn chung(Tương thích với các phiên bản cũ hơn cstool)
        """
        if request_id:
            return os.path.join(self._get_api_base_dir(), "requests", request_id, "request.json")
        return os.path.join(self._get_api_base_dir(), "request.json")

    def _get_api_result_path(self, request_id: str = None) -> str:
        """Lấy API Đường dẫn file kết quả

        Args:
            request_id: hỏi ID, Được sử dụng để cô lập các yêu cầu đồng thời. None sau đó sử dụng đường dẫn chung(Tương thích với các phiên bản cũ hơn cstool)
        """
        if request_id:
            return os.path.join(self._get_api_base_dir(), "requests", request_id, "result.json")
        return os.path.join(self._get_api_base_dir(), "results", "latest.json")

    def _ensure_api_dirs(self, request_id: str = None) -> None:
        """đảm bảo API Thư mục tồn tại

        Args:
            request_id: hỏi ID, Một thư mục yêu cầu riêng sẽ được tạo
        """
        base_dir = self._get_api_base_dir()
        results_dir = os.path.join(base_dir, "results")
        os.makedirs(results_dir, exist_ok=True)
        if request_id:
            request_dir = os.path.join(base_dir, "requests", request_id)
            os.makedirs(request_dir, exist_ok=True)

    def _cleanup_request_dir(self, request_id: str) -> None:
        """Làm sạch thư mục cách ly yêu cầu(nỗ lực tốt nhất, Không có lỗi được báo cáo)"""
        try:
            request_dir = os.path.join(self._get_api_base_dir(), "requests", request_id)
            if os.path.exists(request_dir):
                import shutil
                shutil.rmtree(request_dir, ignore_errors=True)
        except Exception:
            pass

    def _normalize_result(self, result: Dict[str, Any]) -> Dict[str, Any]:
        """thống nhất API Trường hợp trường kết quả(C# sử dụng PascalCase, Python Thói quen camelCase)"""
        field_mapping = {
            "Success": "success",
            "ErrorCode": "errorCode",
            "Message": "message",
            "Data": "data",
            "RequestId": "requestId",
            "CommandName": "commandName",
            "ExecutionTimeMs": "executionTimeMs",
            "Timestamp": "timestamp",
            "CreatedEntityHandles": "createdEntityHandles",
            "ModifiedEntityHandles": "modifiedEntityHandles"
        }

        normalized = {}
        for key, value in result.items():
            # trong trường hợp PascalCase chìa khóa, Chuyển đổi thành camelCase
            new_key = field_mapping.get(key, key)
            normalized[new_key] = value

        return normalized

    def execute_api_command(self, command_name: str, parameters: Dict[str, Any] = None,
                           timeout: float = 10.0, application_context: bool = False) -> Dict[str, Any]:
        """Điều phối lệnh về Native COM (tương thích ngược).

        Giữ tên hàm cũ để batch_executor/direct_command không gãy,
        nhưng mọi lệnh đều đi qua smart_cad_command (Native COM 100%, Zero-DLL).
        """
        return self.smart_cad_command(
            command_name, parameters, need_result=True, timeout=timeout,
        )

    def _execute_api_command_once(self, command_name: str, parameters: Dict[str, Any] = None,
                                  timeout: float = 10.0, _is_retry: bool = False,
                                  application_context: bool = False) -> Dict[str, Any]:
        """Alias tương thích ngược -> smart_cad_command (Native COM)."""
        return self.smart_cad_command(
            command_name, parameters, need_result=True, timeout=timeout,
        )

    def get_api_commands_metadata(self) -> Dict[str, Any]:
        """Nhận tất cả siêu dữ liệu lệnh API từ manifest cục bộ đã Việt hóa."""
        manifest_path = os.path.join(os.path.dirname(__file__), "api_commands_manifest.json")
        if os.path.exists(manifest_path):
            try:
                with open(manifest_path, 'r', encoding='utf-8') as f:
                    manifest = json.load(f)
                commands = []
                for cat in manifest.get("categories", []):
                    commands.extend(cat.get("commands", []))
                return {
                    "success": True,
                    "commands": commands,
                    "count": len(commands)
                }
            except Exception as e:
                return {"success": False, "message": f"Lỗi đọc manifest: {e}", "commands": []}
        return {"success": False, "message": "Không tìm thấy file manifest lệnh API", "commands": []}

    def get_api_command_info(self, command_name: str) -> Dict[str, Any]:
        """Lấy thông tin chi tiết lệnh API chỉ định từ manifest cục bộ.

        Args:
            command_name: Tên lệnh

        Returns:
            dict: Chi tiết lệnh
        """
        all_commands = self.get_api_commands_metadata()
        if not all_commands.get("success"):
            return all_commands

        for cmd in all_commands.get("commands", []):
            if cmd.get("name", "").lower() == command_name.lower():
                return {"success": True, "command": cmd}

        return {
            "success": False,
            "message": f"Không tìm thấy lệnh: {command_name}"
        }

    # ==================== thông minh CAD Lệnh ====================
    # Tự động xác định việc sử dụng SendCommand vẫn API Layer

    def smart_cad_command(self, command_name: str, parameters: Dict[str, Any] = None,
                          need_result: bool = None, timeout: float = 10.0, confirm: bool = False) -> Dict[str, Any]:
        """Thực thi thông minh lệnh CAD"""
        if not self.is_running():
            started = self.start_cad()
            if not started:
                return {
                    "success": False,
                    "errorCode": -8,
                    "message": "Không tìm thấy phần mềm CAD đang chạy (đã quét AutoCAD/GstarCAD/ZWCAD). Vui lòng mở CAD và mở bản vẽ rồi thử lại.",
                    "hint": "MCP sẽ tự động kết nối với bất kỳ phiên bản CAD nào đang mở.",
                }

        # Lệnh ngữ cảnh ứng dụng (Mở/Đóng bản vẽ) Phải đi API Layer + DLApiExecuteSession
        application_context = command_name in APPLICATION_CONTEXT_COMMANDS
        if application_context and timeout < 30.0:
            # Mở/Lưu bản vẽ có thể chậm, tăng giới hạn thời gian chờ
            timeout = 30.0

        # Xác định phương pháp thực hiện nào sẽ sử dụng
        use_api = True if application_context else self._should_use_api(command_name, parameters, need_result)
        method_name = "API Layer(Session)" if application_context else ("API Layer" if use_api else "SendCommand")
        logger.info(f"[ĐIỀU HƯỚNG] Lệnh '{command_name}' thực thi bằng {method_name}")

        # Ưu tiên thực thi trực tiếp qua Native COM Interop 100% không phụ thuộc DLL
        # 1. Các lệnh truy vấn trạng thái và cấu trúc bản vẽ
        if command_name in ("api_get_document_info", "get_document_info", "api_document_info", "document_info"):
            return self.get_document_info()
        elif command_name in ("api_open_document", "open_document", "open_drawing") and parameters:
            fp = parameters.get("filePath") or parameters.get("file_path") or parameters.get("path")
            if fp:
                return self.open_document(fp)
            return {"success": False, "errorCode": -2,
                    "message": "Thiếu tham số filePath (đường dẫn .dwg/.dxf cần mở).",
                    "commandName": command_name, "executionMethod": "direct_com"}
        elif command_name in ("api_close_document", "close_document", "close_drawing"):
            save = False
            if parameters:
                save = bool(parameters.get("save") or parameters.get("saveChanges") or parameters.get("save_changes", False))
            return self.close_document(save_changes=save)
        elif command_name in ("api_get_layers", "get_layers", "api_list_layers", "list_layers", "api_get_layer_list"):
            return self.get_layers()
        elif command_name in ("api_get_active_layer", "get_active_layer"):
            return self.get_active_layer()
        elif command_name in ("api_get_entities_summary", "get_entities_summary", "api_get_entities", "get_entities", "api_list_entities"):
            return self.get_entities_summary()
        elif command_name in ("api_zoom_extents", "zoom_extents", "api_zoom", "zoom"):
            ok = self.zoom_extents()
            return {"success": ok, "message": "Đã zoom extents toàn bộ bản vẽ", "executionMethod": "direct_com"}
        elif command_name in ("api_save_drawing", "save_drawing") and parameters:
            fp = parameters.get("filePath") or parameters.get("file_path") or parameters.get("path")
            if fp:
                ok = self.save_drawing(fp)
                return {"success": ok, "message": f"Lưu bản vẽ {'thành công' if ok else 'thất bại'}: {fp}", "executionMethod": "direct_com"}
        elif command_name in ("api_create_layer", "create_layer") and parameters:
            l_name = parameters.get("layerName") or parameters.get("layer_name") or parameters.get("name")
            if l_name:
                ok = self.create_layer(l_name)
                return {"success": ok, "message": f"Tạo layer '{l_name}' {'thành công' if ok else 'thất bại'}", "executionMethod": "direct_com"}

        # 2. Các lệnh vẽ hình học cơ bản
        elif command_name in ("api_draw_circle", "draw_circle") and parameters:
            center = parameters.get("center", [0, 0, 0])
            radius = parameters.get("radius", 50)
            color = parameters.get("color")
            layer = parameters.get("layer")
            entity = self.draw_circle(center, radius, layer=layer, color=color)
            if entity:
                return {"success": True, "message": "Đã vẽ hình tròn trực tiếp vào AutoCAD", "executionMethod": "direct_com"}
        elif command_name in ("api_draw_line", "draw_line") and parameters:
            p1 = parameters.get("startPoint") or parameters.get("p1", [0, 0, 0])
            p2 = parameters.get("endPoint") or parameters.get("p2", [100, 100, 0])
            entity = self.draw_line(p1, p2, layer=parameters.get("layer"), color=parameters.get("color"))
            if entity:
                return {"success": True, "message": "Đã vẽ đoạn thẳng trực tiếp vào AutoCAD", "executionMethod": "direct_com"}
        elif command_name in ("api_draw_polyline", "draw_polyline") and parameters:
            pts = parameters.get("points", [])
            closed = parameters.get("closed", False)
            entity = self.draw_polyline(pts, closed=closed, layer=parameters.get("layer"), color=parameters.get("color"))
            if entity:
                return {"success": True, "message": "Đã vẽ polyline trực tiếp vào AutoCAD", "executionMethod": "direct_com"}
        elif command_name in ("api_draw_rectangle", "draw_rectangle") and parameters:
            c1 = parameters.get("corner1", [0, 0, 0])
            c2 = parameters.get("corner2", [100, 100, 0])
            entity = self.draw_rectangle(c1, c2, layer=parameters.get("layer"), color=parameters.get("color"))
            if entity:
                return {"success": True, "message": "Đã vẽ hình chữ nhật trực tiếp vào AutoCAD", "executionMethod": "direct_com"}
        elif command_name in ("api_draw_text", "draw_text") and parameters:
            pos = parameters.get("position", [0, 0, 0])
            text_str = parameters.get("text", "")
            h = parameters.get("height", 2.5)
            entity = self.draw_text(pos, text_str, height=h, layer=parameters.get("layer"))
            if entity:
                return {"success": True, "message": f"Đã vẽ text '{text_str}' trực tiếp vào AutoCAD", "executionMethod": "direct_com"}
        elif command_name in ("api_draw_arc", "draw_arc") and parameters:
            center = parameters.get("center", [0, 0, 0])
            radius = parameters.get("radius", 50)
            sa = parameters.get("startAngle", 0)
            ea = parameters.get("endAngle", 90)
            entity = self.draw_arc(center, radius, sa, ea, layer=parameters.get("layer"), color=parameters.get("color"))
            if entity:
                return {"success": True, "message": "Đã vẽ cung tròn trực tiếp vào AutoCAD", "executionMethod": "direct_com"}
        elif command_name in ("api_draw_ellipse", "draw_ellipse") and parameters:
            center = parameters.get("center", [0, 0, 0])
            major = parameters.get("majorAxis", [100, 0, 0])
            ratio = parameters.get("radiusRatio", 0.5)
            entity = self.draw_ellipse(center, major, ratio, layer=parameters.get("layer"), color=parameters.get("color"))
            if entity:
                return {"success": True, "message": "Đã vẽ hình elip trực tiếp vào AutoCAD", "executionMethod": "direct_com"}

        # 3. Các lệnh Bóc tách Khối lượng QS Takeoff (Native COM 100%)
        elif command_name in ("get_units", "api_get_units", "cad_get_units"):
            return self.get_units()
        elif command_name in ("quantity_by_layer", "api_quantity_by_layer", "cad_quantity_by_layer", "takeoff_layers", "cad_takeoff_layers"):
            layers = None
            if parameters:
                layers = parameters.get("layers") or parameters.get("layer_list")
            return self.quantity_by_layer(layers=layers)
        elif command_name in ("count_blocks", "api_count_blocks", "cad_count_blocks"):
            block_names = None
            include_attributes = True
            if parameters:
                block_names = parameters.get("blockNames") or parameters.get("block_names") or parameters.get("names")
                if "includeAttributes" in parameters:
                    include_attributes = bool(parameters.get("includeAttributes"))
                elif "include_attributes" in parameters:
                    include_attributes = bool(parameters.get("include_attributes"))
            return self.count_blocks(block_names=block_names, include_attributes=include_attributes)
        elif command_name in ("export_boq", "api_export_boq", "cad_export_boq", "cad_export_takeoff_excel", "export_takeoff_excel"):
            file_path = None
            fmt = "xlsx"
            title = "BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ)"
            quantities_data = None
            blocks_data = None
            auto_cleanup = False
            if parameters:
                file_path = parameters.get("filePath") or parameters.get("file_path") or parameters.get("path") or parameters.get("output")
                fmt = parameters.get("format", "xlsx")
                title = parameters.get("title") or title
                quantities_data = parameters.get("quantitiesData") or parameters.get("quantities_data")
                blocks_data = parameters.get("blocksData") or parameters.get("blocks_data")
                auto_cleanup = bool(parameters.get("autoCleanup") or parameters.get("auto_cleanup", False))
            return self.export_boq(file_path=file_path, format=fmt, quantities_data=quantities_data, blocks_data=blocks_data, title=title, auto_cleanup=auto_cleanup)
        elif command_name in ("disconnect", "cleanup", "cad_disconnect", "close_session"):
            return self.disconnect()

        if use_api:
            # Lệnh api_* chưa có nhánh Native COM ở trên: báo rõ, không gọi
            # ngược execute_api_command (tránh đệ quy vô hạn), không cầu .NET chết.
            return {
                "success": False,
                "errorCode": -4,
                "message": (
                    f"Lệnh '{command_name}' chưa có cài đặt Native COM trong phiên bản này. "
                    "Các lệnh đã hỗ trợ: document/layer/query/draw (line, circle, arc, ellipse, polyline, rectangle, text), "
                    "zoom_extents, save, QS takeoff (get_units, quantity_by_layer, count_blocks, export_boq), disconnect. "
                    "Gọi query_api_commands để tra cứu, hoặc dùng batch_execute/dry_run để kiểm tra trước."
                ),
                "commandName": command_name,
                "executionMethod": "native_com_unsupported",
            }
        else:
            # Sử dụng SendCommand với kiểm soát an ninh Security Guard và confirm
            params_str = self._build_sendcommand_params(parameters)
            cmd_str = f"{command_name} {params_str}".strip() if params_str else command_name
            result = self.send_command(cmd_str, confirm=confirm)
            result["executionMethod"] = "sendcommand"
            return result

    def _should_use_api(self, command_name: str, parameters: Dict[str, Any],
                        need_result: bool) -> bool:
        """Xác định xem có nên sử dụng API Layer

        Args:
            command_name: Tên lệnh
            parameters: Tham số lệnh
            need_result: Người dùng chỉ định rõ ràng liệu giá trị trả về có cần thiết hay không

        Returns:
            bool: True sử dụng API Layer, False sử dụng SendCommand
        """
        # Các lệnh AutoCAD CLI hoặc AutoCAD tiêu chuẩn (không bắt đầu bằng api_) luôn dùng SendCommand
        if not command_name.lower().startswith("api_"):
            logger.debug(f"[Định tuyến thông minh] Lệnh CAD CLI '{command_name}', sử dụng SendCommand")
            return False

        # 1. Được chỉ định rõ ràng cho lệnh api_*
        if need_result is True:
            logger.debug("[Định tuyến thông minh] need_result=True, sử dụng API Layer")
            return True
        if need_result is False:
            logger.debug("[Định tuyến thông minh] need_result=False, sử dụng SendCommand")
            return False

        # 2. Suy luận tự động cho lệnh api_*
        return True

    def _build_sendcommand_params(self, parameters: Dict[str, Any]) -> str:
        """Chuyển đổi đối số từ điển thành SendCommand Định dạng chuỗi

        Args:
            parameters: từ điển tham số

        Returns:
            str: Chuỗi tham số được phân tách bằng dấu cách
        """
        if not parameters:
            return ""

        # Chuyển đổi các tham số đơn giản thành các chuỗi được phân tách bằng dấu cách
        # Để ý: Chỉ hoạt động với các giá trị đơn giản, Các giá trị phức tạp nên sử dụng API Layer
        parts = []
        for value in parameters.values():
            if value is not None and not isinstance(value, (list, dict, tuple)):
                parts.append(str(value))
        return " ".join(parts)


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    print("=" * 60)
    print("Đang kiểm tra và kết nối tới AutoCAD qua CADController...")
    print("=" * 60)
    controller = CADController()
    connected = controller.start_cad()
    if connected:
        print(f"\n Kết nối THÀNH CÔNG!")
        print(f" - Phần mềm CAD: {controller.active_cad_name} ({controller.active_cad_type})")
        if controller.doc:
            try:
                print(f" - Bản vẽ đang mở: {controller.doc.Name}")
            except Exception:
                print(f" - Đã kết nối nhưng chưa đọc được tên file bản vẽ.")
    else:
        print(f"\n❌ Kết nối THẤT BẠI!")
        print("Lưu ý:")
        print("1. Hãy mở sẵn AutoCAD và tạo/mở một bản vẽ trước khi chạy.")
        print("2. Hoặc chỉnh trong config.json ('type': 'autocad') để tự động mở AutoCAD nếu chưa chạy.")
    print("=" * 60)

