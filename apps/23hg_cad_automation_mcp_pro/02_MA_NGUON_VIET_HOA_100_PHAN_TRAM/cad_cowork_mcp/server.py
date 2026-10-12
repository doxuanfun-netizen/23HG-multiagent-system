import os
import sys
import json
import logging
import tempfile
from pathlib import Path
from mcp.server.models import InitializationOptions
import mcp.types as types
from mcp.server import NotificationOptions, Server
import mcp.server.stdio
from pydantic import AnyUrl
from typing import Any, Dict, List
import sys

sys.dont_write_bytecode = True

from .cad_controller import CADController
from .batch_executor import BatchExecutor
from .metrics import CommandMetrics, SessionMetrics


# Chỉ thiết lập UTF-8 cho stderr
# Không can thiệp stdin/stdout để bảo đảm chuẩn giao thức MCP
# Tránh reconfigure làm hỏng Content-Length handshake
if sys.platform == "win32" and os.environ.get('PYTHONIOENCODING') is None:
    if hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding="utf-8")

# Cấu hình logging
def _make_log_handlers():
    handlers = [logging.StreamHandler(sys.stderr)]
    try:
        log_dir = os.path.join(tempfile.gettempdir(), 'CAD-Cowork')
        os.makedirs(log_dir, exist_ok=True)
        handlers.append(logging.FileHandler(
            os.path.join(log_dir, 'cad_mcp.log'), encoding='utf-8'
        ))
    except OSError:
        pass
    return handlers

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=_make_log_handlers()
)


logger = logging.getLogger('mcp_cad_server')
logger.info("Khởi động máy chủ CAD MCP")


def load_prompt_template(cad_name: str = "CAD", drawing_state: str = "(no CAD activity yet)") -> str:
    """Tải file template prompt hệ thống, thay thế các placeholder động

    Args:
        cad_name: Tên phần mềm CAD để thay thế {{CAD_NAME}}
        drawing_state: Chuỗi trạng thái bản vẽ động để thay thế {{DRAWING_STATE}}

    Returns:
        str: Nội dung prompt sau xử lý
    """
    template_path = os.path.join(os.path.dirname(__file__), 'prompts', 'cad_system_prompt.md')
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            content = f.read()
        content = content.replace('{{CAD_NAME}}', cad_name)
        content = content.replace('{{DRAWING_STATE}}', drawing_state)
        logger.info(f"Đã tải mẫu prompt: {template_path}")
        return content
    except Exception as e:
        logger.warning(f"Tải mẫu prompt thất bại: {e}, sử dụng prompt mặc định")
        return f"""Bạn là trợ lý vẽ CAD, có thể điều khiển {cad_name} thông qua ngôn ngữ tự nhiên.
Vui lòng sử dụng công cụ smart_cad_command để thực thi các lệnh API."""


# NOTE: Prompt template luôn tải động qua load_prompt_template(cad_name, drawing_state)
# để phản ánh đúng CAD đang kết nối + trạng thái bản vẽ. Không cache tĩnh ở đây.


def load_api_commands_manifest():
    """Tải file danh mục lệnh API"""
    manifest_path = os.path.join(os.path.dirname(__file__), 'api_commands_manifest.json')
    try:
        with open(manifest_path, 'r', encoding='utf-8-sig') as f:
            return json.load(f)
    except Exception as e:
        logger.warning(f"Tải danh mục lệnh API thất bại: {e}")
        return None

# Nạp toàn cục danh mục lệnh API
API_COMMANDS_MANIFEST = load_api_commands_manifest()
if API_COMMANDS_MANIFEST:
    logger.info(f"Đã tải danh mục lệnh API CAD, tổng cộng {API_COMMANDS_MANIFEST.get('totalCommands', 0)} lệnh")


# Loại CAD -> Tên hiển thị thân thiện
_CAD_DISPLAY_NAMES = {
    "auto": "CAD",
    "autocad": "AutoCAD",
    "gcad": "GstarCAD",
    "gstarcad": "GstarCAD",
    "zwcad": "ZWCAD",
}


class Config:

    def __init__(self):
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        logger.info("Đã tải cấu hình thành công")

    @property
    def server_name(self) -> str:
        return self.config['server']['name']

    @property
    def server_version(self) -> str:
        return self.config['server']['version']

    @property
    def cad_display_name(self) -> str:
        cad_type = self.config.get("cad", {}).get("type", "CAD")
        return _CAD_DISPLAY_NAMES.get(cad_type.lower(), cad_type)

class DrawingState:
    """Ảnh chụp trạng thái bản vẽ gọn nhẹ, cập nhật sau mỗi lệnh

    Được chèn vào system prompt dưới dạng ngữ cảnh động, giảm truy vấn thừa.
    """

    def __init__(self):
        self.active_document: str | None = None
        self.layer_count: int = 0
        self.recent_handles: list[str] = []  # 20 handle gần nhất
        self.last_command: str = ""
        self.last_error: str | None = None
        self.command_count: int = 0

    def update_from_result(self, command_name: str, result: dict) -> None:
        """Cập nhật trạng thái từ kết quả thực thi lệnh"""
        self.last_command = command_name
        self.command_count += 1

        if not result.get("success"):
            self.last_error = result.get("message", "unknown error")
            return

        self.last_error = None

        # Sau lệnh mở/đóng bản vẽ, handle cũ hết hiệu lực, làm sạch cache
        # (active_document sẽ được cập nhật từ kết nối mới)
        if command_name in ("api_open_document", "api_close_document"):
            self.recent_handles = []
            return

        # Thu thập handle tạo mới và sửa đổi
        new_handles = []
        new_handles.extend(result.get("createdEntityHandles", []))
        new_handles.extend(result.get("modifiedEntityHandles", []))

        # Trích xuất handle đơn lẻ từ data
        data = result.get("data", {})
        if isinstance(data, dict) and data.get("entityHandle"):
            handle = data["entityHandle"]
            if handle not in new_handles:
                new_handles.append(handle)

        if new_handles:
            # Gộp vào recent_handles sau khi khử trùng lặp (mới nhất xếp trước)
            seen = set(self.recent_handles)
            merged = [h for h in new_handles if h not in seen]
            self.recent_handles = (merged + self.recent_handles)[:20]

    def update_document_info(self, controller) -> None:
        """Cập nhật thông tin bản vẽ từ CAD controller (an toàn ngoại lệ)"""
        try:
            if controller.doc:
                self.active_document = controller.doc.Name
        except Exception:
            pass

    def to_context_string(self) -> str:
        """Tạo chuỗi ngữ cảnh động chèn vào system prompt"""
        parts = []
        if self.active_document:
            parts.append(f"Active document: {self.active_document}")
        if self.recent_handles:
            parts.append(f"Recent handles ({len(self.recent_handles)}): {', '.join(self.recent_handles[:10])}")
        if self.last_command:
            status = "✓" if not self.last_error else f"✗ {self.last_error}"
            parts.append(f"Last command: {self.last_command} [{status}]")
        parts.append(f"Commands executed this session: {self.command_count}")
        return "\n".join(parts) if parts else "(no CAD activity yet)"


class CADService:
    def __init__(self):
        """Khởi tạo dịch vụ CAD"""
        self.controller = CADController()
        self.drawing_state = DrawingState()
        self.session_metrics = SessionMetrics()
        logger.info("Dịch vụ CAD đã khởi tạo")

    def start_cad(self):
        """Khởi động CAD"""
        result = self.controller.start_cad()
        if result:
            self.drawing_state.update_document_info(self.controller)
        return result

def _query_api_commands(cad_service: CADService, command_name: str = None, category: str = None) -> dict:
    """Logic tra cứu lệnh API thống nhất

    Args:
        cad_service: Thực thể CAD service
        command_name: Tra cứu chi tiết lệnh chỉ định
        category: Lọc danh sách theo phân loại

    Returns:
        dict: Kết quả tra cứu
    """
    if not API_COMMANDS_MANIFEST:
        return {"success": False, "error": "Danh mục lệnh API cục bộ không khả dụng"}

    # Chế độ 1: Tra cứu chi tiết một lệnh
    if command_name:
        for cat in API_COMMANDS_MANIFEST.get("categories", []):
            for cmd in cat.get("commands", []):
                if cmd.get("name") == command_name:
                    return {
                        "success": True,
                        "source": "local_manifest",
                        "command": cmd,
                        "category": cat.get("name")
                    }
        return {"success": False, "error": f"Không tìm thấy lệnh: {command_name}"}

    # Chế độ 2: Lọc theo phân loại
    if category:
        for cat in API_COMMANDS_MANIFEST.get("categories", []):
            if cat.get("nameEn", "").lower() == category.lower() or cat.get("name", "") == category:
                return {
                    "success": True,
                    "source": "local_manifest",
                    "category": cat.get("name"),
                    "count": cat.get("count", len(cat.get("commands", []))),
                    "commands": [
                        {"name": cmd["name"], "displayName": cmd.get("displayName", ""), "description": cmd.get("description", "")}
                        for cmd in cat.get("commands", [])
                    ]
                }
        return {"success": False, "error": f"Không tìm thấy phân loại: {category}"}

    # Chế độ 3: Trả về danh sách tổng quan các lệnh
    categories_summary = []
    for cat in API_COMMANDS_MANIFEST.get("categories", []):
        categories_summary.append({
            "name": cat.get("name"),
            "nameEn": cat.get("nameEn"),
            "count": cat.get("count", len(cat.get("commands", []))),
            "commands": [cmd["name"] for cmd in cat.get("commands", [])]
        })
    return {
        "success": True,
        "source": "local_manifest",
        "totalCommands": API_COMMANDS_MANIFEST.get("totalCommands", 0),
        "categories": categories_summary
    }


def _format_param_hint(p: dict) -> str:
    """Định dạng gợi ý tham số đơn lẻ"""
    name = p.get("name", "?")
    ptype = p.get("type", "?")
    aliases = p.get("aliases", [])
    if aliases:
        return f"{name}({ptype}, alias: {'/'.join(aliases)})"
    return f"{name}({ptype})"


def _get_command_fix_hint(command_name: str) -> dict | None:
    """Lấy gợi ý định dạng tham số từ manifest cục bộ

    Args:
        command_name: Tên lệnh API

    Returns:
        dict | None: Thông tin gợi ý tham số và ví dụ
    """
    if not API_COMMANDS_MANIFEST:
        return None

    for cat in API_COMMANDS_MANIFEST.get("categories", []):
        for cmd in cat.get("commands", []):
            if cmd.get("name") == command_name:
                params = cmd.get("parameters", [])
                required = [_format_param_hint(p) for p in params if p.get("required")]
                optional = [_format_param_hint(p) for p in params if not p.get("required")]
                hint = {
                    "required_params": ", ".join(required) if required else "(none)",
                    "optional_params": ", ".join(optional) if optional else "(none)",
                }
                example = cmd.get("usageExample")
                if example:
                    hint["example"] = example
                return hint
    return None


async def main():
    """Hàm khởi chạy chính"""
    logger.info("Khởi động máy chủ CAD MCP")

    # Nạp cấu hình
    config = Config()
    cad_service = CADService()
    server = Server(config.server_name)

    # Đăng ký các handler
    logger.debug("Đăng ký các handler")

    @server.list_resources()
    async def handle_list_resources() -> list[types.Resource]:
        logger.debug("Xử lý yêu cầu list_resources")
        return [
            types.Resource(
                uri=AnyUrl("drawing://current"),
                name="Bản vẽ CAD hiện hành",
                description="Trạng thái bản vẽ CAD hiện hành",
                mimeType="application/json",
            )
        ]

    @server.read_resource()
    async def handle_read_resource(uri: AnyUrl) -> str:
        logger.debug(f"Xử lý yêu cầu read_resource, URI: {uri}")
        if uri.scheme != "drawing":
            logger.error(f"Giao thức URI không hỗ trợ: {uri.scheme}")
            raise ValueError(f"Giao thức URI không hỗ trợ: {uri.scheme}")

        path = str(uri).replace("drawing://", "")
        if not path or path != "current":
            logger.error(f"Đường dẫn tài nguyên không xác định: {path}")
            raise ValueError(f"Đường dẫn tài nguyên không xác định: {path}")

        # Làm mới thông tin bản vẽ rồi trả về
        cad_service.drawing_state.update_document_info(cad_service.controller)
        state = {
            "active_document": cad_service.drawing_state.active_document,
            "recent_handles": cad_service.drawing_state.recent_handles,
            "last_command": cad_service.drawing_state.last_command,
            "last_error": cad_service.drawing_state.last_error,
            "command_count": cad_service.drawing_state.command_count,
        }
        return json.dumps(state, ensure_ascii=False)


    @server.list_tools()
    async def handle_list_tools() -> list[types.Tool]:
        """Liệt kê các công cụ khả dụng"""
        return [
            # ===== CÔNG CỤ THỰC THI CỐT LÕI =====
            types.Tool(
                name="smart_cad_command",
                description=(
                    "Execute a single CAD command. This is the PRIMARY tool for ALL CAD operations.\n\n"
                    "WHEN TO USE:\n"
                    "- Any single CAD operation: draw, query, modify, block, layer, view, selection, spatial calc, etc.\n"
                    "- For GsLc domain commands (api_gslc_*): call gslc_route_command FIRST to get the correct command name, then use this tool.\n\n"
                    "WHEN NOT TO USE:\n"
                    "- Multiple dependent commands in sequence → use batch_execute instead.\n"
                    "- To look up available commands or parameter formats → use query_api_commands instead.\n\n"
                    "PARAMETER FORMAT:\n"
                    "- command_name: exact api_* name (e.g. 'api_draw_circle', 'api_move_entity')\n"
                    "- parameters: JSON object with camelCase keys\n"
                    "- Points are [x, y, z] arrays (z defaults to 0 if omitted)\n"
                    "- Entity handles are hex strings from previous command results (e.g. '27E2C0')\n\n"
                    "RETURNS: {success, data, createdEntityHandles[], executionMethod}\n"
                    "- Save createdEntityHandles for subsequent move/modify/delete operations.\n\n"
                    "COMMON MISTAKES:\n"
                    "- Using api_gslc_* commands directly without gslc_route_command → may pick wrong command due to similar names.\n"
                    "- Passing snake_case parameter names → must be camelCase (e.g. startPoint, not start_point).\n"
                    "- Omitting required parameters → check with query_api_commands first if unsure."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "command_name": {
                            "type": "string",
                            "description": "API command name (e.g. 'api_draw_circle', 'api_move_entity', 'api_query_by_layer')"
                        },
                        "parameters": {
                            "type": "object",
                            "description": "Command parameters as JSON object with camelCase keys. Example: {\"center\": [100,100,0], \"radius\": 50}"
                        },
                        "need_result": {
                            "type": "boolean",
                            "description": "Force execution method: true=API Layer (structured return), false=SendCommand (fast, no return data), omit=auto-detect"
                        },
                        "timeout": {
                            "type": "number",
                            "description": "Timeout in seconds (default: 10)"
                        },
                        "confirm": {
                            "type": "boolean",
                            "description": "Cờ xác nhận tường minh từ người dùng (confirm=true). BẮT BUỘC đặt true khi thực hiện các lệnh có nguy cơ phá hoại/thay đổi hàng loạt dữ liệu bản vẽ (ví dụ: ERASE, EXPLODE, ssget, xóa layer, v.v.). Mặc định: false.",
                            "default": False
                        }
                    },
                    "required": ["command_name"],
                },
            ),

            # ===== CÔNG CỤ THỰC THI HÀNG LOẠT =====
            types.Tool(
                name="batch_execute",
                description=(
                    "Execute multiple CAD commands in a single call with DAG dependency management.\n\n"
                    "WHEN TO USE:\n"
                    "- Drawing multiple entities that depend on each other (e.g. draw line → move it → copy it).\n"
                    "- Any workflow requiring 2+ sequential commands where later commands reference earlier results.\n\n"
                    "WHEN NOT TO USE:\n"
                    "- Single command → use smart_cad_command instead (simpler, less overhead).\n"
                    "- Independent commands with no dependencies → multiple smart_cad_command calls may be clearer.\n\n"
                    "DEPENDENCY & REFERENCE SYNTAX:\n"
                    "- depends_on: [\"cmd_id\"] → ensures cmd_id completes before this command runs.\n"
                    "- ${cmd_id.handle} → resolves to the first created entity handle from cmd_id.\n"
                    "- ${cmd_id.data.fieldName} → resolves to a nested field in cmd_id's result data.\n"
                    "- ${cmd_id.createdEntityHandles[0]} → resolves to a specific array element.\n\n"
                    "OPTIONS:\n"
                    "- dry_run: true → validate all parameters and dependencies WITHOUT executing. Returns execution plan.\n"
                    "- use_transaction: true → wrap all commands in a transaction; auto-rollback on any failure. WARNING: transactions expire after 30 minutes.\n"
                    "- detail_level: 'summary' (stats only) | 'handles' (stats + handle map, default) | 'full' (all data).\n\n"
                    "RETURNS: {success, summary: {total, successful, failed, total_time_ms}, handle_map, failed_commands[]}\n\n"
                    "EXAMPLE:\n"
                    "commands: [\n"
                    "  {id: 'c1', command_name: 'api_draw_circle', parameters: {center: [0,0,0], radius: 50}},\n"
                    "  {id: 'c2', command_name: 'api_move_entity', parameters: {entityHandle: '${c1.handle}', fromPoint: [0,0,0], toPoint: [100,0,0]}, depends_on: ['c1']}\n"
                    "]"
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "commands": {
                            "type": "array",
                            "description": "Command list with DAG dependencies",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "id": {
                                        "type": "string",
                                        "description": "Unique command ID for dependency references"
                                    },
                                    "command_name": {
                                        "type": "string",
                                        "description": "API command name (e.g. 'api_draw_circle')"
                                    },
                                    "parameters": {
                                        "type": "object",
                                        "description": "Command parameters. Use ${cmd_id.field} to reference previous results."
                                    },
                                    "depends_on": {
                                        "type": "array",
                                        "description": "IDs of commands that must complete before this one",
                                        "items": {"type": "string"}
                                    },
                                    "confirm": {
                                        "type": "boolean",
                                        "description": "Cờ xác nhận riêng cho lệnh này (nếu là lệnh phá hoại như ERASE, EXPLODE, ssget)"
                                    }
                                },
                                "required": ["id", "command_name"]
                            }
                        },
                        "stop_on_error": {
                            "type": "boolean",
                            "description": "Stop all remaining commands on first failure (default: true)"
                        },
                        "timeout_per_command": {
                            "type": "number",
                            "description": "Timeout per command in seconds (default: 10)"
                        },
                        "detail_level": {
                            "type": "string",
                            "description": "Result verbosity: 'summary', 'handles' (default), 'full'",
                            "enum": ["summary", "handles", "full"],
                            "default": "handles"
                        },
                        "use_transaction": {
                            "type": "boolean",
                            "description": "Wrap all commands in a single transaction with auto-rollback on failure (default: false)",
                            "default": False
                        },
                        "dry_run": {
                            "type": "boolean",
                            "description": "Validate parameters and dependencies only, do NOT execute. Returns execution plan preview. (default: false)",
                            "default": False
                        },
                        "confirm": {
                            "type": "boolean",
                            "description": "Cờ xác nhận toàn cục cho toàn bộ batch (confirm=true). Cho phép các lệnh có nguy cơ phá hoại (ERASE, EXPLODE, ssget) được thực thi nếu được người dùng duyệt. Mặc định: false.",
                            "default": False
                        }
                    },
                    "required": ["commands"],
                },
            ),

            # ===== CÔNG CỤ TRA CỨU LỆNH =====
            types.Tool(
                name="query_api_commands",
                description=(
                    "Look up available CAD API commands and their parameter definitions.\n\n"
                    "WHEN TO USE:\n"
                    "- Before calling smart_cad_command, if you are unsure about the exact command name or required parameters.\n"
                    "- To discover what commands are available in a specific category.\n\n"
                    "MODES (set exactly one):\n"
                    "- Neither command_name nor category → returns full command list grouped by category.\n"
                    "- command_name: 'api_draw_circle' → returns detailed parameter definitions and usage example for that command.\n"
                    "- category: 'drawing' → returns all commands in that category.\n\n"
                    "AVAILABLE CATEGORIES: drawing, query, block, modify, batch, spatial, layer, document, view, selection, transaction, gslc\n\n"
                    "RETURNS: {success, source, commands[] or command: {name, parameters[], usageExample}}"
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "command_name": {
                            "type": "string",
                            "description": "Exact command name to look up (e.g. 'api_draw_circle'). Returns detailed parameter info."
                        },
                        "category": {
                            "type": "string",
                            "description": "Filter by category English name (e.g. 'drawing', 'query', 'gslc'). Returns all commands in that category."
                        }
                    },
                    "required": [],
                },
            ),



            # ===== GsLc Điều hướng sơ đồ quy trình công nghệ P&ID =====
            types.Tool(
                name="gslc_route_command",
                description=(
                    "Điều hướng các thao tác GsLc (Sơ đồ công nghệ / P&ID) sang lệnh api_gslc_* chính xác nhất.\n\n"
                    "KHI NÀO SỬ DỤNG:\n"
                    "- LUÔN GỌI lệnh này TRƯỚC KHI thực thi bất kỳ lệnh api_gslc_* nào.\n"
                    "- Khi người dùng đề cập đến các khái niệm GsLc: Thiết bị (equipment), Đường ống (pipe), Đồng hồ đo (instrument), Khung bản vẽ (frame), Mã vị trí (tag), Van (valve), Mũi tên liên kết (arrow).\n\n"
                    "KHI NÀO KHÔNG SỬ DỤNG:\n"
                    "- Các thao tác CAD thông thường (vẽ, di chuyển, layer, zoom, truy vấn...) → Dùng trực tiếp smart_cad_command.\n\n"
                    "LÝ DO CÔNG CỤ NÀY TỒN TẠI:\n"
                    "GsLc có 36 lệnh chuyên ngành dễ gây nhầm lẫn:\n"
                    "- 'Tạo hàng loạt mã vị trí' → Sai: auto_number (chỉnh sửa cái cũ), Đúng: insert_bottom_equip_tags (tạo mới)\n"
                    "- 'Chèn mũi tên nối bản vẽ' → Sai: insert_outer_pipe_arrow (ống ngoại vi), Đúng: insert_join_draw_arrow (xuyên bản vẽ)\n"
                    "- 'Làm mới hướng đến/đi của ống' → Sai: sync_pipe_data (đồng bộ từ Excel), Đúng: brush_pipe_from_to (làm mới liên kết)\n\n"
                    "QUY TRÌNH: gslc_route_command(intent) → Nhận lệnh đề xuất → smart_cad_command(recommended_command, params)\n\n"
                    "KẾT QUẢ TRẢ VỀ: {recommended: {command, score, reason}, candidates[], disambiguation_hint}"
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "user_intent": {
                            "type": "string",
                            "description": "Mô tả bằng ngôn ngữ tự nhiên về thao tác GsLc, ví dụ: 'Tạo hàng loạt mã vị trí thiết bị', 'Chèn van điều khiển', 'Làm mới hướng đường ống'"
                        },
                        "context": {
                            "type": "object",
                            "description": "Ngữ cảnh bổ sung để tăng độ chính xác điều hướng",
                            "properties": {
                                "has_frame_handle": {"type": "boolean", "description": "Đã có Handle khung bản vẽ hay chưa"},
                                "has_entity_handles": {"type": "boolean", "description": "Đã có danh sách Handle thực thể hay chưa"},
                                "target_entity_type": {"type": "string", "description": "Loại thực thể đích (equipment, instrument, pipe, valve, etc.)"}
                            }
                        }
                    },
                    "required": ["user_intent"]
                }
            ),

            # ===== CÔNG CỤ ĐO BÓC KHỐI LƯỢNG (QS TAKEOFF COM NATIVE) =====
            types.Tool(
                name="cad_get_units",
                description=(
                    "Đọc đơn vị bản vẽ CAD (INSUNITS, MEASUREMENT, DIMLFAC, LUNITS) và tính tỷ lệ quy đổi sang mét (m) và mét vuông (m2).\n\n"
                    "CÔNG DỤNG:\n"
                    "- Xác định bản vẽ đang dùng hệ mm, m, inch hay chưa chỉ định (Unspecified -> mặc định mm tại VN).\n"
                    "- Kiểm tra cờ cảnh báo tỷ lệ DIMLFAC khác 1.0 (tránh sai lệch kích thước hiển thị vs hình học thực tế)."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            ),

            types.Tool(
                name="cad_quantity_by_layer",
                description=(
                    "Duyệt ModelSpace trong CAD, bóc tách và cộng dồn Chiều dài (m) và Diện tích (m2) theo từng Layer và phân loại hình học (Lines, Polylines, Circles, Arcs, Hatches, Ellipses).\n\n"
                    "CÔNG DỤNG:\n"
                    "- Tự động quy đổi từ đơn vị bản vẽ sang mét và m2.\n"
                    "- Hỗ trợ lọc danh sách các Layer cụ thể (ví dụ: ['BETONG', 'TUONG_220', 'TRUC']).\n"
                    "- Thống kê số lượng polyline khép kín (Closed) để tính diện tích m2."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "layers": {
                            "type": "array",
                            "description": "Danh sách tên Layer cần bóc tách (để trống nếu muốn bóc tách toàn bộ ModelSpace)",
                            "items": {"type": "string"}
                        },
                        "area_mode": {
                            "type": "string",
                            "description": "Chế độ tính diện tích: 'separate' (khuyến nghị: tách riêng Polyline kín và Hatch chống cộng trùng), 'polyline_only', 'hatch_only', 'combined'",
                            "enum": ["separate", "polyline_only", "hatch_only", "combined"],
                            "default": "separate"
                        }
                    },
                    "required": []
                }
            ),

            types.Tool(
                name="cad_count_blocks",
                description=(
                    "Thống kê số lượng Block và thiết bị trong ModelSpace theo tên Block và Layer, tự động nhận diện Dynamic Block và trích xuất thuộc tính Attribute.\n\n"
                    "CÔNG DỤNG:\n"
                    "- Tự động phân giải tên hiển thị chuẩn (EffectiveName) của Dynamic Block thay vì tên nặc danh (*U...).\n"
                    "- Đọc giá trị thuộc tính Attribute (Tag & Value), phân bố số lượng theo giá trị thuộc tính (ví dụ: kích thước cửa, công suất thiết bị).\n"
                    "- Lọc theo danh sách tên block chỉ định nếu cần."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "block_names": {
                            "type": "array",
                            "description": "Danh sách tên block cần đếm (để trống để đếm toàn bộ block)",
                            "items": {"type": "string"}
                        },
                        "include_attributes": {
                            "type": "boolean",
                            "description": "Trích xuất giá trị thuộc tính Attribute của block (mặc định: true)",
                            "default": True
                        },
                        "scan_nested": {
                            "type": "boolean",
                            "description": "Quét cả block lồng nhau (nested blocks)",
                            "default": False
                        }
                    },
                    "required": []
                }
            ),

            types.Tool(
                name="cad_export_boq",
                description=(
                    "Xuất bảng tiên lượng khối lượng BoQ ra file Excel (.xlsx) đa sheet chuyên nghiệp với 100% CÔNG THỨC SỐNG (Zero Dead Numbers: =SUM, nhân hệ số quy đổi m/m2) hoặc file CSV.\n\n"
                    "CẤU TRÚC EXCEL:\n"
                    "- Sheet 1 (1_Tong_Hop_BoQ): Bảng tổng hợp tiên lượng chính, liên kết công thức sống từ Sheet 2 và Sheet 3.\n"
                    "- Sheet 2 (2_Chi_Tiet_Layer): Chi tiết từng layer với công thức quy đổi m, m2 và dòng tổng =SUM(...).\n"
                    "- Sheet 3 (3_Thong_Ke_Block): Thống kê block, thiết bị, thuộc tính attribute và tổng số lượng.\n"
                    "- Định dạng đẹp chuẩn AEC: font Segoe UI, viền bảng, tiêu đề Navy Blue, số thập phân rõ ràng."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "Đường dẫn file đầu ra (.xlsx hoặc .csv). Nếu không truyền, hệ thống tự sinh tên file kèm timestamp trong thư mục làm việc."
                        },
                        "format": {
                            "type": "string",
                            "description": "Định dạng file: 'xlsx' (khuyến nghị, 100% công thức sống) hoặc 'csv'",
                            "enum": ["xlsx", "csv"],
                            "default": "xlsx"
                        },
                        "title": {
                            "type": "string",
                            "description": "Tiêu đề bảng tiên lượng (mặc định: 'BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ)')",
                            "default": "BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ)"
                        },
                        "auto_cleanup": {
                            "type": "boolean",
                            "description": "Tự động ngắt kết nối COM và dọn dẹp tài nguyên nền sau khi xuất file xong để tránh tích tụ phiên ngầm (mặc định: false)",
                            "default": False
                        }
                    },
                    "required": []
                }
            ),

            types.Tool(
                name="cad_disconnect",
                description=(
                    "Giải phóng tường minh các tham chiếu COM, dọn dẹp tài nguyên và ngắt phiên kết nối nền với AutoCAD.\n\n"
                    "CÔNG DỤNG:\n"
                    "- Giải phóng RAM và đóng kênh COM RPC.\n"
                    "- Ngăn chặn Autodesk Licensing Service tích tụ các tiến trình con (AdskLicensingAgent.exe) chạy nền.\n"
                    "- Tự động kết nối lại khi có lệnh mới được gọi."
                ),
                inputSchema={
                    "type": "object",
                    "properties": {},
                    "required": []
                }
            ),

        ]

    @server.call_tool()
    async def handle_call_tool(
        name: str, arguments: dict[str, Any] | None
    ) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
        """Xử lý yêu cầu gọi công cụ"""
        try:
            # Đảm bảo arguments là dict, cho phép rỗng
            if arguments is None:
                arguments = {}

            if name == "smart_cad_command":
                command_name = arguments.get("command_name")
                parameters = arguments.get("parameters")
                need_result = arguments.get("need_result")  # Có thể là None
                timeout = arguments.get("timeout", 10.0)
                confirm = arguments.get("confirm", False)

                if not command_name:
                    raise ValueError("Thiếu tham số command_name")

                # Đảm bảo CAD đã kết nối
                if not cad_service.controller.is_running():
                    cad_service.start_cad()

                cmd_metrics = CommandMetrics(command_name)
                result = cad_service.controller.smart_cad_command(
                    command_name=command_name,
                    parameters=parameters,
                    need_result=need_result,
                    timeout=timeout,
                    confirm=confirm
                )
                cmd_metrics.finish(result)
                cad_service.session_metrics.record(cmd_metrics)

                # Sau khi mở/đóng bản vẽ, làm mới kết nối đến bản vẽ hiện hành
                # Đảm bảo đọc đúng tên bản vẽ hiện hành
                if command_name in ("api_open_document", "api_close_document") and result.get("success"):
                    try:
                        cad_service.controller._ensure_valid_connection()
                    except Exception:
                        pass

                # Cập nhật trạng thái bản vẽ
                cad_service.drawing_state.update_from_result(command_name, result)
                cad_service.drawing_state.update_document_info(cad_service.controller)

                # Kèm fix_hint khi có lỗi (lấy từ manifest)
                if not result.get("success") and result.get("errorCode") in (-3, -4):
                    fix_hint = _get_command_fix_hint(command_name)
                    if fix_hint:
                        result["fix_hint"] = fix_hint

                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name == "batch_execute":
                commands = arguments.get("commands")
                stop_on_error = arguments.get("stop_on_error", True)
                timeout_per_command = arguments.get("timeout_per_command", 10.0)
                detail_level = arguments.get("detail_level", "handles")
                use_transaction = arguments.get("use_transaction", False)
                dry_run = arguments.get("dry_run", False)
                confirm = arguments.get("confirm", False)

                if not commands:
                    raise ValueError("Thiếu tham số commands")

                # Chế độ dry_run: chỉ xác thực không thực thi
                if dry_run:
                    batch_executor = BatchExecutor(cad_service.controller)
                    result = batch_executor.dry_run(commands)
                    return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

                # Đảm bảo CAD đã kết nối
                if not cad_service.controller.is_running():
                    cad_service.start_cad()

                batch_executor = BatchExecutor(cad_service.controller)
                result = batch_executor.execute(
                    commands=commands,
                    stop_on_error=stop_on_error,
                    timeout_per_command=timeout_per_command,
                    detail_level=detail_level,
                    use_transaction=use_transaction,
                    confirm=confirm
                )
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name == "query_api_commands":
                command_name = arguments.get("command_name")
                category = arguments.get("category")
                result = _query_api_commands(cad_service, command_name, category)
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name == "gslc_route_command":
                from .domains.gslc.router import GsLcRouter

                user_intent = arguments.get("user_intent", "")
                context = arguments.get("context")

                if not user_intent:
                    raise ValueError("Thiếu tham số user_intent")

                router = GsLcRouter()
                result = router.route(user_intent, context)

                result_dict = {
                    "domain": result.domain,
                    "intent_category": result.intent_category,
                    "intent_confidence": result.intent_confidence,
                    "recommended": result.recommended,
                    "candidates": [
                        {
                            "command": c.command,
                            "score": c.score,
                            "category": c.category,
                            "reason": c.reason,
                            "why_not": c.why_not
                        } for c in result.candidates
                    ],
                    "disambiguation_hint": result.disambiguation_hint
                }
                return [types.TextContent(type="text", text=json.dumps(result_dict, ensure_ascii=False, indent=2))]

            elif name in ("cad_get_units", "get_units"):
                if not cad_service.controller.is_running():
                    cad_service.start_cad()
                result = cad_service.controller.get_units()
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name in ("cad_quantity_by_layer", "quantity_by_layer", "cad_takeoff_layers", "takeoff_layers"):
                if not cad_service.controller.is_running():
                    cad_service.start_cad()
                layers = arguments.get("layers")
                area_mode = arguments.get("area_mode", "separate")
                result = cad_service.controller.quantity_by_layer(layers=layers, area_mode=area_mode)
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name in ("cad_count_blocks", "count_blocks"):
                if not cad_service.controller.is_running():
                    cad_service.start_cad()
                block_names = arguments.get("block_names")
                include_attributes = arguments.get("include_attributes", True)
                scan_nested = arguments.get("scan_nested", False)
                result = cad_service.controller.count_blocks(block_names=block_names, include_attributes=include_attributes, scan_nested=scan_nested)
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name in ("cad_export_boq", "export_boq", "cad_export_takeoff_excel", "export_takeoff_excel"):
                if not cad_service.controller.is_running():
                    cad_service.start_cad()
                file_path = arguments.get("file_path")
                fmt = arguments.get("format", "xlsx")
                title = arguments.get("title", "BẢNG TIÊN LƯỢNG KHỐI LƯỢNG (BOQ)")
                auto_cleanup = arguments.get("auto_cleanup", False)
                result = cad_service.controller.export_boq(file_path=file_path, format=fmt, title=title, auto_cleanup=auto_cleanup)
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            elif name in ("cad_disconnect", "disconnect", "cleanup"):
                result = cad_service.controller.disconnect()
                return [types.TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]

            else:
                raise ValueError(f"Công cụ không xác định: {name}")

        except Exception as e:
            return [types.TextContent(type="text", text=json.dumps({
                "success": False,
                "error": str(e),
                "tool": name,
            }, ensure_ascii=False, indent=2))]

    @server.list_prompts()
    async def handle_list_prompts() -> list[types.Prompt]:
        logger.debug("Xử lý yêu cầu list_prompts")
        return [
            types.Prompt(
                name="cad-assistant",
                description="Trợ lý kỹ thuật điều khiển và tự động hóa CAD bằng ngôn ngữ tự nhiên",
                arguments=[],
            )
        ]

    @server.get_prompt()
    async def handle_get_prompt(name: str, arguments: dict[str, str] | None) -> types.GetPromptResult:
        logger.debug(f"Xử lý yêu cầu get_prompt, tên: {name}, tham số: {arguments}")
        if name != "cad-assistant":
            logger.error(f"Chỉ dẫn không xác định: {name}")
            raise ValueError(f"Chỉ dẫn không xác định: {name}")

        # Tự động cập nhật trạng thái bản vẽ hiện tại
        cad_service.drawing_state.update_document_info(cad_service.controller)
        # Ưu tiên lấy tên CAD thực tế đang kết nối; nếu chưa kết nối thì lấy tên từ config
        cad_name = cad_service.controller.active_cad_name or config.cad_display_name
        prompt = load_prompt_template(cad_name, cad_service.drawing_state.to_context_string())

        return types.GetPromptResult(
            description="Mẫu chỉ dẫn (Prompt Template) Trợ lý CAD tiếng Việt",
            messages=[
                types.PromptMessage(
                    role="user",
                    content=types.TextContent(type="text", text=prompt.strip()),
                )
            ],
        )


    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        logger.info("Máy chủ đang chạy qua kênh stdio")
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name=config.server_name,
                server_version=config.server_version,
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

def main_cli():
    """Hàm khởi chạy đồng bộ"""
    import argparse
    from . import __version__

    parser = argparse.ArgumentParser(description="CAD-Cowork-MCP Server")
    parser.add_argument("--version", action="version", version=f"cad-cowork-mcp {__version__}")
    parser.parse_args()

    import asyncio
    asyncio.run(main())


if __name__ == "__main__":
    main_cli()