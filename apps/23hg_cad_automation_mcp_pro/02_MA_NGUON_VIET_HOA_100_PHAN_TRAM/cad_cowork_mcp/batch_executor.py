"""
người thực thi lệnh hàng loạt

ủng hộ DAG Quản lý phụ thuộc và truyền tham chiếu tham số, Dùng để thực hiện nhiều API Đặt hàng. 
"""

import json
import re
import logging
import time
from typing import Any, Dict, List, Optional, Tuple
from collections import defaultdict

logger = logging.getLogger('batch_executor')


class BatchExecutionError(Exception):
    """Lỗi thực thi hàng loạt"""
    pass


class CyclicDependencyError(BatchExecutionError):
    """lỗi phụ thuộc vòng tròn"""
    pass


class ReferenceResolutionError(BatchExecutionError):
    """Lỗi phân tích tài liệu tham khảo"""
    pass


class BatchExecutor:
    """người thực thi lệnh hàng loạt, ủng hộ DAG Sự phụ thuộc và truyền tham số"""

    # Trích dẫn các quy tắc ngữ pháp: ${cmd_id.field.path} hoặc ${cmd_id.array[0]}
    REFERENCE_PATTERN = re.compile(r'\$\{([^}]+)\}')

    def __init__(self, cad_controller):
        """
        Khởi tạo trình thực thi hàng loạt

        Args:
            cad_controller: CADController Ví dụ
        """
        self.controller = cad_controller
        self.results: Dict[str, Dict] = {}      # Đặt hàngID -> Kết quả thực hiện
        self.handle_map: Dict[str, str] = {}    # Đặt hàngID -> chủ nhà Handle

    def dry_run(self, commands: List[Dict]) -> Dict[str, Any]:
        """Xác minh định dạng của danh sách lệnh, Phụ thuộc và tham chiếu tham số, Thực tế chưa được triển khai

        Được sử dụng để kiểm tra trước khi thực hiện hàng loạt, Xác nhận rằng tất cả các tham số đều hợp pháp, Không có sự phụ thuộc vòng tròn. 

        Args:
            commands: danh sách lệnh(Định dạng giống như execute)

        Returns:
            dict: Bao gồm valid, execution_order, issues[] Kết quả xác minh
        """
        issues: List[str] = []

        # 1. Xác minh định dạng
        validation_error = self._validate_commands(commands)
        if validation_error:
            return {
                "valid": False,
                "message": f"Xác thực định dạng lệnh thất bại: {validation_error}",
                "execution_order": [],
                "issues": [validation_error]
            }

        # 2. sắp xếp tôpô(Phát hiện sự phụ thuộc vòng tròn)
        try:
            execution_order = self._topological_sort(commands)
        except CyclicDependencyError as e:
            return {
                "valid": False,
                "message": f"Phát hiện phụ thuộc vòng tròn: {str(e)}",
                "execution_order": [],
                "issues": [f"Phụ thuộc vòng tròn: {str(e)}"]
            }

        # 3. Kiểm tra xem tham chiếu đối số có đề cập đến lệnh hiện có không ID
        cmd_ids = {cmd["id"] for cmd in commands}
        for cmd in commands:
            params_str = json.dumps(cmd.get("parameters", {}))
            for match in self.REFERENCE_PATTERN.finditer(params_str):
                ref_path = match.group(1)
                ref_cmd_id = ref_path.split(".", 1)[0]
                if ref_cmd_id not in cmd_ids:
                    issues.append(f"Lệnh '{cmd['id']}' tham chiếu đến ID lệnh không tồn tại: '{ref_cmd_id}'")
                # Kiểm tra xem lệnh được tham chiếu có nằm trong chuỗi phụ thuộc của lệnh hiện tại không
                depends_on = set(cmd.get("depends_on", []))
                if ref_cmd_id not in depends_on and ref_cmd_id != cmd["id"]:
                    issues.append(
                        f"Lệnh '{cmd['id']}' tham chiếu đến '{ref_cmd_id}', nhưng chưa khai báo trong depends_on"
                    )

        # 4. Xây dựng kế hoạch thực hiện
        cmd_index = {cmd["id"]: cmd for cmd in commands}
        plan = []
        for cmd_id in execution_order:
            cmd = cmd_index[cmd_id]
            plan.append({
                "id": cmd_id,
                "command_name": cmd["command_name"],
                "depends_on": cmd.get("depends_on", []),
                "has_parameter_refs": bool(self.REFERENCE_PATTERN.search(
                    json.dumps(cmd.get("parameters", {}))
                ))
            })

        return {
            "valid": len(issues) == 0,
            "message": "Xác thực thành công, sẵn sàng thực thi" if not issues else f"Phát hiện {len(issues)} vấn đề",
            "execution_order": execution_order,
            "execution_plan": plan,
            "total_commands": len(commands),
            "issues": issues
        }

    def execute(
        self,
        commands: List[Dict],
        stop_on_error: bool = True,
        timeout_per_command: float = 10.0,
        detail_level: str = "handles",
        use_transaction: bool = False,
        confirm: bool = False
    ) -> Dict[str, Any]:
        """
        Thực hiện các lệnh hàng loạt

        Args:
            commands: danh sách lệnh, Mỗi mục chứa:
                - id: Mã định danh duy nhất của lệnh
                - command_name: API Tên lệnh
                - parameters: Tham số lệnh (Không bắt buộc)
                - depends_on: Lệnh phụ thuộcIDdanh sách (Không bắt buộc)
            stop_on_error: Có dừng thực thi khi gặp lỗi hay không
            timeout_per_command: Hết thời gian chờ cho mỗi lệnh(Thứ hai)
            detail_level: Mức độ chi tiết trong kết quả trả về
                - "summary": Chỉ trả về số liệu thống kê
                - "handles": Trả về số liệu thống kê và danh sách các thẻ điều khiển(mặc định)
                - "full": Trả về dữ liệu đầy đủ
            use_transaction: Có nên sử dụng giao dịch toàn cầu hay không
                - True: Tất cả các lệnh được thực hiện trong cùng một giao dịch, Tự động khôi phục khi thất bại
                - False: Mỗi lệnh là một giao dịch độc lập(mặc định, duy trì sự tương thích)

        Returns:
            Từ điển kết quả thực hiện
        """
        start_time = time.time()
        transaction_id = None

        # Giao dịch toàn cục: dùng backup file thật (Native COM không có transaction DB).
        # Bắt đầu = chụp backup trước khi chạy; commit/rollback ở cuối file này.

        # thiết lập lại trạng thái
        self.results = {}
        self.handle_map = {}

        # Xác minh định dạng lệnh
        validation_error = self._validate_commands(commands)
        if validation_error:
            return {
                "success": False,
                "message": f"Xác thực lệnh thất bại: {validation_error}",
                "summary": {"total": len(commands), "successful": 0, "failed": len(commands)},
                "results": [],
                "handle_map": {},
                "failed_commands": [{"id": cmd.get("id", "unknown"), "reason": validation_error} for cmd in commands]
            }

        # Xây dựng chỉ mục lệnh
        cmd_index = {cmd["id"]: cmd for cmd in commands}

        # Xây dựng biểu đồ phụ thuộc và sắp xếp theo cấu trúc liên kết
        try:
            execution_order = self._topological_sort(commands)
        except CyclicDependencyError as e:
            return {
                "success": False,
                "message": f"Phát hiện phụ thuộc vòng tròn: {str(e)}",
                "summary": {"total": len(commands), "successful": 0, "failed": len(commands)},
                "results": [],
                "handle_map": {},
                "failed_commands": [{"id": cmd["id"], "reason": "Phụ thuộc vòng tròn"} for cmd in commands]
            }

        # Tự động sao lưu bản vẽ (.bak / snapshot) trước khi thực thi hàng loạt để có đường lui an toàn
        backup_path = None
        if hasattr(self.controller, "backup_current_drawing"):
            backup_path = self.controller.backup_current_drawing()

        # Thu thập kết quả thực hiện
        successful_count = 0
        failed_count = 0
        results_list = []
        failed_commands = []

        # Thực hiện theo thứ tự tôpô
        for cmd_id in execution_order:
            cmd = cmd_index[cmd_id]
            cmd_name = cmd["command_name"]
            parameters = cmd.get("parameters", {})
            depends_on = cmd.get("depends_on", [])

            # Kiểm tra xem tất cả các phụ thuộc có thành công không
            dependency_failed = False
            for dep_id in depends_on:
                if dep_id in self.results and not self.results[dep_id].get("success", False):
                    dependency_failed = True
                    break

            if dependency_failed:
                result = {
                    "id": cmd_id,
                    "command_name": cmd_name,
                    "success": False,
                    "message": f"Lệnh phụ thuộc thất bại, bỏ qua thực thi",
                    "handle": None,
                    "data": None,
                    "execution_time_ms": 0
                }
                self.results[cmd_id] = result
                results_list.append(result)
                failed_commands.append({"id": cmd_id, "reason": "Lệnh phụ thuộc thất bại"})
                failed_count += 1

                if stop_on_error:
                    # Đánh dấu các lệnh còn lại là bị bỏ qua
                    remaining = execution_order[execution_order.index(cmd_id) + 1:]
                    for remaining_id in remaining:
                        skip_result = {
                            "id": remaining_id,
                            "command_name": cmd_index[remaining_id]["command_name"],
                            "success": False,
                            "message": "Bỏ qua do lệnh tiên quyết thất bại",
                            "handle": None,
                            "data": None,
                            "execution_time_ms": 0
                        }
                        results_list.append(skip_result)
                        failed_commands.append({"id": remaining_id, "reason": "Lệnh tiên quyết thất bại"})
                        failed_count += 1
                    break
                continue

            # Phân tích tham chiếu tham số
            try:
                resolved_params = self._resolve_references(parameters)
            except ReferenceResolutionError as e:
                result = {
                    "id": cmd_id,
                    "command_name": cmd_name,
                    "success": False,
                    "message": f"Giải mã tham chiếu tham số thất bại: {str(e)}",
                    "handle": None,
                    "data": None,
                    "execution_time_ms": 0
                }
                self.results[cmd_id] = result
                results_list.append(result)
                failed_commands.append({"id": cmd_id, "reason": str(e)})
                failed_count += 1

                if stop_on_error:
                    break
                continue

            # Thực hiện lệnh qua smart_cad_command (ưu tiên Native COM)
            cmd_start = time.time()
            try:
                cmd_confirm = cmd.get("confirm", confirm)
                if hasattr(self.controller, "smart_cad_command"):
                    api_result = self.controller.smart_cad_command(
                        cmd_name,
                        resolved_params,
                        need_result=True,
                        timeout=timeout_per_command,
                        confirm=cmd_confirm
                    )
                else:
                    api_result = self.controller.execute_api_command(
                        cmd_name,
                        resolved_params,
                        timeout=timeout_per_command
                    )
                cmd_time_ms = int((time.time() - cmd_start) * 1000)

                # chiết xuất Handle
                handle = None
                if api_result.get("success"):
                    # Cố gắng lấy nó từ nhiều vị trí có thể Handle
                    handles = api_result.get("createdEntityHandles", [])
                    if handles:
                        handle = handles[0]
                    elif api_result.get("data", {}).get("entityHandle"):
                        handle = api_result["data"]["entityHandle"]

                result = {
                    "id": cmd_id,
                    "command_name": cmd_name,
                    "success": api_result.get("success", False),
                    "message": api_result.get("message", ""),
                    "handle": handle,
                    "data": api_result.get("data"),
                    "createdEntityHandles": api_result.get("createdEntityHandles", []),
                    "execution_time_ms": cmd_time_ms
                }

                self.results[cmd_id] = result
                if handle:
                    self.handle_map[cmd_id] = handle

                results_list.append(result)

                if api_result.get("success"):
                    successful_count += 1
                else:
                    failed_count += 1
                    failed_commands.append({
                        "id": cmd_id,
                        "reason": api_result.get("message", "Thực thi không thành công")
                    })
                    if stop_on_error:
                        break

            except Exception as e:
                cmd_time_ms = int((time.time() - cmd_start) * 1000)
                result = {
                    "id": cmd_id,
                    "command_name": cmd_name,
                    "success": False,
                    "message": f"Ngoại lệ thực thi: {str(e)}",
                    "handle": None,
                    "data": None,
                    "execution_time_ms": cmd_time_ms
                }
                self.results[cmd_id] = result
                results_list.append(result)
                failed_commands.append({"id": cmd_id, "reason": str(e)})
                failed_count += 1

                if stop_on_error:
                    break

        total_time_ms = int((time.time() - start_time) * 1000)

        # Giao dịch toàn cục = backup file thật: thành công thì commit logic,
        # thất bại thì trả đường dẫn backup để người dùng khôi phục thủ công.
        # Native COM không có transaction DB nên không gọi api_transaction_* chết.
        transaction_info = None
        if use_transaction and backup_path:
            if failed_count == 0:
                transaction_info = {"status": "committed", "backup_path": backup_path}
                logger.info(f"Batch thành công, giữ backup tại: {backup_path}")
            else:
                transaction_info = {
                    "status": "rolled_back",
                    "backup_path": backup_path,
                    "hint": f"Mở lại file backup để khôi phục: {backup_path}",
                }
                logger.info(f"Batch thất bại, khôi phục từ backup: {backup_path}")

        result = self._build_result(
            results_list=results_list,
            successful_count=successful_count,
            failed_count=failed_count,
            total_count=len(commands),
            total_time_ms=total_time_ms,
            failed_commands=failed_commands,
            detail_level=detail_level
        )

        # Thêm thông tin sao lưu và giao dịch vào kết quả
        if backup_path:
            result["backup_path"] = backup_path
        if transaction_info:
            result["transaction"] = transaction_info

        return result

    def _validate_commands(self, commands: List[Dict]) -> Optional[str]:
        """
        Xác minh định dạng danh sách lệnh

        Args:
            commands: danh sách lệnh

        Returns:
            thông báo lỗi, Trả lại nếu xác minh thành công None
        """
        if not commands:
            return "Danh sách lệnh trống"

        seen_ids = set()
        for i, cmd in enumerate(commands):
            if not isinstance(cmd, dict):
                return f"Lệnh thứ {i+1} không phải là đối tượng hợp lệ"

            if "id" not in cmd:
                return f"Lệnh thứ {i+1} thiếu trường 'id'"

            if "command_name" not in cmd:
                return f"Lệnh '{cmd.get('id', i+1)}' thiếu trường 'command_name'"

            cmd_id = cmd["id"]
            if cmd_id in seen_ids:
                return f"ID lệnh '{cmd_id}' bị trùng lặp"
            seen_ids.add(cmd_id)

            # Xác minh rằng lệnh phụ thuộc vào tham chiếu tồn tại
            depends_on = cmd.get("depends_on", [])
            if not isinstance(depends_on, list):
                return f"Trường 'depends_on' của lệnh '{cmd_id}' phải là mảng"

        # Xác minh rằng tất cả các tham chiếu phụ thuộc đều tồn tại
        for cmd in commands:
            for dep_id in cmd.get("depends_on", []):
                if dep_id not in seen_ids:
                    return f"Lệnh '{cmd['id']}' phụ thuộc vào '{dep_id}' không tồn tại"

        return None

    def _topological_sort(self, commands: List[Dict]) -> List[str]:
        """
        Sắp xếp cấu trúc liên kết của các lệnh

        Args:
            commands: danh sách lệnh

        Returns:
            Lệnh sắp xếpIDdanh sách

        Raises:
            CyclicDependencyError: Bị ném khi có sự phụ thuộc vòng tròn
        """
        # Xây dựng bảng kề và bảng bậc
        graph = defaultdict(list)  # biểu đồ phụ thuộc: A -> [B, C] thể hiện B, C phụ thuộc A
        in_degree = {}

        for cmd in commands:
            cmd_id = cmd["id"]
            in_degree[cmd_id] = 0

        for cmd in commands:
            cmd_id = cmd["id"]
            for dep_id in cmd.get("depends_on", []):
                graph[dep_id].append(cmd_id)
                in_degree[cmd_id] += 1

        # Kahn's algorithm
        queue = [cmd_id for cmd_id, degree in in_degree.items() if degree == 0]
        result = []

        while queue:
            # Xử lý theo độ theo thứ tự ban đầu là0nút
            node = queue.pop(0)
            result.append(node)

            for neighbor in graph[node]:
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)

        if len(result) != len(commands):
            # Tìm các nút trong vòng lặp
            cycle_nodes = [cmd_id for cmd_id, degree in in_degree.items() if degree > 0]
            raise CyclicDependencyError(f"Các đơn đặt hàng liên quan: {cycle_nodes}")

        return result

    def _resolve_references(self, obj: Any) -> Any:
        """
        Giải quyết đệ quy các tham chiếu trong đối tượng

        Args:
            obj: đối tượng để phân tích(Có thể là từ điển, danh sách, chuỗi vv.)

        Returns:
            đối tượng được phân tích cú pháp

        Raises:
            ReferenceResolutionError: Bị ném khi độ phân giải tham chiếu không thành công
        """
        if isinstance(obj, str):
            return self._resolve_string_references(obj)
        elif isinstance(obj, dict):
            return {k: self._resolve_references(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [self._resolve_references(item) for item in obj]
        else:
            return obj

    def _resolve_string_references(self, s: str) -> Any:
        """
        Phân tích tham chiếu trong chuỗi

        Các định dạng được hỗ trợ: 
        - ${cmd_id.handle} - Nhận lệnh chủ Handle
        - ${cmd_id.data.field} - Lấy data lĩnh vực trong
        - ${cmd_id.createdEntityHandles[0]} - Lấy các phần tử mảng

        Args:
            s: một chuỗi chứa một trích dẫn

        Returns:
            giá trị được phân tích cú pháp(Nếu toàn bộ chuỗi là một tham chiếu, Trả về loại ban đầu; Ngược lại trả về chuỗi)
        """
        # Kiểm tra xem toàn bộ chuỗi có phải là tham chiếu hay không
        match = self.REFERENCE_PATTERN.fullmatch(s)
        if match:
            # Toàn bộ chuỗi được trích dẫn, Trả về loại ban đầu
            return self._get_reference_value(match.group(1))

        # Chuỗi chứa tham chiếu, Thực hiện thay thế(Kết quả là một chuỗi)
        def replace_ref(m):
            value = self._get_reference_value(m.group(1))
            return str(value) if value is not None else ""

        return self.REFERENCE_PATTERN.sub(replace_ref, s)

    def _get_reference_value(self, ref_path: str) -> Any:
        """
        Lấy giá trị tương ứng với đường dẫn tham chiếu

        Args:
            ref_path: Đường dẫn tham khảo, giống "cmd_id.data.field" hoặc "cmd_id.array[0]"

        Returns:
            giá trị trích dẫn

        Raises:
            ReferenceResolutionError: Bị ném khi tham chiếu không thể giải quyết được
        """
        parts = ref_path.split(".", 1)
        if len(parts) < 1:
            raise ReferenceResolutionError(f"Định dạng trích dẫn không hợp lệ: {ref_path}")

        cmd_id = parts[0]

        # Kiểm tra xem lệnh đã được thực thi chưa
        if cmd_id not in self.results:
            raise ReferenceResolutionError(f"Lệnh trích dẫn '{cmd_id}' Chưa được thực thi")

        result = self.results[cmd_id]

        # Kiểm tra lệnh có thành công không
        if not result.get("success"):
            raise ReferenceResolutionError(f"Lệnh trích dẫn '{cmd_id}' Thực thi không thành công")

        # Nếu chỉ có lệnhID, Trả về toàn bộ kết quả
        if len(parts) == 1:
            return result

        # Đường dẫn trường phân tích cú pháp
        field_path = parts[1]
        return self._navigate_path(result, field_path, cmd_id)

    def _navigate_path(self, obj: Any, path: str, cmd_id: str) -> Any:
        """
        Điều hướng đến một giá trị trong một đối tượng dựa trên đường dẫn

        Args:
            obj: đối tượng bắt đầu
            path: con đường, giống "data.field" hoặc "array[0].name"
            cmd_id: Đặt hàngID(cho các thông báo lỗi)

        Returns:
            Giá trị tương ứng với đường dẫn
        """
        current = obj

        # đường dẫn phân tích(ủng hộ . Tách biệt và [index] Truy cập mảng)
        # Ví dụ: "data.items[0].name" -> ["data", "items", "[0]", "name"]
        tokens = re.split(r'\.|\[', path)
        tokens = [t.rstrip(']') for t in tokens if t]

        for token in tokens:
            if current is None:
                raise ReferenceResolutionError(
                    f"Đặt hàng '{cmd_id}' con đường '{path}' Giá trị trung bình là null"
                )

            # chỉ số mảng
            if token.isdigit():
                index = int(token)
                if not isinstance(current, list):
                    raise ReferenceResolutionError(
                        f"Đặt hàng '{cmd_id}' của '{token}' không phải là một mảng"
                    )
                if index >= len(current):
                    raise ReferenceResolutionError(
                        f"Đặt hàng '{cmd_id}' chỉ số mảng của {index} Vượt qua ranh giới"
                    )
                current = current[index]
            # khóa từ điển
            elif isinstance(current, dict):
                if token not in current:
                    raise ReferenceResolutionError(
                        f"Đặt hàng '{cmd_id}' Trường này không tồn tại trong kết quả của '{token}'"
                    )
                current = current[token]
            else:
                raise ReferenceResolutionError(
                    f"Đặt hàng '{cmd_id}' của '{token}' Không phải là đường dẫn hợp lệ"
                )

        return current

    def _build_result(
        self,
        results_list: List[Dict],
        successful_count: int,
        failed_count: int,
        total_count: int,
        total_time_ms: int,
        failed_commands: List[Dict],
        detail_level: str
    ) -> Dict[str, Any]:
        """
        theo detail_level Xây dựng kết quả trả về

        Args:
            results_list: Danh sách đầy đủ các kết quả thực hiện
            successful_count: số lần thành công
            failed_count: Số lần thất bại
            total_count: tổng số lượng
            total_time_ms: Tổng thời gian sử dụng(mili giây)
            failed_commands: Danh sách các lệnh thất bại
            detail_level: Mức độ chi tiết ("summary", "handles", "full")

        Returns:
            Từ điển kết quả giảm theo mức độ chi tiết
        """
        # Kết quả cơ bản(Tất cả các chế độ bao gồm)
        base_result = {
            "success": failed_count == 0,
            "message": f"Thực thi hàng loạt hoàn tất: {successful_count}/{total_count} thành công",
            "summary": {
                "total": total_count,
                "successful": successful_count,
                "failed": failed_count,
                "total_time_ms": total_time_ms
            }
        }

        if detail_level == "summary":
            # Tinh giản nhất: Chỉ trả về số liệu thống kê và thông tin lỗi
            if failed_commands:
                base_result["failed_commands"] = failed_commands
            return base_result

        elif detail_level == "handles":
            # trung bình: Trả về danh sách các tay cầm, Không chứa đầy đủ data
            base_result["handle_map"] = self.handle_map
            base_result["created_handles"] = [
                r["handle"] for r in results_list
                if r.get("success") and r.get("handle")
            ]
            if failed_commands:
                base_result["failed_commands"] = failed_commands
            return base_result

        else:  # "full"
            # trọn: Trả về tất cả dữ liệu(hành vi ban đầu)
            base_result["results"] = results_list
            base_result["handle_map"] = self.handle_map
            base_result["failed_commands"] = failed_commands
            return base_result
