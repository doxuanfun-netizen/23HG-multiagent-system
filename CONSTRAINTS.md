# CONSTRAINTS.md — Khế Ước Chất Lượng Cơ Học (23HG-multiagent-system)

Mọi thay đổi trong repository này phải vượt qua các chốt chặn cơ học dưới đây trước khi được coi là hoàn thành.

---

## 1. Lệnh Kiểm Tra Cơ Học Bắt Buộc (Mechanical Verification Commands)

| Chốt Chặn (Gate) | Lệnh Kiểm Tra (Chạy tại thư mục gốc repo) | Ngưỡng Đạt (Pass Bar) |
| :--- | :--- | :--- |
| **1. Toàn bộ Unit Tests** | `python -m unittest discover -s tests -t . -q` | **PASS 100%** (0 failures, 0 errors) |
| **2. Tự kiểm tra Solver & CPM** | `python run_state_graph.py --solver-test` | **Exit code 0** |
| **3. Vệ sinh Repo & Đường dẫn** | `python -m unittest tests/test_repo_hygiene.py tests/test_examples_consistency.py -v` | **0 đường dẫn cá nhân**, không file > 1 MB trái phép, đường dẫn <= 190 ký tự |
| **4. Quy ước CLAUDE.md & Pháp lý** | `python -m unittest tests/test_project_conventions.py tests/test_legal_status.py -v` | `CLAUDE.md` <= 200 dòng, mọi đường dẫn tồn tại, bảng pháp lý nguyên vẹn |

---

## 2. Ràng Buộc Số Liệu & Nghiệp Vụ (Từ `CLAUDE.md`)
- **Không nới dung sai để ép đạt:** Tuyệt đối không sửa ngưỡng dung sai trong `tests/golden/` để biến một ca FAIL thành PASS.
- **Không trộn quy ước đơn trọng thép:** Hồ sơ dùng bảng đơn trọng của solver (`rebar_unit_weight_kg_m`) và mẫu QS dùng $D^2/162$ (`tools/steel_qs.py`) phải tách bạch theo đúng quyết định của chủ dự án trong `CLAUDE.md`.
- **Không bỏ dữ liệu im lặng:** Mọi dòng bị loại hoặc chưa tính được phải xuất hiện tường minh trong log và sheet `NGHI_VAN`.
- **Tiền tệ:** Bắt buộc dùng `tools/money.py` (`Decimal`, làm tròn `ROUND_HALF_UP`).

---

## 3. Ràng Buộc Tự Động Hóa COM / Excel
- **Cấm Modal Dialog ngầm:** Không gọi `MsgBox` / `InputBox` khi điều khiển Excel COM / AutoCAD MCP ở chế độ nền (`Visible=False`).
- **Bảo toàn Drawing XML:** Không dùng `openpyxl` ghi đè trực tiếp lên file `.xlsm` có chứa VBA Shape/Drawing phức tạp; sử dụng Excel COM Range hoặc bản sao tính toán qua `tools/excel_eval.py`.
