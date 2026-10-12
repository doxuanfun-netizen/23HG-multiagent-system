# CLAUDE.md — quy ước làm việc với repo 23HG-multiagent-system

Hệ thống AEC chạy bằng Python tất định (không có LLM trong lõi): đo bóc, cắt thép, dự toán G_XD,
thanh toán 03a, tiến độ CPM, KCS, As-Built. Giữ file này dưới 200 dòng (có test kiểm tra).

## Lệnh thường dùng

- Cài: `pip install -r requirements.txt`. Test apps cần thêm LibreOffice Calc (`libreoffice-calc`);
  thiếu Calc thì các test đó tự bỏ qua.
- Test gốc (như CI): `python -m unittest discover -s tests -t . -q`
- Tự kiểm tra solver + CPM (như CI): `python run_state_graph.py --solver-test`
- Test apps: `cd apps/23hg_schedule_assistant_pro && python -m unittest discover -s tests -p "test_*.py"`
- Khung đối chiếu hồ sơ duyệt: `python -m tests.golden.harness.runner` (còn FAIL có lý do, xem bên dưới)
- Quét lỗi công thức Excel: `python -m tools.audit_excels_static <thư mục>`
- Kiểm tra trước khi push (một lần cho mỗi bản clone): `git config core.hooksPath .githooks`

## Nguyên tắc số liệu (quan trọng nhất)

- "Đúng" nghĩa là khớp hồ sơ ĐÃ DUYỆT, không chỉ tự nhất quán. Dùng `tests/golden/harness/`:
  - `approved_dossier` mới là bằng chứng; `hand_formula` chỉ kiểm tra công thức.
  - Ca `pending` không bao giờ được tính là đạt. Không nới dung sai để một ca đạt.
- Trạng thái phải được TÍNH, không gõ tay: không ghi "Hợp lệ", "OPTIMAL", "HOÀN TẤT", "100%" nếu chưa đo.
- Không bỏ dữ liệu im lặng: dòng bị loại, thanh chưa lập kế hoạch cắt, ô chưa tính được... phải hiện ra
  trong kết quả (log + file xuất), kèm lý do.
- Không đoán số OCR mờ. Dữ liệu trích từ bản vẽ phải ghi trang (`[Trang N]`) và số hiệu bản vẽ.
- Tiền: dùng `tools/money.py` (Decimal, làm tròn half-up như Excel).
- File Excel trong repo thường KHÔNG có giá trị công thức lưu sẵn. Muốn đọc giá trị: tính lại bản sao
  bằng LibreOffice (`soffice --headless --convert-to xlsx`) hoặc `tools/excel_eval.py`; không sửa file gốc.

## Quyết định của chủ dự án (11/10/2026) và điểm còn mở

- Hồ sơ Km19 dùng kg/m theo bảng của solver cắt thép (`rebar_unit_weight_kg_m`, vd. Ø14 = 1,208).
  `tools/steel_qs.py` (D²/162) giữ cho mẫu QS khác — đừng trộn hai quy ước trong một hồ sơ.
- Thanh dài hơn 11,7 m trong BBS Km19 đã gồm nối chồng 40d: cắt số đoạn tối thiểu, tổng đoạn = chiều dài
  BBS (`noi_trong_chieu_dai=True` trong `examples/generate_rebarcut_dedicated_package.py`).
- BBS, dự toán, 03a sửa theo bản vẽ có dấu thẩm định; sai khác thì đánh dấu NGHI VẤN (cột + sheet
  `NGHI_VAN`, cột "Đã soát / Ghim") để chủ dự án soát. Không tự "chốt" thay chủ dự án.
- Tr.40 Mố M2 bản vẽ ghi 0 kg cho W6/W7/W9: chủ dự án quyết định TÍNH (237,05 kg) — `tinh_dong_ban_ve_ghi_0`.
- Bê tông cọc khoan nhồi: thể tích thiết kế × hệ số nở thành 1,12 (chủ dự án quyết định giữ).
- Còn mở: tỷ lệ hoàn thành 03a chưa có biên bản; hạng mục bản vẽ có mà dự toán chưa có (sheet `NGHI_VAN`).
- Pháp lý: chưa văn bản nào được đối chiếu bản gốc — xem `workflows/LEGAL_STATUS.md`.

## Riêng tư và vệ sinh repo

- Không đưa tên đơn vị/người bên thứ ba, đường dẫn máy cá nhân (`C:\Users\...`, `D:\...`) vào file
  đã track. `tests/test_repo_hygiene.py` có kiểm tra.
- File đã track ≤ 1 MB (ngoại lệ ghi trong `ALLOWED_LARGE*` của test vệ sinh); đường dẫn ≤ 190 ký tự
  (giới hạn Windows); các bản cùng tên trong một dự án `examples/<dự án>/` phải giống hệt nhau.
- Không commit nguyên file mẫu của bên thứ ba; chỉ trích số liệu kỹ thuật đã ẩn danh.

## Bố cục

- `core/` — agents, supervisor, state bus, quality gate. `tools/` — các bộ tính (thép, QS, CPM, KCS...).
- `examples/` — hồ sơ mẫu và script sinh hồ sơ; `examples/data/` — dữ liệu đầu vào dạng CSV/JSON.
- `examples/HO_SO_CAU_KM19_529/05_DU_LIEU_GOC_SCAN_MARKER/` — HSTK đã OCR làm đầu vào. File lớn:
  đừng đọc nguyên văn; dùng `du_lieu.json` theo trang, `chia_doan.jsonl` theo đoạn, `kiem_tra_so_lieu.md`
  cho mục nghi ngờ.
- `apps/23hg_schedule_assistant_pro/` — phân hệ Excel/VBA có bộ test riêng.
- `apps/23hg_cad_automation_mcp_pro/` — phân hệ điều khiển AutoCAD qua MCP Native COM (139 API, QS Takeoff, 74 test).
- `tests/golden/` — số liệu chuẩn và khung đối chiếu. `workflows/` — quy trình, trạng thái pháp lý.

## Git

- Làm trên nhánh tính năng; chỉ mở PR khi được yêu cầu. Không viết lại lịch sử của `main`.
- Commit message tiếng Việt, nêu cả việc CHƯA làm. Chạy hai bộ test trước khi push (hook làm việc này).
- Tài liệu không nói quá: số liệu, thời gian chạy, số test phải là số đo được.
