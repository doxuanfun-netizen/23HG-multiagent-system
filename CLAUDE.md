# CLAUDE.md — quy ước làm việc với repo 23HG-multiagent-system

Hệ thống AEC chạy bằng Python tất định (không có LLM trong lõi): đo bóc, cắt thép, dự toán G_XD,
thanh toán 03a, tiến độ CPM, KCS, As-Built. Giữ file này dưới 200 dòng (có test kiểm tra).

## Lệnh thường dùng

- Cài: `pip install -r requirements.txt`. Test apps cần thêm LibreOffice Calc (`libreoffice-calc`);
  thiếu Calc thì các test đó tự bỏ qua.
- Test gốc (như CI): `python -m unittest discover -s tests -t . -q`
- Tự kiểm tra solver + CPM (như CI): `python run_state_graph.py --solver-test`
- Test apps: `cd apps/23hg_schedule_assistant_pro && python -m unittest discover -s tests -p "test_*.py"`
- Khung đối chiếu hồ sơ duyệt: `python tests/golden/harness/runner.py` (đang FAIL có chủ đích, xem bên dưới)
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

## Điểm đang mở (chờ chủ dự án quyết định — đừng tự chọn)

- Ba quy ước khối lượng đơn vị thép đang cùng tồn tại: `tools/steel_qs.py` dùng D²/162,
  `tools/cutting_stock_solver.py` tra bảng π·d²/4·ρ (vd. Ø14 = 1,208), BBS Km19 dùng bảng làm tròn
  (Ø14 = 1,21). Harness đang báo lệch ở ca `bbs_don_vi_khoi_luong_thep`.
- Thanh dài hơn cây 11,7 m cần vùng cho phép nối; chưa có thì solver để vào `unplanned` và bảng cắt
  thép báo "CHƯA ĐỦ". Không tự đặt vùng nối.
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
- `tests/golden/` — số liệu chuẩn và khung đối chiếu. `workflows/` — quy trình, trạng thái pháp lý.

## Git

- Làm trên nhánh tính năng; chỉ mở PR khi được yêu cầu. Không viết lại lịch sử của `main`.
- Commit message tiếng Việt, nêu cả việc CHƯA làm. Chạy hai bộ test trước khi push (hook làm việc này).
- Tài liệu không nói quá: số liệu, thời gian chạy, số test phải là số đo được.
