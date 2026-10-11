# Lộ trình repo 23HG Schedule Assistant Pro

Cập nhật: 11/10/2026. Mốc thời gian là **ước lượng** theo tuần làm việc, phụ thuộc vào việc kiểm thử trên Windows và các quyết định ở mục 4.

## 1. Nguyên tắc

1. **CPM nằm trong công thức Excel, VBA chỉ lo giao diện.** VBA hiện tại dùng `CreateObject("Scripting.Dictionary")` (chỉ có trên Windows) nên không thể kiểm thử tự động. Công thức Excel thì kiểm thử được bằng LibreOffice và đối chiếu với bộ tính Python. VBA giữ phần vẽ Gantt, ribbon, xuất bản.
2. **Không đổi lịch dự án khi chưa có quyết định của kỹ sư.** Mọi ngày chỉnh tay phải được ghi lại bằng ràng buộc kèm lý do, không để lẫn trong dữ liệu.
3. **Mỗi thay đổi có test.** Phần nào chỉ kiểm thử được trên Windows thì ghi rõ trong `KIEM_THU_TREN_WINDOWS.md`, không tuyên bố đã chạy.
4. **`main` luôn xanh.** Mọi thay đổi đi qua nhánh và pull request; commit theo Conventional Commits.

## 2. Hiện trạng

| Hạng mục | Trạng thái |
|---|---|
| Bộ tính CPM chuẩn (Python) | Xong, có test |
| Bảng CPM công thức Excel | Xong ở dạng bản nháp, khớp bộ tính 37/37 dòng; chưa vào file Master |
| Bộ kiểm thử | 49 test tự động, chạy bằng `python -m unittest discover -s tests -v` |
| CI GitHub Actions | Đã viết, chưa chạy (chưa có repo trên GitHub) |
| Ribbon mới, installer mới | Viết xong, **chưa chạy thử trên Excel/Windows** |
| Ngày trong sheet `TIEN_DO` | **31/37 dòng lệch** so với tính từ tiền nhiệm; cần quyết định |
| Chức năng bản cũ còn thiếu | Danh sách trong CHANGELOG |

## 3. Các mốc

### M0. Đưa lên GitHub và bật CI (tuần 1, 12–18/10)
- Bạn tạo repo rỗng (đề xuất tên `23hg-schedule-assistant`), thêm vào phiên làm việc; tôi đẩy 5 commit hiện có.
- CI chạy trên Ubuntu với LibreOffice. **Xong khi:** CI xanh trên `main`.

### M1. Kiểm thử Windows (tuần 1, song song M0)
- Bạn làm theo `KIEM_THU_TREN_WINDOWS.md` (mục A–D) và gửi bảng kết quả.
- Mỗi lỗi thành một issue. Thêm runner `windows-latest` vào CI cho phần test tĩnh nếu cần.
- **Xong khi:** installer, ribbon và 9 nhóm nút đã được xác nhận hoặc có issue tương ứng.

### M2. Chốt lịch dự án (tuần 2, 19–25/10)
- Kỹ sư quyết định: lấy ngày đang lưu hay ngày tính từ logic tiền nhiệm (kết thúc 02/10/2025 so với 29/10/2025).
- Rà lại tiền nhiệm Phân đoạn A1, vì dự trữ 141–467 ngày công cho thấy logic hiện tại chưa đúng ý đồ thi công.
- Thêm cột "ràng buộc sớm nhất" kèm "lý do" cho các ngày chỉnh tay (mùa mưa, giao mặt bằng...).
- **Xong khi:** có một bộ ngày được kỹ sư xác nhận, ghi lại lý do từng ràng buộc.

### M3. CPM bằng công thức vào file Master (tuần 3–4, 26/10–8/11) — cần Excel
- Đưa bảng công thức vào `rebuild_perfect_master_v3_professional.py`; nút **Cập Nhật Tiến Độ** chỉ tính lại và vẽ Gantt, bỏ `Scripting.Dictionary`.
- Bổ sung hệ số mưa K_tt trong công thức; cân nhắc FF/SF.
- **Xong khi:** `cpm_compare.py` báo 0 dòng lệch trên file Master sau khi bấm nút.

### M4. Khôi phục chức năng bản cũ (tuần 5, 9–15/11) — cần Excel
- Ưu tiên theo mức dùng thực tế: Thêm/Xóa việc, Indent/Outdent, Cập Nhật Định Mức, Gắn Mã Máy, Tính Ngày Từ BoQ, Đối Chiếu Hồ Sơ.
- Mỗi chức năng một pull request, có kịch bản thử tay trong checklist. Chú ý trùng tên hàm với bản hiện tại (lỗi "Ambiguous name").
- **Xong khi:** mỗi nút trong ribbon có macro thật và đã qua checklist.

### M5. Phát hành 4.1 (tuần 6, 16–22/11)
- Chọn giấy phép; ký số macro hoặc hướng dẫn Trusted Location cho người dùng.
- Đối chiếu căn cứ pháp lý trong README với nguồn chính thức.
- Gắn tag, đính kèm `.xlam`, `.xlsm`, bản PDF mẫu vào GitHub Release.
- **Xong khi:** một máy Windows sạch cài và chạy hết checklist mà không cần hỗ trợ.

### M6. Sau phát hành (ý tưởng, chưa cam kết)
- Nhiều lịch làm việc (theo tổ đội), nhập/xuất MS Project hoặc Primavera, cảnh báo trễ so với baseline, bảng điều khiển EVM.

## 4. Quyết định đang chờ

| # | Câu hỏi | Ai quyết | Cần trước |
|---|---|---|---|
| 1 | Ngày nào là đúng: đã lưu hay tính từ logic? | Kỹ sư dự án | M2 |
| 2 | Tên và chủ sở hữu repo GitHub | Bạn | M0 |
| 3 | Giấy phép mã nguồn | Bạn | M5 |
| 4 | Có mua/dùng chứng chỉ ký số macro không? | Bạn | M5 |
| 5 | Chức năng bản cũ nào bắt buộc phải khôi phục | Bạn | M4 |

## 5. Quan hệ với repo `rea`

Tôi đã đọc tài liệu của `rea` (`baotuhg/rea`, gói `rea-agents` 6.3.0): đây là công cụ dịch ngược cho file nhị phân native, ELF, EVM, JavaScript/Electron, website, HAR, .NET và APK. **Tài liệu không nhắc đến VBA, xlsm hay Office**, nên `rea` không phân tích được add-in Excel của dự án này.

| Hướng | Đánh giá |
|---|---|
| Gộp 23HG vào `rea` | Không nên. Khác miền nghiệp vụ, ngôn ngữ và quy tắc đóng góp. |
| Dùng `rea` để phân tích file `.xlam` | Không phù hợp. |
| Dùng `rea` cho ứng dụng Electron của bạn | Phù hợp: `rea` có phân tích Electron/JavaScript. Đây là một nhánh việc riêng, không thuộc repo 23HG. |
| Đóng góp cho `rea` | Làm theo `AGENTS.md` của repo: nhánh riêng, Conventional Commits, `npm run check:fast`. Chỉ làm khi bạn chỉ rõ một issue cụ thể. |

Hai repo giữ độc lập; chia sẻ quy trình (CI, commit, kiểm thử) chứ không chia sẻ mã.
