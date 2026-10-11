# Changelog

## [4.1.0-dev] - 2026-10-11 (Đang hoàn thiện)

### Thêm mới
- **Bộ tiện ích mở rộng** (`1_Scripts_TuDongHoa/vba_utilities.bas`, 332 dòng VBA thuần):
  * `ThemCongTac` & `XoaCongTac`: Thêm/xóa dòng công tác linh hoạt, tự động đánh lại STT toàn bộ bảng và kích hoạt động cơ CPM cập nhật mạng lưới tức thì.
  * `IndentWBS` & `OutdentWBS`: Tăng/giảm thụt lề phân cấp công việc trực quan trên cột C.
  * `TinhNgayTuBoQ`: Tự động đồng bộ thời lượng tính toán từ khối lượng thiết kế chia năng suất ca máy TT 38/2026 từ sheet `BOQ_TIEN_DO` sang bảng tiến độ `TIEN_DO`.
  * `InTienDo`: Tự động thiết lập trang in A3 ngang (Fit to 1 page wide) và kích hoạt Print Preview.
  * `ChuanHoaGiaoDien`: Tự động bật khung lưới ô (Gridlines), zoom 85% và tối ưu độ rộng chuẩn cho tất cả các cột.
  * `ThongTinDuAn`: Bảng tổng quan thông số gói thầu, chiều dài tuyến L=5.50 km, ngày bắt đầu/kết thúc và căn cứ pháp lý.
  * `NgayNghiLe`, `CapNhatDM`, `DoiChieuHoSo`, `TroGiup`: Điều hướng mượt mà đến các bảng tính nghiệp vụ chuyên sâu.
- **Cầu nối 2 chiều MS Project (MSPDI XML) & Primavera P6 (XER)** (`1_Scripts_TuDongHoa/vba_mspdi_xer_bridge.bas`):
  * Thuần 100% VBA, không phụ thuộc Java Runtime hay Python trên máy người dùng cuối.
  * Đã kiểm chứng xuất/nhập thực tế trên Microsoft Excel 16.0 COM (`chay_cau_noi_com.py`).
- **Thanh Ribbon hoàn chỉnh 34 nút** (`customUI14.xml`): Toàn bộ 34/34 nút đều có macro VBA tương ứng, kiểm tra tĩnh 55/55 test đạt 100%.

### Đã giải quyết & Hoàn thiện
- **Lõi CPM VBA chuẩn hóa:** `vba_cpm_core.bas` và `vba_cpm_adapter.bas` tính toán xuôi/ngược, độ trễ lag theo ngày công, dời ES sang ngày làm việc tiếp theo, khớp 100% với bộ tính Python.
- **Đối soát dữ liệu tiến độ:** Khi chạy `cpm_compare.py --rain` (xét hệ số thời tiết mùa mưa K_tt theo TCVN 8819), toàn bộ **37/37 dòng khớp chính xác 100%** (lệch: 0).
- **Đồng bộ giao diện Warm Executive (Tone Nâu Cà phê & Cát Ngà cổ điển):**
  * Tái hiện hoàn hảo phong cách thẩm mỹ executive của file tham chiếu mẫu (`#5A3718` Coffee Brown, `#F7EDE2` Warm Cream, `#EFE2D3` Sand/Ivory, `#FBF8F5` Soft Summary) trên toàn bộ 9 sheet của hệ thống.
  * Thiết kế Header Timeline 2 tầng tại hàng 4 (Năm 2023, 2024, 2025) và hàng 5 (24 Tháng / Quý: T12/23 -> T11/25), kết hợp Sub-banner tiêu đề công việc `A4:M4`.
  * Duy trì tuyệt đối dòng bắt đầu dữ liệu tại Hàng 6 (`FIRST = 6`), bảo toàn 100% tính tương thích của mạng lưới CPM, bộ kiểm thử Python và macro VBA.
  * Nâng cấp động cơ vẽ Gantt Live Canvas: bổ sung ngoặc phân cấp cho thanh Summary (`shpBL`, `shpBR`) và thẻ nhãn ngày tháng (`GTag_S`, `GTag_E`).
- **Thanh Ribbon chuẩn hóa 6 nhóm phong cách Executive:** Tổ chức mạch lạc thành 6 nhóm (`TIẾN ĐỘ CPM & GANTT`, `QUẢN TRỊ 5D`, `ĐIỀU HƯỚNG DỰ ÁN`, `CẦU NỐI TIẾN ĐỘ`, `XUẤT BẢN BÁO CÁO`, `BẢN QUYỀN HỆ THỐNG`) chứa đầy đủ 34 chức năng và bộ icon Microsoft Office chuẩn.
- **Quy chuẩn lịch thi công:** Ô F3 trên sheet `TIEN_DO` hiển thị chuẩn "6 ngày/tuần (Nghỉ CN & Lễ)", ô H3 hiển thị "Tự động (Forward/Backward Pass)".
- **Chuẩn hóa chuỗi VBA:** Chuyển toàn bộ chuỗi tiếng Việt trong các hộp thoại sang ASCII không dấu, ngăn ngừa hoàn toàn lỗi vỡ font mã trang trên Windows.

### Việc còn mở
- Quyết định về dự trữ tự do (Float) lớn của Phân đoạn A1 (chờ quyết định phân đoạn công trường của kỹ sư phụ trách).
- Chọn giấy phép mã nguồn mở (MIT / Proprietary); ký số chứng chỉ (code signing certificate) cho macro VBA nếu phát hành thương mại.

## [4.0] - trước 2026-10-10
- Phiên bản Doanh nghiệp v4.0 (các script build, `.xlsm` và `.xlam`).
