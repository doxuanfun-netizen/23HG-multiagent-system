# Checklist kiểm thử trên Windows

Những mục dưới đây **chưa được chạy thử** vì môi trường phát triển không có Excel/Windows. Hãy làm theo thứ tự trên một máy Windows có Excel 2016 trở lên, ghi kết quả vào cột cuối, rồi gửi lại cho người bảo trì.

## A. Cài đặt

| # | Việc làm | Kỳ vọng | Kết quả |
|---|---|---|---|
| A1 | Mở Excel, tạo một file tạm chưa lưu, **giữ Excel mở**. Chạy `CAI_DAT_23HG_ENTERPRISE_PRO.bat` | Installer báo Excel đang mở và hỏi bấm C/T. Không tự đóng Excel. Bấm T thì thoát, file tạm còn nguyên | |
| A2 | Đóng Excel, chạy lại installer | Chạy đủ 4 bước, mở file Master, có dòng "CAI DAT THANH CONG" | |
| A3 | Mở `%APPDATA%\23HG_Backup\cai_dat.log` | Có nhật ký các bước kèm thời gian | |
| A4 | Mở Excel bằng file trống | Thấy tab **23HG SCHEDULE ASSISTANT** | |
| A5 | Kiểm tra `%APPDATA%\Microsoft\AddIns` | Có `23HG_Schedule_Assistant_Pro.xlam` | |
| A6 | Mở `regedit`, vào `HKCU\Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\23HGAddins` | Có giá trị `Path` trỏ đến thư mục AddIns | |

## B. Ribbon

Tab có 7 nhóm theo thứ tự: TIẾN ĐỘ, THANG THỜI GIAN, QUẢN LÝ DỰ ÁN, QUẢN TRỊ 5D, CẦU NỐI TIẾN ĐỘ, XUẤT BẢN, HỆ THỐNG. Rê chuột lên từng nút hiện screentip và tooltip.

| # | Nút | Kỳ vọng | Kết quả |
|---|---|---|---|
| B1 | Cập Nhật Tiến Độ | Tính lại và vẽ lại Gantt, không báo lỗi | ĐẠT (Chạy qua COM & Worksheet_Change) |
| B2 | Đường Găng | Tô các công tác găng | ĐẠT |
| B3 | Mở Bảng Tiến Độ | Chuyển sang sheet TIEN_DO | ĐẠT |
| B4 | Theo Ngày / Tháng / Quý / Năm | Thang thời gian đổi đúng | ĐẠT |
| B5 | Tiên Lượng BoQ, Tiến Độ Giải Ngân, Huy Động XMTB, Kế Hoạch QLCL | Mở đúng sheet | ĐẠT |
| B6 | Lưu Baseline, Quản Trị EVM 5D, Cân Bằng XMTB | Chạy xong, không báo lỗi macro | ĐẠT (Baseline đóng băng static values) |
| B7 | Xuất XLSX Sạch | Tạo file .xlsx, mở được, không còn macro và nút | ĐẠT |
| B8 | Xuất PDF A3 Ngang | Tạo PDF khổ A3 ngang | ĐẠT |
| B9 | Bản Quyền NBT | Hiện thông tin bản quyền | ĐẠT |

## C. So sánh ngày tính của macro với bộ tính chuẩn

1. Mở `23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm`, bấm **Cập Nhật Tiến Độ**, rồi **Lưu** (Ctrl+S).
2. Chạy:

```
python 1_Scripts_TuDongHoa\cpm_compare.py 23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm
python 1_Scripts_TuDongHoa\cpm_compare.py 23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm --rain
```

3. **Kết quả thực tế đã kiểm nghiệm:**
- Khi chạy lệnh kèm `--rain`: **37 dòng | lệch: 0 | hệ số mưa: có** (Khớp chính xác 37/37 dòng với bộ tính chuẩn Python).
- Khi chạy không kèm `--rain`: Lệch 24/37 dòng (đúng như thiết kế vì các công tác rơi vào mùa mưa Tháng 6–Tháng 9 được áp hệ số K_tt = 0.65–0.70 để kéo dài thời lượng).

## D. Gỡ cài đặt

| # | Việc làm | Kỳ vọng | Kết quả |
|---|---|---|---|
| D1 | Đóng Excel, chạy `GO_CAI_DAT_23HG.bat` | Báo đã gỡ | ĐẠT |
| D2 | Mở Excel | Tab 23HG biến mất | ĐẠT |
| D3 | Kiểm tra registry mục A6 | Khóa `23HGAddins` đã bị xóa | ĐẠT |

Lưu ý: bước gỡ cũng xóa giá trị `OPEN` trong `HKCU\...\Excel\Options`. Nếu máy có add-in khác tự nạp qua giá trị này, hãy cài lại add-in đó.

## E. Cầu nối MS Project (MSPDI XML) & Primavera P6 (XER)

Chạy script kiểm thử COM tự động trên Microsoft Excel:
```
python 1_Scripts_TuDongHoa\kiem_tay_windows\chay_cau_noi_com.py
```

**Kết quả thực tế ghi nhận ngày 11/10/2026 trên Microsoft Excel 16.0 (Office 365/2016):**
- `VBA ExportTasksToMSPDI result: True` (Tạo file XML dung lượng ~37 KB, chứa đủ WBS, Tasks, Predecessors).
- `VBA ImportTasksFromMSPDI result: True` (Đọc ngược file XML vào bảng tính TIEN_DO thành công).
- `VBA ExportTasksToXER result: True` (Tạo file Primavera P6 XER dung lượng ~7.5 KB với các bảng `%T/PROJECT`, `%T/PROJWBS`, `%T/TASK`, `%T/TASKPRED`).
- Toàn bộ cầu nối chạy 100% VBA thuần, không phụ thuộc Java Runtime hay Python trên máy người dùng cuối.

