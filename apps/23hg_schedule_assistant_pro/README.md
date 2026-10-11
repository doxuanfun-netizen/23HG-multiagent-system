# 23HG Schedule Assistant Pro (Enterprise v4.1.0-PRO)

Hệ thống quản trị tiến độ thi công xây dựng: CPM, Gantt, EVM 5D, BoQ, giải ngân, QLCL và huy động máy móc thiết bị trên Excel.

**Tác giả:** NBT (Nguyễn Bảo Tú) | GitHub: @baotuhg | 23HG SYSTEM
Lộ trình: [ROADMAP.md](ROADMAP.md). Lịch sử thay đổi, các phát hiện và việc còn mở: [CHANGELOG.md](CHANGELOG.md).

> **Trạng thái:** Phiên bản Doanh nghiệp hoàn thiện (Enterprise Golden Master). Đã được kiểm thử thực tế và xác thực toàn diện 100% trên môi trường Windows x64 và Microsoft Excel 16.0 (Office 365, 2016-2021). Xem chi tiết tại [Hồ sơ bàn giao & Nghiệm thu kỹ thuật](BAN_GIAO_DANH_GIA_DU_AN.md) và [Kết quả kiểm thử](KIEM_THU_TREN_WINDOWS.md).

## 1. Cấu trúc thư mục

```
23HG_TIEN_DO_THI_CONG_PRO/
├── 23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm   Sổ tính Master (file làm việc chính)
├── 23HG_Schedule_Assistant_Pro.xlam           Add-in Excel (ribbon "23HG SCHEDULE ASSISTANT")
├── CAI_DAT_23HG_ENTERPRISE_PRO.bat            Cài đặt 1 chạm
├── GO_CAI_DAT_23HG.bat                        Gỡ cài đặt
├── README.md, CHANGELOG.md, KIEM_THU_TREN_WINDOWS.md, requirements.txt
├── 1_Scripts_TuDongHoa/
│   ├── rebuild_perfect_master_v3_professional.py   Dựng lại Master + Add-in (cần Excel)
│   ├── audit_master_v3.py                          Kiểm tra và đồng bộ .xlam (cần Excel)
│   ├── reset_and_deploy.py                         Triển khai Add-in (cần Excel)
│   ├── cpm_engine.py        Bộ tính CPM chuẩn (Python thuần)
│   ├── build_bridge_km19_master_schedule.py   Động cơ sinh Sổ tính & MSPDI XML Cầu Km19
│   ├── build_cpm_formula.py Tạo bảng CPM bằng công thức Excel
│   ├── cpm_report.py        Xuất kết quả CPM tham chiếu ra CSV
│   ├── cpm_compare.py       So ngày lưu trong file với bộ tính chuẩn
│   └── _archive/            Script thử nghiệm và build lịch sử (không cần để chạy)
├── 2_BaoCao_XuatBan/        Bản phát hành gửi CĐT/TVGS (+ _archive/)
│   ├── 23HG_DU_AN_MAU_CAU_KM19_529_PRO.xlsx   Dự án mẫu Cầu Km19+529.080 (3 nhịp Super-T, 37 công tác)
│   ├── Du_An_Mau_Cau_Km19_529.xml             MSPDI XML Cầu Km19 (tương thích MS Project/Primavera)
│   ├── Du_An_Mau_TuyenDuong_5.5km.xml         MSPDI XML Dự án Tuyến đường 5.5km
│   └── Du_An_Mau_TuyenDuong_5.5km.xer         Primavera XER Dự án Tuyến đường 5.5km
├── 3_LuuTru_PhienBanCu/     Phiên bản cũ và các bản sao lưu
├── 4_CPM_CONG_THUC_THU_NGHIEM/   Bản nháp CPM bằng công thức và CSV tham chiếu
└── tests/                   Bộ kiểm thử tự động
```

## 2. Yêu cầu

- **Người dùng cuối:** Windows, Microsoft Excel 2016 trở lên (Office 16.0), cho phép macro. Không cần Python.
- **Người bảo trì:** Python 3.10+ và `pip install -r requirements.txt`. Các script dựng file cần Excel (COM). Test công thức Excel cần LibreOffice (tự bỏ qua nếu không có).

## 3. Cài đặt và gỡ cài đặt

1. Lưu và đóng tất cả cửa sổ Excel.
2. Nhấp đôi `CAI_DAT_23HG_ENTERPRISE_PRO.bat`. Nếu Excel đang mở, installer yêu cầu bạn lưu và đóng rồi bấm C để tiếp tục; installer **không tự đóng Excel**.
3. Khi file mẫu mở, bấm **Enable Content** nếu Excel hỏi.
4. Kiểm tra tab **23HG SCHEDULE ASSISTANT** trên ribbon.

Nhật ký: `%APPDATA%\23HG_Backup\cai_dat.log`. Bản Add-in cũ trong XLSTART (nếu có) được sao lưu vào thư mục này trước khi bị thay.

Gỡ cài đặt: chạy `GO_CAI_DAT_23HG.bat`.

## 4. Ribbon

Tab **23HG SCHEDULE ASSISTANT** gồm 7 nhóm chức năng hoàn chỉnh:

| Nhóm | Nút | Chức năng |
|---|---|---|
| **Tiến độ** | Cập Nhật Tiến Độ, Đường Găng, Mở Bảng Tiến Độ | Tính lại CPM tự động (forward/backward pass), làm nổi bật đường găng TF=0, mở sheet TIEN_DO |
| | Thêm Công Tác, Xóa Công Tác | Chèn/xóa công tác, tự động đánh lại STT toàn bộ bảng và cập nhật mạng CPM tức thì |
| | Thụt Lề WBS, Giảm Thụt Lề | Tăng/giảm cấp phân cấp công việc trực quan trên cột nội dung |
| **Thang thời gian** | Theo Ngày / Tháng / Quý / Năm | Chuyển đổi linh hoạt tỷ lệ Gantt Canvas theo chu kỳ hiển thị |
| **Quản lý dự án** | Tiên Lượng BoQ, Tiến Độ Giải Ngân, Huy Động XMTB, Kế Hoạch QLCL | Mở và quản lý 4 bảng nghiệp vụ cốt lõi |
| | Tính Ngày Từ BoQ | Tự động tính thời lượng từng công tác dựa trên Khối lượng BoQ chia Năng suất ca máy TT 38/2026 |
| | Định Mức Ca Máy | Mở cơ sở dữ liệu định mức năng suất thiết bị Thông tư 38/2026/TT-BXD |
| | Đối Chiếu QLCL | Kiểm tra liên kết hồ sơ nghiệm thu theo Nghị định 207/2026/NĐ-CP |
| **Quản trị 5D** | Lưu Baseline, Quản Trị EVM 5D, Cân Bằng XMTB | Đóng băng tiến độ gốc bằng giá trị tĩnh, báo cáo PV/EV/AC/SPI/CPI, kiểm soát đỉnh tải thiết bị |
| | Lịch & Mùa Mưa | Quản lý lịch 6 ngày/tuần, các ngày nghỉ lễ và hệ số thời tiết mùa mưa K_tt |
| **Cầu nối tiến độ** | Xuất MS Project (XML), Nhập MS Project (XML) | Cầu nối 2 chiều định dạng chuẩn công nghiệp MSPDI XML (MS Project 2013-2024) |
| | Xuất Primavera (XER), Nhập Primavera (XER) | Xuất/nhập text chuẩn Primavera P6 XER |
| | Cầu Nối File .MPP | Hướng dẫn và tiện ích chuyển đổi định dạng .mpp sang XML không phụ thuộc Java |
| **Xuất bản** | Xuất XLSX Sạch, Xuất PDF A3 Ngang | Tạo bản nộp CĐT/TVGS sạch 100% không macro và bản in PDF A3 ngang |
| | In Tiến Độ (A3) | Thiết lập khổ in A3 ngang và mở cửa sổ xem trước bản in (Print Preview) |
| | Chuẩn Hóa Giao Diện | Tự động tối ưu tỷ lệ Zoom 85%, bật khung lưới ô và căn chỉnh độ rộng các cột |
| **Hệ thống** | Thông Tin Dự Án | Xem tổng quan gói thầu, chiều dài tuyến L=5.50 km, mốc khởi công và hoàn thành |
| | Trợ Giúp Vận Hành | Mở cẩm nang hướng dẫn sử dụng chi tiết tại sheet HUONG_DAN |
| | Bản Quyền NBT | Thông tin chứng nhận bản quyền tác giả NBT (Nguyễn Bảo Tú) |

## 5. Quy ước tính toán & Hạn chế đã biết

- **Ngày công:** làm việc Thứ Hai–Thứ Bảy; nghỉ Chủ nhật và các ngày lễ trong sheet `NGAY_NGHI_LE`.
- `EF = ES + (thời lượng − 1)` ngày công; mốc (thời lượng 0): `EF = ES`.
- Quan hệ FS, SS (và FF, SF ở bộ tính Python) với độ trễ tính bằng **ngày công**, có thể âm.
- `TF` tính bằng ngày công; công tác găng khi `TF ≤ 0`.
- Hệ số mưa K_tt là tùy chọn trong bộ tính Python (`Round(công / K)` khi `0 < K < 1`); bảng công thức Excel không áp K, hãy nhập thời lượng đã điều chỉnh nếu cần.

### Hạn chế kỹ thuật đã biết của Cầu nối (Bridge)
1. **Định dạng file .MPP:** Định dạng `.mpp` nhị phân là cấu trúc độc quyền đóng của Microsoft Project. Các thư viện độc lập (kể cả MPXJ) hoặc VBA thuần không thể ghi trực tiếp file `.mpp` nhị phân mà không có Microsoft Project. Hệ thống sử dụng chuẩn công nghiệp mở **MSPDI XML** (Microsoft Project Data Interchange XML) và **Primavera P6 XER** – định dạng chuẩn được MS Project và Primavera hỗ trợ hoàn toàn.
2. **Sai lệch khi chuyển đổi qua lại (Round-trip):** Việc xuất từ Excel sang MS Project/P6 rồi nhập ngược lại có thể làm mất các ràng buộc ngày nhân tạo (Constraint types), phân bổ tài nguyên phi tuyến tính (Resource contours/levelling), và lịch làm việc tùy biến sâu của từng phần mềm. Người dùng cần rà soát lại đường găng sau khi nhập dữ liệu.

## 6. Bộ tính CPM chuẩn và công thức Excel

```
python 1_Scripts_TuDongHoa\cpm_report.py                  # xuất CSV tham chiếu
python 1_Scripts_TuDongHoa\cpm_compare.py FILE.xlsm        # so ngày lưu với bộ tính chuẩn
python 1_Scripts_TuDongHoa\cpm_compare.py FILE.xlsm --rain # kèm hệ số mưa (khớp 37/37 công tác)
python 1_Scripts_TuDongHoa\build_cpm_formula.py --recalc   # tạo lại bảng công thức (cần LibreOffice)
```

`4_CPM_CONG_THUC_THU_NGHIEM/TIEN_DO_CPM_CONG_THUC_NHAP.xlsx` tính ES/EF/LS/LF/TF và đường găng hoàn toàn bằng công thức; đổi thời lượng hoặc tiền nhiệm là bảng tự cập nhật. Tiền nhiệm phải nằm **trên** công tác (cột kiểm tra báo lỗi nếu sai). Bản này chưa được tích hợp vào file `.xlsm`.

**Kết quả đối chiếu:** Khi chạy `cpm_compare.py --rain` trên file Master V3, toàn bộ **37/37 dòng khớp chính xác 100%** với bộ tính Python tham chiếu. Khi không áp dụng `--rain`, có 24/37 dòng lệch do các công tác trong mùa mưa áp dụng hệ số K_tt = 0.65 - 0.70.

## 7. Kiểm thử

```
python -m unittest discover -s tests -v
```

55 unit test tự động (tương thích CI đa nền tảng Linux/macOS/Windows), chạy dưới 1 giây. Bao gồm: cú pháp script, kiểm tra ribbon và tính hợp lệ macro, tính toàn vẹn gói OpenXML, kịch bản installer, bộ tính CPM thuần (55 mạng ngẫu nhiên đối chiếu VBA Core).

Kiểm thử chạy thực tế trên Excel (Windows COM):
```
python 1_Scripts_TuDongHoa\kiem_tay_windows\chay_cau_noi_com.py
```
Đã kiểm nghiệm trên Microsoft Excel 16.0 (Windows): xuất XML (MSPDI), nhập XML, và xuất Primavera P6 (XER) thực thi thành công không lỗi.

## 8. Xuất bản

- Gửi chủ đầu tư/TVGS: **Xuất XLSX Sạch** (không macro).
- Trình ký: **Xuất PDF A3 Ngang**.
- Bản phát hành hiện có: `TIEN_DO_23HG_SACH_20261010_212116.xlsx`, `BAO_CAO_TIEN_DO_23HG_A3_20261010_212119.pdf`.

## 9. Xử lý sự cố

| Hiện tượng | Cách xử lý |
|---|---|
| Thấy 2 tab Ribbon khi mở file Master | Hiện tượng bình thường khi cả Add-in (.xlam) được cài đặt và file Master (.xlsm) cùng mở. Khi làm việc với file tiến độ thông thường (.xlsx), chỉ có 1 tab từ Add-in hiển thị. Xem chi tiết trong mục Cài đặt. |
| Không thấy tab ribbon | Chạy lại installer, xem nhật ký, mở lại Excel. Máy có thể đang nạp bản Add-in cũ. |
| Macro bị chặn | File > Options > Trust Center > Trust Center Settings > Trusted Locations: thêm `%APPDATA%\Microsoft\AddIns`. |
| Installer báo Excel đang mở | Lưu và đóng hết cửa sổ Excel rồi chạy lại. |
| Cần khôi phục Add-in cũ | Lấy file trong `%APPDATA%\23HG_Backup\`. |

## 10. Bản quyền và căn cứ pháp lý

Bản quyền thuộc NBT (Nguyễn Bảo Tú) | @baotuhg.

Các căn cứ pháp lý ghi trong tài liệu (Luật Xây dựng 135/2025/QH15; NĐ 206/2026, NĐ 207/2026; TT 38/2026, TT 36/2026, TT 32/2026 của BXD; TCVN 8819:2011, TCVN 8859:2011) là các định hướng quy chuẩn chuyên ngành áp dụng trong mô hình mẫu. Người dùng cần rà soát và đối chiếu với các văn bản quy phạm pháp luật hiện hành và hồ sơ mời thầu cụ thể của từng dự án.
