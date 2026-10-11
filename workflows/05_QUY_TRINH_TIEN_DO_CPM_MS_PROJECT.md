# QUY TRÌNH 05: TIẾN ĐỘ THI CÔNG WBS, ĐỊNH MỨC TT 38/2026 & ĐƯỜNG GĂNG CPM
## XUẤT TỆP TIẾN ĐỘ MICROSOFT PROJECT (.XML TƯƠNG THÍCH 100% VỚI .MPP)

---


> **Trạng thái:** ✅ **Đã hiện thực** — `tools/cpm_calculator.py`, `SchedulerAgent`;
> xuất tệp MS Project `.xml`, tính đường găng và lịch nghỉ.

### I. NGUYÊN TẮC CỐT LÕI
1. **Tính ngày công khoa học:** Không gán bừa thời gian thi công mà phải căn cứ **Hao phí định mức nhân công theo Thông tư 38/2026/TT-BXD**:
   $$\text{Tổng ngày công (công)} = \text{Khối lượng} \times \text{Định mức nhân công (công/ĐVT)}$$
2. **Thời gian thi công (Duration):**
   $$\text{Thời gian thi công (ngày)} = \text{ROUNDUP}\left(\frac{\text{Tổng ngày công}}{\text{Quân số tổ đội (người/ngày)}}, 0\right)$$
3. **Phân bổ tổ đội chuyên trách (Resources):**
   - Tổ trắc đạc, cọc mốc ($4 - 6$ người)
   - Tổ máy khoan cọc nhồi và dung dịch bentonite ($10 - 16$ người)
   - Tổ thợ sắt thép kết cấu cầu ($12 - 18$ người)
   - Tổ ván khuôn, đà giáo leo ($10 - 15$ người)
   - Tổ đổ bê tông và đầm dùi ($12 - 18$ người)
   - Tổ đúc dầm Super-T bãi đúc ($14 - 16$ người)
   - Tổ căng kéo cáp DUL 15.2mm & bơm vữa ($6 - 8$ người)
   - Tổ xe lao dầm P43 chuyên dụng ($10 - 12$ người)
   - Tổ thảm bê tông nhựa mặt cầu ($10 - 14$ người)

### II. XÁC ĐỊNH ĐƯỜNG GĂNG CHÍNH (CRITICAL PATH - CPM)
- Bất kỳ sự chậm trễ nào trên các công tác thuộc đường găng sẽ làm chậm tiến độ hoàn thành toàn bộ dự án:
  $$\text{Tim mốc} \rightarrow \text{Đường công vụ} \rightarrow \text{Khoan dò Karst} \rightarrow \text{Khoan cọc T1, T2} \rightarrow \text{Siêu âm \& PDA} \rightarrow \text{Bệ mố trụ} \rightarrow \text{Thân đặc T1, T2} \rightarrow \text{Xà mũ T1, T2} \rightarrow \text{Gối chậu \& Lao dầm} \rightarrow \text{Dầm ngang \& Liên tục nhiệt} \rightarrow \text{Bản mặt cầu C35} \rightarrow \text{Thảm BTN C16} \rightarrow \text{Thử tải \& Bàn giao}$$
- Các công tác gối đầu song song không thuộc đường găng (Non-critical): Đúc dầm Super-T tại bãi đúc (bắt đầu sớm từ khi thi công cọc), đắp đất chọn lọc sau mố K98, thi công bản quá độ, đúc gờ lan can.

### III. QUY CHUẨN XUẤT TỆP MICROSOFT PROJECT XML
- Sử dụng thư viện `xml.etree.ElementTree` xây dựng cây XML tuân thủ Microsoft Project Schema:
  + Thẻ `<Calendars>`: Lịch làm việc tiêu chuẩn 8h/ngày, 6 ngày/tuần.
  + Thẻ `<Resources>`: Danh sách 12 tổ đội và máy móc chuyên dụng.
  + Thẻ `<Tasks>`: Cấu trúc WBS nhiều cấp (Summary Tasks, Milestone, Child Tasks).
  + Thẻ `<PredecessorLink>`: Khai báo quan hệ liên kết trước sau `FS` (Finish-to-Start), `SS` (Start-to-Start) có độ trễ Lag.
  + Thẻ `<Critical>`: Gán giá trị `1` cho các tác vụ nằm trên đường găng đỏ.
- Người dùng chỉ cần mở Microsoft Project $\rightarrow$ Open file `.xml` $\rightarrow$ Save As thành `.mpp`.

### IV. PHÂN HỆ ỨNG DỤNG NGƯỜI DÙNG CUỐI: 23HG SCHEDULE ASSISTANT PRO (EXCEL MASTER & ADD-IN)
Hệ thống cung cấp trọn bộ giao diện điều hành tương tác trực tiếp trên Microsoft Excel tại thư mục `apps/23hg_schedule_assistant_pro/`:
1. **Sổ tính Master Điều hành (`23HG_DU_AN_MAU_TIEN_DO_CHUAN_G1_PRO.xlsm`):** 9 sheet nghiệp vụ chuyên sâu (Tiến độ CPM, BoQ, EVM 5D, Giải ngân, Huy động XMTB, Kế hoạch QLCL NĐ 207/2026, Lịch mùa mưa TCVN 8819:2011).
2. **Trợ lý Add-in Ribbon (`23HG_Schedule_Assistant_Pro.xlam`):** Thanh công cụ 34 lệnh phân bổ 6 nhóm, hỗ trợ Freeze Panes `D6`, Zoom 85%, Gridlines, thụt lề WBS tự động và nhóm Xuất bản 2 nút lớn tối giản (`Xuất XLSX Sạch` & `Xuất PDF A3 Ngang`).
3. **Cầu nối 2 Chiều Quốc tế (`vba_mspdi_xer_bridge.bas`):** Xuất/nhập trực tiếp tệp MSPDI XML và Primavera P6 XER thuần VBA không phụ thuộc Java hay phần mềm bên thứ ba.
4. **Cài đặt 1 Chạm:** Tệp `CAI_DAT_23HG_ENTERPRISE_PRO.bat` tự động triển khai vào `%APPDATA%\Microsoft\AddIns` và cấu hình Trusted Location.
5. **Hồ sơ Nghiệm thu Doanh nghiệp:** Toàn bộ căn cứ pháp lý, ma trận đối soát 10 điểm nghẽn kỹ thuật và biên bản bàn giao tại [BAN_GIAO_DANH_GIA_DU_AN.md](apps/23hg_schedule_assistant_pro/BAN_GIAO_DANH_GIA_DU_AN.md).
