---
name: aec-method-statement
description: Biên soạn (compile) và soát xét (audit) Thuyết minh Biện pháp thi công (BPTC) cầu/đường — số liệu sống 100% từ QS/CPM/Fleet/BBS qua method_params.json, văn 8 chương theo mẫu Km19+529, kiểm số chết + DAG + Hold Point + HSE QCVN 18 + TCVN. Dùng khi cần sinh BPTC, soát BPTC có sẵn, hoặc nối BPTC vào pipeline State Graph.
---

# Biện pháp thi công — BPTC Compiler & Checker (aec-method-statement)

## 1. Mục tiêu

Sinh và soát Thuyết minh Biện pháp thi công (BPTC) cho cầu/cống/đường
theo triết lý **deterministic-first**: mọi con số do code tính từ
State Bus (QS sống, CPM, Fleet, BBS), LLM chỉ viết văn công nghệ —
không bao giờ tính số, không số chết.

- **Lệnh `compile`** — BPTC Compiler: QS/CPM/BBS + tham số công trình
  → `method_params.json` + `Thuyet_Minh_BPTC_<cong_trinh>.docx/.md` +
  bảng đối soát số liệu.
- **Lệnh `audit`** — BPTC Checker: BPTC có sẵn + QS + CPM + KCS →
  `BAO_CAO_SOAT_BPTC.md` dạng pass/fail từng mục.

## 2. Căn cứ & Tiêu chuẩn

1. `[HỒ SƠ THIẾT KẾ ĐƯỢC DUYỆT]`: QS/BOQ, tiến độ CPM (.xml/.xlsx),
   BBS, Chỉ dẫn kỹ thuật, bản vẽ thi công — nguồn số liệu duy nhất.
2. `[TIÊU CHUẨN ÁP DỤNG CỦA DỰ ÁN]`: TCVN do Chỉ dẫn kỹ thuật chốt
   (ví dụ TCVN 4453:1995, TCVN 11823:2017, TCVN 9395:2012...) —
   chỉ viện dẫn mã hiệu, không bundle toàn văn (bản quyền).
3. `[PHÁP LUẬT/QCVN BẮT BUỘC]`: QCVN 18:2021/BXD (an toàn xây dựng)
   cho checklist HSE rule-based.
4. `[THÔNG LỆ QA/QC NỘI BỘ]`: mẫu 8 chương cầu Km19+529.080
   (`templates/Thuyet_Minh_Bien_Phap_Thi_Cong_Cau_Km19+529.080.md`).

## 3. Lệnh compile — BPTC Compiler

### 3.1 Input

| Tham số | Bắt buộc | Mô tả |
|---|---|---|
| `--qs` | ✓ | File QS/BOQ (.xlsx/.xls/.csv/.json) — đọc bằng `tools/qs_loader.py` |
| `--schedule` | ✓ | Tiến độ CPM (.xml/.xlsx/.csv/.json) — `tools/schedule_loader.py` |
| `--bbs` | | BBS thép (.xlsx/.csv/.json) — `tools/bbs_loader.py` |
| `--fleet-tasks` | | JSON task → `task_norm_key` cho `tools/equipment_fleet_scheduler.py` |
| `--project` | ✓ | JSON tham số công trình: tên, loại (cầu/cống/đường), địa chất, Karst true/false, thiết bị sẵn có, ngày khởi công |
| `--out-dir` | ✓ | Thư mục xuất kết quả |

Chạy qua pipeline: `python run_state_graph.py --phase method --qs ... --schedule ... --bbs ... --method-out <dir>`

### 3.2 Xử lý deterministic (Python, không LLM)

1. **QS → khối lượng WBS**: từ `QSEstimate.items` gom theo hạng mục
   (bê tông m³, ván khuôn m², đào/đắp m³, thép từng Ø tấn).
2. **CPM → trình tự + đường găng**: `CPMCalculator.calculate()` cho
   ES/EF/LS/LF, `critical_path`, `total_duration_days`; trình tự thi
   công trong BPTC phải đúng thứ tự ES, đường găng ghi rõ.
3. **Fleet → máy + ca + dầu**: `EquipmentFleetScheduler.compute_task_allocations()`
   + `build_daily_matrix()` cho từng công tác → danh mục máy
   (máy khoan, cẩu, giá lao...), tổng ca máy, dầu diesel, quân số đỉnh.
4. **BBS + Cutting → kế hoạch thép**: `CuttingStockSolver.solve()` theo
   nhóm (Ø, mác) → số cây 11.7m, tỷ lệ hao hụt < 1.5%, phân kỳ cấp thép.
5. **Sinh `method_params.json`**: mọi con số trên + trace về nguồn
   (dòng QS/CPM nào). Đây là Single Source of Truth của bản BPTC.

### 3.3 Sinh văn bản

- Python render template `.docx` (`templates/BPTC_TEMPLATE_8_CHUONG.docx`,
  placeholder `{{ten_tham_so}}`) bằng `python-docx`, và bản `.md` song song.
- Cấu trúc 8 chương theo mẫu Km19: Giới thiệu → Đường công vụ/bãi đúc →
  Cọc (+Karst nếu có) → Kết cấu phần dưới → Đúc dầm/căng kéo →
  Vận chuyển/lao lắp → Bản mặt cầu/hoàn thiện → HSE.
- LLM (tùy chọn, local Qwen2.5) chỉ viết đoạn văn công nghệ từ khung câu
  chuẩn — tuyệt đối không sửa số trong `method_params.json`.
- Chưa có LLM vẫn chạy: template câu chuẩn + số sống.

### 3.4 Output

- `Thuyet_Minh_BPTC_<cong_trinh>.docx` + `.md` — thuyết minh 8 chương.
- `method_params.json` — toàn bộ tham số số liệu + trace nguồn.
- `DOI_SOAT_SO_LIEU_BPTC.md` — bảng đối soát: mỗi số trong docx
  trỏ về dòng QS/CPM/Fleet/BBS nào.
- Mọi output pháp lý qua **Human Gate** như Supervisor hiện tại.

## 4. Lệnh audit — BPTC Checker

### 4.1 Input

File BPTC có sẵn (.docx/.md) + `--qs` + `--schedule` + danh mục KCS
(`--lab`, `--takeoff` — tái dùng `tools/lab_qaqc.py`,
`tools/kcs_invariants_verifier.py`).

### 4.2 Năm nhóm kiểm tra (rule-based, không LLM)

1. **Số liệu chết**: trích số khối lượng/mác/máy trong BPTC, đối chiếu
   QS/Fleet → liệt kê chênh lệch từng dòng + nguồn đúng.
2. **Logic trình tự (DAG)**: quan hệ công tác vi phạm (đổ bản mặt cầu
   trước khi lao dầm; nghiệm thu R28 chưa đủ 28 ngày) — 6 bất biến KCS.
3. **Hold Point thiếu**: công việc che khuất (cọc, cốt thép móng, Karst,
   BTCT trước khi đổ) chưa có điểm dừng nghiệm thu (GIẢI TỎA/CHỜ/CHẶN).
4. **HSE thiếu (QCVN 18:2021)**: làm việc trên cao >2m, thiết bị nâng
   (giá lao/cẩu), điện, lũ quét — checklist có/không theo công tác CPM.
5. **Viện dẫn TCVN sai/thiếu**: quét mã hiệu `TCVN xxxx:yyyy` — thiếu mã,
   sai định dạng, hoặc viện dẫn mã không có trong Chỉ dẫn kỹ thuật.

### 4.3 Output

`BAO_CAO_SOAT_BPTC.md` — pass/fail từng mục + dòng vi phạm + gợi ý sửa.
Exit code ≠ 0 khi còn lỗi `CHẶN` (fail).

## 5. Quy tắc chống Hallucination

- Không bao giờ tự điền mác bê tông, Ø thép, tải trọng thiết bị nếu
  chưa có trong QS/thiết kế — ghi `[CẦN THIẾT KẾ DUYỆT]`.
- Mọi số trong BPTC phải có trace về `method_params.json`; số không
  trace được = số chết → audit báo lỗi.
- Tiêu chuẩn viện dẫn phải ghi rõ mã hiệu + nguồn áp dụng trong dự án.
- Mọi output pháp lý qua Human Gate; không train/fine-tune LLM;
  không bundle toàn văn TCVN/QCVN.

## 6. Ví dụ

```bash
# Compile BPTC cầu Km19 từ số liệu sống
python run_state_graph.py --phase method --demo \
  --method-out out/BPTC_Km19

# Audit BPTC có sẵn
python -m tools.method_auditor \
  --bptc out/BPTC_Km19/Thuyet_Minh_BPTC_Cau_Km19.docx \
  --qs templates/Ho_So_KCS_QS_TienDo_Cau_Km19+529.080.xlsx \
  --schedule templates/Tien_Do_Thi_Cong_Cau_Km19+529.080.xml \
  --out out/BPTC_Km19/BAO_CAO_SOAT_BPTC.md
```
