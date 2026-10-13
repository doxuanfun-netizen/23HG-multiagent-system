# -*- coding: utf-8 -*-
"""
23HG GOLDEN MASTER REGISTRY — DANH BẠ MẪU ƯU TÚ NHẤT (SINGLE SOURCE OF TRUTH)
Khóa chặt các Engine chuẩn gốc cho 6 phân hệ nghiệp vụ Dream Team (0% trùng lặp file).
"""

from __future__ import annotations

DREAM_TEAM_6_FOLDERS = (
    "00_BAN_CHI_HUY_MASTER",
    "01_HIEN_TRUONG_QLCL_KCS",
    "02_XUONG_TIEN_CHE_COT_THEP",
    "03_KINH_TE_QS_DU_TOAN_THANH_TOAN",
    "04_CO_GIOI_THIET_BI_VA_DAU_DIEZEL",
    "05_DU_LIEU_GOC_SCAN_MARKER",
)

GOLDEN_MASTER_ENGINES = {
    "PACKAGING_DISPATCHER": {
        "module": "tools.package_dispatcher",
        "entrypoint": "AECPackageDispatcher.reorganize_to_dream_team_6_folders",
        "description": "Đóng gói 6 phân hệ nghiệp vụ Dream Team không trùng lặp file (0% duplication)",
    },
    "FLEET_3TIER_5SHEETS": {
        "module": "tools.generate_bep_an_3tier_schedule",
        "entrypoint": "build_bep_an_3tier_fleet_workbook",
        "description": "Bảng Ca Xe - Ca Máy - Nhân Lực - Dầu Diezel 5 Sheets chuẩn Vincons (Sheet 01 tích hợp 3 tầng công thức sống)",
    },
    "REBAR_1D_CSP_MULTI_PHI": {
        "module": "tools.rebarcut_export",
        "entrypoint": "write_rebarcut_workbook",
        "description": "Tối ưu cắt thép 1D CSP Gilmore-Gomory + Xuất bộ file theo từng đường kính Ø và lệnh cắt CNC CSV",
    },
    "REBAR_V6_VBA_NATIVE": {
        "module": "examples.run_rebarcut_v6_native_com",
        "entrypoint": "run_native_rebarcut_v6",
        "description": "Điều khiển tự động RebarCut Pro V6 .xlsm (12 Sheet, vẽ phôi thép GDI 2 màu + xuất 5 CSV xưởng)",
    },
    "TESTING_FREQUENCY_4SHEETS": {
        "module": "tools.generate_khai_hoang_2_elite_qlcl",
        "entrypoint": "build_tan_suat_khai_hoang_2",
        "description": "Bảng theo dõi tần suất thí nghiệm Bê tông, Cốt thép (theo từng Ø), Tổng hợp vật liệu & Nghi vấn bản vẽ",
    },
    "VISUAL_ASBUILT_DASHBOARD": {
        "module": "tools.generate_khai_hoang_2_elite_qlcl",
        "entrypoint": "build_visual_asbuilt_khai_hoang_2",
        "description": "Sơ đồ mặt cắt đứng tiến độ trực quan As-Built bằng lưới ô Excel & HTML Dashboard",
    },
    "PRO_SCHEDULE_EVM_XLSM": {
        "module": "examples.build_khai_hoang_2_pro_schedule",
        "entrypoint": "build_khai_hoang_2_pro_schedule",
        "description": "Tiến độ thi công Pro 6 Sheet VBA (Dashboard EVM, Gantt Pro, Biểu đồ Nhân lực Máy, S-Curve) + MS Project XML",
    },
    "MASTER_15_SHEETS_AND_KCS_WORD": {
        "module": "examples.generate_khai_hoang_2_project",
        "entrypoint": "generate_khai_hoang_2_master_package",
        "description": "Sổ cái Master 15 Sheet liên kết 100% công thức sống + Hồ sơ Word 22 Biên bản Nghiệm thu KCS",
    },
}

DEPRECATED_GENERATORS = {
    "tools/pho_bang_master_builder.py": "Chỉ dùng dữ liệu giả lập 15 dòng thép rút gọn; đã thay bằng generate_khai_hoang_2_project.py",
    "tools/export_cau_khai_hoang_2.py": "Bản mẫu 30 dòng thép rút gọn thời kỳ đầu; đã thay bằng build_khai_hoang_2_full_pipeline.py",
}
