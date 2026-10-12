"""
Script kiểm tra kết nối chi tiết tới AutoCAD từ cad_cowork_mcp.
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

def run_diagnostics():
    print("=" * 65)
    print("        KIỂM TRA KẾT NỐI TỚI AUTOCAD (CAD_COWORK_MCP)")
    print("=" * 65)

    # 1. Kiểm tra thư viện pywin32
    print("\n[1/4] Kiểm tra thư viện pywin32 / pythoncom...")
    try:
        import win32com.client
        import pythoncom
        print("  -> pywin32: ĐÃ CÀI ĐẶT (OK)")
    except ImportError as e:
        print(f"  -> LỖI: Chưa có thư viện pywin32 ({e})")
        print("     Khắc phục: chạy 'pip install pywin32'")
        return

    # 2. Thử kết nối COM trực tiếp tới AutoCAD
    print("\n[2/4] Dò tìm tiến trình AutoCAD qua COM (GetActiveObject)...")
    pythoncom.CoInitialize()
    
    acad_progids = [
        "AutoCAD.Application",
        "AutoCAD.Application.25", # AutoCAD 2025
        "AutoCAD.Application.24", # AutoCAD 2021-2024
        "AutoCAD.Application.23", # AutoCAD 2019-2020
        "AutoCAD.Application.22", # AutoCAD 2018
        "GCAD.Application",      # GstarCAD
        "ZWCAD.Application"       # ZWCAD
    ]
    
    direct_app = None
    found_progid = None
    
    for progid in acad_progids:
        try:
            app = win32com.client.GetActiveObject(progid)
            direct_app = app
            found_progid = progid
            print(f"  -> Tìm thấy ứng dụng đang chạy với ProgID: {progid}")
            break
        except Exception as e:
            # Không tìm thấy với progid này
            pass

    if direct_app is not None:
        try:
            doc_count = direct_app.Documents.Count
            caption = direct_app.Caption
            print(f"  -> Trạng thái: {caption}")
            print(f"  -> Số bản vẽ đang mở: {doc_count}")
            if doc_count > 0:
                print(f"  -> Bản vẽ hiện hành: {direct_app.ActiveDocument.Name}")
            else:
                print("  -> CẢNH BÁO: AutoCAD đang mở nhưng chưa có bản vẽ nào! Hãy nhấn New hoặc mở 1 file .dwg.")
        except Exception as e:
            print(f"  -> Lỗi khi truy cập thuộc tính AutoCAD: {e}")
    else:
        print("  -> Không tìm thấy tiến trình AutoCAD nào đang chạy trong Running Object Table (ROT).")
        print("     Nguyên nhân thường gặp:")
        print("     a) AutoCAD chưa được bật.")
        print("     b) AutoCAD đang chạy bằng quyền 'Run as Administrator', khiến ứng dụng khác bị chặn kết nối (UIPI).")

    # 3. Thử qua CADController
    print("\n[3/4] Kết nối qua CADController.start_cad()...")
    try:
        from .cad_controller import CADController
    except ImportError:
        from cad_controller import CADController
    
    controller = CADController()
    print(f"  -> Cấu hình cad_type: '{controller.cad_type}' (preferred: '{controller.preferred_cad_type}')")
    
    success = controller.start_cad()
    if success:
        print("\n[4/4] KẾT QUẢ: ✅ KẾT NỐI THÀNH CÔNG!")
        print(f"  • Phần mềm: {controller.active_cad_name} ({controller.active_cad_type})")
        if controller.doc:
            try:
                print(f"  • Bản vẽ: {controller.doc.Name}")
            except Exception:
                pass
    else:
        print("\n[4/4] KẾT QUẢ: ❌ KẾT NỐI THẤT BẠI!")
        print("-" * 65)
        print("HƯỚNG DẪN XỬ LÝ:")
        print("1. Hãy mở AutoCAD lên và mở/tạo một bản vẽ mới (.dwg).")
        print("2. Đảm bảo AutoCAD và Python cùng quyền (không bật AutoCAD dưới quyền Run as Administrator nếu Python không phải Admin).")
        print("3. Nếu muốn thư viện TỰ ĐỘNG KHỞI ĐỘNG AutoCAD khi chưa mở:")
        print("   Mở file config.json và sửa:")
        print('   "type": "autocad"')
        print("=" * 65)

if __name__ == "__main__":
    run_diagnostics()
