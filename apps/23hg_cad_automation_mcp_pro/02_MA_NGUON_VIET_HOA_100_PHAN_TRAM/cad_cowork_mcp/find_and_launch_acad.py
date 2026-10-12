"""
Tìm đường dẫn acad.exe trong máy và khởi chạy ứng dụng GUI trực tiếp.
"""
import os
import sys
import glob
import subprocess
import winreg

def find_acad_exe():
    # 1. Tìm trong Program Files
    patterns = [
        r"C:\Program Files\Autodesk\AutoCAD*\acad.exe",
        r"C:\Program Files (x86)\Autodesk\AutoCAD*\acad.exe",
        r"D:\Program Files\Autodesk\AutoCAD*\acad.exe",
        r"E:\Program Files\Autodesk\AutoCAD*\acad.exe"
    ]
    for pattern in patterns:
        matches = glob.glob(pattern)
        if matches:
            # Sắp xếp lấy phiên bản mới nhất
            matches.sort(reverse=True)
            return matches[0]

    # 2. Tìm trong Windows Registry
    reg_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Autodesk\AutoCAD"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Autodesk\AutoCAD")
    ]
    for root_key, sub_path in reg_paths:
        try:
            with winreg.OpenKey(root_key, sub_path) as key:
                num_subkeys = winreg.QueryInfoKey(key)[0]
                for i in range(num_subkeys):
                    ver_key_name = winreg.EnumKey(key, i)
                    try:
                        with winreg.OpenKey(key, f"{ver_key_name}") as ver_key:
                            num_prod_subkeys = winreg.QueryInfoKey(ver_key)[0]
                            for j in range(num_prod_subkeys):
                                prod_key_name = winreg.EnumKey(ver_key, j)
                                try:
                                    with winreg.OpenKey(ver_key, prod_key_name) as prod_key:
                                        acad_loc, _ = winreg.QueryValueEx(prod_key, "AcadLocation")
                                        candidate = os.path.join(acad_loc, "acad.exe")
                                        if os.path.exists(candidate):
                                            return candidate
                                except Exception:
                                    pass
                    except Exception:
                        pass
        except Exception:
            pass

    # 3. Tìm qua App Paths trong registry
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\acad.exe") as key:
            val, _ = winreg.QueryValue(key, None), None
            if val and os.path.exists(val):
                return val
    except Exception:
        pass

    return None

def check_process():
    import subprocess
    cmd = ["tasklist", "/FI", "IMAGENAME eq acad.exe"]
    try:
        out = subprocess.check_output(cmd, shell=False).decode('utf-8', errors='ignore')
        return "acad.exe" in out.lower()
    except Exception:
        return False

def main():
    print("=" * 60)
    print("KIỂM TRA VÀ KHỞI CHẠY GIAO DIỆN AUTOCAD")
    print("=" * 60)

    is_running = check_process()
    print(f"Tiến trình acad.exe đang chạy ngầm: {'CÓ' if is_running else 'KHÔNG'}")

    acad_path = find_acad_exe()
    if acad_path and os.path.exists(acad_path):
        print(f"Đã tìm thấy AutoCAD tại: {acad_path}")
        print("Đang khởi chạy giao diện AutoCAD trên màn hình...")
        # Sử dụng os.startfile để Windows mở như một ứng dụng người dùng thông thường
        os.startfile(acad_path)
        print("✅ Đã gửi lệnh mở AutoCAD! Cửa sổ ứng dụng đang được nạp lên màn hình.")
    else:
        print("Không tìm thấy đường dẫn acad.exe tự động trong Program Files hoặc Registry.")
        if is_running:
            print("AutoCAD đang chạy trong Background (chế độ COM ngầm).")
        print("Vui lòng mở AutoCAD từ Start Menu hoặc Desktop của bạn.")

if __name__ == "__main__":
    main()
