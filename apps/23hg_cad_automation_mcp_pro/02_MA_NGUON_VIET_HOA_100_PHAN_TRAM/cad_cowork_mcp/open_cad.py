"""
Mở và hiển thị cửa sổ AutoCAD lên màn hình.
"""
import sys
import os
import time

def bring_autocad_to_front():
    print("Đang mở và hiển thị AutoCAD lên màn hình...")
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("Lỗi: pywin32 chưa được cài đặt.")
        return False

    pythoncom.CoInitialize()

    app = None
    # 1. Kiểm tra xem AutoCAD đã chạy chưa
    try:
        app = win32com.client.GetActiveObject("AutoCAD.Application")
        print("-> Đã tìm thấy tiến trình AutoCAD đang chạy.")
    except Exception:
        print("-> AutoCAD chưa chạy, đang khởi động AutoCAD mới...")
        try:
            app = win32com.client.Dispatch("AutoCAD.Application")
        except Exception as e:
            print(f"-> Lỗi khi khởi động AutoCAD qua COM: {e}")
            return False

    if app is not None:
        try:
            app.Visible = True
            # WindowState: 1=acNorm, 2=acMin, 3=acMax
            app.WindowState = 3 
            print(f"-> Đã đặt Visible = True, WindowState = Maximized.")
        except Exception as e:
            print(f"-> Chú ý khi đặt thuộc tính hiển thị: {e}")

        # Đảm bảo có ít nhất một bản vẽ đang mở
        try:
            if app.Documents.Count == 0:
                print("-> Đang tạo bản vẽ mới...")
                app.Documents.Add()
            print(f"-> Bản vẽ hiện hành: {app.ActiveDocument.Name}")
        except Exception as e:
            print(f"-> Lỗi khi kiểm tra bản vẽ: {e}")

        # Đưa cửa sổ lên phía trước bằng win32gui nếu có
        try:
            import win32gui
            import win32con

            def enum_windows_callback(hwnd, extra):
                title = win32gui.GetWindowText(hwnd)
                if "autocad" in title.lower() and win32gui.IsWindowVisible(hwnd):
                    extra.append(hwnd)

            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            if windows:
                hwnd = windows[0]
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                win32gui.SetForegroundWindow(hwnd)
                print("-> Đã đưa cửa sổ AutoCAD lên vị trí trên cùng màn hình.")
        except Exception as e:
            pass

        print("✅ AutoCAD đã sẵn sàng trên màn hình!")
        return True

    return False

if __name__ == "__main__":
    bring_autocad_to_front()
