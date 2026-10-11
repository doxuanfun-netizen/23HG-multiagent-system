"""
23HG SYSTEM - BỘ TIỆN ÍCH TÙY CHỌN: CHUYỂN ĐỔI FILE .MPP SANG MSPDI XML
Tác giả: NBT (Nguyễn Bảo Tú) | Hệ thống 23HG

LƯU Ý QUAN TRỌNG TỪ HỘI ĐỒNG THẨM ĐỊNH KỸ THUẬT:
1. Microsoft Project (.mpp) là định dạng nhị phân độc quyền của Microsoft.
2. Thư viện MPXJ đọc được file .mpp nhưng KHÔNG THỂ ghi file .mpp trực tiếp.
   Nó ghi chuẩn Microsoft Project XML (MSPDI) và Primavera XER.
3. Vì vậy, giải pháp tối ưu và an toàn nhất trên công trường (0 dependency) là:
   - Trong MS Project: Chọn File > Save As > 'XML Format (*.xml)'
   - Trong Excel 23HG: Bấm 'Nhập MS Project (XML)' trên Ribbon
4. Tiện ích Python này là công cụ BỔ TRỢ (Bước 2) dành cho người dùng có cài Python/Java.
"""

import os
import sys

def convert_mpp_to_xml(mpp_file, xml_file=None):
    if not os.path.exists(mpp_file):
        print(f"LỖI: Không tìm thấy file nguồn: {mpp_file}")
        return False
        
    if xml_file is None:
        xml_file = os.path.splitext(mpp_file)[0] + ".xml"
        
    try:
        import mpxj
        from net.sf.mpxj.reader import UniversalProjectReader
        from net.sf.mpxj.mspdi import MSPDIWriter
        
        print(f"Đang đọc file MPP: {mpp_file}...")
        reader = UniversalProjectReader()
        project = reader.read(mpp_file)
        
        print(f"Đang ghi file MSPDI XML chuẩn: {xml_file}...")
        writer = MSPDIWriter()
        writer.write(project, xml_file)
        
        print("CHUYỂN ĐỔI THÀNH CÔNG!")
        print(f"File XML đầu ra: {xml_file}")
        print("Bạn có thể dùng nút 'Nhập MS Project (XML)' trên thanh Ribbon 23HG để nạp vào Excel.")
        return True
    except ImportError:
        print("=" * 70)
        print("THÔNG BÁO TỪ 23HG SYSTEM (CHIẾN LƯỢC 2 BƯỚC):")
        print("Thư viện 'mpxj' (Python Java bridge) chưa được cài đặt trên máy này.")
        print("-" * 70)
        print("ĐỂ THỰC HIỆN KHÔNG CẦN CÀI ĐẶT (0 DEPENDENCY):")
        print("1. Mở file .mpp trong Microsoft Project.")
        print("2. Chọn File > Save As > 'XML Format (*.xml)'.")
        print("3. Trong Excel 23HG, bấm 'Nhập MS Project (XML)' trên thanh Ribbon.")
        print("-" * 70)
        print("NẾU MUỐN DÙNG CÔNG CỤ PYTHON NÀY, CÀI ĐẶT:")
        print("   pip install mpxj")
        print("(Lưu ý: MPXJ yêu cầu máy tính phải có Java Runtime Environment JRE)")
        print("=" * 70)
        return False
    except Exception as e:
        print(f"LỖI trong quá trình xử lý: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        convert_mpp_to_xml(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    else:
        print("Sử dụng: python mpp_to_xml_converter.py <duong_dan_file.mpp> [duong_dan_file.xml]")
