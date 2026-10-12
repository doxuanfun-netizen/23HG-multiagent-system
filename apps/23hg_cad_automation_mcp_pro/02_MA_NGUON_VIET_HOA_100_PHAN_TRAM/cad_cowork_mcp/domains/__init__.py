"""
Bộ định tuyến miền doanh nghiệp

Mỗi miền doanh nghiệp triển khai bộ định tuyến riêng của mình một cách độc lập, đã chia sẻ command_router khung cơ bản. 

Cấu trúc thư mục: 
    domains/
    ├── gslc/       # Sơ đồ quy trình(GsLc)
    ├── gsbz/       # Bản vẽ lắp đặt(tương lai)
    ├── gspg/       # Sơ đồ đường ống(tương lai)
    └── ksbz/       # Kinh doanh khác(tương lai)

Mở rộng lĩnh vực kinh doanh mới: 
    1. tạo nên domains/xxx/ Mục lục
    2. hoàn thành router.py (kế thừa BaseCommandRouter)
    3. sự định nghĩa taxonomy.py (Phân loại)
    4. sự định nghĩa rules.py (quy tắc tiêu cực)
    5. hiện hữu server.py đăng ký xxx_route_command dụng cụ
"""
