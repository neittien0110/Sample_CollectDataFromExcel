import argparse

def main():
    """
    Hàm chính để phân tích và hiển thị các tham số dòng lệnh.
    """
    
    # 1. Khởi tạo ArgumentParser
    # Thêm formatter_class để giữ định dạng dòng mới trong phần help
    parser = argparse.ArgumentParser(
        description="Chương trình xử lý file dữ liệu và đặc tả cấu trúc.",
        formatter_class=argparse.RawTextHelpFormatter
    )

    # 2. Định nghĩa tham số BẮT BUỘC: --document
    parser.add_argument(
        '--document', '-d',
        type=str,
        required=True, # Bắt buộc phải có
        help='Tên file nguồn dữ liệu đầu vào. (Ví dụ: abc.xlsx)'
    )

    # 3. Định nghĩa tham số TÙY CHỌN có giá trị MẶC ĐỊNH: --json
    # Hàm set_default_json_name sẽ xử lý giá trị mặc định dựa trên --document
    # Chúng ta dùng None làm placeholder tạm thời
    parser.add_argument(
        '--json', '-j',
        type=str,
        default=None, # Đặt mặc định là None, sẽ được xử lý sau
        help='Tên file JSON đặc tả cấu trúc. \nNếu không có, mặc định lấy theo tên file --document (Ví dụ: abc.json).'
    )

    # 4. Định nghĩa tham số CỜ (Flag) Boolean: --verbose
    # action='store_true' nghĩa là nếu cờ --verbose/-v xuất hiện, giá trị là True, ngược lại là False (mặc định)
    parser.add_argument(
        '--verbose', '-v',
        action='store_true', 
        default=False, # Mặc định là False
        help='Giải thích chi tiết từng bước hoạt động của chương trình.'
    )

    # 5. Phân tích các đối số
    args = parser.parse_args()

    # 6. Xử lý giá trị mặc định đặc biệt cho --json
    # Nếu người dùng không cung cấp --json, lấy tên từ --document và thay đổi đuôi.
    if args.json is None:
        # Lấy phần tên file (không bao gồm đuôi mở rộng)
        file_name_without_ext = args.document.split('.')[0]
        # Thiết lập giá trị mặc định mới
        args.json = f"{file_name_without_ext}.json"


    # 7. Hiển thị kết quả ra stdout
    print("--- Kết Quả Phân Tích Tham Số Dòng Lệnh ---")
    
    # Hiển thị Tên file document
    print(f"Tham số --document (-d): {args.document}")
    
    # Hiển thị Tên file JSON
    print(f"Tham số --json (-j): {args.json}")
    
    # Hiển thị trạng thái Verbose
    print(f"Tham số --verbose (-v): {args.verbose}")

    # Hiển thị thông báo chi tiết nếu --verbose là True
    if args.verbose:
        print("\n--- Chế độ Chi tiết (--verbose) ---")
        print(f"Chương trình đã khởi tạo với File Dữ liệu: {args.document}.")
        print(f"File Đặc tả Cấu trúc được sử dụng: {args.json}.")
        print("Sẵn sàng tiến hành các bước xử lý dữ liệu chi tiết...")


if __name__ == "__main__":
    main()