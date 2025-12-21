import argparse
import os
import json
import subprocess
import glob
import sys
import io

def main():
    # 1. Thiết lập tham số dòng lệnh
    parser = argparse.ArgumentParser(description="Quét thư mục và tổng hợp dữ liệu từ các file Excel.")
    parser.add_argument("-d", "--directory", required=True, help="Thư mục cần quét")
    parser.add_argument("-j", "--json", required=True, help="File JSON đặc tả cấu trúc")
    parser.add_argument("-o", "--output", help="File JSON tổng hợp đầu ra (ví dụ: all.json)")
    parser.add_argument("-r", "--recursive", action="store_true", help="Quét đệ quy vào thư mục con")
    parser.add_argument("-v", "--verbose", action="store_true", help="Hiển thị chi tiết quá trình")

    args = parser.parse_args()

    # 2. Cấu hình môi trường UTF-8 để tránh lỗi 'charmap' trên Windows
    if sys.platform == "win32":
        # Ép stdout của chính script này dùng utf-8
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    # Thiết lập biến môi trường cho tiến trình con
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"

    # Kiểm tra sự tồn tại của các thành phần cần thiết
    if not os.path.isdir(args.directory):
        print(f"Lỗi: Thư mục '{args.directory}' không tồn tại.")
        return
    
    script_path = "CollectDataFromExcel.py"
    if not os.path.exists(script_path):
        print(f"Lỗi: Không tìm thấy {script_path} trong cùng thư mục.")
        return

    # 3. Tìm kiếm danh sách file Excel
    patterns = ["*.xlsx", "*.xls"]
    excel_files = []
    
    for pattern in patterns:
        path_pattern = os.path.join(args.directory, "**" if args.recursive else "", pattern)
        excel_files.extend(glob.glob(path_pattern, recursive=args.recursive))

    if not excel_files:
        if args.verbose: print("Không tìm thấy file Excel nào phù hợp.")
        return

    # 4. Thực thi và thu thập dữ liệu
    all_results = []
    
    for excel_file in excel_files:
        if args.verbose:
            # Gửi thông báo lỗi qua stderr để không làm hỏng dòng dữ liệu stdout nếu cần bắt pipe
            sys.stderr.write(f"[LOG] Đang xử lý: {excel_file}\n")

        # Lệnh gọi script con
        cmd = [
            sys.executable, script_path,
            "--document", excel_file,
            "--json", args.json
        ]

        try:
            # Gọi tiến trình con, bắt stdout với mã hóa UTF-8
            process = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                env=env,
                check=True
            )

            # Phân tích nội dung stdout nhận được thành JSON
            stdout_content = process.stdout.strip()
            if stdout_content:
                try:
                    data = json.loads(stdout_content)
                    all_results.append({
                        "source_file": excel_file,
                        "data": data
                    })
                except json.JSONDecodeError:
                    sys.stderr.write(f"[LỖI] Output từ {excel_file} không phải JSON hợp lệ.\n")
            
        except subprocess.CalledProcessError as e:
            sys.stderr.write(f"[LỖI] Tiến trình con thất bại cho file {excel_file}: {e.stderr}\n")

    # 5. Xuất kết quả cuối cùng
    final_output = json.dumps(all_results, indent=4, ensure_ascii=False)

    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(final_output)
            if args.verbose:
                print(f"THÀNH CÔNG: Đã gộp {len(all_results)} kết quả vào {args.output}")
        except Exception as e:
            print(f"Lỗi khi ghi file output: {e}")
    else:
        # Xuất ra stdout
        sys.stdout.write(final_output + "\n")

if __name__ == "__main__":
    main()