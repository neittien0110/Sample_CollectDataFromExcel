# THU THẬP DỮ LIỆU TRONG FILE MICRSOFT EXCEL

![Nguồn văn bản](./assets/example.png) ==> ![Trich xuất](./assets/ouput.png)

## Mục tiêu chương trình

Viêt chương trình python để lấy nội dung trong file excel và ghi ra file, với các tham số dòng lệnh.

- File xlsx theo định dạng .xlsx /.xls, chỉ định bởi tham số --document. Kiểm tra sự tồn tại của file này.
- Vị trí các cell cần đọc được mô tả trong file --json, định vị theo chỉ số cột-dòng dạng A3, B5, và đồng thời cho phép định vị theo Range Name.
- Chỉ định chính xác cả tên sheet, theo cấu trúc lồng "sheets": [{"Sheet 1", "cells": [ ]}]
- Các tham số dòng lệnh dài có tham số ngắn kèm theo.
- Xuất nội dung phân tích được ra màn hình hoặc ra file
- Thêm tham số dòng lệnh --verbose.

__Version 2__: Tiến trình quét các file excel trong 1 thư mục

## Cấu trúc File Đặc tả JSON Bắt buộc

File JSON đặc tả của bạn phải tuân theo cấu trúc sau (ví dụ: [example.json](./example.json)):

```json
{
    "sheets": [
        {
            "name": "TrangBia",
            "cells": [
                { "name": "TenBaoCao", "location": "A1" },
                { "name": "NgayPhatHanh", "location": "NgayPhatHanh" }
            ]
        },
        {
            "name": "BaoCaoThongKe",
            "cells": [
                { "name": "TongSoDonHang", "location": "C10" }
            ]
        }
    ]
}
```

## Tham số dòng lệnh

Cú pháp:

```shell
python .\CollectDataFromExcel.py -h         
usage: CollectDataFromExcel.py [-h] --document DOCUMENT [--json JSON] [--output OUTPUT] [--verbose]

Chương trình xử lý file dữ liệu và đặc tả cấu trúc.

options:
  -h, --help            show this help message and exit
  --document, -d DOCUMENT
                        Tên file nguồn dữ liệu đầu vào. (Ví dụ: abc.xlsx)
  --json, -j JSON       Tên file JSON đặc tả cấu trúc.
                        Nếu không có, mặc định lấy theo tên file --document (Ví dụ: abc.json).
  --output, -o OUTPUT   File JSON đầu ra. Nếu trống sẽ in ra màn hình.
  --verbose, -v         Giải thích chi tiết từng bước hoạt động của chương trình.
```

Ví dụ:

```shell
python .\CollectDataFromExcel.py --document example.xlsx
python .\CollectDataFromExcel.py -d example.xlsx -j example.json -o output.json -v
```

## Đọc các file excel trong một thư mục

__Chạy quét thư mục hiện tại và in ra màn hình__

```shell
   python ScanFolder.py -d ./ -j example.json
```

__Quét đệ quy toàn bộ thư mục con và lưu vào file__

```shell
   python ScanFolder.py -d ./myfolder -j example.json -o all.json -v
```

### Cách thức hoạt động của chương trình:

- __Tìm kiếm File__: Chương trình sử dụng thư viện glob để tìm các tệp tin có đuôi .xlsx và .xls. Nếu có tham số --recursive (-r), nó sẽ tìm xuyên suốt qua các thư mục con.
- Giao tiếp giữa các Process: Chương trình gọi CollectDataFromExcel.py thông qua subprocess.run. Nó sử dụng capture_output=True và encoding='utf-8' để thu thập toàn bộ nội dung JSON mà script con in ra màn hình.
- __Hỗ trợ UTF-8__: Việc thiết lập encoding='utf-8' trong cả subprocess và khi ghi file đảm bảo các ký tự tiếng Việt hoặc ký tự đặc biệt được xử lý chính xác.
- __Tổng hợp Dữ liệu__: Mỗi kết quả từ một file Excel sẽ được đóng gói thành một đối tượng trong mảng all_results, bao gồm đường dẫn file nguồn và dữ liệu đã trích xuất.
- __Quản lý Output__:
  - Nếu bạn cung cấp -o output.json, chương trình sẽ tạo ra một file duy nhất chứa dữ liệu của tất cả các file Excel đã quét.
  - Nếu không có -o, chương trình sẽ đẩy mảng JSON tổng hợp ra màn hình stdout để bạn có thể sử dụng kết quả cho các tiến trình khác.

### Xử lý mã hóa (Encoding):

Sử dụng env["PYTHONUTF8"] = "1" để ép Python (phiên bản 3.7+) luôn chạy ở chế độ UTF-8 bất kể cài đặt vùng (Locale) của Windows.

Sử dụng io.TextIOWrapper để ép sys.stdout của script cha ghi dữ liệu theo chuẩn UTF-8, tránh lỗi charmap khi print tiếng Việt ra màn hình.

### Tách biệt kênh thông tin:

Các dòng thông báo trạng thái (verbose) được đẩy qua sys.stderr.

Điều này rất quan trọng: Nếu bạn sử dụng lệnh này để nối với một công cụ khác (ví dụ: python ScanFolder.py ... > result.json), các dòng log sẽ không bị trộn lẫn vào nội dung file JSON.

### Cấu trúc dữ liệu gộp:

Mỗi phần tử trong mảng kết quả sẽ có dạng: {"source_file": "đường/dẫn/file.xlsx", "data": {...}}. Cách này giúp bạn biết chính xác dữ liệu nào đến từ file nào sau khi gộp.