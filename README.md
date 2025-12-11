# CHƯƠNG TRÌNH

## Tham số dòng lệnh

Cú pháp:

```shell
python .\CollectDataFromExcel.py -h         
usage: CollectDataFromExcel.py [-h] --document DOCUMENT [--json JSON] [--verbose]

Chương trình xử lý file dữ liệu và đặc tả cấu trúc.

options:
  -h, --help            show this help message and exit
  --document, -d DOCUMENT
                        Tên file nguồn dữ liệu đầu vào. (Ví dụ: abc.xlsx)
  --json, -j JSON       Tên file JSON đặc tả cấu trúc.
                        Nếu không có, mặc định lấy theo tên file --document (Ví dụ: abc.json).
  --verbose, -v         Giải thích chi tiết từng bước hoạt động của chương trình.
```

Ví dụ:

```shell
python .\CollectDataFromExcel.py -d abc.xlsx
python .\CollectDataFromExcel.py -d abc.xlsx -j config.json -v
```
