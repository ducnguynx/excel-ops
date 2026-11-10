# Tổng hợp hồ sơ thanh toán Excel (IPC Summary Extractor)

Script Python tự động đọc dữ liệu từ sheet `Summary` của các file thanh toán nhà thầu phụ (`.xls`, `.xlsx`, `.xlsm`) theo từng đợt và tổng hợp vào `Sheet1` của file Master.

## 1. Yêu cầu & Cài đặt

Cần cài đặt **Python 3.8+** và các thư viện phụ thuộc:

```bash
# Tạo môi trường ảo (tuỳ chọn)
python3 -m venv .venv
source .venv/bin/activate  # macOS/Linux
# .venv\Scripts\activate   # Windows

# Cài đặt thư viện
pip install openpyxl xlrd thefuzz
```

## 2. Cấu trúc thư mục

Đặt `main.py`, file Master gốc và các thư mục đợt thanh toán (đã giải nén) trong cùng một thư mục:

```text
.
├── main.py
├── Danh mục SA 08 - NOVATION.xlsx
├── Dot 7/
├── Dot 8/
│   ├── Bao Tran/
│   ├── CMT/
│   └── ...
├── Dot 9/
├── Dot 10/
├── Dot 11/
└── Dot 12/
```

## 3. Cách chạy

```bash
python3 main.py
```

- Script sẽ tự động quét tất cả các thư mục đợt thanh toán nằm cùng cấp với `main.py` (sắp xếp theo số thứ tự đợt), tạo bản sao từ file gốc `Danh mục SA 08 - NOVATION.xlsx` sang `Danh mục SA 08 - NOVATION - test.xlsx` và ghi dữ liệu tổng hợp vào `Sheet1`.
- Những dòng/ô có cấu trúc bất thường hoặc lệch số liệu đối chiếu sẽ được đánh dấu `"check"`.
- Để thay đổi tên file Master đầu vào/đầu ra, chỉnh sửa biến `MASTER_FILE`, `TEST_MASTER_FILE` ở cuối file `main.py`.
