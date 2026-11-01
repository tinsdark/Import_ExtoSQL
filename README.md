# Import Excel → SQL Server

Tool web chạy local: nhập Server/User/Password → chọn DB + bảng đích → đẩy file Excel → map cột → import (1 transaction, lỗi là rollback).

## Dùng bản exe (Windows)
1. GitHub → tab **Actions** → run **Build Windows exe** mới nhất → tải artifact `ImportExcelSQL`.
2. Giải nén, chạy `ImportExcelSQL.exe` → trình duyệt tự mở `http://127.0.0.1:5000`.
3. Máy phải cài **ODBC Driver 17 hoặc 18 for SQL Server**.

## Chạy từ source
```
pip install -r requirements.txt
python app.py
```
Hoặc chạy `run.bat`.

Chỉ hỗ trợ `.xlsx`. Ô Excel đọc nguyên text (giữ số 0 đầu), numeric dùng Decimal, smalldatetime check range.
