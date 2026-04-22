# 🗂️ Tool Chia File Excel

Chia nội dung file Excel (.xls / .xlsx) thành các file Excel (.xls) có tối đa **8000 dòng dữ liệu**, giữ nguyên dòng tiêu đề (header) và format.

## ✨ Tính năng

- Hỗ trợ đọc file `.xls` và `.xlsx`
- Output file `.xls` chuẩn format: border, font Arial, header bold + nền xanh
- Giữ nguyên header ở mỗi file con
- Freeze header row để dễ xem
- Tùy chỉnh số dòng tối đa mỗi file
- Giao diện GUI (tkinter) + hỗ trợ command line
- Thanh tiến trình + nhật ký hoạt động

## 🚀 Cách sử dụng

### Cách 1: Chạy file `.exe` (không cần cài Python)

1. Tải file `ChiaFileExcel.exe` từ [Releases](../../releases)
2. Click đúp để chạy
3. Chọn file Excel → nhấn **"Chia File"**

### Cách 2: Chạy bằng Python

```bash
# Cài thư viện
pip install -r requirements.txt

# Chạy GUI
python chia_file_excel.py

# Hoặc chạy command line
python chia_file_excel.py "duong_dan_file.xls"
```

### Cách 3: Click đúp file `.bat` (Windows)

- `setup_and_run.bat` — Tự cài thư viện + mở GUI
- `build_exe.bat` — Đóng gói thành file `.exe`

## 📦 Đóng gói thành `.exe`

Yêu cầu: Python 3.8+ đã cài trên máy Windows

```bash
# Chạy build script
build_exe.bat

# Hoặc chạy thủ công
pip install pyinstaller
pyinstaller --onefile --windowed --name ChiaFileExcel chia_file_excel.py
```

File `.exe` sẽ xuất hiện trong thư mục `dist/`.

## 📋 Yêu cầu

- Python 3.8+
- xlrd==1.2.0
- xlwt==1.3.0
- openpyxl==3.1.2
