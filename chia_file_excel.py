#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tool Chia File Excel - Split Excel files into smaller chunks
============================================================
Chia nội dung file Excel (.xls / .xlsx) thành các file Excel (.xls)
có tối đa 8000 dòng dữ liệu, giữ nguyên dòng tiêu đề (header).

Yêu cầu cài đặt:
    pip install xlrd xlwt xlutils openpyxl

Cách sử dụng:
    - Chạy trực tiếp: python chia_file_excel.py
    - Giao diện GUI sẽ hiện lên, chọn file và nhấn "Chia File"
    - Hoặc chạy từ command line:
        python chia_file_excel.py "đường_dẫn_file.xls"
"""

import os
import sys
import math
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

# ============================================================
# CONSTANTS
# ============================================================
MAX_ROWS_PER_FILE = 8000  # Số dòng dữ liệu tối đa mỗi file (không tính header)

# ============================================================
# CORE LOGIC
# ============================================================

def read_source_file(filepath):
    """
    Đọc file Excel nguồn (.xls hoặc .xlsx).
    Trả về: (headers, data_rows, col_widths, header_styles)
        - headers: list các giá trị tiêu đề
        - data_rows: list các dòng dữ liệu (mỗi dòng là list giá trị)
        - col_widths: list chiều rộng cột (int)
        - header_style_info: dict thông tin style header (font bold, etc.)
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext == '.xls':
        return _read_xls(filepath)
    elif ext in ('.xlsx', '.xlsm'):
        return _read_xlsx(filepath)
    else:
        raise ValueError(f"Định dạng file không được hỗ trợ: {ext}\nChỉ hỗ trợ: .xls, .xlsx")


def _read_xls(filepath):
    """Đọc file .xls bằng xlrd."""
    import xlrd

    workbook = xlrd.open_workbook(filepath, formatting_info=True)
    sheet = workbook.sheet_by_index(0)

    if sheet.nrows == 0:
        raise ValueError("File Excel trống, không có dữ liệu!")

    # Đọc header (dòng đầu tiên)
    headers = []
    for col in range(sheet.ncols):
        val = sheet.cell_value(0, col)
        headers.append(val)

    # Đọc dữ liệu (từ dòng 2 trở đi)
    data_rows = []
    for row_idx in range(1, sheet.nrows):
        row_data = []
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            val = cell.value
            # Xử lý date
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    val = xlrd.xldate_as_datetime(val, workbook.datemode)
                    val = val.strftime('%d/%m/%Y')
                except Exception:
                    pass
            row_data.append(val)
        data_rows.append(row_data)

    # Lấy chiều rộng cột
    col_widths = []
    for col_idx in range(sheet.ncols):
        try:
            width = sheet.col(col_idx).width if hasattr(sheet.col(col_idx), 'width') else 3000
        except Exception:
            width = 3000
        col_widths.append(width)

    return headers, data_rows, col_widths


def _read_xlsx(filepath):
    """Đọc file .xlsx bằng openpyxl."""
    from openpyxl import load_workbook

    wb = load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("File Excel trống, không có dữ liệu!")

    headers = list(rows[0])
    data_rows = [list(row) for row in rows[1:]]

    # Mặc định chiều rộng
    col_widths = [4000] * len(headers)

    wb.close()
    return headers, data_rows, col_widths


def write_xls_file(filepath, headers, data_rows, col_widths=None):
    """
    Ghi dữ liệu ra file .xls với format đẹp.
    """
    import xlwt

    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('Sheet1')

    # === STYLE cho Header ===
    header_style = xlwt.XFStyle()

    # Font: Bold, size 11
    header_font = xlwt.Font()
    header_font.bold = True
    header_font.name = 'Arial'
    header_font.height = 220  # 11pt * 20
    header_style.font = header_font

    # Border
    header_borders = xlwt.Borders()
    header_borders.left = xlwt.Borders.THIN
    header_borders.right = xlwt.Borders.THIN
    header_borders.top = xlwt.Borders.THIN
    header_borders.bottom = xlwt.Borders.THIN
    header_style.borders = header_borders

    # Background color (xanh nhạt)
    header_pattern = xlwt.Pattern()
    header_pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    header_pattern.pattern_fore_colour = xlwt.Style.colour_map['light_blue']
    header_style.pattern = header_pattern

    # Alignment
    header_align = xlwt.Alignment()
    header_align.horz = xlwt.Alignment.HORZ_CENTER
    header_align.vert = xlwt.Alignment.VERT_CENTER
    header_align.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    header_style.alignment = header_align

    # === STYLE cho Data ===
    data_style = xlwt.XFStyle()

    data_font = xlwt.Font()
    data_font.name = 'Arial'
    data_font.height = 200  # 10pt
    data_style.font = data_font

    data_borders = xlwt.Borders()
    data_borders.left = xlwt.Borders.THIN
    data_borders.right = xlwt.Borders.THIN
    data_borders.top = xlwt.Borders.THIN
    data_borders.bottom = xlwt.Borders.THIN
    data_style.borders = data_borders

    data_align = xlwt.Alignment()
    data_align.vert = xlwt.Alignment.VERT_CENTER
    data_align.wrap = xlwt.Alignment.WRAP_AT_RIGHT
    data_style.alignment = data_align

    # === Ghi Header ===
    for col_idx, header in enumerate(headers):
        sheet.write(0, col_idx, str(header) if header is not None else '', header_style)

    # === Ghi Data ===
    for row_idx, row in enumerate(data_rows):
        for col_idx, value in enumerate(row):
            if value is None:
                value = ''
            sheet.write(row_idx + 1, col_idx, value, data_style)

    # === Đặt chiều rộng cột ===
    if col_widths:
        for col_idx, width in enumerate(col_widths):
            # Đảm bảo chiều rộng hợp lý (min 2000, max 15000)
            w = max(2000, min(15000, width))
            sheet.col(col_idx).width = w
    else:
        # Tự tính chiều rộng dựa trên nội dung header
        for col_idx, header in enumerate(headers):
            w = max(3000, len(str(header)) * 350 + 500)
            w = min(15000, w)
            sheet.col(col_idx).width = w

    # Freeze header row
    sheet.set_panes_frozen(True)
    sheet.set_horz_split_pos(1)

    workbook.save(filepath)


def split_excel(filepath, max_rows=MAX_ROWS_PER_FILE, progress_callback=None):
    """
    Chia file Excel thành các file nhỏ hơn.

    Args:
        filepath: Đường dẫn file nguồn
        max_rows: Số dòng dữ liệu tối đa mỗi file
        progress_callback: Hàm callback(current, total, message) để cập nhật tiến trình

    Returns:
        list: Danh sách đường dẫn các file đã tạo
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Không tìm thấy file: {filepath}")

    # Đọc file nguồn
    if progress_callback:
        progress_callback(0, 100, "Đang đọc file nguồn...")

    headers, data_rows, col_widths = read_source_file(filepath)

    total_data_rows = len(data_rows)
    if total_data_rows == 0:
        raise ValueError("File không có dữ liệu (chỉ có header)!")

    # Tính số file cần tạo
    num_files = math.ceil(total_data_rows / max_rows)

    if num_files == 1:
        if progress_callback:
            progress_callback(100, 100, f"File chỉ có {total_data_rows} dòng, không cần chia!")
        return []

    # Tạo thư mục output
    source_dir = os.path.dirname(filepath)
    source_name = os.path.splitext(os.path.basename(filepath))[0]
    output_dir = os.path.join(source_dir, f"{source_name}_split")
    os.makedirs(output_dir, exist_ok=True)

    created_files = []

    for file_idx in range(num_files):
        start_row = file_idx * max_rows
        end_row = min(start_row + max_rows, total_data_rows)
        chunk = data_rows[start_row:end_row]

        # Tên file output
        output_filename = f"{source_name}_part{file_idx + 1}.xls"
        output_path = os.path.join(output_dir, output_filename)

        # Cập nhật tiến trình
        if progress_callback:
            pct = int((file_idx + 1) / num_files * 100)
            msg = (f"Đang tạo file {file_idx + 1}/{num_files}: {output_filename} "
                   f"(dòng {start_row + 1} → {end_row})")
            progress_callback(pct, 100, msg)

        write_xls_file(output_path, headers, chunk, col_widths)
        created_files.append(output_path)

    if progress_callback:
        progress_callback(100, 100,
            f"Hoàn tất! Đã chia thành {num_files} file trong thư mục:\n{output_dir}")

    return created_files


# ============================================================
# GUI
# ============================================================

class ExcelSplitterApp:
    """Giao diện đồ họa cho tool chia file Excel."""

    def __init__(self, root):
        self.root = root
        self.root.title("🗂️ Chia File Excel")
        self.root.geometry("680x520")
        self.root.resizable(True, True)

        # Không cho phép chạy nhiều lần cùng lúc
        self.is_running = False
        self.selected_file = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        # === Frame chính ===
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === Tiêu đề ===
        title_label = ttk.Label(
            main_frame,
            text="TOOL CHIA FILE EXCEL",
            font=('Arial', 16, 'bold')
        )
        title_label.pack(pady=(0, 5))

        desc_label = ttk.Label(
            main_frame,
            text=f"Chia file Excel thành các file .xls có tối đa {MAX_ROWS_PER_FILE:,} dòng dữ liệu",
            font=('Arial', 10)
        )
        desc_label.pack(pady=(0, 15))

        # === Chọn file ===
        file_frame = ttk.LabelFrame(main_frame, text="File nguồn", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        file_entry = ttk.Entry(file_frame, textvariable=self.selected_file, state='readonly')
        file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        browse_btn = ttk.Button(file_frame, text="📂 Chọn file...", command=self._browse_file)
        browse_btn.pack(side=tk.RIGHT)

        # === Cấu hình ===
        config_frame = ttk.LabelFrame(main_frame, text="Cấu hình", padding=10)
        config_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(config_frame, text="Số dòng tối đa mỗi file:").pack(side=tk.LEFT)
        self.max_rows_var = tk.StringVar(value=str(MAX_ROWS_PER_FILE))
        max_rows_entry = ttk.Entry(config_frame, textvariable=self.max_rows_var, width=10)
        max_rows_entry.pack(side=tk.LEFT, padx=(10, 0))

        # === Nút hành động ===
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 10))

        self.split_btn = ttk.Button(
            btn_frame, text="✂️ CHIA FILE", command=self._start_split
        )
        self.split_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.open_folder_btn = ttk.Button(
            btn_frame, text="📁 Mở thư mục kết quả",
            command=self._open_output_folder, state=tk.DISABLED
        )
        self.open_folder_btn.pack(side=tk.LEFT)

        # === Progress bar ===
        self.progress = ttk.Progressbar(main_frame, mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X, pady=(0, 10))

        # === Log ===
        log_frame = ttk.LabelFrame(main_frame, text="Nhật ký", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, state=tk.DISABLED,
                                font=('Consolas', 9))
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Biến lưu output_dir
        self.output_dir = None

    def _browse_file(self):
        filepath = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[
                ("Excel files", "*.xls *.xlsx *.xlsm"),
                ("Excel 97-2003", "*.xls"),
                ("Excel 2007+", "*.xlsx"),
                ("All files", "*.*")
            ]
        )
        if filepath:
            self.selected_file.set(filepath)
            self._log(f"Đã chọn: {filepath}")

    def _log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_progress(self, current, total, message):
        self.root.after(0, lambda: self._do_update_progress(current, total, message))

    def _do_update_progress(self, current, total, message):
        self.progress['value'] = current
        self._log(message)

    def _start_split(self):
        filepath = self.selected_file.get()
        if not filepath:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel trước!")
            return

        # Validate max_rows
        try:
            max_rows = int(self.max_rows_var.get())
            if max_rows < 100:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Cảnh báo", "Số dòng tối đa phải là số nguyên >= 100!")
            return

        if self.is_running:
            return

        self.is_running = True
        self.split_btn.configure(state=tk.DISABLED)
        self.progress['value'] = 0

        # Chạy trong thread riêng để không block GUI
        thread = threading.Thread(target=self._do_split, args=(filepath, max_rows), daemon=True)
        thread.start()

    def _do_split(self, filepath, max_rows):
        try:
            created_files = split_excel(filepath, max_rows, self._update_progress)

            if created_files:
                self.output_dir = os.path.dirname(created_files[0])

                def on_done():
                    self.open_folder_btn.configure(state=tk.NORMAL)
                    self._log(f"\n{'='*50}")
                    self._log(f"✅ Tổng cộng: {len(created_files)} file đã được tạo")
                    self._log(f"📂 Thư mục: {self.output_dir}")
                    self._log(f"{'='*50}")
                    messagebox.showinfo("Thành công",
                        f"Đã chia thành {len(created_files)} file!\n\n"
                        f"Thư mục kết quả:\n{self.output_dir}")

                self.root.after(0, on_done)
            else:
                def on_no_split():
                    self._log("ℹ️ File không cần chia (ít hơn hoặc bằng số dòng tối đa).")
                    messagebox.showinfo("Thông báo",
                        f"File có ít hơn {max_rows:,} dòng dữ liệu, không cần chia!")

                self.root.after(0, on_no_split)

        except Exception as e:
            def on_error():
                self._log(f"❌ Lỗi: {str(e)}")
                messagebox.showerror("Lỗi", f"Đã xảy ra lỗi:\n{str(e)}")

            self.root.after(0, on_error)

        finally:
            def reset():
                self.is_running = False
                self.split_btn.configure(state=tk.NORMAL)

            self.root.after(0, reset)

    def _open_output_folder(self):
        if self.output_dir and os.path.exists(self.output_dir):
            if sys.platform == 'win32':
                os.startfile(self.output_dir)
            elif sys.platform == 'darwin':
                os.system(f'open "{self.output_dir}"')
            else:
                os.system(f'xdg-open "{self.output_dir}"')


# ============================================================
# COMMAND LINE MODE
# ============================================================

def cli_mode(filepath):
    """Chạy ở chế độ command line."""
    print(f"{'='*60}")
    print(f"  TOOL CHIA FILE EXCEL")
    print(f"  Tối đa {MAX_ROWS_PER_FILE:,} dòng mỗi file")
    print(f"{'='*60}")
    print(f"\n📄 File nguồn: {filepath}")

    def cli_progress(current, total, message):
        print(f"  [{current:3d}%] {message}")

    try:
        created_files = split_excel(filepath, MAX_ROWS_PER_FILE, cli_progress)
        if created_files:
            print(f"\n✅ Hoàn tất! Đã tạo {len(created_files)} file:")
            for f in created_files:
                print(f"   → {f}")
        else:
            print(f"\nℹ️  File không cần chia.")
    except Exception as e:
        print(f"\n❌ Lỗi: {e}")
        sys.exit(1)


# ============================================================
# MAIN
# ============================================================

if __name__ == '__main__':
    if len(sys.argv) > 1:
        # Command line mode
        cli_mode(sys.argv[1])
    else:
        # GUI mode
        root = tk.Tk()
        app = ExcelSplitterApp(root)
        root.mainloop()
