# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ - EMBEDDED SERVER
Viện kiểm sát nhân dân Thành phố Hồ Chí Minh

Nhúng Backend vào ứng dụng, tự động phục vụ Frontend và cung cấp REST API nội bộ.
Chạy hoàn toàn Offline trên 127.0.0.1, bảo mật tuyệt đối 100%.
"""

import os
import sys
import json
import socket
import threading
import webbrowser
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse

# Thêm đường dẫn backend vào sys.path
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(CURRENT_DIR)
FRONTEND_DIR = os.path.join(BASE_DIR, "frontend")
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from core import buoc_1_tong_hop_va_loc, buoc_2_cap_nhat_va_bao_cao

def clean_initial_dir(d):
    """Đảm bảo initialdir luôn là một THƯ MỤC hợp lệ, không bao giờ là file hoặc chuỗi không hợp lệ"""
    try:
        if not d:
            desktop = os.path.expanduser("~/Desktop")
            return desktop if os.path.isdir(desktop) else os.getcwd()
        if isinstance(d, (list, tuple)) and d:
            d = d[0]
        if isinstance(d, str):
            # Nếu chứa dấu chấm phẩy do người dùng chọn nhiều file
            d = d.split(";")[0].strip().strip('"').strip("'")

        # Nếu là file, trỏ về thư mục cha chứa file đó
        if os.path.isfile(d):
            d = os.path.dirname(d)

        if os.path.isdir(d):
            return os.path.abspath(d)

        desktop = os.path.expanduser("~/Desktop")
        return desktop if os.path.isdir(desktop) else os.getcwd()
    except Exception:
        return os.getcwd()

def open_native_folder_dialog(initialdir=None):
    """Mở hộp thoại chọn thư mục Windows nguyên bản"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        init_dir = clean_initial_dir(initialdir)
        folder = filedialog.askdirectory(title="Chọn thư mục chứa các file của các đơn vị", initialdir=init_dir)
        root.destroy()
        return folder or ""
    except Exception as e:
        print(f"[Dialog Error] {e}")
        return ""

def open_native_file_dialog(title="Chọn file Excel", filetypes=None, initialdir=None):
    """Mở hộp thoại chọn file Excel Windows nguyên bản"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ft = filetypes or [("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")]
        init_dir = clean_initial_dir(initialdir)
        fpath = filedialog.askopenfilename(title=title, filetypes=ft, initialdir=init_dir)
        root.destroy()
        return fpath or ""
    except Exception as e:
        print(f"[Dialog Error] {e}")
        return ""

def open_native_files_dialog(title="Chọn 1 hoặc nhiều file Excel của các đơn vị", filetypes=None, initialdir=None):
    """Mở hộp thoại chọn 1 hoặc NHIỀU file Excel Windows nguyên bản (giữ Ctrl để chọn)"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ft = filetypes or [("Excel files", "*.xlsx;*.xlsm"), ("All files", "*.*")]
        init_dir = clean_initial_dir(initialdir)
        fpaths = filedialog.askopenfilenames(title=title, filetypes=ft, initialdir=init_dir)
        root.destroy()
        if isinstance(fpaths, (list, tuple)):
            return list(fpaths)
        return [fpaths] if fpaths else []
    except Exception as e:
        print(f"[Dialog Error] {e}")
        return []

class VKSAPIRequestHandler(SimpleHTTPRequestHandler):
    """Bộ xử lý HTTP vừa phục vụ Frontend tĩnh vừa phản hồi REST API nội bộ"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=FRONTEND_DIR, **kwargs)

    def end_headers(self):
        # Thiết lập header chống cache và cho phép truy cập local
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def send_json(self, status_code, data):
        """Phản hồi dữ liệu JSON cho Frontend"""
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/ping":
            self.send_json(200, {"status": "ok", "app": "VKSND-TPHCM-Case-Monitor"})
            return

        # Phục vụ các file giao diện tĩnh (HTML/CSS/JS)
        super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length).decode("utf-8") if content_length > 0 else "{}"

        try:
            req_data = json.loads(body) if body else {}
        except Exception:
            req_data = {}

        # 1. API mở hộp thoại chọn thư mục Windows
        if parsed.path == "/api/browse-folder":
            cur_path = req_data.get("current_path")
            selected_path = open_native_folder_dialog(cur_path)
            self.send_json(200, {"success": True, "path": selected_path})
            return

        # 2. API mở hộp thoại chọn file Excel Windows (đơn lẻ)
        elif parsed.path == "/api/browse-file":
            title = req_data.get("title", "Chọn file Excel")
            cur_path = req_data.get("current_path")
            selected_path = open_native_file_dialog(title=title, initialdir=cur_path)
            self.send_json(200, {"success": True, "path": selected_path})
            return

        # 2b. API mở hộp thoại chọn 1 hoặc nhiều file Excel của các đơn vị
        elif parsed.path == "/api/browse-files":
            title = req_data.get("title", "Chọn 1 hoặc nhiều file Excel của các đơn vị (giữ Ctrl để chọn nhiều)")
            cur_path = req_data.get("current_path")
            selected_paths = open_native_files_dialog(title=title, initialdir=cur_path)
            self.send_json(200, {"success": True, "paths": selected_paths})
            return

        # 3. API thực thi Bước 1 (Tổng hợp & Lọc án >= 16 ngày)
        elif parsed.path == "/api/step1":
            folder_path = req_data.get("folder_path", "").strip()
            file_paths = req_data.get("file_paths", [])
            existing_master = req_data.get("existing_master", "").strip()
            check_days = int(req_data.get("check_days", 16))

            input_source = file_paths if (file_paths and isinstance(file_paths, list)) else folder_path
            if not input_source:
                self.send_json(400, {
                    "success": False,
                    "error": "Chưa chọn file hoặc thư mục!",
                    "message": "Vui lòng bấm 'Chọn thư mục' hoặc 'Chọn file đơn vị' để chọn nguồn dữ liệu."
                })
                return

            use_timestamp = bool(req_data.get("use_timestamp", False))
            existing_master_clean = existing_master if (existing_master and os.path.exists(existing_master)) else None
            ok, result = buoc_1_tong_hop_va_loc(
                input_source,
                check_days=check_days,
                existing_master_path=existing_master_clean,
                use_timestamp=use_timestamp
            )
            if ok:
                self.send_json(200, result)
            else:
                self.send_json(400, result)
            return

        # 4. API thực thi Bước 2 (Cập nhật kết quả & Xuất báo cáo Infographic)
        elif parsed.path == "/api/step2":
            file_tong = req_data.get("file_tong", "").strip()
            file_kt = req_data.get("file_kt", "").strip()
            use_timestamp = bool(req_data.get("use_timestamp", False))

            if not file_tong or not os.path.exists(file_tong) or not file_kt or not os.path.exists(file_kt):
                self.send_json(400, {
                    "success": False,
                    "error": "File không tồn tại!",
                    "message": "Vui lòng kiểm tra lại đường dẫn của cả 2 file:\n- File 1_FILE_TONG.xlsx\n- File 2_DANH_SACH_CAN_KIEM_TRA.xlsx"
                })
                return

            ok, result = buoc_2_cap_nhat_va_bao_cao(file_tong, file_kt, use_timestamp=use_timestamp)
            if ok:
                self.send_json(200, result)
            else:
                self.send_json(400, result)
            return

        # 5. API mở thư mục trong Windows Explorer
        elif parsed.path == "/api/open-folder":
            target = req_data.get("path", "").strip()
            if not target:
                self.send_json(400, {"success": False, "error": "Đường dẫn rỗng!"})
                return

            if os.path.isfile(target):
                target = os.path.dirname(target)

            if os.path.exists(target):
                os.startfile(target)
                self.send_json(200, {"success": True, "message": f"Đã mở thư mục: {target}"})
            else:
                self.send_json(400, {"success": False, "error": "Thư mục không tồn tại!"})
            return

        else:
            self.send_json(404, {"error": "Endpoint không tồn tại!"})

def find_available_port(start_port=5050, max_port=5080):
    """Tìm cổng mạng khả dụng trên máy tính"""
    for p in range(start_port, max_port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return start_port

def start_server(port=None, auto_open_browser=True):
    """Khởi động máy chủ nhúng và mở trình duyệt tự động"""
    if port is None:
        port = find_available_port(5050)

    server_address = ("127.0.0.1", port)
    httpd = ThreadingHTTPServer(server_address, VKSAPIRequestHandler)

    url = f"http://127.0.0.1:{port}"
    print("=" * 68)
    print("  HỆ THỐNG TỔNG HỢP & GIÁM SÁT TIẾN ĐỘ ÁN HÌNH SỰ - VKSND TP.HCM")
    print("=" * 68)
    print(f"[*] Backend nhúng đang hoạt động tại: {url}")
    print(f"[*] Thư mục Frontend: {FRONTEND_DIR}")
    print("[*] Chế độ: OFFLINE nội bộ 100%, bảo mật dữ liệu ngành.")
    print("=" * 68)

    if auto_open_browser:
        def open_browser():
            import time
            time.sleep(0.8)
            webbrowser.open(url)
        threading.Thread(target=open_browser, daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n[*] Đang tắt hệ thống...")
        httpd.server_close()

if __name__ == "__main__":
    start_server()
